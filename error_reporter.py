"""
error_reporter.py
=================
Real-time Slack alerts for tech-side splice-report failures.

Why this exists
---------------
A tech who hits a runtime error sees a red banner and a traceback they can
copy/paste; without active reporting we only learn about it when they call us.
This module ships those errors to the same Slack channel Secret Sauce uses
(commit 44bce61 of the secret-sauce repo) so a tech-side failure is visible in
real time — and the message is tagged by app so one channel can serve every
desktop / web app the user runs.

Webhook discipline
------------------
The webhook URL is NEVER committed.  The repo is public and Slack auto-revokes
any webhook it finds in source.  Instead:

  CI step "Bake error-report webhook" writes the SLACK_ERROR_WEBHOOK repo secret
  to desktop/_webhook.cfg just before the PyInstaller build.  The .spec bundles
  _webhook.cfg only if it exists.  At launch the launcher reads it into
  os.environ["SS_ERROR_WEBHOOK"].  A build without the secret simply ships
  reporting OFF (this module no-ops).

For the web / Streamlit Cloud deployment, set SS_ERROR_WEBHOOK as a platform
secret on the host; report_error reads it the same way.

Safety contract
---------------
  * No-op when SS_ERROR_WEBHOOK is unset.
  * Never raises — reporting must never break a tech's run.
  * Send happens on a daemon thread with a 4-second timeout; the main thread
    never blocks on Slack.
  * Hourly dedup per (where, type(exc), str(exc)) signature so a repeat can't
    flood the channel.  Distinct errors fire immediately.
  * NEVER includes customer / fiber / trace data — just counts, format,
    span length, mode.  Context dict from the caller must respect that.
"""
from __future__ import annotations

import os


APP_NAME = "Splice Report"


# error signature -> last-sent epoch (in-process, hourly dedup)
_ERR_LAST: dict[str, float] = {}


def report_error(where: str, exc: BaseException, context: dict | None = None) -> None:
    """Post a scrubbed tech-side error to the Slack webhook in
    ``SS_ERROR_WEBHOOK``.  No-op when unset/offline.

    Parameters
    ----------
    where : str
        Short tag for the call site, e.g. "Generate report" or "Launcher
        startup".  Used in the message header and the dedup signature.
    exc : BaseException
        The raised exception.  Type + str(exc) are in the dedup signature, so
        a repeat of the same error inside an hour is suppressed.
    context : dict, optional
        Small bag of {key: value} pairs included verbatim in the alert.
        Counts / format / mode only — NEVER trace data.
    """
    try:
        url = os.environ.get("SS_ERROR_WEBHOOK")
        if not url:
            return
        import hashlib
        import platform
        import time
        import traceback

        sig = hashlib.md5(
            f"{where}|{type(exc).__name__}|{exc}".encode()).hexdigest()
        now = time.time()
        if now - _ERR_LAST.get(sig, 0) < 3600:
            return
        _ERR_LAST[sig] = now

        try:
            import getpass
            import socket
            who = f"{socket.gethostname()} / {getpass.getuser()}"
        except Exception:
            who = "?"
        ctx = "".join(
            f"\n• {k}: {v}" for k, v in (context or {}).items())
        text = (
            f":rotating_light: *{APP_NAME} error* — {where}\n"
            f"*{type(exc).__name__}*: {exc}\n"
            f"tech: `{who}`  |  os: {platform.platform()}  |  "
            f"engine: {os.environ.get('SS_ENGINE_SOURCE', '?')}{ctx}\n"
            f"```{traceback.format_exc()[-1400:]}```"
        )

        import json as _json
        import threading
        import urllib.request

        def _send():
            try:
                req = urllib.request.Request(
                    url,
                    data=_json.dumps({"text": text}).encode(),
                    headers={"Content-Type": "application/json"})
                urllib.request.urlopen(req, timeout=4)
            except Exception:
                pass

        threading.Thread(target=_send, daemon=True).start()
    except Exception:
        # Reporting must NEVER break the run.
        pass
