"""애플리케이션 로깅 / 감사 로그 설정."""
import json
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

from amisafe.config import Config


def setup_logging(app):
    """RotatingFileHandler 로 app.log 로테이션."""
    app.logger.setLevel(logging.INFO)

    already = any(
        isinstance(h, RotatingFileHandler) and getattr(h, "baseFilename", "").endswith("app.log")
        for h in app.logger.handlers
    )
    if already:
        return

    handler = RotatingFileHandler(
        Config.APP_LOG_PATH,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setLevel(logging.INFO)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    app.logger.addHandler(handler)


def write_audit_log(action, user=None, form_id=None, document_id=None, extra=None):
    """감사 로그 (JSON line) 기록."""
    payload = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "action": action,
        "user_id": (user or {}).get("id"),
        "user_name": (user or {}).get("name"),
        "group": (user or {}).get("group"),
        "role": (user or {}).get("role"),
        "form_id": form_id,
        "document_id": document_id,
        "extra": extra or {},
    }
    with open(Config.AUDIT_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


def log_exception(app, where, e):
    """앱 로거를 통한 예외 기록 (스택트레이스 포함)."""
    app.logger.exception("[%s] %s", where, e)
