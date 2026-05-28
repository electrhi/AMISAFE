"""AMISAFE 애플리케이션 팩토리.

create_app() 호출 시 Flask 앱을 구성하고 Blueprint를 등록한다.
"""
import json
import traceback

from flask import Flask, Response, request
from werkzeug.exceptions import HTTPException

from amisafe.config import Config
from amisafe.extensions.logging_setup import setup_logging
from amisafe.extensions.db import init_db
from amisafe.routes import auth, admin, forms, api, files, designer


def create_app():
    app = Flask(
        __name__,
        template_folder="templates",
        static_folder="static",
    )
    app.secret_key = Config.SECRET_KEY

    Config.ensure_folders()
    setup_logging(app)

    # 시작 시 DB 테이블 자동 생성 (멱등). 실패해도 앱 기동은 막지 않는다.
    # (fresh 배포/DB 초기화 후 첫 요청에서 테이블 미존재로 500 나는 것 방지)
    try:
        init_db()
        app.logger.info("DB init OK")
    except Exception as e:
        app.logger.warning("DB init skipped: %s", e)

    # Blueprint 등록 순서 (의미상 일반적인 라우트가 먼저, API는 뒤)
    app.register_blueprint(auth.bp)
    app.register_blueprint(admin.bp)
    app.register_blueprint(designer.bp)
    app.register_blueprint(forms.bp)
    app.register_blueprint(api.bp)
    app.register_blueprint(files.bp)

    # ===== [임시 진단] 처리되지 않은 예외를 브라우저/로그에 트레이스백으로 노출 =====
    # 양식 디자이너 저장 후 /admin 500 원인 추적용. 원인 확인되면 이 블록 전체 삭제할 것.
    @app.errorhandler(Exception)
    def _debug_unhandled(e):
        # 404/403/401 등 정상 HTTP 응답은 그대로 통과시킨다.
        if isinstance(e, HTTPException):
            return e
        tb = traceback.format_exc()
        app.logger.error("UNHANDLED EXCEPTION on %s %s\n%s", request.method, request.path, tb)
        body = json.dumps({
            "ok": False,
            "debug": True,
            "path": request.path,
            "method": request.method,
            "error": f"{type(e).__name__}: {e}",
            "traceback": tb.splitlines(),
        }, ensure_ascii=False, indent=2)
        return Response(body, status=500, mimetype="application/json")
    # ===== [임시 진단] 끝 =====

    return app
