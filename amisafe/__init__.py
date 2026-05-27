"""AMISAFE 애플리케이션 팩토리.

create_app() 호출 시 Flask 앱을 구성하고 Blueprint를 등록한다.
"""
from flask import Flask

from amisafe.config import Config
from amisafe.extensions.logging_setup import setup_logging
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

    # Blueprint 등록 순서 (의미상 일반적인 라우트가 먼저, API는 뒤)
    app.register_blueprint(auth.bp)
    app.register_blueprint(admin.bp)
    app.register_blueprint(designer.bp)
    app.register_blueprint(forms.bp)
    app.register_blueprint(api.bp)
    app.register_blueprint(files.bp)

    return app
