"""중앙 설정: 환경변수와 경로 상수를 한 곳에서 관리한다."""
import os
from dotenv import load_dotenv

load_dotenv()


class Config:
    # 보안
    SECRET_KEY = os.getenv("SECRET_KEY", "default-secret")

    # DB
    DATABASE_URL = os.getenv("DATABASE_URL")

    # 파일 경로
    USERS_XLSX_PATH = os.getenv("USERS_XLSX_PATH", "users.xlsx")
    USERS_SHEET_NAME = os.getenv("USERS_SHEET_NAME", "users")
    FORM_CONFIG_PATH = os.getenv("FORM_CONFIG_PATH", "form_config.json")
    FORMS_FOLDER = os.getenv("FORMS_FOLDER", "forms")
    DATA_FOLDER = os.getenv("DATA_FOLDER", "DATA")
    LOGS_FOLDER = os.getenv("LOGS_FOLDER", "logs")

    APP_LOG_PATH = os.path.join(LOGS_FOLDER, "app.log")
    AUDIT_LOG_PATH = os.path.join(LOGS_FOLDER, "audit.log")

    # 이미지 최적화 목표 용량 (KB)
    TARGET_IMAGE_MAX_KB = int(os.getenv("TARGET_IMAGE_MAX_KB", "300"))

    # 서버 포트 (gunicorn 환경에서는 사용 안 함, dev 모드용)
    PORT = int(os.getenv("PORT", "5000"))

    @classmethod
    def ensure_folders(cls):
        """필수 폴더 자동 생성."""
        for folder in (cls.FORMS_FOLDER, cls.DATA_FOLDER, cls.LOGS_FOLDER):
            os.makedirs(folder, exist_ok=True)
