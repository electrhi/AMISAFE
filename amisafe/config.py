"""중앙 설정: 환경변수와 경로 상수를 한 곳에서 관리한다."""
import os
from dotenv import load_dotenv

load_dotenv()

# 프로젝트 루트 (이 파일은 <root>/amisafe/config.py 이므로 두 단계 상위)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _abs(path):
    """상대 경로는 프로젝트 루트 기준 절대 경로로 변환한다.

    Flask send_from_directory 는 상대 directory 를 app.root_path(=amisafe 패키지)
    기준으로 해석하므로, forms/DATA 등을 프로젝트 루트에 고정하려면 절대 경로가 필요하다.
    """
    return path if os.path.isabs(path) else os.path.join(BASE_DIR, path)


class Config:
    # 보안
    SECRET_KEY = os.getenv("SECRET_KEY", "default-secret")

    # DB
    DATABASE_URL = os.getenv("DATABASE_URL")

    # 파일 경로 (모두 프로젝트 루트 기준 절대 경로)
    USERS_XLSX_PATH = _abs(os.getenv("USERS_XLSX_PATH", "users.xlsx"))
    USERS_SHEET_NAME = os.getenv("USERS_SHEET_NAME", "users")
    FORM_CONFIG_PATH = _abs(os.getenv("FORM_CONFIG_PATH", "form_config.json"))
    FORMS_FOLDER = _abs(os.getenv("FORMS_FOLDER", "forms"))
    DATA_FOLDER = _abs(os.getenv("DATA_FOLDER", "DATA"))
    LOGS_FOLDER = _abs(os.getenv("LOGS_FOLDER", "logs"))

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
