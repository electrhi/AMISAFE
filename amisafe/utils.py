"""공통 헬퍼: 날짜, JSON 응답, 파일명 정리."""
import re
from datetime import date
from flask import jsonify


def get_today_str() -> str:
    """오늘 날짜 YYYY-MM-DD."""
    return date.today().isoformat()


def make_json_response(ok: bool = True, **kwargs):
    """일관된 JSON 응답 헬퍼."""
    data = {"ok": ok}
    data.update(kwargs)
    return jsonify(data)


def sanitize_file_part(s) -> str:
    """파일명에 쓰일 수 없는 문자 제거 후 연속 _ 압축."""
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", str(s or ""))
    return re.sub(r"_+", "_", cleaned).strip("_")


def allowed_image_filename(filename: str) -> bool:
    """업로드 허용 이미지 확장자 검사."""
    lower = filename.lower()
    return any(lower.endswith(ext) for ext in (".jpg", ".jpeg", ".png", ".webp"))
