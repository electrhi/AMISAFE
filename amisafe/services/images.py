"""완료된 양식 이미지의 최적화 및 DATA 폴더 저장."""
import base64
import binascii
import os
from io import BytesIO

from PIL import Image
from flask import url_for

from amisafe.config import Config
from amisafe.utils import sanitize_file_part, get_today_str


def build_server_image_filename(document, form, user, ext="jpg") -> str:
    """서버 저장용 파일명 규칙 - 공동: 조_YYYYMMDD_양식.확장자 / 개인: 사번_YYYYMMDD_양식.확장자"""
    form_name = sanitize_file_part(form.get("form_name") or form.get("form_id") or "form")
    work_date = str(document.get("work_date") or get_today_str())
    ymd = work_date.replace("-", "")

    if form.get("form_type") == "group":
        group_name = sanitize_file_part(document.get("group_name") or user.get("group") or "group")
        return f"{group_name}_{ymd}_{form_name}.{ext}"

    user_id = sanitize_file_part(document.get("user_id") or user.get("id") or "user")
    return f"{user_id}_{ymd}_{form_name}.{ext}"


def optimize_image_bytes(image_bytes: bytes, target_max_kb: int = None):
    """JPEG로 재인코딩하며 목표 용량 이하로 축소. (data, ext, size) 반환."""
    if target_max_kb is None:
        target_max_kb = Config.TARGET_IMAGE_MAX_KB
    target_bytes = int(target_max_kb * 1024)

    with Image.open(BytesIO(image_bytes)) as im:
        im = im.convert("RGB")

        # 처음 한 번 안전 축소
        max_side = max(im.size)
        if max_side > 2200:
            ratio = 2200 / max_side
            im = im.resize(
                (max(1, int(im.width * ratio)), max(1, int(im.height * ratio))),
                Image.LANCZOS,
            )

        best_bytes = None
        for max_side_try in [2200, 1800, 1600, 1400, 1200, 1000]:
            temp = im.copy()
            current_max = max(temp.size)
            if current_max > max_side_try:
                ratio = max_side_try / current_max
                temp = temp.resize(
                    (max(1, int(temp.width * ratio)), max(1, int(temp.height * ratio))),
                    Image.LANCZOS,
                )

            for quality in [88, 82, 76, 70, 64, 58, 52, 46]:
                buf = BytesIO()
                temp.save(buf, format="JPEG", quality=quality, optimize=True)
                data = buf.getvalue()

                if best_bytes is None:
                    best_bytes = data

                if len(data) <= target_bytes:
                    return data, "jpg", len(data)

                if len(data) < len(best_bytes):
                    best_bytes = data

        return best_bytes, "jpg", len(best_bytes)


def save_completed_image_to_data_folder(document, form, user, image_data_url):
    """클라이언트가 png base64로 보낸 완성 이미지를 최적화 후 저장."""
    if not image_data_url or not isinstance(image_data_url, str):
        raise ValueError("이미지 데이터가 없습니다.")

    prefix = "data:image/png;base64,"
    if not image_data_url.startswith(prefix):
        raise ValueError("지원하지 않는 이미지 형식입니다.")

    try:
        raw_image_bytes = base64.b64decode(image_data_url[len(prefix):], validate=True)
    except (binascii.Error, ValueError) as e:
        raise ValueError("이미지 디코딩에 실패했습니다.") from e

    optimized_bytes, ext, final_size = optimize_image_bytes(raw_image_bytes, Config.TARGET_IMAGE_MAX_KB)

    work_date = str(document.get("work_date") or get_today_str())
    save_dir = os.path.join(Config.DATA_FOLDER, work_date)
    os.makedirs(save_dir, exist_ok=True)

    filename = build_server_image_filename(document, form, user, ext=ext)
    save_path = os.path.join(save_dir, filename)
    with open(save_path, "wb") as f:
        f.write(optimized_bytes)

    return {
        "filename": filename,
        "relative_path": os.path.join(work_date, filename),
        "absolute_path": os.path.abspath(save_path),
        "saved_size_bytes": final_size,
        "saved_size_kb": round(final_size / 1024, 1),
        "target_max_kb": Config.TARGET_IMAGE_MAX_KB,
    }


def get_saved_image_info(document, form, user):
    """이미 저장된 결과 이미지가 있으면 메타데이터 반환."""
    if not document:
        return None

    work_date = str(document.get("work_date") or get_today_str())

    for ext in ("jpg", "png", "webp"):
        filename = build_server_image_filename(document, form, user, ext=ext)
        save_path = os.path.join(Config.DATA_FOLDER, work_date, filename)
        if os.path.exists(save_path):
            return {
                "filename": filename,
                "relative_path": os.path.join(work_date, filename),
                "absolute_path": os.path.abspath(save_path),
                "view_url": url_for("files.data_file", work_date=work_date, filename=filename),
                "download_url": url_for("files.data_file", work_date=work_date, filename=filename, download=1),
            }
    return None
