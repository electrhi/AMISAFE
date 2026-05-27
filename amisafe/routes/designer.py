"""웹 기반 양식 디자이너 (form_designer_pro_ui_excel.py의 웹 포팅).

데스크톱 Tkinter 앱을 HTML5 Canvas + JS로 재구현한다.
관리자만 접근 가능하다.

엔드포인트:
- GET  /designer                                 디자이너 UI
- GET  /designer/api/config                      현재 form_config.json
- POST /designer/api/config                      전체 form_config 저장
- POST /designer/api/import-category-excel       대/소분류 엑셀에서 dropdown options 추출
- GET  /designer/api/excel-template              엑셀 템플릿 다운로드
"""
import json
import os
from io import BytesIO

from flask import (
    Blueprint, render_template, request, redirect, url_for, session,
    jsonify, send_file,
)

from amisafe.config import Config
from amisafe.utils import allowed_image_filename
from amisafe.services.form_config import load_form_config, save_form_config

bp = Blueprint("designer", __name__, url_prefix="/designer")


def _require_admin_json():
    user = session.get("user")
    if not user:
        return None, (jsonify({"ok": False, "message": "로그인이 필요합니다."}), 401)
    if not user.get("is_admin"):
        return None, (jsonify({"ok": False, "message": "관리자만 접근 가능합니다."}), 403)
    return user, None


@bp.route("")
def designer_page():
    """디자이너 메인 페이지."""
    user = session.get("user")
    if not user:
        return redirect(url_for("auth.home"))
    if not user.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

    image_files = sorted(os.listdir(Config.FORMS_FOLDER)) if os.path.exists(Config.FORMS_FOLDER) else []
    image_files = [f for f in image_files if allowed_image_filename(f)]

    return render_template("designer.html", user=user, image_files=image_files)


@bp.route("/api/config", methods=["GET"])
def api_get_config():
    """현재 form_config.json 반환."""
    user, err = _require_admin_json()
    if err is not None:
        return err
    return jsonify({"ok": True, "config": load_form_config()})


@bp.route("/api/config", methods=["POST"])
def api_save_config():
    """form_config.json 전체 저장 (구조 검증 후)."""
    user, err = _require_admin_json()
    if err is not None:
        return err

    payload = request.get_json(silent=True) or {}
    if "forms" not in payload or not isinstance(payload["forms"], list):
        return jsonify({"ok": False, "message": "forms 배열이 필요합니다."}), 400

    try:
        save_form_config(payload)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "message": f"저장 실패: {e}"}), 500


@bp.route("/api/upload-image", methods=["POST"])
def api_upload_image():
    """디자이너에서 배경 이미지 업로드."""
    user, err = _require_admin_json()
    if err is not None:
        return err

    file = request.files.get("image_file")
    if not file or not file.filename:
        return jsonify({"ok": False, "message": "파일이 없습니다."}), 400

    original_name = os.path.basename(file.filename)
    if not allowed_image_filename(original_name):
        return jsonify({"ok": False, "message": "이미지 파일만 업로드 가능합니다."}), 400

    save_path = os.path.join(Config.FORMS_FOLDER, original_name)
    file.save(save_path)
    return jsonify({"ok": True, "filename": original_name})


@bp.route("/api/import-category-excel", methods=["POST"])
def api_import_category_excel():
    """엑셀(대분류/소분류)에서 dropdown options 추출.

    원본 form_designer.py의 _import_category_excel / _read_category_excel 와 동일 동작.
    A열=대분류, B열=소분류 형태를 읽어 parent_options[]와 option_map{대분류: [소분류]}을 반환.
    """
    user, err = _require_admin_json()
    if err is not None:
        return err

    file = request.files.get("excel_file")
    if not file or not file.filename:
        return jsonify({"ok": False, "message": "엑셀 파일이 없습니다."}), 400

    try:
        from openpyxl import load_workbook
    except ImportError:
        return jsonify({"ok": False, "message": "openpyxl이 설치되지 않았습니다."}), 500

    try:
        wb = load_workbook(BytesIO(file.read()), data_only=True)
        ws = wb.active

        parent_options = []
        seen_parents = set()
        option_map = {}
        rows_read = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or all(c is None for c in row):
                continue
            parent = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
            child = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            if not parent:
                continue

            rows_read += 1
            if parent not in seen_parents:
                seen_parents.add(parent)
                parent_options.append({"option_label": parent, "option_value": parent})
                option_map[parent] = []

            if child and child not in option_map[parent]:
                option_map[parent].append(child)

        # 소분류는 option_label/option_value 두 키로 변환
        normalized_map = {
            parent: [{"option_label": c, "option_value": c} for c in children]
            for parent, children in option_map.items()
        }

        return jsonify({
            "ok": True,
            "parent_options": parent_options,
            "option_map": normalized_map,
            "stats": {
                "rows_read": rows_read,
                "parent_count": len(parent_options),
                "child_total": sum(len(v) for v in normalized_map.values()),
            },
        })
    except Exception as e:
        return jsonify({"ok": False, "message": f"엑셀 읽기 실패: {e}"}), 400


@bp.route("/api/excel-template", methods=["GET"])
def api_excel_template():
    """대/소분류 엑셀 템플릿 다운로드."""
    user = session.get("user")
    if not user or not user.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        return "openpyxl이 설치되지 않았습니다.", 500

    wb = Workbook()
    ws = wb.active
    ws.title = "categories"

    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    ws["A1"] = "대분류"
    ws["B1"] = "소분류"
    for cell in (ws["A1"], ws["B1"]):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    samples = [
        ("점검항목", "정상"), ("점검항목", "이상"), ("점검항목", "해당없음"),
        ("작업유형", "고소작업"), ("작업유형", "용접작업"), ("작업유형", "전기작업"),
    ]
    for row_idx, (a, b) in enumerate(samples, start=2):
        ws.cell(row=row_idx, column=1, value=a)
        ws.cell(row=row_idx, column=2, value=b)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 24

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="대소분류_목록박스_양식.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
