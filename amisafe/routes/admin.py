"""관리자 라우트: 대시보드, 양식 열람, 조편성 저장, JSON 설정 편집, 이미지 업로드."""
import json
import os

from flask import Blueprint, render_template, request, redirect, url_for, session, jsonify

from amisafe.config import Config
from amisafe.utils import get_today_str, allowed_image_filename
from amisafe.extensions.logging_setup import write_audit_log
from amisafe.services.users import (
    get_all_active_non_admin_users, get_group_map, apply_group_assignments,
)
from amisafe.services.form_config import (
    load_form_config, save_form_config, get_form_by_id,
)
from amisafe.services.documents import (
    find_existing_document, build_admin_status_rows,
    sync_group_participants, fetch_participants, fetch_document_values,
)
from amisafe.services.submission import (
    build_resolved_values_map, build_label_values_map, recalc_document_status,
)

bp = Blueprint("admin", __name__, url_prefix="/admin")


def _require_admin():
    """관리자 여부 확인. 응답 반환되면 라우트에서 그대로 반환."""
    user = session.get("user")
    if not user:
        return None, redirect(url_for("auth.home"))
    if not user.get("is_admin"):
        return user, ("관리자만 접근 가능합니다.", 403)
    return user, None


@bp.route("")
def admin_page():
    user, response = _require_admin()
    if response is not None:
        return response

    selected_date = request.args.get("work_date", "").strip() or get_today_str()

    config = load_form_config()
    forms = [f for f in config.get("forms", []) if f.get("active", False)]
    status_rows = build_admin_status_rows(forms, selected_date)
    group_map = get_group_map()

    return render_template(
        "admin.html",
        user=user,
        forms=forms,
        selected_date=selected_date,
        status_rows=status_rows,
        group_map=group_map,
    )


@bp.route("/open/<form_id>/<target_user_id>/<work_date>")
def admin_open_form(form_id, target_user_id, work_date):
    viewer, response = _require_admin()
    if response is not None:
        return response

    form = get_form_by_id(form_id)
    if not form:
        return "양식을 찾을 수 없습니다.", 404

    target_user = next(
        (u for u in get_all_active_non_admin_users() if u["id"] == target_user_id),
        None,
    )
    if not target_user:
        return "대상 사용자를 찾을 수 없습니다.", 404

    allowed_roles = form.get("allowed_roles", [])
    if allowed_roles and target_user["role"] not in allowed_roles:
        return "이 사용자는 해당 양식 대상이 아닙니다.", 403

    document = find_existing_document(form, target_user, work_date)
    if not document:
        return "선택한 날짜에 작성된 문서가 없습니다.", 404

    if form["form_type"] == "group":
        sync_group_participants(document["id"], target_user["group"])

    participants = fetch_participants(document["id"])
    values = fetch_document_values(document["id"])
    status = recalc_document_status(document, form, target_user)

    resolved_values = build_resolved_values_map(form, target_user, values)
    label_values = build_label_values_map(form, target_user, document, participants)
    image_url = url_for("files.form_image", filename=form.get("image_file", ""))

    return render_template(
        "form_run.html",
        user=target_user,
        form=form,
        image_url=image_url,
        document_id=document["id"],
        work_date=document["work_date"],
        doc_status=status["document_status_text"],
        form_json=json.dumps(form, ensure_ascii=False),
        user_json=json.dumps(target_user, ensure_ascii=False),
        doc_json=json.dumps({
            "document_id": document["id"],
            "work_date": document["work_date"],
            "group_name": document.get("group_name"),
            "user_id": document.get("user_id"),
            "admin_view": True,
        }, ensure_ascii=False),
        resolved_values_json=json.dumps(resolved_values, ensure_ascii=False),
        label_values_json=json.dumps(label_values, ensure_ascii=False),
        status_json=json.dumps(status, ensure_ascii=False),
    )


@bp.route("/groups/save", methods=["POST"])
def admin_groups_save():
    user, response = _require_admin()
    if response is not None:
        # _require_admin returns a redirect or tuple; for an API endpoint, send JSON
        if not session.get("user"):
            return jsonify({"ok": False, "message": "로그인이 필요합니다."}), 401
        return jsonify({"ok": False, "message": "관리자만 접근 가능합니다."}), 403

    data = request.get_json(silent=True) or {}
    assignments = data.get("assignments", [])

    if not isinstance(assignments, list) or not assignments:
        return jsonify({"ok": False, "message": "저장할 조편성 데이터가 없습니다."}), 400

    try:
        apply_group_assignments(assignments)
        write_audit_log("group_assignment_saved", user=user, extra={"count": len(assignments)})
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "message": f"저장 실패: {e}"}), 500


@bp.route("/config", methods=["GET", "POST"])
def admin_config_page():
    user, response = _require_admin()
    if response is not None:
        return response

    success = None
    error = None

    if request.method == "POST":
        config_text = request.form.get("config_text", "")
        try:
            parsed = json.loads(config_text)
            if "forms" not in parsed or not isinstance(parsed["forms"], list):
                raise ValueError("JSON 최상위에 forms 배열이 있어야 합니다.")
            save_form_config(parsed)
            success = "form_config.json 저장 완료"
        except Exception as e:
            error = f"저장 실패: {e}"

    try:
        with open(Config.FORM_CONFIG_PATH, "r", encoding="utf-8") as f:
            config_text = f.read()
    except Exception:
        config_text = json.dumps({"forms": []}, ensure_ascii=False, indent=2)

    image_files = sorted(os.listdir(Config.FORMS_FOLDER)) if os.path.exists(Config.FORMS_FOLDER) else []

    return render_template(
        "admin_config.html",
        config_text=config_text,
        success=success,
        error=error,
        image_files=image_files,
    )


@bp.route("/upload-image", methods=["POST"])
def admin_upload_image():
    user, response = _require_admin()
    if response is not None:
        return response

    file = request.files.get("image_file")
    if not file or not file.filename:
        return redirect(url_for("admin.admin_config_page"))

    original_name = os.path.basename(file.filename)
    if not allowed_image_filename(original_name):
        return "이미지 파일만 업로드 가능합니다.", 400

    save_path = os.path.join(Config.FORMS_FOLDER, original_name)
    file.save(save_path)
    return redirect(url_for("admin.admin_config_page"))
