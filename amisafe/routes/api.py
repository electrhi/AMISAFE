"""양식 데이터 입출력 API."""
from flask import Blueprint, request, session, current_app

from amisafe.utils import get_today_str, make_json_response
from amisafe.extensions.logging_setup import write_audit_log, log_exception
from amisafe.services.form_config import get_form_by_id
from amisafe.services.documents import get_or_create_document
from amisafe.services.submission import (
    validate_submission_for_current_user, save_submission, recalc_document_status,
)
from amisafe.services.images import save_completed_image_to_data_folder

bp = Blueprint("api", __name__, url_prefix="/api")


def _check_form_access(form_id, user):
    """양식 조회 + 접근 권한 검사. (form, error_response) 반환."""
    form = get_form_by_id(form_id)
    if not form:
        return None, (make_json_response(False, message="양식을 찾을 수 없습니다."), 404)

    allowed_roles = form.get("allowed_roles", [])
    if allowed_roles and user["role"] not in allowed_roles:
        return None, (make_json_response(False, message="권한이 없습니다."), 403)

    return form, None


@bp.route("/form/<form_id>/save", methods=["POST"])
def api_save_form(form_id):
    user = session.get("user")
    if not user:
        return make_json_response(False, message="로그인이 필요합니다."), 401

    form, err = _check_form_access(form_id, user)
    if err is not None:
        return err

    document = get_or_create_document(form, user, get_today_str())

    payload = request.get_json(silent=True) or {}
    submitted_values = payload.get("values", {}) or {}

    missing = validate_submission_for_current_user(form, user, submitted_values)
    if missing:
        return make_json_response(
            False, message="필수 입력이 누락되었습니다.", missing=missing,
        ), 400

    try:
        save_submission(document, form, user, submitted_values)
        status = recalc_document_status(document, form, user)
    except Exception as e:
        log_exception(current_app, "api_save_form", e)
        return make_json_response(False, message=f"저장 처리 중 오류: {e}"), 500

    write_audit_log(
        "form_saved",
        user=user, form_id=form_id, document_id=document["id"],
        extra={
            "document_completed": status["document_completed"],
            "participant_done": status["participant_done"],
        },
    )

    msg = "저장되었습니다."
    if form["form_type"] == "group":
        if status["document_completed"]:
            msg = "저장되었습니다. 조 전체 문서가 완료되었습니다."
        elif status["participant_done"]:
            msg = "저장되었습니다. 현재 사용자 구역은 완료되었습니다."
        else:
            msg = "저장되었습니다. 아직 현재 사용자 구역의 필수 입력이 남아 있습니다."
    elif status["document_completed"]:
        msg = "저장되었습니다. 개인양식이 완료되었습니다."

    return make_json_response(
        True,
        message=msg,
        document_completed=status["document_completed"],
        participant_done=status["participant_done"],
        document_status_text=status["document_status_text"],
        missing_common=status["missing_common"],
    )


@bp.route("/form/<form_id>/save-image", methods=["POST"])
def api_save_form_image(form_id):
    user = session.get("user")
    if not user:
        return make_json_response(False, message="로그인이 필요합니다."), 401

    form, err = _check_form_access(form_id, user)
    if err is not None:
        return err

    document = get_or_create_document(form, user, get_today_str())
    status = recalc_document_status(document, form, user)

    if not status["document_completed"]:
        return make_json_response(False, message="문서 완료 상태에서만 서버 저장이 가능합니다."), 400

    payload = request.get_json(silent=True) or {}
    image_data_url = payload.get("image_data_url", "")

    try:
        saved = save_completed_image_to_data_folder(document, form, user, image_data_url)
        write_audit_log(
            "form_image_saved",
            user=user, form_id=form_id, document_id=document["id"],
            extra={
                "filename": saved["filename"],
                "saved_size_kb": saved.get("saved_size_kb"),
            },
        )
    except Exception as e:
        log_exception(current_app, "api_save_form_image", e)
        return make_json_response(False, message=f"서버 저장 실패: {e}"), 400

    return make_json_response(
        True,
        message="서버 DATA 폴더에 저장되었습니다.",
        saved_filename=saved["filename"],
        saved_relative_path=saved["relative_path"],
        saved_absolute_path=saved["absolute_path"],
        saved_size_bytes=saved.get("saved_size_bytes"),
        saved_size_kb=saved.get("saved_size_kb"),
        target_max_kb=saved.get("target_max_kb"),
    )


@bp.route("/form/<form_id>/status")
def api_form_status(form_id):
    user = session.get("user")
    if not user:
        return make_json_response(False, message="로그인이 필요합니다."), 401

    form, err = _check_form_access(form_id, user)
    if err is not None:
        return err

    document = get_or_create_document(form, user, get_today_str())
    status = recalc_document_status(document, form, user)

    return make_json_response(
        True,
        document_completed=status["document_completed"],
        participant_done=status["participant_done"],
        document_status_text=status["document_status_text"],
        missing_common=status["missing_common"],
    )
