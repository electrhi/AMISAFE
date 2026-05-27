"""사용자용 양식 라우트: 목록과 양식 열기."""
import json

from flask import Blueprint, render_template, redirect, url_for, session

from amisafe.utils import get_today_str
from amisafe.services.form_config import get_form_by_id, get_available_forms_for_user
from amisafe.services.documents import (
    get_or_create_document, sync_group_participants,
    fetch_participants, fetch_document_values,
)
from amisafe.services.submission import (
    build_resolved_values_map, build_label_values_map, recalc_document_status,
)

bp = Blueprint("forms", __name__)


@bp.route("/forms")
def forms_page():
    user = session.get("user")
    if not user:
        return redirect(url_for("auth.home"))

    forms = get_available_forms_for_user(user)
    return render_template("forms.html", user=user, forms=forms)


@bp.route("/form/<form_id>")
def open_form(form_id):
    user = session.get("user")
    if not user:
        return redirect(url_for("auth.home"))

    form = get_form_by_id(form_id)
    if not form:
        return "양식을 찾을 수 없습니다.", 404
    if not form.get("active", False):
        return "비활성화된 양식입니다.", 403

    allowed_roles = form.get("allowed_roles", [])
    if allowed_roles and user["role"] not in allowed_roles:
        return "이 양식에 접근할 권한이 없습니다.", 403

    document = get_or_create_document(form, user, get_today_str())
    if form["form_type"] == "group":
        sync_group_participants(document["id"], user["group"])

    participants = fetch_participants(document["id"])
    values = fetch_document_values(document["id"])
    status = recalc_document_status(document, form, user)

    resolved_values = build_resolved_values_map(form, user, values)
    label_values = build_label_values_map(form, user, document, participants)
    image_url = url_for("files.form_image", filename=form.get("image_file", ""))

    return render_template(
        "form_run.html",
        user=user,
        form=form,
        image_url=image_url,
        document_id=document["id"],
        work_date=document["work_date"],
        doc_status=status["document_status_text"],
        form_json=json.dumps(form, ensure_ascii=False),
        user_json=json.dumps(user, ensure_ascii=False),
        doc_json=json.dumps({
            "document_id": document["id"],
            "work_date": document["work_date"],
            "group_name": document.get("group_name"),
            "user_id": document.get("user_id"),
        }, ensure_ascii=False),
        resolved_values_json=json.dumps(resolved_values, ensure_ascii=False),
        label_values_json=json.dumps(label_values, ensure_ascii=False),
        status_json=json.dumps(status, ensure_ascii=False),
    )
