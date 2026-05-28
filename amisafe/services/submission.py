"""양식 제출/검증/상태 재계산.

- resolve_single_field_value : 필드별 최신 값 조회
- build_resolved_values_map : 화면 초기화용 값 맵
- build_label_values_map : 자동 채움 label 필드 처리
- validate_submission_for_current_user : 필수 입력 검증
- save_submission : DB 저장
- recalc_document_status : 완료 상태 재계산
"""
import json

from amisafe.extensions.db import get_conn
from amisafe.services.documents import fetch_document_values, fetch_participants
from amisafe.services.form_config import field_key, is_field_editable_for_user


def resolve_single_field_value(form, field, user, values):
    """한 필드의 가장 최근 저장값을 type에 맞춰 반환."""
    target = None

    if form["form_type"] == "individual":
        matches = [v for v in values if v["field_id"] == field["field_id"]]
        if matches:
            target = matches[-1]
    else:
        slot_index = field.get("slot_index")
        if slot_index is not None and str(slot_index) != "":
            matches = [
                v for v in values
                if v["field_id"] == field["field_id"]
                and str(v["slot_index"]) == str(slot_index)
            ]
            if matches:
                target = matches[-1]
        else:
            matches = [v for v in values if v["field_id"] == field["field_id"]]
            if matches:
                target = matches[-1]

    if not target:
        return None

    ftype = field["type"]
    if ftype == "signature":
        return target["value_image"]
    if ftype == "checkbox":
        return target["value_text"] in ("true", "checked", "1", True)
    return target["value_text"]


def build_resolved_values_map(form, user, values):
    """visible & 비-label 필드들의 (field_key → 값) 맵."""
    resolved = {}
    for field in form.get("fields", []):
        if not field.get("visible", True):
            continue
        if field.get("type") == "label":
            continue
        resolved[field_key(field)] = resolve_single_field_value(form, field, user, values)
    return resolved


def build_label_values_map(form, user, document, participants):
    """label 필드의 자동 바인딩 값 (이름/조/역할/날짜/ID)."""
    label_values = {}
    for field in form.get("fields", []):
        if field.get("type") != "label":
            continue

        bind_key = field.get("bind_key")
        key = field_key(field)
        slot_index = field.get("slot_index")
        target_role = field.get("target_role", "공통")

        participant = None
        if (
            form["form_type"] == "group"
            and slot_index is not None
            and str(slot_index) != ""
        ):
            for p in participants:
                if str(p["slot_index"]) == str(slot_index):
                    if target_role == "공통" or p["role"] == target_role:
                        participant = p
                        break

        if bind_key == "name":
            label_values[key] = participant["user_name"] if participant else user["name"]
        elif bind_key == "group":
            label_values[key] = document.get("group_name") or user["group"]
        elif bind_key == "role":
            label_values[key] = participant["role"] if participant else user["role"]
        elif bind_key == "date":
            label_values[key] = str(document.get("work_date"))
        elif bind_key == "id":
            label_values[key] = participant["user_id"] if participant else user["id"]
        else:
            label_values[key] = field.get("label", "")

    return label_values


def validate_submission_for_current_user(form, user, submitted_values):
    """현재 사용자가 책임진 필수 입력값 누락 여부 검사."""
    missing = []
    for field in form.get("fields", []):
        if not field.get("visible", True):
            continue
        if field.get("type") == "label":
            continue
        if not field.get("required", False):
            continue
        if not is_field_editable_for_user(form, field, user):
            continue

        key = field_key(field)
        val = submitted_values.get(key)
        ftype = field["type"]

        label = field.get("label") or field.get("field_id")
        if ftype == "text" and (val is None or str(val).strip() == ""):
            missing.append(label)
        elif ftype == "checkbox" and not bool(val):
            missing.append(label)
        elif ftype in ("choice_group", "dropdown") and (val is None or str(val).strip() == ""):
            missing.append(label)
        elif ftype == "signature" and (val is None or str(val).strip() == ""):
            missing.append(label)
        elif ftype in ("date", "datetime") and (val is None or str(val).strip() == ""):
            missing.append(label)

    return missing


def save_submission(document, form, user, submitted_values):
    """제출값을 DB에 저장 (기존 값 삭제 후 재삽입)."""
    editable_fields = [
        f for f in form.get("fields", [])
        if is_field_editable_for_user(form, f, user) and f.get("type") != "label"
    ]

    with get_conn() as conn:
        with conn.cursor() as cur:
            for field in editable_fields:
                f_id = field["field_id"]
                slot_index = field.get("slot_index")

                if slot_index is None or str(slot_index) == "":
                    cur.execute(
                        """
                        DELETE FROM document_values
                        WHERE document_id = %s AND field_id = %s
                          AND user_id = %s AND slot_index IS NULL
                        """,
                        (document["id"], f_id, user["id"]),
                    )
                else:
                    cur.execute(
                        """
                        DELETE FROM document_values
                        WHERE document_id = %s AND field_id = %s
                          AND user_id = %s AND slot_index = %s
                        """,
                        (document["id"], f_id, user["id"], slot_index),
                    )

                key = field_key(field)
                val = submitted_values.get(key)

                value_text = None
                value_json = None
                value_image = None
                is_completed = False

                ftype = field["type"]
                if ftype == "text" and val is not None and str(val).strip() != "":
                    value_text = str(val)
                    is_completed = True
                elif ftype == "checkbox" and bool(val):
                    value_text = "true"
                    is_completed = True
                elif ftype in ("choice_group", "dropdown") and val is not None and str(val).strip() != "":
                    value_text = str(val)
                    is_completed = True
                elif ftype == "signature" and val is not None and str(val).strip() != "":
                    value_image = str(val)
                    is_completed = True
                elif ftype in ("date", "datetime") and val is not None and str(val).strip() != "":
                    value_text = str(val)
                    is_completed = True

                if value_text is None and value_json is None and value_image is None:
                    continue

                cur.execute(
                    """
                    INSERT INTO document_values (
                        document_id, field_id, user_id, role, slot_index,
                        value_text, value_json, value_image, is_completed, updated_at
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s::jsonb, %s, %s, CURRENT_TIMESTAMP)
                    """,
                    (
                        document["id"], f_id, user["id"], user["role"],
                        field.get("slot_index"),
                        value_text,
                        json.dumps(value_json) if value_json is not None else None,
                        value_image,
                        is_completed,
                    ),
                )
        conn.commit()


def participant_required_fields(form, participant):
    """특정 참여자(역할+슬롯)에 적용되는 필수 필드."""
    result = []
    for field in form.get("fields", []):
        if not field.get("visible", True):
            continue
        if field.get("type") == "label":
            continue
        if not field.get("required", False):
            continue

        slot_index = field.get("slot_index")
        target_role = field.get("target_role", "공통")

        if form["form_type"] == "individual":
            if target_role in ("공통", participant["role"]):
                result.append(field)
            continue

        if slot_index is None or str(slot_index) == "":
            continue

        if str(slot_index) == str(participant["slot_index"]) and target_role in ("공통", participant["role"]):
            result.append(field)

    return result


def common_required_fields(form):
    """group 양식의 슬롯과 무관한 공통 필수 필드."""
    if form["form_type"] != "group":
        return []

    result = []
    for field in form.get("fields", []):
        if not field.get("visible", True):
            continue
        if field.get("type") == "label":
            continue
        if not field.get("required", False):
            continue
        slot_index = field.get("slot_index")
        if slot_index is None or str(slot_index) == "":
            result.append(field)
    return result


def field_has_completed_value(form, field, user, all_values):
    """필드에 의미 있는 값이 있는가."""
    val = resolve_single_field_value(form, field, user, all_values)
    ftype = field["type"]
    if ftype == "text":
        return val is not None and str(val).strip() != ""
    if ftype == "checkbox":
        return bool(val)
    if ftype in ("choice_group", "dropdown"):
        return val is not None and str(val).strip() != ""
    if ftype == "signature":
        return val is not None and str(val).strip() != ""
    if ftype in ("date", "datetime"):
        return val is not None and str(val).strip() != ""
    return True


def recalc_document_status(document, form, current_user):
    """문서 상태 재계산 후 DB 업데이트. {document_completed, ...} 반환."""
    all_values = fetch_document_values(document["id"])

    # 개인 양식
    if form["form_type"] == "individual":
        participant = {
            "user_id": current_user["id"],
            "user_name": current_user["name"],
            "group_name": current_user["group"],
            "role": current_user["role"],
            "slot_index": current_user["slot_index"],
        }
        req_fields = participant_required_fields(form, participant)
        missing = [f for f in req_fields if not field_has_completed_value(form, f, current_user, all_values)]
        completed = len(missing) == 0

        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE documents
                    SET status = %s, updated_at = CURRENT_TIMESTAMP,
                        completed_at = CASE WHEN %s THEN CURRENT_TIMESTAMP ELSE NULL END
                    WHERE id = %s
                    """,
                    ("completed" if completed else "in_progress", completed, document["id"]),
                )
            conn.commit()

        return {
            "document_completed": completed,
            "document_status_text": "completed" if completed else "in_progress",
            "participant_done": completed,
            "missing_common": [],
        }

    # 공동 양식
    participants = fetch_participants(document["id"])
    common_missing = []

    with get_conn() as conn:
        with conn.cursor() as cur:
            for p in participants:
                req_fields = participant_required_fields(form, p)
                missing = [f for f in req_fields if not field_has_completed_value(form, f, current_user, all_values)]
                is_done = len(missing) == 0
                cur.execute(
                    """
                    UPDATE document_participants
                    SET is_done = %s, updated_at = CURRENT_TIMESTAMP
                    WHERE document_id = %s AND user_id = %s AND role = %s AND slot_index = %s
                    """,
                    (is_done, document["id"], p["user_id"], p["role"], p["slot_index"]),
                )

            for field in common_required_fields(form):
                if not field_has_completed_value(form, field, current_user, all_values):
                    common_missing.append(field.get("label") or field.get("field_id"))
        conn.commit()

    participants = fetch_participants(document["id"])
    all_done = all(p["is_done"] for p in participants) if participants else False
    doc_completed = all_done and len(common_missing) == 0

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE documents
                SET status = %s, updated_at = CURRENT_TIMESTAMP,
                    completed_at = CASE WHEN %s THEN CURRENT_TIMESTAMP ELSE NULL END
                WHERE id = %s
                """,
                ("completed" if doc_completed else "in_progress", doc_completed, document["id"]),
            )
        conn.commit()

    current_participant = next(
        (
            p for p in participants
            if p["user_id"] == current_user["id"]
            and str(p["slot_index"]) == str(current_user["slot_index"])
            and p["role"] == current_user["role"]
        ),
        None,
    )

    return {
        "document_completed": doc_completed,
        "document_status_text": "completed" if doc_completed else "in_progress",
        "participant_done": bool(current_participant["is_done"]) if current_participant else False,
        "missing_common": common_missing,
    }
