"""문서/참여자 DB 연산.

- 문서(documents) 생성/조회
- 참여자(document_participants) 동기화
- 값(document_values) 조회
- 관리자 화면 status row 빌드
"""
from flask import url_for

from amisafe.extensions.db import get_conn
from amisafe.services.users import get_all_active_non_admin_users, get_group_users
from amisafe.services.images import get_saved_image_info
from amisafe.utils import get_today_str


def _row_to_doc(row):
    """fetchone() 결과 row를 doc dict로."""
    return {
        "id": row[0],
        "form_id": row[1],
        "form_type": row[2],
        "group_name": row[3],
        "user_id": row[4],
        "work_date": str(row[5]),
        "status": row[6],
    }


def find_existing_document(form, target_user, work_date):
    """주어진 조건의 기존 문서 조회. 없으면 None."""
    form_id = form["form_id"]
    form_type = form["form_type"]

    with get_conn() as conn:
        with conn.cursor() as cur:
            if form_type == "group":
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'group'
                      AND group_name = %s AND work_date = %s
                    """,
                    (form_id, target_user["group"], work_date),
                )
            else:
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'individual'
                      AND user_id = %s AND work_date = %s
                    """,
                    (form_id, target_user["id"], work_date),
                )
            row = cur.fetchone()

    return _row_to_doc(row) if row else None


def get_or_create_document(form, user, work_date=None):
    """문서 조회/생성. group 양식은 참여자 동기화까지 수행."""
    if not work_date:
        work_date = get_today_str()

    form_id = form["form_id"]
    form_type = form["form_type"]
    created_group_doc = False

    with get_conn() as conn:
        with conn.cursor() as cur:
            if form_type == "group":
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'group'
                      AND group_name = %s AND work_date = %s
                    """,
                    (form_id, user["group"], work_date),
                )
                row = cur.fetchone()
                if row:
                    doc = _row_to_doc(row)
                else:
                    cur.execute(
                        """
                        INSERT INTO documents (form_id, form_type, group_name, user_id, work_date, status)
                        VALUES (%s, 'group', %s, NULL, %s, 'in_progress')
                        RETURNING id, form_id, form_type, group_name, user_id, work_date, status
                        """,
                        (form_id, user["group"], work_date),
                    )
                    doc = _row_to_doc(cur.fetchone())
                    created_group_doc = True
            else:
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'individual'
                      AND user_id = %s AND work_date = %s
                    """,
                    (form_id, user["id"], work_date),
                )
                row = cur.fetchone()
                if row:
                    doc = _row_to_doc(row)
                else:
                    cur.execute(
                        """
                        INSERT INTO documents (form_id, form_type, group_name, user_id, work_date, status)
                        VALUES (%s, 'individual', %s, %s, %s, 'in_progress')
                        RETURNING id, form_id, form_type, group_name, user_id, work_date, status
                        """,
                        (form_id, user["group"], user["id"], work_date),
                    )
                    doc = _row_to_doc(cur.fetchone())
        conn.commit()

    if created_group_doc:
        sync_group_participants(doc["id"], user["group"])

    return doc


def sync_group_participants(document_id, group_name):
    """그룹 문서에 현재 사용자(users.xlsx) 기준 참여자 추가."""
    users = [
        u for u in get_group_users(group_name)
        if not u["is_admin"] and u["role"] in ("작업원", "작업책임자")
    ]

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT user_id, role, slot_index FROM document_participants WHERE document_id = %s",
                (document_id,),
            )
            existing = {(r[0], r[1], r[2]) for r in cur.fetchall()}

            for u in users:
                key = (u["id"], u["role"], u["slot_index"])
                if key in existing:
                    continue
                cur.execute(
                    """
                    INSERT INTO document_participants
                        (document_id, user_id, user_name, group_name, role, slot_index, is_done)
                    VALUES (%s, %s, %s, %s, %s, %s, FALSE)
                    """,
                    (document_id, u["id"], u["name"], u["group"], u["role"], u["slot_index"]),
                )
        conn.commit()


def fetch_participants(document_id):
    """문서의 참여자 목록."""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT user_id, user_name, group_name, role, slot_index, is_done
                FROM document_participants
                WHERE document_id = %s
                ORDER BY role, slot_index, user_id
                """,
                (document_id,),
            )
            rows = cur.fetchall()

    return [
        {
            "user_id": r[0],
            "user_name": r[1],
            "group_name": r[2],
            "role": r[3],
            "slot_index": r[4],
            "is_done": r[5],
        }
        for r in rows
    ]


def fetch_document_values(document_id):
    """문서에 저장된 모든 값."""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, document_id, field_id, user_id, role, slot_index,
                       value_text, value_json, value_image, is_completed, updated_at
                FROM document_values
                WHERE document_id = %s
                ORDER BY updated_at ASC, id ASC
                """,
                (document_id,),
            )
            rows = cur.fetchall()

    return [
        {
            "id": r[0],
            "document_id": r[1],
            "field_id": r[2],
            "user_id": r[3],
            "role": r[4],
            "slot_index": r[5],
            "value_text": r[6],
            "value_json": r[7],
            "value_image": r[8],
            "is_completed": r[9],
            "updated_at": str(r[10]),
        }
        for r in rows
    ]


def build_admin_status_rows(forms, work_date):
    """관리자 대시보드용 사용자×양식 격자 데이터."""
    rows = []
    users = get_all_active_non_admin_users()

    for target_user in users:
        form_items = []

        for form in forms:
            allowed_roles = form.get("allowed_roles", [])
            if allowed_roles and target_user["role"] not in allowed_roles:
                form_items.append({
                    "form_id": form["form_id"],
                    "form_name": form["form_name"],
                    "open_url": None,
                    "dot_class": "no-doc",
                })
                continue

            doc = find_existing_document(form, target_user, work_date)

            saved_view_url = None
            saved_download_url = None
            if doc:
                status_done = doc.get("status") == "completed"
                open_url = url_for(
                    "admin.admin_open_form",
                    form_id=form["form_id"],
                    target_user_id=target_user["id"],
                    work_date=work_date,
                )
                dot_class = "done" if status_done else "not-done"
                saved_info = get_saved_image_info(doc, form, target_user) if status_done else None
                if saved_info:
                    saved_view_url = saved_info["view_url"]
                    saved_download_url = saved_info["download_url"]
            else:
                open_url = None
                dot_class = "not-done"

            form_items.append({
                "form_id": form["form_id"],
                "form_name": form["form_name"],
                "open_url": open_url,
                "dot_class": dot_class,
                "saved_view_url": saved_view_url,
                "saved_download_url": saved_download_url,
            })

        rows.append({"user": target_user, "form_items": form_items})

    return rows
