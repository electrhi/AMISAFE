"""users.xlsx 기반 사용자/조 관리.

- 로그인 시 ID/PW 매칭
- 사용자/조/슬롯 정보 조회
- 조편성 변경 저장
"""
import os
import pandas as pd

from amisafe.config import Config


REQUIRED_COLUMNS = ["ID", "PW", "이름", "조", "구분", "슬롯순서", "관리자여부", "사용여부"]
ALLOWED_ROLES = {"작업원", "작업책임자"}


def load_users() -> pd.DataFrame:
    """users.xlsx에서 사용자 시트를 dtype=str로 로드한다."""
    if not os.path.exists(Config.USERS_XLSX_PATH):
        raise FileNotFoundError(f"{Config.USERS_XLSX_PATH} 파일을 찾을 수 없습니다.")

    df = pd.read_excel(
        Config.USERS_XLSX_PATH,
        sheet_name=Config.USERS_SHEET_NAME,
        dtype=str,
    ).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_users_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """필수 컬럼 존재 및 중복 검사."""
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"users 시트에 필요한 컬럼이 없습니다: {', '.join(missing)}")

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].astype(str).str.strip()

    # ID 중복
    duplicated_ids = df[df["ID"].duplicated() & (df["ID"] != "")]
    if not duplicated_ids.empty:
        dup_list = duplicated_ids["ID"].tolist()
        raise ValueError(f"중복된 ID가 있습니다: {dup_list}")

    # 같은 조/구분/슬롯순서 중복
    temp = df[(df["조"] != "") & (df["구분"] != "") & (df["슬롯순서"] != "")]
    dup_slots = temp[temp.duplicated(subset=["조", "구분", "슬롯순서"], keep=False)]
    if not dup_slots.empty:
        sample = dup_slots[["조", "구분", "슬롯순서"]].drop_duplicates().to_dict("records")
        raise ValueError(f"같은 조/구분/슬롯순서 중복이 있습니다: {sample}")

    return df


def _row_to_user(row, default_admin=None):
    """엑셀 행 → 사용자 dict 변환."""
    return {
        "id": row["ID"],
        "name": row["이름"],
        "group": row["조"],
        "role": row["구분"],
        "slot_index": int(row["슬롯순서"]) if str(row["슬롯순서"]).strip() else None,
        "is_admin": (
            default_admin if default_admin is not None
            else str(row["관리자여부"]).strip().upper() == "Y"
        ),
    }


def find_user(user_id: str, password: str):
    """ID/PW로 사용자 검색. 사용 중지 계정은 제외."""
    df = validate_users_dataframe(load_users())
    df["사용여부"] = df["사용여부"].str.upper()
    df["관리자여부"] = df["관리자여부"].str.upper()

    matched = df[
        (df["ID"] == user_id.strip())
        & (df["PW"] == password.strip())
        & (df["사용여부"] == "Y")
    ]
    if matched.empty:
        return None

    return _row_to_user(matched.iloc[0])


def get_group_users(group_name: str):
    """특정 조의 활성 사용자 목록 (역할/슬롯순 정렬)."""
    df = validate_users_dataframe(load_users())
    df["사용여부"] = df["사용여부"].str.upper()
    group_df = df[(df["조"] == group_name) & (df["사용여부"] == "Y")].copy()

    users = [_row_to_user(row) for _, row in group_df.iterrows()]
    users.sort(key=lambda x: (x["role"], x["slot_index"] if x["slot_index"] is not None else 9999, x["id"]))
    return users


def get_all_active_non_admin_users():
    """관리자가 아닌 활성 사용자 전체."""
    df = validate_users_dataframe(load_users())
    df["사용여부"] = df["사용여부"].str.upper()
    df["관리자여부"] = df["관리자여부"].str.upper()

    active_df = df[(df["사용여부"] == "Y") & (df["관리자여부"] != "Y")].copy()
    users = [_row_to_user(row, default_admin=False) for _, row in active_df.iterrows()]
    users.sort(key=lambda x: (
        x["group"], x["role"],
        x["slot_index"] if x["slot_index"] is not None else 9999,
        x["id"],
    ))
    return users


def get_group_map():
    """조 이름 → 사용자 리스트 매핑."""
    users = get_all_active_non_admin_users()
    group_map = {}
    for u in users:
        group_map.setdefault(str(u["group"]).strip(), []).append(u)

    for group_name in group_map:
        group_map[group_name].sort(
            key=lambda x: (x["slot_index"] if x["slot_index"] is not None else 9999, x["id"])
        )
    return dict(sorted(group_map.items(), key=lambda x: x[0]))


def save_users_dataframe(df: pd.DataFrame):
    """DataFrame을 users 시트에만 반영하여 저장(다른 시트 보존)."""
    safe_df = df.copy()
    for col in safe_df.columns:
        safe_df[col] = safe_df[col].astype("string")

    with pd.ExcelWriter(
        Config.USERS_XLSX_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        safe_df.to_excel(writer, sheet_name=Config.USERS_SHEET_NAME, index=False)


def apply_group_assignments(assignments):
    """관리자 화면에서 변경된 조/슬롯/역할 일괄 반영."""
    df = validate_users_dataframe(load_users()).copy()

    for col in REQUIRED_COLUMNS[:1] + REQUIRED_COLUMNS[2:]:  # PW 외 모든 메타 컬럼
        if col in df.columns:
            df[col] = df[col].astype("string")

    id_to_idx = {str(row["ID"]).strip(): idx for idx, row in df.iterrows()}

    for item in assignments:
        user_id = str(item.get("user_id", "")).strip()
        group_name = str(item.get("group", "")).strip()
        slot_index = str(item.get("slot_index", "")).strip()
        role = str(item.get("role", "")).strip()

        if user_id not in id_to_idx:
            continue

        idx = id_to_idx[user_id]
        df.at[idx, "조"] = group_name
        df.at[idx, "슬롯순서"] = slot_index
        if role in ALLOWED_ROLES:
            df.at[idx, "구분"] = role

    save_users_dataframe(df)
