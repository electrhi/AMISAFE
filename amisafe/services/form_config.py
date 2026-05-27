"""form_config.json 로드/저장 및 양식 단위 조회."""
import json
import os

from amisafe.config import Config


def load_form_config() -> dict:
    """양식 설정 전체 로드. 파일이 없거나 잘못되면 빈 forms."""
    if not os.path.exists(Config.FORM_CONFIG_PATH):
        return {"forms": []}

    with open(Config.FORM_CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    if "forms" not in data or not isinstance(data["forms"], list):
        return {"forms": []}
    return data


def save_form_config(config_data: dict):
    """양식 설정 전체 저장 (UTF-8, 들여쓰기 2)."""
    with open(Config.FORM_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)


def get_form_by_id(form_id: str):
    """ID로 단일 양식 조회."""
    config = load_form_config()
    for form in config.get("forms", []):
        if form.get("form_id") == form_id:
            return form
    return None


def get_available_forms_for_user(user: dict):
    """사용자(역할) 기준으로 사용 가능한 양식 목록."""
    config = load_form_config()
    result = []
    for form in config.get("forms", []):
        if not form.get("active", False):
            continue
        allowed_roles = form.get("allowed_roles", [])
        if allowed_roles and user["role"] not in allowed_roles:
            continue
        result.append(form)
    return result


def field_key(field: dict) -> str:
    """필드 식별 키 (slot_index 있으면 결합)."""
    slot_index = field.get("slot_index")
    if slot_index is not None and str(slot_index) != "":
        return f"{field.get('field_id')}__slot_{slot_index}"
    return str(field.get("field_id"))


def is_field_editable_for_user(form: dict, field: dict, user: dict) -> bool:
    """현재 사용자가 이 필드를 편집할 수 있는지."""
    if not field.get("visible", True):
        return False

    target_role = field.get("target_role", "공통")
    form_type = form.get("form_type", "individual")
    slot_index = field.get("slot_index")

    if form_type == "individual":
        return target_role in ("공통", user["role"])

    if slot_index is not None and str(slot_index) != "":
        return (
            str(slot_index) == str(user["slot_index"])
            and target_role in ("공통", user["role"])
        )

    return target_role in ("공통", user["role"])
