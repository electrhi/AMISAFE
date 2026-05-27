"""로그인/로그아웃/대시보드 라우트."""
from flask import Blueprint, render_template, request, redirect, url_for, session

from amisafe.services.users import find_user
from amisafe.extensions.db import init_db
from amisafe.extensions.logging_setup import write_audit_log

bp = Blueprint("auth", __name__)


@bp.route("/init-db")
def init_db_route():
    try:
        init_db()
        return "DB 초기화 완료"
    except Exception as e:
        return f"DB 초기화 실패: {e}", 500


@bp.route("/", methods=["GET", "POST"])
def home():
    if session.get("user"):
        return redirect(url_for("auth.dashboard"))

    if request.method == "POST":
        user_id = request.form.get("user_id", "").strip()
        password = request.form.get("password", "").strip()

        try:
            user = find_user(user_id, password)
        except Exception as e:
            return render_template("login.html", error=f"로그인 처리 중 오류: {e}")

        if not user:
            return render_template(
                "login.html",
                error="아이디 또는 비밀번호가 올바르지 않거나 사용 중지된 계정입니다.",
            )

        session["user"] = user
        write_audit_log("login_success", user=user)
        return redirect(url_for("auth.dashboard"))

    return render_template("login.html", error=None)


@bp.route("/dashboard")
def dashboard():
    user = session.get("user")
    if not user:
        return redirect(url_for("auth.home"))

    if user.get("is_admin"):
        return redirect(url_for("admin.admin_page"))

    return redirect(url_for("forms.forms_page"))


@bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.home"))
