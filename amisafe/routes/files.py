"""정적 자산 라우트: 양식 배경 이미지, 저장된 결과 이미지."""
import os
import re

from flask import Blueprint, send_from_directory, request, redirect, url_for, session

from amisafe.config import Config

bp = Blueprint("files", __name__)


@bp.route("/form-image/<path:filename>")
def form_image(filename):
    """양식의 배경 이미지 제공."""
    return send_from_directory(Config.FORMS_FOLDER, filename)


@bp.route("/data/<work_date>/<path:filename>")
def data_file(work_date, filename):
    """DATA 폴더에 저장된 완료 이미지 제공. download=1이면 첨부 다운로드."""
    user = session.get("user")
    if not user:
        return redirect(url_for("auth.home"))

    safe_work_date = re.sub(r"[^0-9\-]", "", work_date)
    directory = os.path.join(Config.DATA_FOLDER, safe_work_date)
    return send_from_directory(
        directory,
        filename,
        as_attachment=bool(request.args.get("download")),
    )
