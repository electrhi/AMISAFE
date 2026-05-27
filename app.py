"""AMISAFE 진입점.

Procfile / gunicorn 가 `app:app` 으로 참조한다.
- 개발 실행: `python app.py`
- 운영 실행: `gunicorn --bind :$PORT app:app`
"""
import os

from amisafe import create_app

app = create_app()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
