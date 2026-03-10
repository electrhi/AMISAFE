import os
import re
import json
from datetime import date

import pandas as pd
import psycopg

from dotenv import load_dotenv
from flask import (
    Flask,
    request,
    session,
    redirect,
    url_for,
    render_template_string,
    jsonify,
    send_from_directory,
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "default-secret")

DATABASE_URL = os.getenv("DATABASE_URL")
USERS_XLSX_PATH = "users.xlsx"
USERS_SHEET_NAME = "users"
FORM_CONFIG_PATH = "form_config.json"
FORMS_FOLDER = "forms"

os.makedirs(FORMS_FOLDER, exist_ok=True)


LOGIN_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AMISAFE 로그인</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f5f7fb; margin: 0; padding: 0; }
        .wrap { max-width: 420px; margin: 60px auto; background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
        h2 { margin-top: 0; margin-bottom: 20px; }
        .error { color: #d93025; margin-bottom: 12px; }
        .field { margin-bottom: 14px; }
        label { display: block; margin-bottom: 6px; font-weight: bold; }
        input { width: 100%; box-sizing: border-box; padding: 10px; border: 1px solid #d0d7de; border-radius: 8px; }
        button { width: 100%; padding: 11px; border: 0; border-radius: 8px; background: #1f6feb; color: white; font-weight: bold; cursor: pointer; }
        button:hover { background: #1859be; }
        .info { margin-top: 16px; font-size: 13px; color: #666; }
    </style>
</head>
<body>
    <div class="wrap">
        <h2>AMISAFE 로그인</h2>
        {% if error %}
            <div class="error">{{ error }}</div>
        {% endif %}
        <form method="post">
            <div class="field">
                <label for="user_id">ID</label>
                <input id="user_id" type="text" name="user_id" required>
            </div>
            <div class="field">
                <label for="password">PW</label>
                <input id="password" type="password" name="password" required>
            </div>
            <button type="submit">로그인</button>
        </form>
        <div class="info">users.xlsx 의 users 시트를 기준으로 로그인합니다.</div>
    </div>
</body>
</html>
"""

ADMIN_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AMISAFE 관리자</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f5f7fb; margin: 0; padding: 0; }
        .wrap { max-width: 1100px; margin: 40px auto; background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
        h2, h3 { margin-top: 0; }
        .box { margin-top: 16px; padding: 16px; border: 1px solid #e5e7eb; border-radius: 10px; background: #f8fafc; }
        .menu a { display: inline-block; margin-right: 10px; margin-bottom: 10px; padding: 10px 14px; background: #1f6feb; color: white; text-decoration: none; border-radius: 8px; }
        .menu a.gray { background: #6b7280; }
        .form-card { border: 1px solid #e5e7eb; border-radius: 10px; padding: 14px; margin-bottom: 12px; background: white; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 12px; font-weight: bold; margin-left: 8px; }
        .badge.group { background: #e8f0fe; color: #1a73e8; }
        .badge.individual { background: #e6f4ea; color: #188038; }
    </style>
</head>
<body>
    <div class="wrap">
        <h2>관리자 화면</h2>
        <p><strong>{{ user.name }}</strong> 님으로 로그인됨</p>

        <div class="box">
            <h3>관리 메뉴</h3>
            <div class="menu">
                <a href="/init-db">DB 초기화 테스트</a>
                <a href="/admin">관리자 홈</a>
                <a href="/admin/config">양식 JSON 편집</a>
                <a class="gray" href="/logout">로그아웃</a>
            </div>
        </div>

        <div class="box">
            <h3>현재 등록된 양식</h3>
            {% if forms %}
                {% for form in forms %}
                <div class="form-card">
                    <div>
                        <strong>{{ form.form_name }}</strong>
                        {% if form.form_type == 'group' %}
                            <span class="badge group">공동양식</span>
                        {% else %}
                            <span class="badge individual">개인양식</span>
                        {% endif %}
                    </div>
                    <div style="margin-top:6px;">form_id: {{ form.form_id }}</div>
                    <div style="margin-top:6px;">이미지: {{ form.image_file }}</div>
                    <div style="margin-top:6px;">설명: {{ form.description }}</div>
                    <div style="margin-top:6px;">필드 수: {{ form.fields|length }}</div>
                </div>
                {% endfor %}
            {% else %}
                <p>등록된 양식이 없습니다.</p>
            {% endif %}
        </div>
    </div>
</body>
</html>
"""

ADMIN_CONFIG_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AMISAFE 양식 설정 편집</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f5f7fb; margin: 0; padding: 0; }
        .wrap { max-width: 1100px; margin: 40px auto; background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
        .box { margin-top: 16px; padding: 16px; border: 1px solid #e5e7eb; border-radius: 10px; background: #f8fafc; }
        textarea { width: 100%; min-height: 420px; box-sizing: border-box; border-radius: 8px; border: 1px solid #ccc; padding: 12px; font-family: Consolas, monospace; }
        button, a.btn { display: inline-block; padding: 10px 14px; background: #1f6feb; color: white; text-decoration: none; border-radius: 8px; border: none; cursor: pointer; margin-right: 8px; }
        a.gray { background: #6b7280; }
        .success { color: #0b8043; margin-bottom: 10px; }
        .error { color: #d93025; margin-bottom: 10px; }
    </style>
</head>
<body>
    <div class="wrap">
        <h2>양식 설정 JSON 편집</h2>

        <div class="box">
            {% if success %}
                <div class="success">{{ success }}</div>
            {% endif %}
            {% if error %}
                <div class="error">{{ error }}</div>
            {% endif %}

            <form method="post">
                <textarea name="config_text">{{ config_text }}</textarea>
                <div style="margin-top:12px;">
                    <button type="submit">JSON 저장</button>
                    <a class="btn gray" href="/admin">뒤로</a>
                </div>
            </form>
        </div>

        <div class="box">
            <h3>양식 이미지 업로드</h3>
            <form method="post" action="/admin/upload-image" enctype="multipart/form-data">
                <input type="file" name="image_file" accept=".jpg,.jpeg,.png,.webp" required>
                <button type="submit">업로드</button>
            </form>

            {% if image_files %}
                <div style="margin-top:14px;">
                    <strong>현재 forms 폴더 파일</strong>
                    <ul>
                    {% for f in image_files %}
                        <li>{{ f }}</li>
                    {% endfor %}
                    </ul>
                </div>
            {% endif %}
        </div>
    </div>
</body>
</html>
"""

FORMS_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AMISAFE 양식 목록</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f5f7fb; margin: 0; padding: 0; }
        .wrap { max-width: 1000px; margin: 40px auto; background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
        h2 { margin-top: 0; }
        .user-box { padding: 16px; border: 1px solid #e5e7eb; border-radius: 10px; background: #f8fafc; margin-bottom: 20px; }
        .form-card { border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; margin-bottom: 14px; background: #fff; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 12px; font-weight: bold; margin-left: 8px; }
        .badge.group { background: #e8f0fe; color: #1a73e8; }
        .badge.individual { background: #e6f4ea; color: #188038; }
        .btn { display: inline-block; padding: 10px 14px; background: #1f6feb; color: white; text-decoration: none; border-radius: 8px; margin-top: 10px; }
        .logout { display: inline-block; margin-top: 20px; color: #6b7280; text-decoration: none; }
    </style>
</head>
<body>
    <div class="wrap">
        <h2>내 양식 목록</h2>

        <div class="user-box">
            <div><strong>ID:</strong> {{ user.id }}</div>
            <div><strong>이름:</strong> {{ user.name }}</div>
            <div><strong>조:</strong> {{ user.group }}</div>
            <div><strong>구분:</strong> {{ user.role }}</div>
            <div><strong>슬롯순서:</strong> {{ user.slot_index }}</div>
        </div>

        {% if forms %}
            {% for form in forms %}
            <div class="form-card">
                <div>
                    <strong>{{ form.form_name }}</strong>
                    {% if form.form_type == 'group' %}
                        <span class="badge group">공동양식</span>
                    {% else %}
                        <span class="badge individual">개인양식</span>
                    {% endif %}
                </div>
                <div style="margin-top:8px;">설명: {{ form.description }}</div>
                <div style="margin-top:8px;">이미지: {{ form.image_file }}</div>
                <a class="btn" href="/form/{{ form.form_id }}">열기</a>
            </div>
            {% endfor %}
        {% else %}
            <p>표시할 양식이 없습니다.</p>
        {% endif %}

        <a class="logout" href="/logout">로그아웃</a>
    </div>
</body>
</html>
"""

FORM_RUN_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ form.form_name }}</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f5f7fb;
            margin: 0;
            padding: 0;
        }

        .wrap {
            max-width: 1400px;
            margin: 20px auto;
            background: #fff;
            border-radius: 12px;
            padding: 16px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.08);
            box-sizing: border-box;
        }

        .topbar {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 14px;
        }

        .btn {
            display: inline-block;
            padding: 10px 14px;
            background: #1f6feb;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            box-sizing: border-box;
        }

        .btn.gray { background: #6b7280; }
        .btn.green { background: #188038; }

        .info-box {
            padding: 12px;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            background: #f8fafc;
            margin-bottom: 12px;
            box-sizing: border-box;
        }

        #statusText {
            font-weight: bold;
        }

        .stage-wrap {
            overflow: auto;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            background: #fafafa;
            padding: 10px;
            -webkit-overflow-scrolling: touch;
        }

        .stage {
            position: relative;
            display: inline-block;
            transform-origin: top left;
        }

        .stage img {
            display: block;
            max-width: 100%;
            height: auto;
        }

        .overlay {
            position: absolute;
            left: 0;
            top: 0;
            right: 0;
            bottom: 0;
        }

        .field-item {
            position: absolute;
            box-sizing: border-box;
        }

        .field-item.readonly {
            pointer-events: none;
        }

        .field-text {
            width: 100%;
            height: 100%;
            box-sizing: border-box;
            border: 1px solid #1f6feb;
            border-radius: 4px;
            background: rgba(255,255,255,0.88);
            padding: 2px 4px;
            resize: none;
            overflow: hidden;
        }

        .field-checkbox {
            width: 100%;
            height: 100%;
            margin: 0;
            box-sizing: border-box;
        }

        .field-label {
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: flex-start;
            background: transparent;
            overflow: hidden;
            word-break: keep-all;
        }

        .readonly-box {
            width: 100%;
            height: 100%;
            border: 1px dashed #999;
            background: rgba(255,255,255,0.55);
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            box-sizing: border-box;
            text-align: center;
            word-break: break-word;
        }

        .signature-wrap {
            width: 100%;
            height: 100%;
            position: relative;
            border: 1px solid #1f6feb;
            border-radius: 4px;
            background: rgba(255,255,255,0.92);
            box-sizing: border-box;
            overflow: hidden;
        }

        .signature-canvas {
            width: 100%;
            height: 100%;
            display: block;
            touch-action: none;
        }

        .signature-clear {
            position: absolute;
            right: 2px;
            top: 2px;
            border: none;
            border-radius: 4px;
            background: #d93025;
            color: white;
            cursor: pointer;
            line-height: 1;
        }

        .choice-group {
            width: 100%;
            height: 100%;
            position: relative;
        }

        .choice-option {
            position: absolute;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 1px solid #1f6feb;
            background: rgba(255,255,255,0.9);
            box-sizing: border-box;
            border-radius: 999px;
            overflow: hidden;
        }

        .choice-option input {
            width: 100%;
            height: 100%;
            margin: 0;
        }

        .choice-option.readonly-choice {
            background: rgba(255,255,255,0.6);
            border: 1px dashed #999;
        }

        .tools {
            margin: 10px 0;
        }

        .small {
            color: #666;
            font-size: 13px;
        }

        @media (max-width: 768px) {
            .wrap {
                margin: 8px;
                padding: 10px;
                border-radius: 10px;
            }

            .topbar {
                align-items: stretch;
            }

            .topbar > div:last-child {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
            }

            .btn {
                padding: 9px 12px;
                font-size: 14px;
            }

            .info-box {
                font-size: 14px;
            }

            .small {
                font-size: 12px;
            }
        }
    </style>
</head>
<body>
<div class="wrap">
    <div class="topbar">
        <div>
            <h2 style="margin:0;">{{ form.form_name }}</h2>
            <div class="small">양식유형: {{ form.form_type }} / 문서상태: <span id="statusText">{{ doc_status }}</span></div>
        </div>
        <div>
            <button class="btn" id="saveBtn">저장</button>
            <button class="btn green" id="downloadBtn">다운로드</button>
            <a class="btn gray" href="/forms">목록</a>
            <a class="btn gray" href="/logout">로그아웃</a>
        </div>
    </div>

    <div class="info-box">
        <div><strong>사용자:</strong> {{ user.name }} ({{ user.id }}) / <strong>조:</strong> {{ user.group }} / <strong>구분:</strong> {{ user.role }} / <strong>슬롯:</strong> {{ user.slot_index }}</div>
        <div><strong>작업일자:</strong> {{ work_date }} / <strong>문서 ID:</strong> {{ document_id }}</div>
        <div class="small">공동양식은 같은 조가 한 장을 공유합니다. 개인양식은 본인 것만 저장합니다.</div>
    </div>

    <div class="tools">
        <label>배율:
            <input type="range" id="zoomRange" min="50" max="150" value="100">
            <span id="zoomLabel">100%</span>
        </label>
    </div>

    <div class="stage-wrap">
        <div id="captureArea" class="stage">
            <img id="formImage" src="{{ image_url }}" alt="{{ form.form_name }}">
            <div id="overlay" class="overlay"></div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<script>
const FORM_DEF = {{ form_json|safe }};
const USER = {{ user_json|safe }};
const DOC_META = {{ doc_json|safe }};
const RESOLVED_VALUES = {{ resolved_values_json|safe }};
const LABEL_VALUES = {{ label_values_json|safe }};
let CURRENT_STATUS = {{ status_json|safe }};

const img = document.getElementById("formImage");
const overlay = document.getElementById("overlay");
const captureArea = document.getElementById("captureArea");
const zoomRange = document.getElementById("zoomRange");
const zoomLabel = document.getElementById("zoomLabel");
const statusText = document.getElementById("statusText");

function fieldKey(field) {
    if (field.slot_index !== null && field.slot_index !== undefined && field.slot_index !== "") {
        return `${field.field_id}__slot_${field.slot_index}`;
    }
    return field.field_id;
}

function isEditable(field) {
    if (!field.visible) return false;
    const targetRole = field.target_role || "공통";

    if (FORM_DEF.form_type === "individual") {
        return targetRole === "공통" || targetRole === USER.role;
    }

    if (field.slot_index !== null && field.slot_index !== undefined && field.slot_index !== "") {
        return String(field.slot_index) === String(USER.slot_index)
            && (targetRole === "공통" || targetRole === USER.role);
    }

    return targetRole === "공통" || targetRole === USER.role;
}

function setStageSize() {
    overlay.style.width = img.clientWidth + "px";
    overlay.style.height = img.clientHeight + "px";
    captureArea.style.width = img.clientWidth + "px";
    captureArea.style.height = img.clientHeight + "px";
}

function px(v) {
    return `${v}px`;
}

function createReadonlyText(field, value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "readonly-box";
    wrap.textContent = value || "";
    wrap.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    return wrap;
}

function createEditableText(field, value, uiScale) {
    const input = document.createElement("textarea");
    input.className = "field-text";
    input.value = value || "";
    input.dataset.fieldId = field.field_id;
    input.dataset.fieldType = field.type;
    input.dataset.fieldKey = fieldKey(field);
    if (field.max_length) input.maxLength = field.max_length;
    if (field.placeholder) input.placeholder = field.placeholder;
    input.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    input.style.padding = `${Math.max(1, 3 * uiScale)}px ${Math.max(2, 4 * uiScale)}px`;
    return input;
}

function createEditableCheckbox(field, value, uiScale) {
    const input = document.createElement("input");
    input.type = "checkbox";
    input.className = "field-checkbox";
    input.checked = value === true || value === "true" || value === "checked" || value === 1;
    input.dataset.fieldId = field.field_id;
    input.dataset.fieldType = field.type;
    input.dataset.fieldKey = fieldKey(field);
    input.style.transform = `scale(${Math.max(0.75, uiScale)})`;
    input.style.transformOrigin = "center center";
    return input;
}

function createReadonlyCheckbox(value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "readonly-box";
    wrap.textContent = (value === true || value === "true" || value === "checked" || value === 1) ? "☑" : "☐";
    wrap.style.fontSize = `${Math.max(10, 18 * uiScale)}px`;
    return wrap;
}

function createEditableChoiceGroup(field, value, ratioX, ratioY, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "choice-group";
    const key = fieldKey(field);
    const options = field.options || [];

    options.forEach((opt) => {
        const div = document.createElement("div");
        div.className = "choice-option";
        div.style.left = px((opt.dx || 0) * ratioX);
        div.style.top = px((opt.dy || 0) * ratioY);
        div.style.width = px((opt.width || 20) * ratioX);
        div.style.height = px((opt.height || 20) * ratioY);

        const input = document.createElement("input");
        input.type = "radio";
        input.name = "cg_" + key;
        input.value = opt.option_value;
        input.checked = String(value || "") === String(opt.option_value);
        input.dataset.fieldId = field.field_id;
        input.dataset.fieldType = field.type;
        input.dataset.fieldKey = key;
        input.style.transform = `scale(${Math.max(0.7, uiScale)})`;
        input.style.transformOrigin = "center center";

        div.appendChild(input);
        wrap.appendChild(div);
    });

    return wrap;
}

function createReadonlyChoiceGroup(field, value, ratioX, ratioY, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "choice-group";
    const options = field.options || [];

    options.forEach((opt) => {
        const div = document.createElement("div");
        div.className = "choice-option readonly-choice";
        div.style.left = px((opt.dx || 0) * ratioX);
        div.style.top = px((opt.dy || 0) * ratioY);
        div.style.width = px((opt.width || 20) * ratioX);
        div.style.height = px((opt.height || 20) * ratioY);
        div.style.fontSize = `${Math.max(8, 12 * uiScale)}px`;
        div.textContent = String(value || "") === String(opt.option_value) ? "●" : "";
        wrap.appendChild(div);
    });

    return wrap;
}

function fitCanvasResolution(canvas) {
    const rect = canvas.getBoundingClientRect();
    const ratio = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.floor(rect.width * ratio));
    canvas.height = Math.max(1, Math.floor(rect.height * ratio));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    return ctx;
}

function loadDataUrlOnCanvas(canvas, dataUrl) {
    if (!dataUrl) return;
    const ctx = fitCanvasResolution(canvas);
    const rect = canvas.getBoundingClientRect();
    const image = new Image();
    image.onload = () => {
        ctx.clearRect(0, 0, rect.width, rect.height);
        ctx.drawImage(image, 0, 0, rect.width, rect.height);
    };
    image.src = dataUrl;
}

function attachSignatureDraw(canvas) {
    const ctx = fitCanvasResolution(canvas);
    let drawing = false;
    let lastX = 0;
    let lastY = 0;

    function getPos(e) {
        const rect = canvas.getBoundingClientRect();
        if (e.touches && e.touches[0]) {
            return {
                x: e.touches[0].clientX - rect.left,
                y: e.touches[0].clientY - rect.top
            };
        }
        return {
            x: e.clientX - rect.left,
            y: e.clientY - rect.top
        };
    }

    function start(e) {
        drawing = true;
        const p = getPos(e);
        lastX = p.x;
        lastY = p.y;
        e.preventDefault();
    }

    function move(e) {
        if (!drawing) return;
        const p = getPos(e);
        ctx.lineWidth = 2;
        ctx.lineCap = "round";
        ctx.strokeStyle = "#111";
        ctx.beginPath();
        ctx.moveTo(lastX, lastY);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
        lastX = p.x;
        lastY = p.y;
        e.preventDefault();
    }

    function end(e) {
        drawing = false;
        e.preventDefault();
    }

    canvas.addEventListener("mousedown", start);
    canvas.addEventListener("mousemove", move);
    canvas.addEventListener("mouseup", end);
    canvas.addEventListener("mouseleave", end);
    canvas.addEventListener("touchstart", start, { passive: false });
    canvas.addEventListener("touchmove", move, { passive: false });
    canvas.addEventListener("touchend", end, { passive: false });
}

function createEditableSignature(field, value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "signature-wrap";

    const canvas = document.createElement("canvas");
    canvas.className = "signature-canvas";
    canvas.dataset.fieldId = field.field_id;
    canvas.dataset.fieldType = field.type;
    canvas.dataset.fieldKey = fieldKey(field);

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "signature-clear";
    clearBtn.textContent = "지우기";
    clearBtn.style.fontSize = `${Math.max(8, 11 * uiScale)}px`;
    clearBtn.style.padding = `${Math.max(1, 2 * uiScale)}px ${Math.max(2, 4 * uiScale)}px`;

    clearBtn.onclick = () => {
        const ctx = fitCanvasResolution(canvas);
        const rect = canvas.getBoundingClientRect();
        ctx.clearRect(0, 0, rect.width, rect.height);
    };

    wrap.appendChild(canvas);
    wrap.appendChild(clearBtn);

    requestAnimationFrame(() => {
        fitCanvasResolution(canvas);
        attachSignatureDraw(canvas);
        if (value) loadDataUrlOnCanvas(canvas, value);
    });

    return wrap;
}

function createReadonlySignature(value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "readonly-box";
    if (value) {
        const img = document.createElement("img");
        img.src = value;
        img.style.maxWidth = "100%";
        img.style.maxHeight = "100%";
        wrap.appendChild(img);
    } else {
        wrap.textContent = "";
    }
    return wrap;
}

function createLabel(field, value, uiScale) {
    const div = document.createElement("div");
    div.className = "field-label";
    div.textContent = value || "";
    div.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    return div;
}

function renderFields() {
    setStageSize();
    overlay.innerHTML = "";

    const naturalW = img.naturalWidth || img.clientWidth;
    const naturalH = img.naturalHeight || img.clientHeight;
    const ratioX = img.clientWidth / naturalW;
    const ratioY = img.clientHeight / naturalH;
    const uiScale = Math.min(ratioX, ratioY);

    (FORM_DEF.fields || []).forEach((field) => {
        if (!field.visible) return;

        const key = fieldKey(field);
        const value = RESOLVED_VALUES[key];
        const labelValue = LABEL_VALUES[key];
        const editable = isEditable(field);

        const item = document.createElement("div");
        item.className = "field-item" + (editable ? "" : " readonly");
        item.style.left = px((field.x || 0) * ratioX);
        item.style.top = px((field.y || 0) * ratioY);
        item.style.width = px((field.width || 24) * ratioX);
        item.style.height = px((field.height || 24) * ratioY);

        let node = null;

        if (field.type === "label") {
            node = createLabel(field, labelValue, uiScale);
        } else if (field.type === "text") {
            node = editable ? createEditableText(field, value, uiScale) : createReadonlyText(field, value, uiScale);
        } else if (field.type === "checkbox") {
            node = editable ? createEditableCheckbox(field, value, uiScale) : createReadonlyCheckbox(value, uiScale);
        } else if (field.type === "choice_group") {
            node = editable ? createEditableChoiceGroup(field, value, ratioX, ratioY, uiScale) : createReadonlyChoiceGroup(field, value, ratioX, ratioY, uiScale);
        } else if (field.type === "signature") {
            node = editable ? createEditableSignature(field, value, uiScale) : createReadonlySignature(value, uiScale);
        } else {
            node = createReadonlyText(field, value, uiScale);
        }

        item.appendChild(node);
        overlay.appendChild(item);
    });
}

function getCurrentFieldValues() {
    const result = {};

    (FORM_DEF.fields || []).forEach((field) => {
        const key = fieldKey(field);
        if (!isEditable(field)) return;

        if (field.type === "label") return;

        if (field.type === "text") {
            const el = overlay.querySelector(`[data-field-key="${key}"]`);
            result[key] = el ? el.value : "";
            return;
        }

        if (field.type === "checkbox") {
            const el = overlay.querySelector(`[data-field-key="${key}"]`);
            result[key] = el ? !!el.checked : false;
            return;
        }

        if (field.type === "choice_group") {
            const checked = overlay.querySelector(`input[name="cg_${key}"]:checked`);
            result[key] = checked ? checked.value : "";
            return;
        }

        if (field.type === "signature") {
            const canvas = overlay.querySelector(`canvas[data-field-key="${key}"]`);
            if (canvas) {
                const blank = isCanvasBlank(canvas);
                result[key] = blank ? "" : canvas.toDataURL("image/png");
            } else {
                result[key] = "";
            }
        }
    });

    return result;
}

function isCanvasBlank(canvas) {
    const temp = document.createElement("canvas");
    temp.width = canvas.width;
    temp.height = canvas.height;
    return canvas.toDataURL() === temp.toDataURL();
}

function validateClientSide() {
    const values = getCurrentFieldValues();
    const missing = [];

    (FORM_DEF.fields || []).forEach((field) => {
        if (!field.visible || !field.required) return;
        if (!isEditable(field)) return;
        if (field.type === "label") return;

        const key = fieldKey(field);
        const v = values[key];

        if (field.type === "text" && (!v || String(v).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "checkbox" && !v) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "choice_group" && (!v || String(v).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "signature" && (!v || String(v).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
    });

    return missing;
}

async function saveForm() {
    const missing = validateClientSide();
    if (missing.length > 0) {
        alert("필수 입력이 누락되었습니다:\\n- " + missing.join("\\n- "));
        return { ok: false };
    }

    const values = getCurrentFieldValues();
    const resp = await fetch(`/api/form/${FORM_DEF.form_id}/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ values })
    });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
        alert(data.message || "저장에 실패했습니다.");
        return data;
    }

    CURRENT_STATUS = data;
    statusText.textContent = data.document_status_text;
    alert(data.message || "저장되었습니다.");
    return data;
}

async function refreshStatus() {
    const resp = await fetch(`/api/form/${FORM_DEF.form_id}/status`);
    const data = await resp.json();
    if (resp.ok && data.ok) {
        CURRENT_STATUS = data;
        statusText.textContent = data.document_status_text;
    }
    return data;
}

function sanitizeFilePart(s) {
    return String(s || "").replace(/[\\\\/:*?"<>|\\s]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
}

function getDownloadBaseFilename() {
    const formName = sanitizeFilePart(FORM_DEF.form_name);
    const ymd = String(DOC_META.work_date).replace(/-/g, "");

    if (FORM_DEF.form_type === "group") {
        return `${sanitizeFilePart(USER.group)}_${ymd}_${formName}`;
    }
    return `${sanitizeFilePart(USER.id)}_${ymd}_${formName}`;
}

function getNextDownloadFilename() {
    const base = getDownloadBaseFilename();
    const key = "download_counter::" + base;
    const count = Number(localStorage.getItem(key) || "0");
    const filename = count === 0 ? `${base}.png` : `${base}_${count}.png`;
    localStorage.setItem(key, String(count + 1));
    return filename;
}

async function downloadImage() {
    const saveResult = await saveForm();
    if (!saveResult.ok) return;

    const status = await refreshStatus();

    if (FORM_DEF.form_type === "group" && !status.document_completed) {
        alert("공동양식은 조 전체 필수 입력이 완료되어야 다운로드할 수 있습니다.");
        return;
    }

    if (!status.document_completed) {
        alert("필수 입력이 완료되지 않아 다운로드할 수 없습니다.");
        return;
    }

    const canvas = await html2canvas(document.getElementById("captureArea"), {
        useCORS: true,
        backgroundColor: "#ffffff",
        scale: 2
    });

    const link = document.createElement("a");
    link.href = canvas.toDataURL("image/png");
    link.download = getNextDownloadFilename();
    link.click();
}

function applyInitialMobileZoom() {
    const vw = window.innerWidth || 400;
    if (vw <= 480) {
        zoomRange.value = "85";
    } else if (vw <= 768) {
        zoomRange.value = "92";
    } else {
        zoomRange.value = "100";
    }
    const v = Number(zoomRange.value);
    zoomLabel.textContent = v + "%";
    captureArea.style.transform = `scale(${v / 100})`;
}

document.getElementById("saveBtn").addEventListener("click", saveForm);
document.getElementById("downloadBtn").addEventListener("click", downloadImage);

zoomRange.addEventListener("input", () => {
    const v = Number(zoomRange.value);
    zoomLabel.textContent = v + "%";
    captureArea.style.transform = `scale(${v / 100})`;
});

img.addEventListener("load", () => {
    setStageSize();
    renderFields();
    applyInitialMobileZoom();
    refreshStatus();
});

window.addEventListener("resize", () => {
    setStageSize();
    renderFields();
});
</script>
</body>
</html>
"""


def get_conn():
    if not DATABASE_URL:
        raise ValueError("DATABASE_URL 환경변수가 설정되지 않았습니다.")
    return psycopg.connect(DATABASE_URL)


def get_today_str():
    return date.today().isoformat()


def allowed_image_filename(filename: str) -> bool:
    lower = filename.lower()
    return lower.endswith(".jpg") or lower.endswith(".jpeg") or lower.endswith(".png") or lower.endswith(".webp")


def make_json_response(ok=True, **kwargs):
    data = {"ok": ok}
    data.update(kwargs)
    return jsonify(data)


def init_db():
    sql_list = [
        """
        CREATE TABLE IF NOT EXISTS documents (
            id SERIAL PRIMARY KEY,
            form_id TEXT NOT NULL,
            form_type TEXT NOT NULL,
            group_name TEXT,
            user_id TEXT,
            work_date DATE NOT NULL,
            status TEXT NOT NULL DEFAULT 'in_progress',
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            completed_at TIMESTAMP
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS document_participants (
            id SERIAL PRIMARY KEY,
            document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
            user_id TEXT NOT NULL,
            user_name TEXT NOT NULL,
            group_name TEXT,
            role TEXT NOT NULL,
            slot_index INTEGER,
            is_done BOOLEAN NOT NULL DEFAULT FALSE,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS document_values (
            id SERIAL PRIMARY KEY,
            document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
            field_id TEXT NOT NULL,
            user_id TEXT,
            role TEXT,
            slot_index INTEGER,
            value_text TEXT,
            value_json JSONB,
            value_image TEXT,
            is_completed BOOLEAN NOT NULL DEFAULT FALSE,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """,
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_group_document
        ON documents (form_id, group_name, work_date)
        WHERE form_type = 'group';
        """,
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_individual_document
        ON documents (form_id, user_id, work_date)
        WHERE form_type = 'individual';
        """
    ]

    with get_conn() as conn:
        with conn.cursor() as cur:
            for sql in sql_list:
                cur.execute(sql)
        conn.commit()


def load_users():
    if not os.path.exists(USERS_XLSX_PATH):
        raise FileNotFoundError(f"{USERS_XLSX_PATH} 파일을 찾을 수 없습니다.")

    df = pd.read_excel(
        USERS_XLSX_PATH,
        sheet_name=USERS_SHEET_NAME,
        dtype=str
    ).fillna("")

    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_users_dataframe(df: pd.DataFrame):
    required_columns = [
        "ID", "PW", "이름", "조", "구분",
        "슬롯순서", "관리자여부", "사용여부"
    ]

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"users 시트에 필요한 컬럼이 없습니다: {', '.join(missing)}")

    for col in required_columns:
        df[col] = df[col].astype(str).str.strip()

    duplicated_ids = df[df["ID"].duplicated() & (df["ID"] != "")]
    if not duplicated_ids.empty:
        dup_list = duplicated_ids["ID"].tolist()
        raise ValueError(f"중복된 ID가 있습니다: {dup_list}")

    temp = df[(df["조"] != "") & (df["구분"] != "") & (df["슬롯순서"] != "")]
    dup_slots = temp[temp.duplicated(subset=["조", "구분", "슬롯순서"], keep=False)]
    if not dup_slots.empty:
        sample = dup_slots[["조", "구분", "슬롯순서"]].drop_duplicates().to_dict("records")
        raise ValueError(f"같은 조/구분/슬롯순서 중복이 있습니다: {sample}")

    return df


def find_user(user_id: str, password: str):
    df = load_users()
    df = validate_users_dataframe(df)
    df["사용여부"] = df["사용여부"].str.upper()
    df["관리자여부"] = df["관리자여부"].str.upper()

    matched = df[
        (df["ID"] == user_id.strip()) &
        (df["PW"] == password.strip()) &
        (df["사용여부"] == "Y")
    ]

    if matched.empty:
        return None

    row = matched.iloc[0]
    return {
        "id": row["ID"],
        "name": row["이름"],
        "group": row["조"],
        "role": row["구분"],
        "slot_index": int(row["슬롯순서"]) if str(row["슬롯순서"]).strip() else None,
        "is_admin": row["관리자여부"] == "Y"
    }


def get_group_users(group_name: str):
    df = load_users()
    df = validate_users_dataframe(df)
    df["사용여부"] = df["사용여부"].str.upper()
    group_df = df[(df["조"] == group_name) & (df["사용여부"] == "Y")].copy()

    users = []
    for _, row in group_df.iterrows():
        users.append({
            "id": row["ID"],
            "name": row["이름"],
            "group": row["조"],
            "role": row["구분"],
            "slot_index": int(row["슬롯순서"]) if str(row["슬롯순서"]).strip() else None,
            "is_admin": str(row["관리자여부"]).strip().upper() == "Y"
        })

    users.sort(key=lambda x: (x["role"], x["slot_index"] if x["slot_index"] is not None else 9999, x["id"]))
    return users


def load_form_config():
    if not os.path.exists(FORM_CONFIG_PATH):
        return {"forms": []}

    with open(FORM_CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    if "forms" not in data or not isinstance(data["forms"], list):
        return {"forms": []}

    return data


def save_form_config(config_data):
    with open(FORM_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)


def get_form_by_id(form_id):
    config = load_form_config()
    for form in config.get("forms", []):
        if form.get("form_id") == form_id:
            return form
    return None


def get_available_forms_for_user(user):
    config = load_form_config()
    forms = []
    for form in config.get("forms", []):
        if not form.get("active", False):
            continue
        allowed_roles = form.get("allowed_roles", [])
        if allowed_roles and user["role"] not in allowed_roles:
            continue
        forms.append(form)
    return forms


def field_key(field):
    slot_index = field.get("slot_index")
    if slot_index is not None and str(slot_index) != "":
        return f"{field.get('field_id')}__slot_{slot_index}"
    return str(field.get("field_id"))


def is_field_editable_for_user(form, field, user):
    if not field.get("visible", True):
        return False

    target_role = field.get("target_role", "공통")
    form_type = form.get("form_type", "individual")
    slot_index = field.get("slot_index")

    if form_type == "individual":
        return target_role in ("공통", user["role"])

    if slot_index is not None and str(slot_index) != "":
        return str(slot_index) == str(user["slot_index"]) and target_role in ("공통", user["role"])

    return target_role in ("공통", user["role"])


def get_or_create_document(form, user, work_date=None):
    if not work_date:
        work_date = get_today_str()

    form_id = form["form_id"]
    form_type = form["form_type"]

    with get_conn() as conn:
        with conn.cursor() as cur:
            if form_type == "group":
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'group' AND group_name = %s AND work_date = %s
                    """,
                    (form_id, user["group"], work_date)
                )
                row = cur.fetchone()
                if row:
                    doc = {
                        "id": row[0],
                        "form_id": row[1],
                        "form_type": row[2],
                        "group_name": row[3],
                        "user_id": row[4],
                        "work_date": str(row[5]),
                        "status": row[6],
                    }
                else:
                    cur.execute(
                        """
                        INSERT INTO documents (form_id, form_type, group_name, user_id, work_date, status)
                        VALUES (%s, 'group', %s, NULL, %s, 'in_progress')
                        RETURNING id, form_id, form_type, group_name, user_id, work_date, status
                        """,
                        (form_id, user["group"], work_date)
                    )
                    row = cur.fetchone()
                    doc = {
                        "id": row[0],
                        "form_id": row[1],
                        "form_type": row[2],
                        "group_name": row[3],
                        "user_id": row[4],
                        "work_date": str(row[5]),
                        "status": row[6],
                    }
                    sync_group_participants(doc["id"], user["group"])
            else:
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'individual' AND user_id = %s AND work_date = %s
                    """,
                    (form_id, user["id"], work_date)
                )
                row = cur.fetchone()
                if row:
                    doc = {
                        "id": row[0],
                        "form_id": row[1],
                        "form_type": row[2],
                        "group_name": row[3],
                        "user_id": row[4],
                        "work_date": str(row[5]),
                        "status": row[6],
                    }
                else:
                    cur.execute(
                        """
                        INSERT INTO documents (form_id, form_type, group_name, user_id, work_date, status)
                        VALUES (%s, 'individual', %s, %s, %s, 'in_progress')
                        RETURNING id, form_id, form_type, group_name, user_id, work_date, status
                        """,
                        (form_id, user["group"], user["id"], work_date)
                    )
                    row = cur.fetchone()
                    doc = {
                        "id": row[0],
                        "form_id": row[1],
                        "form_type": row[2],
                        "group_name": row[3],
                        "user_id": row[4],
                        "work_date": str(row[5]),
                        "status": row[6],
                    }
            conn.commit()
    return doc


def sync_group_participants(document_id, group_name):
    users = [u for u in get_group_users(group_name) if not u["is_admin"] and u["role"] in ("작업원", "작업책임자")]

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT user_id, role, slot_index FROM document_participants WHERE document_id = %s",
                (document_id,)
            )
            existing = {(r[0], r[1], r[2]) for r in cur.fetchall()}

            for u in users:
                key = (u["id"], u["role"], u["slot_index"])
                if key in existing:
                    continue
                cur.execute(
                    """
                    INSERT INTO document_participants (document_id, user_id, user_name, group_name, role, slot_index, is_done)
                    VALUES (%s, %s, %s, %s, %s, %s, FALSE)
                    """,
                    (document_id, u["id"], u["name"], u["group"], u["role"], u["slot_index"])
                )
        conn.commit()


def fetch_participants(document_id):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT user_id, user_name, group_name, role, slot_index, is_done
                FROM document_participants
                WHERE document_id = %s
                ORDER BY role, slot_index, user_id
                """,
                (document_id,)
            )
            rows = cur.fetchall()

    participants = []
    for r in rows:
        participants.append({
            "user_id": r[0],
            "user_name": r[1],
            "group_name": r[2],
            "role": r[3],
            "slot_index": r[4],
            "is_done": r[5],
        })
    return participants


def fetch_document_values(document_id):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, document_id, field_id, user_id, role, slot_index, value_text, value_json, value_image, is_completed, updated_at
                FROM document_values
                WHERE document_id = %s
                ORDER BY updated_at ASC, id ASC
                """,
                (document_id,)
            )
            rows = cur.fetchall()

    values = []
    for r in rows:
        values.append({
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
        })
    return values


def resolve_single_field_value(form, field, user, values):
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
                if v["field_id"] == field["field_id"] and str(v["slot_index"]) == str(slot_index)
            ]
            if matches:
                target = matches[-1]
        else:
            matches = [v for v in values if v["field_id"] == field["field_id"]]
            if matches:
                target = matches[-1]

    if not target:
        return None

    if field["type"] == "signature":
        return target["value_image"]
    if field["type"] == "checkbox":
        return target["value_text"] in ("true", "checked", "1", True)
    if field["type"] == "choice_group":
        return target["value_text"]
    if field["type"] == "text":
        return target["value_text"]
    return target["value_text"]


def build_resolved_values_map(form, user, values):
    resolved = {}
    for field in form.get("fields", []):
        if not field.get("visible", True):
            continue
        if field.get("type") == "label":
            continue
        resolved[field_key(field)] = resolve_single_field_value(form, field, user, values)
    return resolved


def build_label_values_map(form, user, document, participants):
    label_values = {}
    for field in form.get("fields", []):
        if field.get("type") != "label":
            continue
        bind_key = field.get("bind_key")
        key = field_key(field)
        slot_index = field.get("slot_index")
        target_role = field.get("target_role", "공통")

        participant = None
        if form["form_type"] == "group" and slot_index is not None and str(slot_index) != "":
            for p in participants:
                if str(p["slot_index"]) == str(slot_index):
                    if target_role == "공통" or p["role"] == target_role:
                        participant = p
                        break

        if bind_key == "name":
            if participant:
                label_values[key] = participant["user_name"]
            else:
                label_values[key] = user["name"]
        elif bind_key == "group":
            label_values[key] = document.get("group_name") or user["group"]
        elif bind_key == "role":
            if participant:
                label_values[key] = participant["role"]
            else:
                label_values[key] = user["role"]
        elif bind_key == "date":
            label_values[key] = str(document.get("work_date"))
        elif bind_key == "id":
            if participant:
                label_values[key] = participant["user_id"]
            else:
                label_values[key] = user["id"]
        else:
            label_values[key] = field.get("label", "")
    return label_values


def validate_submission_for_current_user(form, user, submitted_values):
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

        if field["type"] == "text":
            if val is None or str(val).strip() == "":
                missing.append(field.get("label") or field.get("field_id"))
        elif field["type"] == "checkbox":
            if not bool(val):
                missing.append(field.get("label") or field.get("field_id"))
        elif field["type"] == "choice_group":
            if val is None or str(val).strip() == "":
                missing.append(field.get("label") or field.get("field_id"))
        elif field["type"] == "signature":
            if val is None or str(val).strip() == "":
                missing.append(field.get("label") or field.get("field_id"))

    return missing


def save_submission(document, form, user, submitted_values):
    editable_fields = [f for f in form.get("fields", []) if is_field_editable_for_user(form, f, user) and f.get("type") != "label"]

    with get_conn() as conn:
        with conn.cursor() as cur:
            for field in editable_fields:
                f_id = field["field_id"]
                slot_index = field.get("slot_index")

                if slot_index is None or str(slot_index) == "":
                    cur.execute(
                        """
                        DELETE FROM document_values
                        WHERE document_id = %s AND field_id = %s AND user_id = %s AND slot_index IS NULL
                        """,
                        (document["id"], f_id, user["id"])
                    )
                else:
                    cur.execute(
                        """
                        DELETE FROM document_values
                        WHERE document_id = %s AND field_id = %s AND user_id = %s AND slot_index = %s
                        """,
                        (document["id"], f_id, user["id"], slot_index)
                    )

                key = field_key(field)
                val = submitted_values.get(key)

                value_text = None
                value_json = None
                value_image = None
                is_completed = False

                if field["type"] == "text":
                    if val is not None and str(val).strip() != "":
                        value_text = str(val)
                        is_completed = True
                elif field["type"] == "checkbox":
                    if bool(val):
                        value_text = "true"
                        is_completed = True
                elif field["type"] == "choice_group":
                    if val is not None and str(val).strip() != "":
                        value_text = str(val)
                        is_completed = True
                elif field["type"] == "signature":
                    if val is not None and str(val).strip() != "":
                        value_image = str(val)
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
                        document["id"],
                        f_id,
                        user["id"],
                        user["role"],
                        field.get("slot_index"),
                        value_text,
                        json.dumps(value_json) if value_json is not None else None,
                        value_image,
                        is_completed,
                    )
                )

        conn.commit()


def participant_required_fields(form, participant):
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
    result = []
    if form["form_type"] != "group":
        return result

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
    val = resolve_single_field_value(form, field, user, all_values)
    if field["type"] == "text":
        return val is not None and str(val).strip() != ""
    if field["type"] == "checkbox":
        return bool(val)
    if field["type"] == "choice_group":
        return val is not None and str(val).strip() != ""
    if field["type"] == "signature":
        return val is not None and str(val).strip() != ""
    return True


def recalc_document_status(document, form, current_user):
    all_values = fetch_document_values(document["id"])

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
                    ("completed" if completed else "in_progress", completed, document["id"])
                )
            conn.commit()

        return {
            "document_completed": completed,
            "document_status_text": "completed" if completed else "in_progress",
            "participant_done": completed,
            "missing_common": [],
        }

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
                    (is_done, document["id"], p["user_id"], p["role"], p["slot_index"])
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
                ("completed" if doc_completed else "in_progress", doc_completed, document["id"])
            )
        conn.commit()

    current_participant = next(
        (p for p in participants if p["user_id"] == current_user["id"] and str(p["slot_index"]) == str(current_user["slot_index"]) and p["role"] == current_user["role"]),
        None
    )

    return {
        "document_completed": doc_completed,
        "document_status_text": "completed" if doc_completed else "in_progress",
        "participant_done": bool(current_participant["is_done"]) if current_participant else False,
        "missing_common": common_missing,
    }


@app.route("/init-db")
def init_db_route():
    try:
        init_db()
        return "DB 초기화 완료"
    except Exception as e:
        return f"DB 초기화 실패: {e}", 500


@app.route("/", methods=["GET", "POST"])
def home():
    if session.get("user"):
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        user_id = request.form.get("user_id", "").strip()
        password = request.form.get("password", "").strip()

        try:
            user = find_user(user_id, password)
        except Exception as e:
            return render_template_string(LOGIN_HTML, error=f"로그인 처리 중 오류: {e}")

        if not user:
            return render_template_string(LOGIN_HTML, error="아이디 또는 비밀번호가 올바르지 않거나 사용 중지된 계정입니다.")

        session["user"] = user
        return redirect(url_for("dashboard"))

    return render_template_string(LOGIN_HTML, error=None)


@app.route("/dashboard")
def dashboard():
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))

    if user.get("is_admin"):
        return redirect(url_for("admin_page"))

    return redirect(url_for("forms_page"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


@app.route("/admin")
def admin_page():
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))
    if not user.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

    config = load_form_config()
    forms = config.get("forms", [])
    return render_template_string(ADMIN_HTML, user=user, forms=forms)


@app.route("/admin/config", methods=["GET", "POST"])
def admin_config_page():
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))
    if not user.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

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
        with open(FORM_CONFIG_PATH, "r", encoding="utf-8") as f:
            config_text = f.read()
    except Exception:
        config_text = json.dumps({"forms": []}, ensure_ascii=False, indent=2)

    image_files = sorted(os.listdir(FORMS_FOLDER)) if os.path.exists(FORMS_FOLDER) else []

    return render_template_string(
        ADMIN_CONFIG_HTML,
        config_text=config_text,
        success=success,
        error=error,
        image_files=image_files
    )


@app.route("/admin/upload-image", methods=["POST"])
def admin_upload_image():
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))
    if not user.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

    file = request.files.get("image_file")
    if not file or not file.filename:
        return redirect(url_for("admin_config_page"))

    original_name = os.path.basename(file.filename)
    if not allowed_image_filename(original_name):
        return "이미지 파일만 업로드 가능합니다.", 400

    save_path = os.path.join(FORMS_FOLDER, original_name)
    file.save(save_path)
    return redirect(url_for("admin_config_page"))


@app.route("/form-image/<path:filename>")
def form_image(filename):
    return send_from_directory(FORMS_FOLDER, filename)


@app.route("/forms")
def forms_page():
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))

    forms = get_available_forms_for_user(user)
    return render_template_string(FORMS_HTML, user=user, forms=forms)


@app.route("/form/<form_id>")
def open_form(form_id):
    user = session.get("user")
    if not user:
        return redirect(url_for("home"))

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

    image_url = url_for("form_image", filename=form.get("image_file", ""))

    return render_template_string(
        FORM_RUN_HTML,
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
            "user_id": document.get("user_id")
        }, ensure_ascii=False),
        resolved_values_json=json.dumps(resolved_values, ensure_ascii=False),
        label_values_json=json.dumps(label_values, ensure_ascii=False),
        status_json=json.dumps(status, ensure_ascii=False),
    )


@app.route("/api/form/<form_id>/save", methods=["POST"])
def api_save_form(form_id):
    user = session.get("user")
    if not user:
        return make_json_response(False, message="로그인이 필요합니다."), 401

    form = get_form_by_id(form_id)
    if not form:
        return make_json_response(False, message="양식을 찾을 수 없습니다."), 404

    allowed_roles = form.get("allowed_roles", [])
    if allowed_roles and user["role"] not in allowed_roles:
        return make_json_response(False, message="권한이 없습니다."), 403

    document = get_or_create_document(form, user, get_today_str())

    payload = request.get_json(silent=True) or {}
    submitted_values = payload.get("values", {}) or {}

    missing = validate_submission_for_current_user(form, user, submitted_values)
    if missing:
        return make_json_response(False, message="필수 입력이 누락되었습니다.", missing=missing), 400

    save_submission(document, form, user, submitted_values)
    status = recalc_document_status(document, form, user)

    msg = "저장되었습니다."
    if form["form_type"] == "group":
        if status["document_completed"]:
            msg = "저장되었습니다. 조 전체 문서가 완료되었습니다."
        elif status["participant_done"]:
            msg = "저장되었습니다. 현재 사용자 구역은 완료되었습니다."
        else:
            msg = "저장되었습니다. 아직 현재 사용자 구역의 필수 입력이 남아 있습니다."
    else:
        if status["document_completed"]:
            msg = "저장되었습니다. 개인양식이 완료되었습니다."

    return make_json_response(
        True,
        message=msg,
        document_completed=status["document_completed"],
        participant_done=status["participant_done"],
        document_status_text=status["document_status_text"],
        missing_common=status["missing_common"],
    )


@app.route("/api/form/<form_id>/status")
def api_form_status(form_id):
    user = session.get("user")
    if not user:
        return make_json_response(False, message="로그인이 필요합니다."), 401

    form = get_form_by_id(form_id)
    if not form:
        return make_json_response(False, message="양식을 찾을 수 없습니다."), 404

    document = get_or_create_document(form, user, get_today_str())
    status = recalc_document_status(document, form, user)

    return make_json_response(
        True,
        document_completed=status["document_completed"],
        participant_done=status["participant_done"],
        document_status_text=status["document_status_text"],
        missing_common=status["missing_common"],
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
