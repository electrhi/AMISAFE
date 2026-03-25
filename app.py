import os
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
        * { box-sizing: border-box; }
        body {
            margin: 0;
            font-family: Arial, sans-serif;
            background: #eef3f8;
            color: #1f2937;
        }
        .wrap {
            max-width: 1380px;
            margin: 16px auto;
            padding: 0 12px 24px;
        }
        .page-head {
            background: linear-gradient(135deg, #0f172a, #1d4ed8);
            color: white;
            border-radius: 18px;
            padding: 18px 20px;
            box-shadow: 0 12px 28px rgba(15, 23, 42, 0.18);
        }
        .page-head h2 {
            margin: 0 0 6px 0;
            font-size: 24px;
        }
        .page-head .sub {
            opacity: 0.9;
            font-size: 14px;
        }
        .top-actions {
            margin-top: 14px;
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            height: 42px;
            padding: 0 14px;
            border: 0;
            border-radius: 12px;
            background: #2563eb;
            color: white;
            font-weight: 700;
            text-decoration: none;
            cursor: pointer;
        }
        .btn.gray { background: #64748b; }
        .btn.green { background: #059669; }

        .card {
            margin-top: 16px;
            background: white;
            border-radius: 18px;
            padding: 18px;
            box-shadow: 0 10px 26px rgba(15, 23, 42, 0.08);
        }
        .card h3 {
            margin: 0 0 14px 0;
            font-size: 20px;
        }

        .filter-row {
            display: flex;
            flex-wrap: wrap;
            align-items: end;
            gap: 10px;
            margin-bottom: 14px;
        }
        .filter-row label {
            display: flex;
            flex-direction: column;
            gap: 6px;
            font-size: 14px;
            font-weight: 700;
        }
        .filter-row input[type="date"] {
            height: 42px;
            padding: 0 12px;
            border: 1px solid #cbd5e1;
            border-radius: 12px;
            font-size: 15px;
            background: white;
        }

        .legend {
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 14px;
            color: #475569;
            font-size: 13px;
        }
        .legend-item {
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }

        .status-dot {
            width: 16px;
            height: 16px;
            border-radius: 50%;
            display: inline-block;
            flex: 0 0 16px;
            box-shadow: inset 0 0 0 2px rgba(255,255,255,0.75), 0 0 0 1px rgba(0,0,0,0.08);
        }
        .status-dot.done { background: #16a34a; }
        .status-dot.not-done { background: #dc2626; }
        .status-dot.no-doc { background: #cbd5e1; }

        .desktop-only { display: block; }
        .mobile-only { display: none; }

        .admin-grid-wrap {
            overflow: auto;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            background: white;
        }
        table.admin-grid {
            width: 100%;
            border-collapse: collapse;
            min-width: 1100px;
        }
        .admin-grid th,
        .admin-grid td {
            border-bottom: 1px solid #edf2f7;
            border-right: 1px solid #f1f5f9;
            padding: 12px 10px;
            text-align: center;
            vertical-align: middle;
            background: white;
        }
        .admin-grid th {
            position: sticky;
            top: 0;
            z-index: 3;
            background: #f8fafc;
            font-size: 14px;
        }
        .admin-grid th:first-child,
        .admin-grid td:first-child {
            position: sticky;
            left: 0;
            z-index: 2;
            background: white;
            text-align: left;
            min-width: 220px;
        }
        .user-name {
            font-weight: 800;
            font-size: 15px;
            margin-bottom: 4px;
        }
        .user-meta {
            font-size: 12px;
            color: #64748b;
            line-height: 1.5;
        }
        .form-link {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            text-decoration: none;
            color: #111827;
            font-weight: 700;
            line-height: 1.35;
            padding: 4px 6px;
            border-radius: 10px;
        }
        .form-link:hover {
            background: #eff6ff;
            color: #1d4ed8;
        }

        .mobile-status-list {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        .mobile-user-card {
            background: #fff;
            border: 1px solid #e2e8f0;
            border-radius: 18px;
            padding: 12px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
        }
        .mobile-user-title {
            font-size: 17px;
            font-weight: 800;
            margin-bottom: 4px;
            color: #0f172a;
        }
        .mobile-user-meta {
            font-size: 12px;
            color: #64748b;
            margin-bottom: 10px;
            line-height: 1.5;
        }
        .mobile-form-list {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        .mobile-form-item,
        .mobile-form-item:visited {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            padding: 10px 12px;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            background: #f8fafc;
            text-decoration: none;
            color: #111827;
            font-weight: 700;
            min-height: unset;
            height: auto;
        }
        .mobile-form-left {
            display: flex;
            align-items: center;
            gap: 10px;
            min-width: 0;
        }
        .mobile-form-name {
            word-break: keep-all;
            line-height: 1.35;
        }

        .group-board {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 14px;
        }
        .group-col {
            background: #f8fafc;
            border: 1px solid #dbe5ef;
            border-radius: 16px;
            padding: 12px;
        }
        .group-col h4 {
            margin: 0 0 10px 0;
            font-size: 16px;
        }
        .dropzone {
            min-height: 70px;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .member-card {
            background: white;
            border: 1px solid #dbe5ef;
            border-radius: 14px;
            padding: 10px 12px;
            cursor: grab;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
        }
        .member-card:active {
            cursor: grabbing;
        }
        .member-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
        }
        .member-name {
            font-weight: 800;
            font-size: 15px;
        }
        .member-id {
            font-size: 12px;
            color: #64748b;
            margin-top: 4px;
        }
        .pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 11px;
            font-weight: 800;
            background: #e0ecff;
            color: #1d4ed8;
            white-space: nowrap;
        }

        .save-row {
            margin-top: 14px;
            display: flex;
            justify-content: flex-end;
        }

        @media (max-width: 768px) {
            .wrap { padding: 0 8px 18px; }
            .page-head { border-radius: 14px; padding: 14px; }
            .page-head h2 { font-size: 20px; }
            .card { border-radius: 14px; padding: 14px; }
            .btn { height: 40px; padding: 0 12px; font-size: 14px; }
            .filter-row input[type="date"] { width: 100%; }
            .desktop-only { display: none; }
            .mobile-only { display: block; }
            .legend {
                gap: 10px;
                font-size: 12px;
                align-items: center;
            }
            .legend-item {
                gap: 6px;
                white-space: nowrap;
            }
            .group-board {
                grid-template-columns: 1fr;
                gap: 12px;
            }
            .group-col {
                padding: 12px;
                border-radius: 14px;
            }
            .member-card {
                padding: 10px 12px;
                border-radius: 12px;
            }
            .member-name {
                font-size: 16px;
            }
            .member-id {
                font-size: 12px;
            }
        }
    </style>
</head>
<body>
    <div class="wrap">
        <div class="page-head">
            <h2>관리자 화면</h2>
            <div class="sub">{{ user.name }} 님으로 로그인됨</div>
            <div class="top-actions">
                <a class="btn" href="/admin">제출 현황</a>
                <a class="btn green" href="#group-board">조편성</a>
                <a class="btn gray" href="/logout">로그아웃</a>
            </div>
        </div>

        <div class="card">
            <h3>날짜별 제출 현황</h3>
            <form method="get" action="/admin" class="filter-row">
                <label>
                    조회 날짜
                    <input type="date" name="work_date" value="{{ selected_date }}">
                </label>
                <button class="btn" type="submit">조회</button>
            </form>

            <div class="legend">
                <div class="legend-item"><span class="status-dot done"></span> 작성 완료</div>
                <div class="legend-item"><span class="status-dot not-done"></span> 미작성 / 미완료</div>
                <div class="legend-item"><span class="status-dot no-doc"></span> 대상 아님 / 문서 없음</div>
            </div>

            <div class="desktop-only">
                <div class="admin-grid-wrap">
                    <table class="admin-grid">
                        <thead>
                            <tr>
                                <th>사용자</th>
                                {% for form in forms %}
                                    <th>{{ form.form_name }}</th>
                                {% endfor %}
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in status_rows %}
                            <tr>
                                <td>
                                    <div class="user-name">{{ row.user.group }}조 · {{ row.user.name }}</div>
                                    <div class="user-meta">
                                        ID: {{ row.user.id }}<br>
                                        구분: {{ row.user.role }} / 슬롯: {{ row.user.slot_index }}
                                    </div>
                                </td>
                                {% for item in row.form_items %}
                                <td>
                                    {% if item.open_url %}
                                        <a class="form-link" href="{{ item.open_url }}">
                                            <span class="status-dot {{ item.dot_class }}"></span>
                                            <span>{{ item.form_name }}</span>
                                        </a>
                                    {% else %}
                                        <div class="form-link" style="cursor:default;">
                                            <span class="status-dot {{ item.dot_class }}"></span>
                                            <span>{{ item.form_name }}</span>
                                        </div>
                                    {% endif %}
                                </td>
                                {% endfor %}
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="mobile-only">
                <div class="mobile-status-list">
                    {% for row in status_rows %}
                    <div class="mobile-user-card">
                        <div class="mobile-user-title">{{ row.user.group }}조 · {{ row.user.name }}</div>
                        <div class="mobile-user-meta">
                            ID: {{ row.user.id }} / {{ row.user.role }} / 슬롯 {{ row.user.slot_index }}
                        </div>
                        <div class="mobile-form-list">
                            {% for item in row.form_items %}
                                {% if item.open_url %}
                                <a class="mobile-form-item" href="{{ item.open_url }}">
                                    <div class="mobile-form-left">
                                        <span class="status-dot {{ item.dot_class }}"></span>
                                        <span class="mobile-form-name">{{ item.form_name }}</span>
                                    </div>
                                </a>
                                {% else %}
                                <div class="mobile-form-item">
                                    <div class="mobile-form-left">
                                        <span class="status-dot {{ item.dot_class }}"></span>
                                        <span class="mobile-form-name">{{ item.form_name }}</span>
                                    </div>
                                </div>
                                {% endif %}
                            {% endfor %}
                        </div>
                    </div>
                    {% endfor %}
                </div>
            </div>
        </div>

        <div class="card" id="group-board">
            <h3>조편성 인원 조정</h3>
            <div class="group-board">
                {% for group_name, members in group_map.items() %}
                <div class="group-col">
                    <h4>{{ group_name }}조</h4>
                    <div class="dropzone" data-group="{{ group_name }}">
                        {% for m in members %}
                        <div class="member-card"
                             data-user-id="{{ m.id }}"
                             data-role="{{ m.role }}"
                             data-slot-index="{{ m.slot_index }}">
                            <div class="member-top">
                                <div class="member-name">{{ m.group }}조 · {{ m.name }}</div>
                                <div class="pill">{{ m.role }}</div>
                            </div>
                            <div class="member-id">{{ m.id }} / 슬롯 {{ m.slot_index }}</div>
                        </div>
                        {% endfor %}
                    </div>
                </div>
                {% endfor %}
            </div>

            <div class="save-row">
                <button class="btn" type="button" id="saveGroupBtn">조편성 저장</button>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.2/Sortable.min.js"></script>
    <script>
        const zones = document.querySelectorAll(".dropzone");

        zones.forEach((zone) => {
            new Sortable(zone, {
                group: "groups",
                animation: 180,
                ghostClass: "drag-ghost",
            });
        });

        document.getElementById("saveGroupBtn").addEventListener("click", async () => {
            const assignments = [];

            document.querySelectorAll(".dropzone").forEach((zone) => {
                const groupName = zone.dataset.group;
                [...zone.querySelectorAll(".member-card")].forEach((card, idx) => {
                    assignments.push({
                        user_id: card.dataset.userId,
                        group: groupName,
                        slot_index: idx + 1
                    });
                });
            });

            const resp = await fetch("/admin/groups/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ assignments })
            });

            const data = await resp.json();
            if (!resp.ok || !data.ok) {
                alert(data.message || "저장에 실패했습니다.");
                return;
            }
            alert("조편성이 저장되었습니다.");
            location.reload();
        });
    </script>
</body>
</html>
"""
FORM_RUN_HTML = """
<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <title>{{ form.form_name }}</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: Arial, sans-serif;
            background: #f5f7fb;
            margin: 0;
            padding: 0;
        }

        .wrap {
            max-width: 1400px;
            margin: 16px auto;
            background: #fff;
            border-radius: 12px;
            padding: 16px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.08);
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
            min-height: 42px;
        }

        .btn.gray { background: #6b7280; }
        .btn.green { background: #188038; }
        .btn.red { background: #d93025; }

        .info-box {
            padding: 14px;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            background: #f8fafc;
            margin-bottom: 12px;
            line-height: 1.5;
        }

        #statusText { font-weight: bold; }

        .tools {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: center;
            margin: 10px 0 14px 0;
        }

        .tools input[type="range"] {
            vertical-align: middle;
            width: 240px;
            max-width: 60vw;
        }

        .small {
            color: #666;
            font-size: 13px;
        }

        .stage-wrap {
            overflow: auto;
            border-radius: 18px;
            background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
            padding: 12px;
            -webkit-overflow-scrolling: touch;
            touch-action: pan-x pan-y;
        }

        .stage {
            position: relative;
            display: inline-block;
            transform-origin: top left;
            line-height: 0;
        }

        .stage img {
            display: block;
            width: 100%;
            height: auto;
            user-select: none;
            -webkit-user-drag: none;
            pointer-events: none;
            border-radius: 12px;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.12);
        }

        .overlay {
            position: absolute;
            inset: 0;
        }

        .field-item {
            position: absolute;
        }

        .field-item.readonly {
            pointer-events: none;
        }

        .field-label {
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: flex-start;
            overflow: hidden;
            white-space: pre-wrap;
            word-break: break-word;
            line-height: 1.2;
            background: transparent;
        }

        .readonly-box {
            width: 100%;
            height: 100%;
            border: 1px solid rgba(15, 23, 42, 0.12);
            background: rgba(255,255,255,0.55);
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            text-align: center;
            word-break: break-word;
            line-height: 1.2;
            padding: 2px 4px;
            color: #111827;
            box-shadow: 0 1px 4px rgba(15, 23, 42, 0.08);
        }

        .tap-box {
            width: 100%;
            height: 100%;
            border: 1px solid rgba(15, 23, 42, 0.12);
            background: rgba(255,255,255,0.72);
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: flex-start;
            overflow: hidden;
            position: relative;
            cursor: pointer;
            padding: 3px 6px;
            line-height: 1.2;
            white-space: pre-wrap;
            word-break: break-word;
            color: #111827;
            box-shadow: 0 1px 4px rgba(15, 23, 42, 0.08);
            backdrop-filter: blur(2px);
        }

        .tap-box.required-empty {
            border-color: #d93025;
            background: rgba(255, 244, 244, 0.95);
        }

        .tap-box .placeholder {
            color: #999;
        }

        .checkbox-box {
            width: 100%;
            height: 100%;
            border: 2px solid #94a3b8;
            border-radius: 8px;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            user-select: none;
            padding: 0;
            position: relative;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.08);
            transition: all 0.16s ease;
            -webkit-tap-highlight-color: transparent;
        }

        .checkbox-box::after {
            content: "";
            width: 34%;
            height: 58%;
            border-right: 3px solid transparent;
            border-bottom: 3px solid transparent;
            transform: rotate(45deg) scale(0.7);
            margin-top: -6%;
            transition: all 0.14s ease;
        }

        .checkbox-box.checked {
            background: linear-gradient(180deg, #2563eb 0%, #1d4ed8 100%);
            border-color: #1d4ed8;
            box-shadow: 0 4px 10px rgba(37, 99, 235, 0.28);
        }

        .checkbox-box.checked::after {
            border-right-color: #ffffff;
            border-bottom-color: #ffffff;
            transform: rotate(45deg) scale(1);
        }

        .checkbox-box.required-empty {
            border-color: #ef4444;
            background: linear-gradient(180deg, #fff5f5 0%, #ffe8e8 100%);
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
            border: 2px solid #1f6feb;
            background: rgba(255,255,255,0.92);
            border-radius: 999px;
            cursor: pointer;
            user-select: none;
            font-weight: bold;
        }

        .choice-option.selected::after {
            content: "●";
        }

        .choice-option.readonly-choice {
            background: rgba(255,255,255,0.6);
            border: 1px dashed #999;
        }

        .signature-preview {
            width: 100%;
            height: 100%;
            border: 1px dashed rgba(15, 23, 42, 0.18);
            border-radius: 10px;
            background: rgba(255,255,255,0.55);
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            cursor: pointer;
            position: relative;
            box-shadow: 0 1px 4px rgba(15, 23, 42, 0.08);
            backdrop-filter: blur(2px);
        }

        .signature-preview.required-empty {
            border-color: #ef4444;
            background: rgba(255, 244, 244, 0.9);
        }

        .signature-preview img {
            max-width: 100%;
            max-height: 100%;
            display: block;
            object-fit: contain;
        }

        .signature-preview .signature-placeholder {
            color: #666;
            text-align: center;
            padding: 4px;
            line-height: 1.2;
        }


        .capture-mode .tap-box,
        .capture-mode .readonly-box,
        .capture-mode .signature-preview,
        .capture-mode .choice-group-wrap,
        .capture-mode .checkbox-box {
            background: transparent !important;
            border-color: transparent !important;
            box-shadow: none !important;
            backdrop-filter: none !important;
        }

        .capture-mode .tap-box .placeholder,
        .capture-mode .signature-placeholder {
            display: none !important;
        }

        .capture-mode .readonly-box img,
        .capture-mode .signature-preview img {
            background: transparent !important;
            box-shadow: none !important;
        }

        .capture-mode .checkbox-box {
            border: none !important;
        }

        .capture-mode .checkbox-box::after {
            border-right-color: #111827 !important;
            border-bottom-color: #111827 !important;
        }

        @media (max-width: 768px) {
            .topbar {
                position: sticky;
                top: 0;
                z-index: 100;
                background: rgba(255,255,255,0.94);
                backdrop-filter: blur(10px);
                padding-bottom: 10px;
                margin-bottom: 10px;
            }

            .tools {
                position: sticky;
                top: 72px;
                z-index: 99;
                background: rgba(255,255,255,0.94);
                backdrop-filter: blur(10px);
                padding: 10px 0;
            }

            .tap-box,
            .readonly-box,
            .signature-preview {
                min-height: 28px;
                border-radius: 12px;
            }

            .readonly-box {
                font-weight: 700;
            }

            .checkbox-box {
                border-width: 2.5px;
                border-radius: 10px;
                box-shadow: 0 3px 10px rgba(15, 23, 42, 0.1);
            }

            .checkbox-box::after {
                border-right-width: 3.5px;
                border-bottom-width: 3.5px;
            }
        }

        .modal-backdrop {
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.55);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 9999;
            padding: 16px;
        }

        .modal-backdrop.show {
            display: flex;
        }

        .modal-card {
            width: min(96vw, 900px);
            max-height: 92vh;
            overflow: auto;
            background: white;
            border-radius: 14px;
            padding: 14px;
            box-shadow: 0 12px 40px rgba(0,0,0,0.25);
        }

        .modal-title {
            font-size: 18px;
            font-weight: bold;
            margin: 0 0 10px 0;
        }

        .modal-actions {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            justify-content: flex-end;
            margin-top: 12px;
        }

        .modal-textarea {
            width: 100%;
            min-height: 180px;
            resize: vertical;
            border: 1px solid #cfd8e3;
            border-radius: 10px;
            padding: 12px;
            font-size: 16px;
        }

        .signature-modal-box {
            width: 100%;
            height: min(58vh, 420px);
            border: 1px solid #d0d7de;
            border-radius: 10px;
            overflow: hidden;
            background: #fff;
            touch-action: none;
        }

        #signatureModalCanvas {
            width: 100%;
            height: 100%;
            display: block;
            touch-action: none;
            background: #fff;
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

            .tools {
                align-items: flex-start;
            }

            .tools input[type="range"] {
                width: 180px;
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
            {% if doc_json and '"admin_view": true' in doc_json %}
                <a class="btn gray" href="/admin?work_date={{ work_date }}">관리자 목록</a>
            {% else %}
                <button class="btn" id="saveBtn" type="button">저장</button>
                <button class="btn green" id="downloadBtn" type="button">다운로드</button>
                <a class="btn gray" href="/forms">목록</a>
            {% endif %}
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
            <input type="range" id="zoomRange" min="50" max="200" value="100">
            <span id="zoomLabel">100%</span>
        </label>
        <div class="small">모바일에서는 칸을 직접 편집하지 않고, 칸을 누르면 큰 입력창이 열립니다.</div>
    </div>

    <div class="stage-wrap" id="stageWrap">
        <div id="captureArea" class="stage">
            <img id="formImage" src="{{ image_url }}" alt="{{ form.form_name }}">
            <div id="overlay" class="overlay"></div>
        </div>
    </div>
</div>

<div id="textModal" class="modal-backdrop">
    <div class="modal-card">
        <div class="modal-title" id="textModalTitle">텍스트 입력</div>
        <textarea id="textModalInput" class="modal-textarea" placeholder="내용을 입력하세요"></textarea>
        <div class="modal-actions">
            <button type="button" class="btn gray" id="textModalCancel">취소</button>
            <button type="button" class="btn" id="textModalConfirm">확인</button>
        </div>
    </div>
</div>

<div id="signatureModal" class="modal-backdrop">
    <div class="modal-card">
        <div class="modal-title" id="signatureModalTitle">서명 입력</div>
        <div class="signature-modal-box">
            <canvas id="signatureModalCanvas"></canvas>
        </div>
        <div class="modal-actions">
            <button type="button" class="btn red" id="signatureModalClear">지우기</button>
            <button type="button" class="btn gray" id="signatureModalCancel">취소</button>
            <button type="button" class="btn" id="signatureModalConfirm">확인</button>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<script>
const FORM_DEF = {{ form_json|safe }};
const USER = {{ user_json|safe }};
const DOC_META = {{ doc_json|safe }};
const ADMIN_VIEW = !!(DOC_META && DOC_META.admin_view);
const IS_MOBILE = window.innerWidth <= 768;
const RESOLVED_VALUES = {{ resolved_values_json|safe }};
const LABEL_VALUES = {{ label_values_json|safe }};
let CURRENT_STATUS = {{ status_json|safe }};

const img = document.getElementById("formImage");
const overlay = document.getElementById("overlay");
const captureArea = document.getElementById("captureArea");
const stageWrap = document.getElementById("stageWrap");
const zoomRange = document.getElementById("zoomRange");
const zoomLabel = document.getElementById("zoomLabel");
const statusText = document.getElementById("statusText");

const textModal = document.getElementById("textModal");
const textModalTitle = document.getElementById("textModalTitle");
const textModalInput = document.getElementById("textModalInput");
const signatureModal = document.getElementById("signatureModal");
const signatureModalTitle = document.getElementById("signatureModalTitle");
const signatureModalCanvas = document.getElementById("signatureModalCanvas");

let ACTIVE_TEXT_KEY = null;
let ACTIVE_SIGNATURE_KEY = null;
let SIGNATURE_DRAW_BOUND = false;

const EDIT_STATE = {};

function fieldKey(field) {
    if (field.slot_index !== null && field.slot_index !== undefined && field.slot_index !== "") {
        return `${field.field_id}__slot_${field.slot_index}`;
    }
    return field.field_id;
}

function isEditable(field) {
    if (ADMIN_VIEW) return false;
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

function px(v) {
    return `${v}px`;
}

function initializeState() {
    (FORM_DEF.fields || []).forEach((field) => {
        if (!field.visible || field.type === "label") return;
        const key = fieldKey(field);
        const initVal = RESOLVED_VALUES[key];
        if (field.type === "checkbox") {
            EDIT_STATE[key] = !!initVal;
        } else {
            EDIT_STATE[key] = initVal || "";
        }
    });
}

function setStageSize() {
    overlay.style.width = img.clientWidth + "px";
    overlay.style.height = img.clientHeight + "px";
    captureArea.style.height = img.clientHeight + "px";
}

function getCurrentScaleRatios() {
    const naturalW = img.naturalWidth || img.clientWidth || 1;
    const naturalH = img.naturalHeight || img.clientHeight || 1;
    return {
        naturalW,
        naturalH,
        ratioX: img.clientWidth / naturalW,
        ratioY: img.clientHeight / naturalH,
        uiScale: Math.min(img.clientWidth / naturalW, img.clientHeight / naturalH),
    };
}

function isEmptyValue(field, value) {
    if (field.type === "checkbox") return !value;
    return value === null || value === undefined || String(value).trim() === "";
}

function createReadonlyText(value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "readonly-box";
    wrap.textContent = value || "";
    wrap.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    return wrap;
}

function createEditableText(field, key, value, uiScale) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "tap-box";
    if (field.required && isEmptyValue(field, value)) {
        btn.classList.add("required-empty");
    }
    btn.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    btn.style.padding = `${Math.max(1, 3 * uiScale)}px ${Math.max(2, 6 * uiScale)}px`;
    btn.dataset.fieldKey = key;

    if (value && String(value).trim() !== "") {
        btn.textContent = value;
    } else {
        const span = document.createElement("span");
        span.className = "placeholder";
        span.textContent = field.placeholder || "눌러서 입력";
        btn.appendChild(span);
    }

    btn.addEventListener("click", (e) => {
        e.preventDefault();
        openTextModal(field, key, value || "");
    });

    return btn;
}

function createEditableCheckbox(field, key, value, uiScale) {
    const box = document.createElement("button");
    box.type = "button";
    box.className = "checkbox-box";
    if (value) box.classList.add("checked");
    if (field.required && !value) box.classList.add("required-empty");
    box.dataset.fieldKey = key;

    const minSize = window.innerWidth <= 768 ? 28 : 22;
    box.style.minWidth = `${minSize}px`;
    box.style.minHeight = `${minSize}px`;

    box.addEventListener("click", (e) => {
        e.preventDefault();
        EDIT_STATE[key] = !EDIT_STATE[key];
        renderFields();
    });

    return box;
}

function createReadonlyCheckbox(value, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "checkbox-box" + (value ? " checked" : "");
    wrap.style.cursor = "default";
    wrap.style.minWidth = `${Math.max(18, 22 * uiScale)}px`;
    wrap.style.minHeight = `${Math.max(18, 22 * uiScale)}px`;
    return wrap;
}

function createEditableChoiceGroup(field, key, value, ratioX, ratioY, uiScale) {
    const wrap = document.createElement("div");
    wrap.className = "choice-group";
    const options = field.options || [];

    options.forEach((opt) => {
        const div = document.createElement("button");
        div.type = "button";
        div.className = "choice-option";
        if (String(value || "") === String(opt.option_value)) {
            div.classList.add("selected");
        }
        div.style.left = px((opt.dx || 0) * ratioX);
        div.style.top = px((opt.dy || 0) * ratioY);
        div.style.width = px((opt.width || 20) * ratioX);
        div.style.height = px((opt.height || 20) * ratioY);
        div.style.fontSize = `${Math.max(8, 12 * uiScale)}px`;

        div.addEventListener("click", (e) => {
            e.preventDefault();
            EDIT_STATE[key] = opt.option_value;
            renderFields();
        });

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

function createEditableSignature(field, key, value, uiScale) {
    const box = document.createElement("button");
    box.type = "button";
    box.className = "signature-preview";
    if (field.required && isEmptyValue(field, value)) {
        box.classList.add("required-empty");
    }
    box.dataset.fieldKey = key;

    if (value) {
        const signImg = document.createElement("img");
        signImg.src = value;
        box.appendChild(signImg);
    } else {
        const place = document.createElement("div");
        place.className = "signature-placeholder";
        place.style.fontSize = `${Math.max(8, 12 * uiScale)}px`;
        place.textContent = "눌러서 서명";
        box.appendChild(place);
    }

    box.addEventListener("click", (e) => {
        e.preventDefault();
        openSignatureModal(field, key, value || "");
    });

    return box;
}

function createReadonlySignature(value) {
    const wrap = document.createElement("div");
    wrap.className = "readonly-box";
    if (value) {
        const image = document.createElement("img");
        image.src = value;
        image.style.maxWidth = "100%";
        image.style.maxHeight = "100%";
        image.style.objectFit = "contain";
        wrap.appendChild(image);
    }
    return wrap;
}

function createLabel(value, uiScale) {
    const div = document.createElement("div");
    div.className = "field-label";
    div.textContent = value || "";
    div.style.fontSize = `${Math.max(8, 14 * uiScale)}px`;
    return div;
}

function renderFields() {
    setStageSize();
    overlay.innerHTML = "";

    const { ratioX, ratioY, uiScale } = getCurrentScaleRatios();

    (FORM_DEF.fields || []).forEach((field) => {
        if (!field.visible) return;

        const key = fieldKey(field);
        const editable = isEditable(field);
        const value = EDIT_STATE[key];
        const labelValue = LABEL_VALUES[key];

        const item = document.createElement("div");
        item.className = "field-item" + (editable ? "" : " readonly");
        item.style.left = px((field.x || 0) * ratioX);
        item.style.top = px((field.y || 0) * ratioY);
        item.style.width = px((field.width || 24) * ratioX);
        item.style.height = px((field.height || 24) * ratioY);

        let node = null;

        if (field.type === "label") {
            node = createLabel(labelValue, uiScale);
        } else if (field.type === "text") {
            node = editable ? createEditableText(field, key, value, uiScale) : createReadonlyText(value, uiScale);
        } else if (field.type === "checkbox") {
            node = editable ? createEditableCheckbox(field, key, value, uiScale) : createReadonlyCheckbox(value, uiScale);
        } else if (field.type === "choice_group") {
            node = editable
                ? createEditableChoiceGroup(field, key, value, ratioX, ratioY, uiScale)
                : createReadonlyChoiceGroup(field, value, ratioX, ratioY, uiScale);
        } else if (field.type === "signature") {
            node = editable ? createEditableSignature(field, key, value, uiScale) : createReadonlySignature(value);
        } else {
            node = createReadonlyText(value, uiScale);
        }

        item.appendChild(node);
        overlay.appendChild(item);
    });
}

function getCurrentFieldValues() {
    return JSON.parse(JSON.stringify(EDIT_STATE));
}

function validateClientSide() {
    const missing = [];

    (FORM_DEF.fields || []).forEach((field) => {
        if (!field.visible || !field.required) return;
        if (!isEditable(field)) return;
        if (field.type === "label") return;

        const key = fieldKey(field);
        const value = EDIT_STATE[key];

        if (field.type === "text" && (!value || String(value).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "checkbox" && !value) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "choice_group" && (!value || String(value).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
        if (field.type === "signature" && (!value || String(value).trim() === "")) {
            missing.push(field.label || field.field_id);
        }
    });

    return missing;
}

async function saveForm() {
    const missing = validateClientSide();
    if (missing.length > 0) {
        alert("필수 입력이 누락되었습니다.\\n- " + missing.join("\\n- "));
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

async async function downloadImage() {
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

    try {
        document.body.classList.add("capture-mode");

        const canvas = await html2canvas(document.getElementById("captureArea"), {
            useCORS: true,
            backgroundColor: null,
            scale: 2
        });

        const link = document.createElement("a");
        link.href = canvas.toDataURL("image/png");
        link.download = getNextDownloadFilename();
        link.click();
    } catch (e) {
        alert("이미지 저장에 실패했습니다.");
    } finally {
        document.body.classList.remove("capture-mode");
    }
}

function setZoom(v) {
    if (!img.naturalWidth) return;
    zoomRange.value = String(v);
    zoomLabel.textContent = v + "%";
    captureArea.style.width = Math.max(100, Math.round(img.naturalWidth * (v / 100))) + "px";

    requestAnimationFrame(() => {
        setStageSize();
        renderFields();
    });
}

function applyInitialMobileZoom() {
    const wrapWidth = stageWrap.clientWidth - 24;
    if (!img.naturalWidth || wrapWidth <= 0) {
        setZoom(IS_MOBILE ? 130 : 100);
        return;
    }

    const fitPercent = Math.floor((wrapWidth / img.naturalWidth) * 100);
    const initial = IS_MOBILE ? Math.max(85, Math.min(140, fitPercent + 18)) : Math.max(60, Math.min(100, fitPercent));
    setZoom(initial);
}

function openTextModal(field, key, value) {
    ACTIVE_TEXT_KEY = key;
    textModalTitle.textContent = field.label || "텍스트 입력";
    textModalInput.value = value || "";
    textModalInput.placeholder = field.placeholder || "";
    if (field.max_length) {
        textModalInput.maxLength = field.max_length;
    } else {
        textModalInput.removeAttribute("maxLength");
    }
    textModal.classList.add("show");
    setTimeout(() => textModalInput.focus(), 50);
}

function closeTextModal() {
    ACTIVE_TEXT_KEY = null;
    textModal.classList.remove("show");
}

function fitCanvasResolution(canvas) {
    const rect = canvas.getBoundingClientRect();
    const ratio = Math.max(1, window.devicePixelRatio || 1);
    canvas.width = Math.max(1, Math.floor(rect.width * ratio));
    canvas.height = Math.max(1, Math.floor(rect.height * ratio));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    return ctx;
}

function clearSignatureCanvas() {
    const ctx = fitCanvasResolution(signatureModalCanvas);
    const rect = signatureModalCanvas.getBoundingClientRect();
    ctx.clearRect(0, 0, rect.width, rect.height);
}

function drawDataUrlOnCanvas(canvas, dataUrl) {
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

function bindSignatureDrawingOnce() {
    if (SIGNATURE_DRAW_BOUND) return;
    SIGNATURE_DRAW_BOUND = true;

    let drawing = false;
    let lastX = 0;
    let lastY = 0;

    function getPos(e) {
        const rect = signatureModalCanvas.getBoundingClientRect();
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
        const ctx = signatureModalCanvas.getContext("2d");
        const p = getPos(e);
        ctx.lineWidth = 2.2;
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
        if (e) e.preventDefault();
    }

    signatureModalCanvas.addEventListener("mousedown", start);
    signatureModalCanvas.addEventListener("mousemove", move);
    signatureModalCanvas.addEventListener("mouseup", end);
    signatureModalCanvas.addEventListener("mouseleave", end);
    signatureModalCanvas.addEventListener("touchstart", start, { passive: false });
    signatureModalCanvas.addEventListener("touchmove", move, { passive: false });
    signatureModalCanvas.addEventListener("touchend", end, { passive: false });
}

function openSignatureModal(field, key, value) {
    ACTIVE_SIGNATURE_KEY = key;
    signatureModalTitle.textContent = field.label || "서명 입력";
    signatureModal.classList.add("show");

    requestAnimationFrame(() => {
        fitCanvasResolution(signatureModalCanvas);
        bindSignatureDrawingOnce();
        if (value) {
            drawDataUrlOnCanvas(signatureModalCanvas, value);
        } else {
            clearSignatureCanvas();
        }
    });
}

function closeSignatureModal() {
    ACTIVE_SIGNATURE_KEY = null;
    signatureModal.classList.remove("show");
}

function isCanvasBlank(canvas) {
    const temp = document.createElement("canvas");
    temp.width = canvas.width;
    temp.height = canvas.height;
    return canvas.toDataURL() === temp.toDataURL();
}

const saveBtn = document.getElementById("saveBtn");
const downloadBtn = document.getElementById("downloadBtn");

if (!ADMIN_VIEW) {
    if (saveBtn) saveBtn.addEventListener("click", saveForm);
    if (downloadBtn) downloadBtn.addEventListener("click", downloadImage);
}

zoomRange.addEventListener("input", () => {
    setZoom(Number(zoomRange.value));
});

document.getElementById("textModalCancel").addEventListener("click", closeTextModal);
document.getElementById("textModalConfirm").addEventListener("click", () => {
    if (!ACTIVE_TEXT_KEY) return;
    EDIT_STATE[ACTIVE_TEXT_KEY] = textModalInput.value || "";
    closeTextModal();
    renderFields();
});

document.getElementById("signatureModalCancel").addEventListener("click", closeSignatureModal);
document.getElementById("signatureModalClear").addEventListener("click", clearSignatureCanvas);
document.getElementById("signatureModalConfirm").addEventListener("click", () => {
    if (!ACTIVE_SIGNATURE_KEY) return;
    EDIT_STATE[ACTIVE_SIGNATURE_KEY] = isCanvasBlank(signatureModalCanvas)
        ? ""
        : signatureModalCanvas.toDataURL("image/png");
    closeSignatureModal();
    renderFields();
});

textModal.addEventListener("click", (e) => {
    if (e.target === textModal) closeTextModal();
});

signatureModal.addEventListener("click", (e) => {
    if (e.target === signatureModal) closeSignatureModal();
});

window.addEventListener("resize", () => {
    if (!img.naturalWidth) return;
    const currentZoom = Number(zoomRange.value || "100");
    setZoom(currentZoom);
});

function initAfterImageLoad() {
    initializeState();
    applyInitialMobileZoom();
    refreshStatus();
}

if (img.complete) {
    initAfterImageLoad();
} else {
    img.addEventListener("load", initAfterImageLoad);
}
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


def get_all_active_non_admin_users():
    df = load_users()
    df = validate_users_dataframe(df)
    df["사용여부"] = df["사용여부"].str.upper()
    df["관리자여부"] = df["관리자여부"].str.upper()

    active_df = df[(df["사용여부"] == "Y") & (df["관리자여부"] != "Y")].copy()

    users = []
    for _, row in active_df.iterrows():
        users.append({
            "id": row["ID"],
            "name": row["이름"],
            "group": row["조"],
            "role": row["구분"],
            "slot_index": int(row["슬롯순서"]) if str(row["슬롯순서"]).strip() else None,
            "is_admin": False,
        })

    users.sort(key=lambda x: (
        x["group"],
        x["role"],
        x["slot_index"] if x["slot_index"] is not None else 9999,
        x["id"]
    ))
    return users


def get_group_map():
    users = get_all_active_non_admin_users()
    group_map = {}

    for user in users:
        group_name = str(user["group"]).strip()
        group_map.setdefault(group_name, []).append(user)

    for group_name in group_map:
        group_map[group_name].sort(
            key=lambda x: (x["slot_index"] if x["slot_index"] is not None else 9999, x["id"])
        )

    return dict(sorted(group_map.items(), key=lambda x: x[0]))


def save_users_dataframe(df):
    safe_df = df.copy()
    for col in safe_df.columns:
        safe_df[col] = safe_df[col].astype("string")

    with pd.ExcelWriter(USERS_XLSX_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        safe_df.to_excel(writer, sheet_name=USERS_SHEET_NAME, index=False)


def apply_group_assignments(assignments):
    df = load_users()
    df = validate_users_dataframe(df).copy()

    for col in ["ID", "이름", "조", "구분", "슬롯순서", "관리자여부", "사용여부"]:
        if col in df.columns:
            df[col] = df[col].astype("string")

    id_to_idx = {str(row["ID"]).strip(): idx for idx, row in df.iterrows()}

    for item in assignments:
        user_id = str(item.get("user_id", "")).strip()
        group_name = str(item.get("group", "")).strip()
        slot_index = str(item.get("slot_index", "")).strip()

        if user_id not in id_to_idx:
            continue

        idx = id_to_idx[user_id]
        df.at[idx, "조"] = group_name
        df.at[idx, "슬롯순서"] = slot_index

    save_users_dataframe(df)

def find_existing_document(form, target_user, work_date):
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
                    (form_id, target_user["group"], work_date)
                )
            else:
                cur.execute(
                    """
                    SELECT id, form_id, form_type, group_name, user_id, work_date, status
                    FROM documents
                    WHERE form_id = %s AND form_type = 'individual' AND user_id = %s AND work_date = %s
                    """,
                    (form_id, target_user["id"], work_date)
                )

            row = cur.fetchone()

    if not row:
        return None

    return {
        "id": row[0],
        "form_id": row[1],
        "form_type": row[2],
        "group_name": row[3],
        "user_id": row[4],
        "work_date": str(row[5]),
        "status": row[6],
    }


def build_admin_status_rows(forms, work_date):
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

            if doc:
                status_done = doc.get("status") == "completed"
                open_url = url_for(
                    "admin_open_form",
                    form_id=form["form_id"],
                    target_user_id=target_user["id"],
                    work_date=work_date
                )
                dot_class = "done" if status_done else "not-done"
            else:
                open_url = None
                dot_class = "not-done"

            form_items.append({
                "form_id": form["form_id"],
                "form_name": form["form_name"],
                "open_url": open_url,
                "dot_class": dot_class,
            })

        rows.append({
            "user": target_user,
            "form_items": form_items,
        })

    return rows

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

    selected_date = request.args.get("work_date", "").strip() or get_today_str()

    config = load_form_config()
    forms = [f for f in config.get("forms", []) if f.get("active", False)]
    status_rows = build_admin_status_rows(forms, selected_date)
    group_map = get_group_map()

    return render_template_string(
        ADMIN_HTML,
        user=user,
        forms=forms,
        selected_date=selected_date,
        status_rows=status_rows,
        group_map=group_map,
    )


@app.route("/admin/open/<form_id>/<target_user_id>/<work_date>")
def admin_open_form(form_id, target_user_id, work_date):
    viewer = session.get("user")
    if not viewer:
        return redirect(url_for("home"))
    if not viewer.get("is_admin"):
        return "관리자만 접근 가능합니다.", 403

    form = get_form_by_id(form_id)
    if not form:
        return "양식을 찾을 수 없습니다.", 404

    target_user = next((u for u in get_all_active_non_admin_users() if u["id"] == target_user_id), None)
    if not target_user:
        return "대상 사용자를 찾을 수 없습니다.", 404

    allowed_roles = form.get("allowed_roles", [])
    if allowed_roles and target_user["role"] not in allowed_roles:
        return "이 사용자는 해당 양식 대상이 아닙니다.", 403

    document = find_existing_document(form, target_user, work_date)
    if not document:
        return "선택한 날짜에 작성된 문서가 없습니다.", 404

    if form["form_type"] == "group":
        sync_group_participants(document["id"], target_user["group"])

    participants = fetch_participants(document["id"])
    values = fetch_document_values(document["id"])
    status = recalc_document_status(document, form, target_user)

    resolved_values = build_resolved_values_map(form, target_user, values)
    label_values = build_label_values_map(form, target_user, document, participants)
    image_url = url_for("form_image", filename=form.get("image_file", ""))

    return render_template_string(
        FORM_RUN_HTML,
        user=target_user,
        form=form,
        image_url=image_url,
        document_id=document["id"],
        work_date=document["work_date"],
        doc_status=status["document_status_text"],
        form_json=json.dumps(form, ensure_ascii=False),
        user_json=json.dumps(target_user, ensure_ascii=False),
        doc_json=json.dumps({
            "document_id": document["id"],
            "work_date": document["work_date"],
            "group_name": document.get("group_name"),
            "user_id": document.get("user_id"),
            "admin_view": True,
        }, ensure_ascii=False),
        resolved_values_json=json.dumps(resolved_values, ensure_ascii=False),
        label_values_json=json.dumps(label_values, ensure_ascii=False),
        status_json=json.dumps(status, ensure_ascii=False),
    )


@app.route("/admin/groups/save", methods=["POST"])
def admin_groups_save():
    user = session.get("user")
    if not user:
        return jsonify({"ok": False, "message": "로그인이 필요합니다."}), 401
    if not user.get("is_admin"):
        return jsonify({"ok": False, "message": "관리자만 접근 가능합니다."}), 403

    data = request.get_json(silent=True) or {}
    assignments = data.get("assignments", [])

    if not isinstance(assignments, list) or not assignments:
        return jsonify({"ok": False, "message": "저장할 조편성 데이터가 없습니다."}), 400

    try:
        apply_group_assignments(assignments)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "message": f"저장 실패: {e}"}), 500

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
