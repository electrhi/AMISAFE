/* AMISAFE 양식 디자이너 (웹 포팅)
 * 원본 form_designer_pro_ui_excel.py 의 주요 기능을 HTML5 Canvas + JS로 재구현.
 *
 * 주요 기능:
 * - 양식 목록 관리 (CRUD)
 * - 캔버스 위 필드 배치 (드래그/드롭, 클릭 추가)
 * - 7가지 필드 타입 (checkbox, choice_group, dropdown, date, signature, text, label)
 * - 필드 속성 편집 (좌표/크기/역할/슬롯/필수 여부/옵션)
 * - 엑셀 대/소분류 자동 적용
 * - JSON 직접 편집
 * - 줌 조정
 */

(function () {
  'use strict';

  // ===== State =====
  const state = {
    config: { forms: [] },
    currentFormIdx: -1,
    selectedFieldIdx: -1,
    zoom: 1.0,
    pendingFieldType: null,  // type 버튼을 누른 후 캔버스 클릭 대기 상태
    dragging: null,  // { fieldIdx, startX, startY, origX, origY }
    imageNaturalSize: { w: 0, h: 0 },
  };

  // ===== Defaults =====
  const FIELD_TYPE_DEFAULTS = {
    checkbox:     { width: 24,  height: 24, label: '체크' },
    choice_group: { width: 100, height: 24, label: '확인/해당없음' },
    dropdown:     { width: 170, height: 34, label: '목록선택' },
    date:         { width: 150, height: 34, label: '날짜' },
    datetime:     { width: 210, height: 34, label: '날짜시간' },
    signature:    { width: 170, height: 45, label: '서명' },
    text:         { width: 180, height: 32, label: '입력' },
    label:        { width: 150, height: 28, label: '안내문구' },
  };

  const DEFAULT_CHOICE_OPTIONS = [
    { option_label: '확인',     option_value: 'confirm', dx: 0,  dy: 0, width: 24, height: 24 },
    { option_label: '해당없음', option_value: 'na',      dx: 58, dy: 0, width: 24, height: 24 },
  ];

  const DEFAULT_DROPDOWN_OPTIONS = [
    { option_label: '정상',     option_value: 'normal' },
    { option_label: '이상',     option_value: 'abnormal' },
    { option_label: '해당없음', option_value: 'na' },
  ];

  // ===== Utilities =====
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  function setStatus(msg, isError = false) {
    const el = $('#statusbar');
    el.textContent = msg;
    el.style.color = isError ? '#dc2626' : '#6b7280';
  }

  function showHint(msg) {
    const el = $('#canvas-hint');
    el.textContent = msg;
    el.classList.add('show');
  }
  function hideHint() { $('#canvas-hint').classList.remove('show'); }

  function currentForm() {
    return state.currentFormIdx >= 0 ? state.config.forms[state.currentFormIdx] : null;
  }
  function currentField() {
    const f = currentForm();
    if (!f || state.selectedFieldIdx < 0) return null;
    return f.fields[state.selectedFieldIdx];
  }

  // Auto-generate unique field_id within a form
  function genFieldId(form, baseType) {
    let n = 1;
    const ids = new Set(form.fields.map((f) => f.field_id));
    while (ids.has(`${baseType}_${n}`)) n++;
    return `${baseType}_${n}`;
  }
  function genFormId() {
    let n = 1;
    const ids = new Set(state.config.forms.map((f) => f.form_id));
    while (ids.has(`form_${n}`)) n++;
    return `form_${n}`;
  }

  // ===== Server I/O =====
  async function loadConfig() {
    try {
      const resp = await fetch('/designer/api/config');
      const data = await resp.json();
      if (!data.ok) throw new Error(data.message || '로드 실패');
      state.config = data.config || { forms: [] };
      if (!Array.isArray(state.config.forms)) state.config.forms = [];
      refreshFormList();
      if (state.config.forms.length > 0) selectForm(0);
      else { state.currentFormIdx = -1; renderCanvas(); renderMetaForm(); }
      setStatus(`로드 완료: ${state.config.forms.length}개 양식`);
    } catch (e) {
      setStatus('설정 로드 실패: ' + e.message, true);
    }
  }

  async function saveAll() {
    try {
      const resp = await fetch('/designer/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(state.config),
      });
      const data = await resp.json();
      if (!data.ok) throw new Error(data.message || '저장 실패');
      setStatus('전체 저장 완료 ✔');
    } catch (e) {
      setStatus('저장 실패: ' + e.message, true);
      alert('저장 실패: ' + e.message);
    }
  }

  async function uploadImage(file) {
    const fd = new FormData();
    fd.append('image_file', file);
    const resp = await fetch('/designer/api/upload-image', { method: 'POST', body: fd });
    const data = await resp.json();
    if (!data.ok) throw new Error(data.message || '업로드 실패');
    return data.filename;
  }

  // ===== Render: form list =====
  function refreshFormList() {
    const ul = $('#form-list');
    ul.innerHTML = '';
    state.config.forms.forEach((form, i) => {
      const li = document.createElement('li');
      li.textContent = `${form.form_name || form.form_id} ${form.active ? '✓' : '○'}`;
      li.className = (i === state.currentFormIdx) ? 'active' : '';
      li.title = `[${form.form_type}] ${form.form_id}`;
      li.addEventListener('click', () => selectForm(i));
      ul.appendChild(li);
    });
    if (state.config.forms.length === 0) {
      const li = document.createElement('li');
      li.textContent = '(양식이 없습니다)';
      li.style.color = '#9ca3af';
      li.style.cursor = 'default';
      ul.appendChild(li);
    }
  }

  function selectForm(idx) {
    state.currentFormIdx = idx;
    state.selectedFieldIdx = -1;
    refreshFormList();
    renderMetaForm();
    renderCanvas();
    renderFieldProps();
    renderRawJson();
    setStatus(`양식 선택: ${currentForm()?.form_name || ''}`);
  }

  // ===== Form meta editor =====
  function renderMetaForm() {
    const f = currentForm();
    if (!f) {
      ['meta-form-id', 'meta-form-name', 'meta-image-file', 'meta-allowed-roles'].forEach(id => { $('#' + id).value = ''; });
      $('#meta-form-type').value = 'individual';
      $('#meta-active').checked = false;
      return;
    }
    $('#meta-form-id').value = f.form_id || '';
    $('#meta-form-name').value = f.form_name || '';
    $('#meta-form-type').value = f.form_type || 'individual';
    $('#meta-image-file').value = f.image_file || '';
    $('#meta-active').checked = !!f.active;
    $('#meta-allowed-roles').value = (f.allowed_roles || []).join(',');
  }

  function applyMetaForm() {
    const f = currentForm();
    if (!f) return alert('선택된 양식이 없습니다.');

    const newId = $('#meta-form-id').value.trim();
    if (!newId) return alert('form_id를 입력하세요.');

    // ID 중복 검사 (자기 자신은 OK)
    const duplicates = state.config.forms.filter((x, i) => x.form_id === newId && i !== state.currentFormIdx);
    if (duplicates.length > 0) return alert(`form_id가 중복됩니다: ${newId}`);

    f.form_id = newId;
    f.form_name = $('#meta-form-name').value.trim() || newId;
    f.form_type = $('#meta-form-type').value;
    f.image_file = $('#meta-image-file').value.trim();
    f.active = $('#meta-active').checked;
    const rolesText = $('#meta-allowed-roles').value.trim();
    f.allowed_roles = rolesText ? rolesText.split(',').map((s) => s.trim()).filter(Boolean) : [];

    refreshFormList();
    renderCanvas();
    setStatus(`양식 정보 적용: ${f.form_name}`);
  }

  // ===== Canvas rendering =====
  function renderCanvas() {
    const f = currentForm();
    const empty = $('#canvas-empty');
    const stage = $('#canvas-stage');
    const img = $('#canvas-img');

    if (!f) {
      empty.style.display = 'block';
      stage.style.display = 'none';
      return;
    }
    empty.style.display = 'none';
    stage.style.display = 'inline-block';

    // 기존 필드 오버레이 제거
    stage.querySelectorAll('.field-overlay').forEach((n) => n.remove());

    if (f.image_file) {
      img.src = `/form-image/${encodeURIComponent(f.image_file)}`;
      img.onload = () => {
        state.imageNaturalSize = { w: img.naturalWidth, h: img.naturalHeight };
        applyZoom();
        drawFields();
      };
      img.onerror = () => { setStatus('이미지를 불러올 수 없습니다: ' + f.image_file, true); drawFields(); };
    } else {
      img.removeAttribute('src');
      state.imageNaturalSize = { w: 800, h: 600 };
      stage.style.width = '800px';
      stage.style.height = '600px';
      drawFields();
    }
  }

  function applyZoom() {
    const img = $('#canvas-img');
    const stage = $('#canvas-stage');
    const w = state.imageNaturalSize.w * state.zoom;
    const h = state.imageNaturalSize.h * state.zoom;
    img.style.width = w + 'px';
    img.style.height = h + 'px';
    stage.style.width = w + 'px';
    stage.style.height = h + 'px';
    $('#zoom-label').textContent = Math.round(state.zoom * 100) + '%';
  }

  function drawFields() {
    const f = currentForm();
    if (!f) return;
    const stage = $('#canvas-stage');
    stage.querySelectorAll('.field-overlay').forEach((n) => n.remove());

    (f.fields || []).forEach((field, i) => {
      const div = document.createElement('div');
      div.className = `field-overlay t-${field.type || 'text'}`;
      if (i === state.selectedFieldIdx) div.classList.add('selected');
      div.dataset.fieldIdx = String(i);

      const x = (field.x || 0) * state.zoom;
      const y = (field.y || 0) * state.zoom;
      const w = (field.width || 24) * state.zoom;
      const h = (field.height || 24) * state.zoom;
      div.style.left = x + 'px';
      div.style.top = y + 'px';
      div.style.width = w + 'px';
      div.style.height = h + 'px';
      const labelText = field.label || field.field_id || '';
      div.textContent = labelText.length > 14 ? labelText.slice(0, 12) + '…' : labelText;
      div.title = `${field.field_id} (${field.type})`;

      // choice_group 옵션 미리보기 추가
      if (field.type === 'choice_group' && field.options) {
        field.options.forEach((opt) => {
          const o = document.createElement('div');
          o.style.cssText = 'position:absolute;background:rgba(255,255,255,0.85);border:1px dashed #1f6feb;font-size:9px;text-align:center;line-height:1.1;color:#1f6feb;';
          o.style.left = ((opt.dx || 0) * state.zoom) + 'px';
          o.style.top = ((opt.dy || 0) * state.zoom) + 'px';
          o.style.width = ((opt.width || 24) * state.zoom) + 'px';
          o.style.height = ((opt.height || 24) * state.zoom) + 'px';
          o.textContent = opt.option_label || '';
          o.style.pointerEvents = 'none';
          div.appendChild(o);
        });
      }

      div.addEventListener('mousedown', (e) => onFieldMouseDown(e, i));
      stage.appendChild(div);
    });
  }

  // ===== Field manipulation =====
  function onCanvasClick(e) {
    if (!state.pendingFieldType) return;
    const f = currentForm();
    if (!f) return;
    const rect = $('#canvas-stage').getBoundingClientRect();
    const x = Math.round((e.clientX - rect.left) / state.zoom);
    const y = Math.round((e.clientY - rect.top) / state.zoom);

    const type = state.pendingFieldType;
    const defs = FIELD_TYPE_DEFAULTS[type] || FIELD_TYPE_DEFAULTS.text;
    const newField = {
      field_id: genFieldId(f, type),
      label: defs.label,
      type: type,
      x: x,
      y: y,
      width: defs.width,
      height: defs.height,
      required: type !== 'label',
      target_role: '공통',
      slot_index: null,
      visible: true,
    };

    if (type === 'choice_group') {
      newField.options = JSON.parse(JSON.stringify(DEFAULT_CHOICE_OPTIONS));
    } else if (type === 'dropdown') {
      newField.options = JSON.parse(JSON.stringify(DEFAULT_DROPDOWN_OPTIONS));
      newField.placeholder = '선택하세요';
    } else if (type === 'date') {
      newField.placeholder = 'YYYY-MM-DD';
    } else if (type === 'datetime') {
      newField.placeholder = 'YYYY-MM-DD HH:MM';
    } else if (type === 'text') {
      newField.placeholder = '입력하세요';
    } else if (type === 'label') {
      newField.bind_key = '';
    }

    f.fields = f.fields || [];
    f.fields.push(newField);
    state.selectedFieldIdx = f.fields.length - 1;
    state.pendingFieldType = null;
    hideHint();
    drawFields();
    renderFieldProps();
    setStatus(`필드 추가됨: ${newField.field_id}`);
  }

  function onFieldMouseDown(e, idx) {
    e.preventDefault();
    e.stopPropagation();
    state.selectedFieldIdx = idx;
    drawFields();
    renderFieldProps();
    // 활성 탭이 폼 메타라면 필드 속성 탭으로
    activateTab('field-prop-tab');

    const f = currentForm();
    const field = f.fields[idx];
    state.dragging = {
      fieldIdx: idx,
      startX: e.clientX,
      startY: e.clientY,
      origX: field.x,
      origY: field.y,
    };
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd);
  }

  function onDragMove(e) {
    if (!state.dragging) return;
    const f = currentForm();
    if (!f) return;
    const field = f.fields[state.dragging.fieldIdx];
    const dx = (e.clientX - state.dragging.startX) / state.zoom;
    const dy = (e.clientY - state.dragging.startY) / state.zoom;
    field.x = Math.max(0, Math.round(state.dragging.origX + dx));
    field.y = Math.max(0, Math.round(state.dragging.origY + dy));
    drawFields();
    // 속성 패널의 x/y만 빠르게 갱신
    $('#fp-x').value = field.x;
    $('#fp-y').value = field.y;
  }

  function onDragEnd() {
    if (state.dragging) {
      setStatus(`이동: ${currentField()?.field_id}`);
    }
    state.dragging = null;
    document.removeEventListener('mousemove', onDragMove);
    document.removeEventListener('mouseup', onDragEnd);
  }

  // ===== Field properties panel =====
  function renderFieldProps() {
    const field = currentField();
    const wrap = $('#field-props');
    const noFieldMsg = $('#no-field-msg');
    if (!field) {
      wrap.style.display = 'none';
      noFieldMsg.style.display = 'block';
      return;
    }
    wrap.style.display = 'block';
    noFieldMsg.style.display = 'none';

    $('#fp-field-id').value = field.field_id || '';
    $('#fp-label').value = field.label || '';
    $('#fp-type').value = field.type || 'text';
    $('#fp-x').value = field.x || 0;
    $('#fp-y').value = field.y || 0;
    $('#fp-width').value = field.width || 0;
    $('#fp-height').value = field.height || 0;
    $('#fp-target-role').value = field.target_role || '공통';
    $('#fp-slot-index').value = (field.slot_index === null || field.slot_index === undefined) ? '' : field.slot_index;
    $('#fp-required').checked = !!field.required;
    $('#fp-visible').checked = field.visible !== false;

    // label 전용
    if (field.type === 'label') {
      $('#fp-bind-row').style.display = 'block';
      $('#fp-bind-key').value = field.bind_key || '';
    } else {
      $('#fp-bind-row').style.display = 'none';
    }

    // choice_group / dropdown 옵션
    renderOptionsEditor(field);
  }

  function renderOptionsEditor(field) {
    const block = $('#fp-options-block');
    const title = $('#fp-options-title');
    const tbody = $('#fp-options-table tbody');
    tbody.innerHTML = '';

    if (field.type !== 'choice_group' && field.type !== 'dropdown') {
      block.style.display = 'none';
      return;
    }
    block.style.display = 'block';
    title.textContent = field.type === 'choice_group' ? '선택 옵션 (확인/해당없음 등)' : '드롭다운 옵션';

    (field.options || []).forEach((opt, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input type="text" data-opt-key="option_label" data-idx="${idx}" value="${escapeHtml(opt.option_label || '')}"></td>
        <td><input type="text" data-opt-key="option_value" data-idx="${idx}" value="${escapeHtml(opt.option_value || '')}"></td>
        <td><button class="btn red small" data-del-opt="${idx}">×</button></td>
      `;
      tbody.appendChild(tr);
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function applyFieldProps() {
    const field = currentField();
    if (!field) return alert('선택된 필드가 없습니다.');

    const newId = $('#fp-field-id').value.trim();
    if (!newId) return alert('field_id는 비울 수 없습니다.');

    // ID 중복 검사 (자기 자신은 OK)
    const f = currentForm();
    const dupIdx = f.fields.findIndex((x, i) => x.field_id === newId && i !== state.selectedFieldIdx);
    if (dupIdx >= 0) return alert(`field_id가 중복됩니다: ${newId}`);

    const oldType = field.type;
    field.field_id = newId;
    field.label = $('#fp-label').value;
    field.type = $('#fp-type').value;
    field.x = parseInt($('#fp-x').value) || 0;
    field.y = parseInt($('#fp-y').value) || 0;
    field.width = parseInt($('#fp-width').value) || 24;
    field.height = parseInt($('#fp-height').value) || 24;
    field.target_role = $('#fp-target-role').value;
    const slotText = $('#fp-slot-index').value.trim();
    field.slot_index = slotText === '' ? null : (parseInt(slotText) || slotText);
    field.required = $('#fp-required').checked;
    field.visible = $('#fp-visible').checked;
    if (field.type === 'label') field.bind_key = $('#fp-bind-key').value;

    // 옵션 수집
    if (field.type === 'choice_group' || field.type === 'dropdown') {
      const inputs = $$('#fp-options-table tbody input');
      const optMap = {};
      inputs.forEach((inp) => {
        const idx = inp.dataset.idx;
        const key = inp.dataset.optKey;
        if (!optMap[idx]) optMap[idx] = {};
        optMap[idx][key] = inp.value;
      });
      field.options = Object.keys(optMap)
        .sort((a, b) => parseInt(a) - parseInt(b))
        .map((idx) => {
          const orig = (field.options || [])[parseInt(idx)] || {};
          return Object.assign({}, orig, optMap[idx]);
        });
    }

    // 타입이 바뀌었을 때 기본 옵션 채워주기
    if (oldType !== field.type) {
      if (field.type === 'choice_group' && (!field.options || field.options.length === 0)) {
        field.options = JSON.parse(JSON.stringify(DEFAULT_CHOICE_OPTIONS));
      }
      if (field.type === 'dropdown' && (!field.options || field.options.length === 0)) {
        field.options = JSON.parse(JSON.stringify(DEFAULT_DROPDOWN_OPTIONS));
      }
    }

    drawFields();
    renderFieldProps();
    renderRawJson();
    setStatus(`필드 적용: ${field.field_id}`);
  }

  function deleteSelectedField() {
    const f = currentForm();
    if (!f || state.selectedFieldIdx < 0) return;
    if (!confirm('선택한 필드를 삭제할까요?')) return;
    f.fields.splice(state.selectedFieldIdx, 1);
    state.selectedFieldIdx = -1;
    drawFields();
    renderFieldProps();
    renderRawJson();
    setStatus('필드 삭제됨');
  }

  // ===== Form CRUD =====
  function newForm() {
    const fid = genFormId();
    const newF = {
      form_id: fid,
      form_name: fid,
      form_type: 'individual',
      image_file: '',
      active: true,
      fields: [],
      allowed_roles: [],
    };
    state.config.forms.push(newF);
    selectForm(state.config.forms.length - 1);
    setStatus(`새 양식 생성: ${fid}`);
  }

  function deleteCurrentForm() {
    const f = currentForm();
    if (!f) return alert('선택된 양식이 없습니다.');
    if (!confirm(`'${f.form_name || f.form_id}' 양식을 삭제할까요?`)) return;
    state.config.forms.splice(state.currentFormIdx, 1);
    state.currentFormIdx = -1;
    state.selectedFieldIdx = -1;
    refreshFormList();
    renderCanvas();
    renderMetaForm();
    renderFieldProps();
    if (state.config.forms.length > 0) selectForm(0);
    setStatus('양식 삭제됨');
  }

  // ===== Tabs =====
  function activateTab(tabId) {
    $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === tabId));
    $$('.tab-content').forEach((c) => c.classList.toggle('active', c.id === tabId));
  }

  // ===== Excel import (대/소분류) =====
  async function importCategoryExcel() {
    const file = $('#excel-import-file').files[0];
    if (!file) return alert('엑셀 파일을 선택하세요.');

    const field = currentField();
    if (!field || field.type !== 'dropdown') {
      return alert('drop-down 필드를 먼저 선택해 주세요. (다른 타입은 적용 불가)');
    }

    const fd = new FormData();
    fd.append('excel_file', file);

    try {
      const resp = await fetch('/designer/api/import-category-excel', { method: 'POST', body: fd });
      const data = await resp.json();
      if (!data.ok) throw new Error(data.message || '실패');

      // dropdown 옵션을 대분류로 일단 채움
      field.options = data.parent_options;
      // 소분류 매핑은 field.option_map 에 저장 (원본 디자이너의 의도와 동일)
      field.option_map = data.option_map;
      drawFields();
      renderFieldProps();
      renderRawJson();
      setStatus(`엑셀 적용: 대분류 ${data.stats.parent_count}개 / 소분류 ${data.stats.child_total}개`);
      alert(`적용 완료\n대분류 ${data.stats.parent_count}개, 소분류 총 ${data.stats.child_total}개`);
    } catch (e) {
      alert('엑셀 적용 실패: ' + e.message);
    }
  }

  // ===== Raw JSON tab =====
  function renderRawJson() {
    const f = currentForm();
    $('#raw-json-editor').value = f ? JSON.stringify(f, null, 2) : '';
  }

  function applyRawJson() {
    const f = currentForm();
    if (!f) return alert('선택된 양식이 없습니다.');
    try {
      const parsed = JSON.parse($('#raw-json-editor').value);
      if (!parsed.form_id) throw new Error('form_id가 필요합니다.');
      Object.keys(f).forEach((k) => delete f[k]);
      Object.assign(f, parsed);
      refreshFormList();
      renderCanvas();
      renderMetaForm();
      renderFieldProps();
      setStatus('JSON 적용 완료');
    } catch (e) {
      alert('JSON 파싱 오류: ' + e.message);
    }
  }

  // ===== Wire up =====
  function init() {
    // Top buttons
    $('#btn-new-form').addEventListener('click', newForm);
    $('#btn-save-all').addEventListener('click', saveAll);
    $('#btn-delete-form').addEventListener('click', deleteCurrentForm);
    $('#btn-apply-meta').addEventListener('click', applyMetaForm);

    // Image upload (meta tab)
    $('#meta-image-upload').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const filename = await uploadImage(file);
        // refresh dropdown
        const sel = $('#meta-image-file');
        if (!Array.from(sel.options).some((o) => o.value === filename)) {
          const opt = document.createElement('option');
          opt.value = filename;
          opt.textContent = filename;
          sel.appendChild(opt);
        }
        sel.value = filename;
        setStatus(`이미지 업로드: ${filename}`);
      } catch (err) {
        alert('업로드 실패: ' + err.message);
      }
    });

    // Field add buttons
    $$('[data-add-field]').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (!currentForm()) return alert('먼저 양식을 선택하거나 새로 만드세요.');
        state.pendingFieldType = btn.dataset.addField;
        showHint(`'${btn.textContent}' 필드 - 캔버스 위 원하는 위치를 클릭하세요`);
      });
    });

    // Canvas click
    $('#canvas-stage').addEventListener('click', onCanvasClick);

    // Zoom controls
    $('#btn-zoom-in').addEventListener('click', () => { state.zoom = Math.min(3, state.zoom + 0.1); applyZoom(); drawFields(); });
    $('#btn-zoom-out').addEventListener('click', () => { state.zoom = Math.max(0.2, state.zoom - 0.1); applyZoom(); drawFields(); });
    $('#btn-zoom-fit').addEventListener('click', () => {
      const host = $('#canvas-host');
      const availW = host.clientWidth - 40;
      if (state.imageNaturalSize.w > 0) {
        state.zoom = Math.min(1.5, availW / state.imageNaturalSize.w);
        applyZoom();
        drawFields();
      }
    });

    // Field property buttons
    $('#btn-apply-field').addEventListener('click', applyFieldProps);
    $('#btn-delete-field').addEventListener('click', deleteSelectedField);
    $('#fp-add-option').addEventListener('click', () => {
      const field = currentField();
      if (!field) return;
      if (!field.options) field.options = [];
      const newOpt = field.type === 'choice_group'
        ? { option_label: '새옵션', option_value: 'new', dx: 0, dy: 0, width: 24, height: 24 }
        : { option_label: '새옵션', option_value: 'new' };
      field.options.push(newOpt);
      renderOptionsEditor(field);
    });

    // Delegated: option delete buttons
    $('#fp-options-table').addEventListener('click', (e) => {
      const target = e.target.closest('[data-del-opt]');
      if (!target) return;
      const idx = parseInt(target.dataset.delOpt);
      const field = currentField();
      if (!field || !field.options) return;
      field.options.splice(idx, 1);
      renderOptionsEditor(field);
    });

    // Field type change updates options editor
    $('#fp-type').addEventListener('change', () => {
      const field = currentField();
      if (!field) return;
      const newType = $('#fp-type').value;
      $('#fp-bind-row').style.display = newType === 'label' ? 'block' : 'none';
      if ((newType === 'choice_group' || newType === 'dropdown') && (!field.options || field.options.length === 0)) {
        // temporarily render placeholder options
        const tempField = Object.assign({}, field, {
          type: newType,
          options: newType === 'choice_group'
            ? JSON.parse(JSON.stringify(DEFAULT_CHOICE_OPTIONS))
            : JSON.parse(JSON.stringify(DEFAULT_DROPDOWN_OPTIONS)),
        });
        renderOptionsEditor(tempField);
      } else {
        renderOptionsEditor(Object.assign({}, field, { type: newType }));
      }
    });

    // Tabs
    $$('.tab').forEach((t) => t.addEventListener('click', () => activateTab(t.dataset.tab)));

    // Excel
    $('#btn-import-excel').addEventListener('click', importCategoryExcel);
    $('#btn-download-template').addEventListener('click', () => {
      window.location.href = '/designer/api/excel-template';
    });

    // Raw JSON
    $('#btn-apply-raw-json').addEventListener('click', applyRawJson);

    // Initial load
    loadConfig();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
