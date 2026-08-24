"""Server-rendered templates for the local desktop workspace."""

from __future__ import annotations

from jinja2 import DictLoader

TEMPLATES = {
    "base.html": r"""<!doctype html>
<html lang="{{ language }}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ page_title }} · WorkbookLens</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f3f5f7;
      --surface: #ffffff;
      --surface-soft: #f8fafb;
      --ink: #17202a;
      --muted: #5e6975;
      --line: #d8dee5;
      --line-strong: #b9c2cc;
      --accent: #1769aa;
      --accent-hover: #10558c;
      --accent-soft: #eaf4fb;
      --success: #18794e;
      --success-soft: #e9f7ef;
      --warning: #9a6700;
      --warning-soft: #fff7d6;
      --danger: #b42318;
      --danger-soft: #fff0ee;
      --focus: #8cbde1;
      --shadow: 0 8px 24px rgba(28, 39, 49, .08);
    }
    * { box-sizing: border-box; }
    html { min-width: 320px; background: var(--bg); }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font: 14px/1.5 "Segoe UI", "Microsoft YaHei UI", Arial, sans-serif;
      letter-spacing: 0;
    }
    button, input, select { font: inherit; letter-spacing: 0; }
    button, .button, select { min-height: 38px; }
    a { color: var(--accent); }
    a:hover { color: var(--accent-hover); }
    :focus-visible { outline: 3px solid var(--focus); outline-offset: 2px; }
    .shell { min-height: 100vh; display: grid; grid-template-rows: auto 1fr auto; }
    .topbar {
      position: sticky;
      top: 0;
      z-index: 10;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      min-height: 64px;
      padding: 10px clamp(16px, 3vw, 36px);
      background: rgba(255, 255, 255, .97);
      border-bottom: 1px solid var(--line);
    }
    .brand { display: flex; align-items: center; gap: 11px; min-width: 0; color: var(--ink); text-decoration: none; }
    .brand-mark {
      display: grid;
      place-items: center;
      flex: 0 0 38px;
      width: 38px;
      height: 38px;
      border-radius: 7px;
      background: var(--accent);
      color: #fff;
      font-size: 19px;
      font-weight: 700;
    }
    .brand-copy { min-width: 0; }
    .brand-name { display: block; font-size: 16px; line-height: 1.2; }
    .brand-subtitle { display: block; margin-top: 2px; color: var(--muted); font-size: 12px; white-space: nowrap; }
    .topbar-actions { display: flex; align-items: center; gap: 12px; }
    .local-status { display: inline-flex; align-items: center; gap: 7px; color: var(--success); font-size: 13px; white-space: nowrap; }
    .local-status::before { content: ""; width: 8px; height: 8px; border-radius: 50%; background: var(--success); }
    .language-form { display: flex; align-items: center; gap: 8px; }
    .language-form label { color: var(--muted); font-size: 13px; }
    select {
      max-width: 150px;
      padding: 7px 30px 7px 10px;
      color: var(--ink);
      background: var(--surface);
      border: 1px solid var(--line-strong);
      border-radius: 6px;
    }
    .content { width: min(1160px, 100%); margin: 0 auto; padding: 26px clamp(16px, 3vw, 36px) 40px; }
    .page-heading { display: flex; align-items: flex-start; justify-content: space-between; gap: 20px; margin-bottom: 20px; }
    .page-heading-copy { min-width: 0; }
    h1, h2, h3, p { overflow-wrap: anywhere; }
    h1 { margin: 0; font-size: clamp(24px, 3vw, 32px); line-height: 1.25; font-weight: 700; }
    h2 { margin: 0; font-size: 18px; line-height: 1.35; }
    h3 { margin: 0; font-size: 15px; line-height: 1.4; }
    .lede { max-width: 720px; margin: 7px 0 0; color: var(--muted); font-size: 15px; }
    .back-link { display: inline-flex; align-items: center; gap: 6px; margin-bottom: 13px; font-weight: 600; text-decoration: none; }
    .notice {
      display: grid;
      grid-template-columns: auto 1fr;
      gap: 11px;
      padding: 12px 14px;
      margin: 0 0 20px;
      background: var(--accent-soft);
      border: 1px solid #c7dfef;
      border-radius: 6px;
    }
    .notice.warning { background: var(--warning-soft); border-color: #ead58a; }
    .notice.error { background: var(--danger-soft); border-color: #efbbb5; }
    .notice strong { display: block; margin-bottom: 2px; }
    .notice p { margin: 0; color: var(--muted); }
    .notice-icon { font-size: 17px; line-height: 1.3; font-weight: 700; }
    .tool-grid { display: grid; grid-template-columns: minmax(0, 1.25fr) minmax(300px, .75fr); gap: 18px; align-items: start; }
    .panel { background: var(--surface); border: 1px solid var(--line); border-radius: 7px; box-shadow: var(--shadow); }
    .panel-header { padding: 18px 20px 14px; border-bottom: 1px solid var(--line); }
    .panel-header p { margin: 5px 0 0; color: var(--muted); }
    .panel-body { padding: 20px; }
    .upload-field { display: grid; gap: 9px; }
    .field-label { font-weight: 600; }
    .upload-box {
      display: grid;
      gap: 8px;
      padding: 18px;
      background: var(--surface-soft);
      border: 1px dashed var(--line-strong);
      border-radius: 6px;
    }
    .file-picker-row { display: flex; align-items: center; gap: 10px; min-width: 0; flex-wrap: wrap; }
    .file-picker-control {
      position: relative;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 38px;
      padding: 7px 12px;
      color: var(--ink);
      background: var(--surface);
      border: 1px solid var(--line-strong);
      border-radius: 5px;
      font-weight: 600;
      cursor: pointer;
    }
    .file-picker-control:hover { background: var(--accent-soft); border-color: var(--focus); }
    .file-picker-control:focus-within { outline: 3px solid var(--focus); outline-offset: 2px; }
    .file-picker-control input[type="file"] {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      opacity: 0;
      cursor: pointer;
    }
    .file-name { min-width: 0; color: var(--muted); overflow-wrap: anywhere; }
    .field-hint, .status-text { margin: 0; color: var(--muted); font-size: 13px; }
    .form-actions { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-top: 16px; }
    .selection-actions { display: flex; align-items: center; gap: 9px; flex-wrap: wrap; margin: 0 0 12px; }
    button, .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 7px;
      padding: 8px 14px;
      color: #fff;
      background: var(--accent);
      border: 1px solid var(--accent);
      border-radius: 6px;
      font-weight: 650;
      text-decoration: none;
      cursor: pointer;
    }
    button:hover, .button:hover { color: #fff; background: var(--accent-hover); border-color: var(--accent-hover); }
    button.secondary, .button.secondary { color: var(--ink); background: var(--surface); border-color: var(--line-strong); }
    button.secondary:hover, .button.secondary:hover { color: var(--ink); background: var(--surface-soft); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    .trust-copy { margin: 15px 0 0; padding-top: 14px; color: var(--muted); border-top: 1px solid var(--line); font-size: 13px; }
    .trust-copy strong { color: var(--ink); }
    .summary-bar { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); margin: 0 0 18px; background: var(--surface); border: 1px solid var(--line); border-radius: 7px; }
    .metric { min-width: 0; padding: 14px 16px; border-right: 1px solid var(--line); }
    .metric:last-child { border-right: 0; }
    .metric-value { display: block; font-size: 22px; line-height: 1.1; font-weight: 700; }
    .metric-label { display: block; margin-top: 5px; color: var(--muted); font-size: 12px; }
    .toolbar { display: flex; gap: 9px; flex-wrap: wrap; margin-bottom: 18px; }
    .section { margin-top: 24px; }
    .section-title { display: flex; align-items: baseline; justify-content: space-between; gap: 16px; margin-bottom: 10px; }
    .section-title p { margin: 0; color: var(--muted); font-size: 13px; }
    .patch-list, .finding-list { display: grid; gap: 9px; }
    .patch-row, .finding-row { display: block; background: var(--surface); border: 1px solid var(--line); border-radius: 6px; }
    .patch-row { position: relative; padding: 14px 16px 14px 46px; cursor: pointer; }
    .patch-row:hover { border-color: var(--focus); }
    .patch-row > input { position: absolute; top: 17px; left: 17px; width: 17px; height: 17px; accent-color: var(--accent); }
    .patch-heading { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .patch-location { font-weight: 700; }
    .badge { display: inline-flex; align-items: center; min-height: 23px; padding: 2px 7px; border-radius: 5px; background: var(--surface-soft); border: 1px solid var(--line); color: var(--muted); font-size: 12px; font-weight: 650; }
    .badge.safe, .badge.info { color: var(--success); background: var(--success-soft); border-color: #b8dfc9; }
    .badge.layout_review, .badge.warning { color: var(--warning); background: var(--warning-soft); border-color: #ead58a; }
    .badge.error, .badge.critical { color: var(--danger); background: var(--danger-soft); border-color: #efbbb5; }
    .patch-description { margin: 7px 0 0; color: var(--muted); }
    .change { display: grid; grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr); gap: 9px; align-items: center; margin-top: 10px; }
    code { padding: 4px 6px; background: var(--surface-soft); border: 1px solid var(--line); border-radius: 4px; font: 12px/1.45 Consolas, monospace; white-space: pre-wrap; overflow-wrap: anywhere; }
    .arrow { color: var(--muted); font-weight: 700; }
    .consent { display: flex; gap: 10px; align-items: flex-start; margin: 12px 0; padding: 13px 14px; background: var(--warning-soft); border: 1px solid #ead58a; border-radius: 6px; }
    .consent input { flex: 0 0 auto; width: 17px; height: 17px; margin-top: 2px; accent-color: var(--accent); }
    .consent strong { display: block; }
    .consent span { display: block; margin-top: 2px; color: var(--muted); }
    .finding-row { padding: 15px 17px; border-left-width: 5px; border-left-color: var(--accent); }
    .finding-row.warning { border-left-color: var(--warning); }
    .finding-row.error, .finding-row.critical { border-left-color: var(--danger); }
    .finding-meta { display: flex; align-items: center; gap: 7px; flex-wrap: wrap; margin-bottom: 7px; }
    .finding-row p { margin: 6px 0 0; color: var(--muted); }
    details { margin-top: 10px; }
    summary { color: var(--accent); font-weight: 650; cursor: pointer; }
    .evidence-grid { display: grid; gap: 7px; margin-top: 8px; }
    .empty-state { padding: 24px; text-align: center; color: var(--muted); background: var(--surface); border: 1px solid var(--line); border-radius: 6px; }
    .success-layout, .error-layout { width: min(760px, 100%); margin: 4vh auto 0; }
    .result-mark { display: grid; place-items: center; width: 46px; height: 46px; border-radius: 50%; color: #fff; background: var(--success); font-size: 24px; font-weight: 700; }
    .result-mark.error { background: var(--danger); }
    .result-heading { display: flex; gap: 15px; align-items: flex-start; margin-bottom: 18px; }
    .result-heading p { margin: 5px 0 0; color: var(--muted); }
    .result-actions { display: flex; gap: 9px; flex-wrap: wrap; margin-top: 18px; }
    .diagnostic { margin-top: 16px; padding-top: 13px; border-top: 1px solid var(--line); }
    .diagnostic dl { display: grid; grid-template-columns: max-content minmax(0, 1fr); gap: 7px 12px; margin: 10px 0 0; }
    .diagnostic dt { color: var(--muted); }
    .diagnostic dd { min-width: 0; margin: 0; overflow-wrap: anywhere; }
    .footer { padding: 12px 20px 18px; color: var(--muted); text-align: center; font-size: 12px; }
    .sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0; }
    @media (max-width: 760px) {
      .topbar { align-items: flex-start; }
      .brand-subtitle, .local-status, .language-form label { display: none; }
      .tool-grid { grid-template-columns: 1fr; }
      .summary-bar { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .metric:nth-child(2) { border-right: 0; }
      .metric:nth-child(-n + 2) { border-bottom: 1px solid var(--line); }
      .page-heading { display: block; }
      .change { grid-template-columns: 1fr; }
      .arrow { transform: rotate(90deg); width: max-content; }
    }
    @media (max-width: 420px) {
      .brand-mark { display: none; }
      .topbar { gap: 8px; }
      .topbar-actions { gap: 8px; }
      select { max-width: 118px; }
      .content { padding-top: 20px; }
      .panel-body, .panel-header { padding-left: 15px; padding-right: 15px; }
    }
    @media (prefers-reduced-motion: reduce) { *, *::before, *::after { scroll-behavior: auto !important; } }
  </style>
</head>
<body data-view="{{ view }}">
<div class="shell">
  <header class="topbar">
    <a class="brand" href="/?lang={{ language|urlencode }}">
      <span class="brand-mark" aria-hidden="true">W</span>
      <span class="brand-copy">
        <strong class="brand-name">WorkbookLens</strong>
        <span class="brand-subtitle">{{ t("web.brand_subtitle") }}</span>
      </span>
    </a>
    <div class="topbar-actions">
      <span class="local-status">{{ t("web.local_status") }}</span>
      <form class="language-form" method="get" action="{{ language_action }}">
        <label for="language">{{ t("web.language_label") }}</label>
        <select id="language" name="lang" onchange="this.form.submit()">
          <option value="zh-CN"{% if language == "zh-CN" %} selected{% endif %}>简体中文</option>
          <option value="en"{% if language == "en" %} selected{% endif %}>English</option>
        </select>
        <noscript><button class="secondary" type="submit">{{ t("web.language_apply") }}</button></noscript>
      </form>
    </div>
  </header>
  <main class="content" id="main-content">{% block content %}{% endblock %}</main>
  <footer class="footer">WorkbookLens {{ version }} · {{ t("web.footer_local") }}</footer>
</div>
{% block script %}{% endblock %}
</body>
</html>""",
    "index.html": r"""{% extends "base.html" %}{% block content %}
<div class="page-heading">
  <div class="page-heading-copy">
    <h1>{{ t("web.home_title") }}</h1>
    <p class="lede">{{ t("web.home_intro") }}</p>
  </div>
</div>
<aside class="notice" aria-label="{{ t('web.privacy_title') }}">
  <span class="notice-icon" aria-hidden="true">i</span>
  <div><strong>{{ t("web.privacy_title") }}</strong><p>{{ t("web.privacy_body") }}</p></div>
</aside>
<div class="tool-grid">
  <form id="scan-form" class="panel" action="/scan" method="post" enctype="multipart/form-data">
    <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
    <input type="hidden" name="language" value="{{ language }}">
    <div class="panel-header"><h2>{{ t("web.scan_title") }}</h2><p>{{ t("web.scan_intro") }}</p></div>
    <div class="panel-body">
      <div class="upload-field">
        <span id="workbook-label" class="field-label">{{ t("web.scan_label") }}</span>
        <div class="upload-box">
          <div class="file-picker-row">
            <label class="file-picker-control"><span id="workbook-choose">{{ t("web.file_choose") }}</span><input id="workbook" name="workbook" type="file" accept=".xlsx,.xlsm" aria-labelledby="workbook-label workbook-choose" aria-describedby="workbook-file-name workbook-hint" required></label>
            <span id="workbook-file-name" class="file-name" aria-live="polite">{{ t("web.file_none") }}</span>
          </div>
          <p id="workbook-hint" class="field-hint">{{ t("web.scan_hint") }}</p>
        </div>
      </div>
      <div class="form-actions">
        <button id="scan-button" type="submit">{{ t("web.scan_button") }}</button>
        <p id="scan-status" class="status-text" aria-live="polite">{{ t("web.scan_limit", max_mb=max_mb) }}</p>
      </div>
    </div>
  </form>
  <form id="convert-form" class="panel" action="/convert" method="post" enctype="multipart/form-data">
    <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
    <input type="hidden" name="language" value="{{ language }}">
    <div class="panel-header"><h2>{{ t("web.convert_title") }}</h2><p>{{ t("web.convert_intro") }}</p></div>
    <div class="panel-body">
      <div class="upload-field">
        <span id="legacy-workbook-label" class="field-label">{{ t("web.convert_label") }}</span>
        <div class="upload-box">
          <div class="file-picker-row">
            <label class="file-picker-control"><span id="legacy-workbook-choose">{{ t("web.file_choose") }}</span><input id="legacy-workbook" name="legacy_workbook" type="file" accept=".xls" aria-labelledby="legacy-workbook-label legacy-workbook-choose" aria-describedby="legacy-workbook-file-name legacy-workbook-hint" required></label>
            <span id="legacy-workbook-file-name" class="file-name" aria-live="polite">{{ t("web.file_none") }}</span>
          </div>
          <p id="legacy-workbook-hint" class="field-hint">{{ t("web.convert_hint") }}</p>
        </div>
      </div>
      <p class="trust-copy"><strong>{{ t("web.convert_trust_title") }}</strong> {{ t("web.convert_trust_body") }}</p>
      <div class="form-actions">
      {% if converter_names %}
        <button id="convert-button" type="submit">{{ t("web.convert_button") }}</button>
        <p id="convert-status" class="status-text" aria-live="polite">{{ t("web.convert_provider", providers=converter_names|join(provider_separator)) }}</p>
      {% else %}
        <button id="convert-button" type="submit" disabled>{{ t("web.convert_unavailable") }}</button>
        <p id="convert-status" class="status-text" aria-live="polite">{{ t("web.convert_install_provider") }}</p>
      {% endif %}
      </div>
    </div>
  </form>
</div>
{% endblock %}{% block script %}
<script>
  const bindBusyState = (formId, buttonId, statusId, busyLabel, busyStatus) => {
    const form = document.getElementById(formId);
    if (!form) return;
    form.addEventListener('submit', () => {
      const button = document.getElementById(buttonId);
      const status = document.getElementById(statusId);
      if (button) { button.disabled = true; button.textContent = busyLabel; }
      if (status) status.textContent = busyStatus;
    });
  };
  const bindFileSelection = (inputId, statusId, emptyLabel) => {
    const input = document.getElementById(inputId);
    const status = document.getElementById(statusId);
    if (!input || !status) return;
    const update = () => {
      status.textContent = input.files && input.files.length ? input.files[0].name : emptyLabel;
    };
    input.addEventListener('change', update);
    update();
  };
  bindFileSelection('workbook', 'workbook-file-name', {{ t("web.file_none")|tojson }});
  bindFileSelection('legacy-workbook', 'legacy-workbook-file-name', {{ t("web.file_none")|tojson }});
  bindBusyState('scan-form', 'scan-button', 'scan-status', {{ t("web.scan_running")|tojson }}, {{ t("web.scan_running_detail")|tojson }});
  bindBusyState('convert-form', 'convert-button', 'convert-status', {{ t("web.convert_running")|tojson }}, {{ t("web.convert_running_detail")|tojson }});
</script>
{% endblock %}""",
    "results.html": r"""{% extends "base.html" %}{% block content %}
<a class="back-link" href="/?lang={{ language|urlencode }}">&larr; {{ t("web.results_new_scan") }}</a>
<div class="page-heading">
  <div class="page-heading-copy"><h1>{{ t("web.results_title") }}</h1><p class="lede"><strong>{{ filename }}</strong> · {{ t("web.results_intro") }}</p></div>
</div>
<div class="summary-bar" aria-label="{{ t('web.results_summary') }}">
  <div class="metric"><span class="metric-value">{{ findings|length }}</span><span class="metric-label">{{ t("web.metric_findings") }}</span></div>
  <div class="metric"><span class="metric-value">{{ patches|length }}</span><span class="metric-label">{{ t("web.metric_patches") }}</span></div>
  <div class="metric"><span class="metric-value">{{ severity_counts.error + severity_counts.critical }}</span><span class="metric-label">{{ t("web.metric_errors") }}</span></div>
  <div class="metric"><span class="metric-value">{{ severity_counts.warning }}</span><span class="metric-label">{{ t("web.metric_warnings") }}</span></div>
</div>
<nav class="toolbar" aria-label="{{ t('web.results_actions') }}">
  <a class="button secondary" href="/sessions/{{ session_id }}/report">{{ t("web.results_report") }}</a>
  <a class="button secondary" href="/sessions/{{ session_id }}/plan">{{ t("web.results_plan") }}</a>
</nav>
<section class="section" aria-labelledby="patches-heading">
  <div class="section-title"><div><h2 id="patches-heading">{{ t("web.results_review_title") }}</h2><p>{{ t("web.results_review_help") }}</p></div></div>
  <form id="apply-form" action="/sessions/{{ session_id }}/apply" method="post">
    <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
    <input type="hidden" name="language" value="{{ language }}">
    {% if patches %}<div class="selection-actions">
      <button id="select-all-patches" class="secondary" type="button" aria-controls="patch-list">{{ t("web.results_select_all") }}</button>
      <button id="clear-all-patches" class="secondary" type="button" aria-controls="patch-list" disabled>{{ t("web.results_clear_all") }}</button>
      <p id="patch-selection-status" class="status-text" aria-live="polite">{{ t("web.results_selection_count", selected=0, total=patches|length) }}</p>
    </div><div id="patch-list" class="patch-list">
    {% for patch in patches %}
      <label class="patch-row">
        <input type="checkbox" name="patch_id" value="{{ patch.id }}" data-risk="{{ patch.risk.value }}">
        <span class="patch-heading"><span class="patch-location">{{ patch.sheet }}!{{ patch.cell }}</span><span class="badge">{{ t("patch_kind." ~ patch.kind.value) }}</span><span class="badge {{ patch.risk.value }}">{{ t("risk." ~ patch.risk.value) }}</span><span class="badge">{{ t("web.confidence", percent='%.0f'|format(patch.confidence.root * 100)) }}</span></span>
        <span class="patch-description">{{ patch.description }}</span>
        <span class="change"><code>{{ display_value(patch.before) }}</code><span class="arrow" aria-hidden="true">&rarr;</span><code>{{ display_value(patch.after) }}</code></span>
      </label>
    {% endfor %}</div>
    {% if has_layout_review %}<label class="consent"><input type="checkbox" name="accept_layout_risk" value="true"><span><strong>{{ t("web.results_layout_consent") }}</strong><span>{{ t("web.layout_consent_body") }}</span></span></label>{% endif %}
    <div class="form-actions"><button id="apply-button" type="submit">{{ t("web.results_apply") }}</button><p id="apply-status" class="status-text" aria-live="polite"></p></div>
    {% else %}<div class="empty-state">{{ t("web.results_no_patches") }}</div>{% endif %}
  </form>
</section>
<section class="section" aria-labelledby="findings-heading">
  <div class="section-title"><h2 id="findings-heading">{{ t("web.results_findings") }} ({{ findings|length }})</h2></div>
  {% if findings %}<div class="finding-list">
  {% for finding in findings %}<article class="finding-row {{ finding.severity.value }}">
    <div class="finding-meta"><span class="badge {{ finding.severity.value }}">{{ t("severity." ~ finding.severity.value) }}</span><span class="badge">{{ finding.rule_id }}</span><span>{{ finding.sheet or t("web.results_workbook_scope") }}{% if finding.location %}!{{ finding.location }}{% endif %}</span></div>
    <h3>{{ finding.title }}</h3><p>{{ finding.explanation }}</p>
    <details><summary>{{ t("web.results_evidence") }}</summary><div class="evidence-grid"><span>{{ finding.evidence.summary }}</span><code>{{ display_value(finding.evidence.observed) }}</code></div></details>
  </article>{% endfor %}
  </div>{% else %}<div class="empty-state">{{ t("web.no_findings") }}</div>{% endif %}
</section>
{% endblock %}{% block script %}
<script>
  const applyForm = document.getElementById('apply-form');
  const patchCheckboxes = applyForm ? Array.from(
    applyForm.querySelectorAll('#patch-list input[name="patch_id"]:not(:disabled)')
  ) : [];
  const selectAllButton = document.getElementById('select-all-patches');
  const clearAllButton = document.getElementById('clear-all-patches');
  const patchSelectionStatus = document.getElementById('patch-selection-status');
  const selectionCountTemplate = {{ t("web.results_selection_count", selected="__SELECTED__", total="__TOTAL__")|tojson }};

  const updatePatchSelectionControls = () => {
    const selectedCount = patchCheckboxes.filter((checkbox) => checkbox.checked).length;
    if (selectAllButton) selectAllButton.disabled = selectedCount === patchCheckboxes.length;
    if (clearAllButton) clearAllButton.disabled = selectedCount === 0;
    if (patchSelectionStatus) {
      patchSelectionStatus.textContent = selectionCountTemplate
        .replace('__SELECTED__', String(selectedCount))
        .replace('__TOTAL__', String(patchCheckboxes.length));
    }
  };

  if (selectAllButton) selectAllButton.addEventListener('click', () => {
    patchCheckboxes.forEach((checkbox) => { checkbox.checked = true; });
    updatePatchSelectionControls();
  });
  if (clearAllButton) clearAllButton.addEventListener('click', () => {
    patchCheckboxes.forEach((checkbox) => { checkbox.checked = false; });
    updatePatchSelectionControls();
  });
  patchCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener('change', updatePatchSelectionControls);
  });
  updatePatchSelectionControls();

  if (applyForm) applyForm.addEventListener('submit', () => {
    const button = document.getElementById('apply-button');
    const status = document.getElementById('apply-status');
    if (selectAllButton) selectAllButton.disabled = true;
    if (clearAllButton) clearAllButton.disabled = true;
    if (button) { button.disabled = true; button.textContent = {{ t("web.apply_running")|tojson }}; }
    if (status) status.textContent = {{ t("web.apply_running_detail")|tojson }};
  });
</script>
{% endblock %}""",
    "applied.html": r"""{% extends "base.html" %}{% block content %}
<div class="success-layout">
  <a class="back-link" href="/?lang={{ language|urlencode }}">&larr; {{ t("web.results_new_scan") }}</a>
  <div class="panel panel-body">
    <div class="result-heading"><span class="result-mark" aria-hidden="true">&#10003;</span><div><h1>{{ t("web.applied_title") }}</h1><p>{{ t("web.applied_intro") }}</p></div></div>
    <div class="summary-bar">
      <div class="metric"><span class="metric-value">{{ result.applied_patch_ids|length }}</span><span class="metric-label">{{ t("web.metric_applied") }}</span></div>
      <div class="metric"><span class="metric-value">{{ result.resolved_finding_ids|length }}</span><span class="metric-label">{{ t("web.metric_resolved") }}</span></div>
      <div class="metric"><span class="metric-value">{{ result.new_finding_ids|length }}</span><span class="metric-label">{{ t("web.metric_new_findings") }}</span></div>
      <div class="metric"><span class="metric-value">{{ result.output_sha256[:8] }}</span><span class="metric-label">{{ t("web.metric_output_hash") }}</span></div>
    </div>
    <p>{{ t("web.applied_validation") }}</p>
    <div class="result-actions"><a class="button" href="/sessions/{{ session_id }}/fixed">{{ t("web.applied_download") }}</a><a class="button secondary" href="/sessions/{{ session_id }}/diff">{{ t("web.applied_diff") }}</a><a class="button secondary" href="/sessions/{{ session_id }}/apply-report">{{ t("web.applied_apply_report") }}</a></div>
  </div>
</div>
{% endblock %}""",
    "error.html": r"""{% extends "base.html" %}{% block content %}
<div class="error-layout">
  <div class="panel panel-body">
    <div class="result-heading"><span class="result-mark error" aria-hidden="true">!</span><div><h1>{{ error.title }}</h1><p>{{ error.message }}</p></div></div>
    <aside class="notice warning"><span class="notice-icon" aria-hidden="true">i</span><div><strong>{{ t("web.error_suggestion") }}</strong><p>{{ error.suggestion }}</p></div></aside>
    <p>{{ t("web.error_source_safe") }}</p>
    <details class="diagnostic"><summary>{{ t("web.error_details") }}</summary><dl><dt>{{ t("web.error_code") }}</dt><dd><code>{{ error.code }}</code></dd><dt>{{ t("web.error_diagnostic_id") }}</dt><dd><code>{{ error.diagnostic_id }}</code></dd></dl></details>
    <div class="result-actions"><a class="button" href="/?lang={{ language|urlencode }}">{{ t("web.error_retry") }}</a></div>
  </div>
</div>
{% endblock %}""",
}


def template_loader() -> DictLoader:
    """Return an in-package loader that also works in frozen builds."""

    return DictLoader(TEMPLATES)
