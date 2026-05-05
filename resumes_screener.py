#!/usr/bin/env python3
"""Resumes Screener web service."""

from __future__ import annotations

import csv
import io
import json
import os
import re
import tempfile
import uuid
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from flask import Flask, Response, abort, jsonify, render_template, request, send_file, url_for
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename


APP_NAME = "Resumes Screener"
ORG_NAME = os.environ.get("RESUMES_SCREENER_ORG_NAME", "Samsung Semiconductor India Research")
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = Path(os.environ.get("RESUMES_SCREENER_UPLOAD_DIR", BASE_DIR / "uploads"))
EXPORT_ROOT = Path(os.environ.get("RESUMES_SCREENER_EXPORT_DIR", BASE_DIR / "exports"))
SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx", ".ppt", ".pptx", ".txt"}
OLLAMA_TIMEOUT = int(os.environ.get("RESUMES_SCREENER_OLLAMA_TIMEOUT", "600"))
DEFAULT_OLLAMA_HOST = os.environ.get("RESUMES_SCREENER_OLLAMA_HOST", "localhost")
DEFAULT_OLLAMA_PORT = os.environ.get("RESUMES_SCREENER_OLLAMA_PORT", "11434")
DEFAULT_MODEL = os.environ.get("RESUMES_SCREENER_MODEL", "llama3.1")


ANALYSIS_PROMPT = """\
You are a seasoned HR analyst with more than 20 years of recruitment experience
across the industry. You have screened thousands of resumes, interviewed candidates
at every level, and have a sharp eye for role fit, career trajectory, and red flags.
Apply that depth of judgement when evaluating this resume against the job description.

Respond with ONLY a valid JSON object. No markdown fences, no explanation, just raw JSON.

JOB DESCRIPTION:
{jd}

RESUME:
{resume}

Return this exact JSON (fill in real values):
{{
  "name": "<candidate full name>",
  "education": [
    {{"degree": "<e.g., B.Tech>", "college": "<e.g., NIT Warangal>", "year": <year of passing as int, e.g., 2018>}}
  ],
  "current_organization": "<current employer, or empty string if not currently employed>",
  "past_organizations": ["<previous employer 1>", "<previous employer 2>"],
  "total_experience": {{"years": <int>, "months": <int 0-11>}},
  "relevant_experience": {{"years": <int>, "months": <int 0-11>}},
  "highlights": ["<strength 1>", "<strength 2>", "<strength 3>"],
  "lowlights": ["<gap 1>", "<gap 2>"],
  "score": <integer 1-10>,
  "decision": "<Select if score > 6, else Reject>"
}}

Rules:
- education: include ONLY degree, college name, and year of passing.
  Do NOT include location, GPA, percentage, specialisation, coursework, or any other detail.
  List every distinct qualification (e.g., both B.Tech and M.Tech).
  Do NOT repeat the same qualification more than once.
- current_organization: the single employer the candidate works at right now.
  Use an empty string if the candidate is not currently employed.
- past_organizations: every prior employer in reverse chronological order (most recent first).
  Deduplicate strictly. Each company name must appear only once even if the candidate
  worked there in multiple stints. Do NOT include the current organization in this list.
- score is 1-10 based on how well the resume matches the job description
- decision MUST be "Select" when score > 6, otherwise "Reject"
- Extract the candidate name directly from the resume text
- Be specific in highlights and lowlights (3 highlights, 2 lowlights minimum)
"""


INDEX_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ app_name }}</title>
  <style>
    :root {
      --ink: #17211f;
      --muted: #5f6f6a;
      --subtle: #8b9994;
      --bg: #f5f7f4;
      --panel: #ffffff;
      --line: #dbe3de;
      --line-strong: #c6d3cd;
      --brand: #0f766e;
      --brand-deep: #115e59;
      --brand-soft: #d9f3ef;
      --warn-soft: #fff1d6;
      --danger: #b42318;
      --danger-soft: #fee4e2;
      --ok: #047857;
      --ok-soft: #dff7ea;
      --shadow: 0 12px 30px rgba(30, 41, 36, .08);
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--ink);
      background: var(--bg);
      font-family: "Segoe UI", Arial, sans-serif;
      line-height: 1.45;
    }

    header {
      height: 76px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 28px;
      color: white;
      background: #17211f;
      border-bottom: 4px solid var(--brand);
    }

    h1 {
      margin: 0;
      font-size: 24px;
      line-height: 1.1;
      letter-spacing: 0;
    }

    .status-pill {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      min-height: 32px;
      padding: 6px 12px;
      border: 1px solid rgba(255,255,255,.18);
      border-radius: 999px;
      color: #d5ded9;
      font-size: 13px;
      white-space: nowrap;
    }

    .status-dot {
      width: 9px;
      height: 9px;
      border-radius: 50%;
      background: var(--subtle);
    }

    main {
      width: min(1360px, 100%);
      margin: 0 auto;
      padding: 24px;
    }

    form {
      display: grid;
      gap: 16px;
    }

    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    .panel-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 16px;
      padding: 16px 18px 10px;
      border-bottom: 1px solid var(--line);
    }

    .panel-title {
      margin: 0;
      font-size: 16px;
      line-height: 1.2;
    }

    .panel-note {
      margin: 4px 0 0;
      color: var(--muted);
      font-size: 13px;
    }

    .panel-body {
      padding: 16px 18px 18px;
    }

    .grid {
      display: grid;
      grid-template-columns: minmax(180px, 2fr) minmax(90px, .7fr) minmax(180px, 2fr) auto auto;
      gap: 12px;
      align-items: end;
    }

    label {
      display: grid;
      gap: 6px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0;
      text-transform: uppercase;
    }

    input[type="text"],
    input[type="number"],
    input[type="file"],
    textarea,
    select {
      width: 100%;
      min-height: 42px;
      border: 1px solid var(--line-strong);
      border-radius: 6px;
      padding: 9px 10px;
      color: var(--ink);
      background: white;
      font: inherit;
      outline: none;
    }

    textarea {
      min-height: 160px;
      resize: vertical;
    }

    input:focus,
    textarea:focus,
    select:focus {
      border-color: var(--brand);
      box-shadow: 0 0 0 3px rgba(15, 118, 110, .14);
    }

    .button-row {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 10px;
    }

    button,
    .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 40px;
      border: 1px solid transparent;
      border-radius: 6px;
      padding: 9px 14px;
      color: white;
      background: var(--brand);
      font: inherit;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
      white-space: nowrap;
    }

    button:hover,
    .button:hover {
      background: var(--brand-deep);
    }

    .button.secondary,
    button.secondary {
      color: var(--ink);
      background: #edf2ef;
      border-color: var(--line);
    }

    .button.secondary:hover,
    button.secondary:hover {
      background: #e1e9e5;
    }

    .button.ghost,
    button.ghost {
      color: var(--brand-deep);
      background: transparent;
      border-color: var(--line-strong);
    }

    .button.ghost:hover,
    button.ghost:hover {
      background: var(--brand-soft);
    }

    .action-panel {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 16px 18px;
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      color: var(--muted);
      font-size: 14px;
    }

    .metric {
      display: inline-flex;
      align-items: center;
      min-height: 30px;
      padding: 5px 10px;
      border-radius: 999px;
      background: #edf2ef;
      color: var(--ink);
      font-weight: 700;
    }

    .message {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px 14px;
      background: white;
      color: var(--muted);
    }

    .message.error {
      color: var(--danger);
      border-color: #f5b4ae;
      background: var(--danger-soft);
    }

    .message.ok {
      color: var(--ok);
      border-color: #b7e8cb;
      background: var(--ok-soft);
    }

    .table-wrap {
      overflow: auto;
      border-top: 1px solid var(--line);
    }

    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 1120px;
    }

    th,
    td {
      padding: 12px 10px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      text-align: left;
      font-size: 13px;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      color: white;
      background: #17211f;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0;
    }

    tr.select td { background: var(--ok-soft); }
    tr.reject td { background: var(--danger-soft); }
    .mono { font-family: Consolas, "Liberation Mono", monospace; }
    .muted { color: var(--muted); }
    .score {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 46px;
      height: 28px;
      border-radius: 999px;
      color: var(--ink);
      background: var(--warn-soft);
      font-weight: 800;
    }

    details {
      margin: 0;
    }

    summary {
      cursor: pointer;
      color: var(--brand-deep);
      font-weight: 700;
    }

    .detail-body {
      display: grid;
      gap: 8px;
      margin-top: 10px;
      color: var(--ink);
    }

    .kv {
      display: grid;
      grid-template-columns: 130px minmax(0, 1fr);
      gap: 10px;
    }

    .kv b {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
    }

    ul {
      margin: 0;
      padding-left: 18px;
    }

    footer {
      width: min(1360px, 100%);
      margin: 0 auto;
      padding: 0 24px 24px;
      color: var(--muted);
      font-size: 12px;
    }

    @media (max-width: 900px) {
      header {
        height: auto;
        min-height: 76px;
        align-items: flex-start;
        flex-direction: column;
        justify-content: center;
        gap: 10px;
        padding: 16px 20px;
      }

      main {
        padding: 16px;
      }

      .grid {
        grid-template-columns: 1fr;
      }

      .action-panel {
        align-items: stretch;
      }

      .button-row,
      .action-panel > .button-row {
        width: 100%;
      }

      button,
      .button {
        width: 100%;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>{{ app_name }}</h1>
    <div class="status-pill" id="connectionStatus"><span class="status-dot"></span><span>Ready</span></div>
  </header>

  <main>
    {% if message %}
      <div class="message {{ message_type }}">{{ message }}</div>
    {% endif %}

    <form id="screeningForm" action="{{ url_for('index') }}" method="post" enctype="multipart/form-data">
      <section class="panel">
        <div class="panel-head">
          <div>
            <h2 class="panel-title">Model Connection</h2>
            <p class="panel-note">Configure the model endpoint used for screening.</p>
          </div>
        </div>
        <div class="panel-body">
          <div class="grid">
            <label>Host / IP
              <input name="ollama_host" id="ollamaHost" type="text" value="{{ form.ollama_host }}" required>
            </label>
            <label>Port
              <input name="ollama_port" id="ollamaPort" type="number" min="1" max="65535" value="{{ form.ollama_port }}" required>
            </label>
            <label>Model
              <input name="model" id="modelName" type="text" value="{{ form.model }}" list="modelOptions" required>
              <datalist id="modelOptions">
                {% for model in known_models %}
                  <option value="{{ model }}">
                {% endfor %}
              </datalist>
            </label>
            <button class="secondary" type="button" id="refreshModels">Refresh</button>
            <button class="ghost" type="button" id="testConnection">Test</button>
          </div>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2 class="panel-title">Job Description</h2>
            <p class="panel-note">Upload a PDF, DOCX, PPTX, or TXT role document.</p>
          </div>
        </div>
        <div class="panel-body">
          <label>Role Document
            <input name="jd_file" type="file" accept=".pdf,.doc,.docx,.ppt,.pptx,.txt" required>
          </label>
        </div>
      </section>

      <section class="panel">
        <div class="panel-head">
          <div>
            <h2 class="panel-title">Resumes</h2>
            <p class="panel-note">Upload one or more candidate documents.</p>
          </div>
        </div>
        <div class="panel-body">
          <label>Candidate Documents
            <input name="resumes" type="file" accept=".pdf,.doc,.docx,.ppt,.pptx,.txt" multiple required>
          </label>
        </div>
      </section>

      <section class="panel action-panel">
        <div class="summary">
          {% if summary %}
            <span class="metric">{{ summary.total }} analysed</span>
            <span class="metric">{{ summary.selected }} selected</span>
            <span class="metric">{{ summary.rejected }} rejected</span>
            {% if errors %}<span class="metric">{{ errors|length }} error{{ '' if errors|length == 1 else 's' }}</span>{% endif %}
          {% else %}
            <span>Ready to analyse</span>
          {% endif %}
        </div>
        <div class="button-row">
          {% if job_id and results %}
            <a class="button secondary" href="{{ url_for('download_csv', job_id=job_id) }}">CSV</a>
            <a class="button secondary" href="{{ url_for('download_excel', job_id=job_id) }}">Excel</a>
          {% endif %}
          <button type="submit" id="submitButton">Run Analysis</button>
        </div>
      </section>
    </form>

    {% if errors %}
      <section class="panel" style="margin-top: 16px;">
        <div class="panel-head">
          <div>
            <h2 class="panel-title">Errors</h2>
          </div>
        </div>
        <div class="panel-body">
          {% for item in errors %}
            <div class="message error" style="margin-bottom: 10px;"><b>{{ item.file }}</b>: {{ item.error }}</div>
          {% endfor %}
        </div>
      </section>
    {% endif %}

    <section class="panel" style="margin-top: 16px;">
      <div class="panel-head">
        <div>
          <h2 class="panel-title">Analysis Results</h2>
          <p class="panel-note">Expand a row for candidate detail.</p>
        </div>
      </div>
      {% if results %}
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Name</th>
                <th>Education</th>
                <th>Current Org.</th>
                <th>Past Orgs.</th>
                <th>Total Exp.</th>
                <th>Relevant Exp.</th>
                <th>Highlights</th>
                <th>Lowlights</th>
                <th>Score</th>
                <th>Decision</th>
                <th>Detail</th>
              </tr>
            </thead>
            <tbody>
              {% for row in results %}
                {% set is_select = row.decision|lower == 'select' %}
                <tr class="{{ 'select' if is_select else 'reject' }}">
                  <td class="mono">{{ loop.index }}</td>
                  <td>{{ row.name or 'Unknown' }}<br><span class="muted">{{ row._file }}</span></td>
                  <td>{{ format_education(row.education) }}</td>
                  <td>{{ row.current_organization or '-' }}</td>
                  <td>{{ format_orgs(row.past_organizations, row.current_organization) }}</td>
                  <td>{{ format_exp(row.total_experience) }}</td>
                  <td>{{ format_exp(row.relevant_experience) }}</td>
                  <td><ul>{% for item in row.highlights or [] %}<li>{{ item }}</li>{% endfor %}</ul></td>
                  <td><ul>{% for item in row.lowlights or [] %}<li>{{ item }}</li>{% endfor %}</ul></td>
                  <td><span class="score">{{ row.score }}</span></td>
                  <td><b>{{ row.decision }}</b></td>
                  <td>
                    <details>
                      <summary>Open</summary>
                      <div class="detail-body">
                        <div class="kv"><b>File</b><span>{{ row._file }}</span></div>
                        <div class="kv"><b>Education</b><span>{{ format_education(row.education) }}</span></div>
                        <div class="kv"><b>Current</b><span>{{ row.current_organization or '-' }}</span></div>
                        <div class="kv"><b>Past</b><span>{{ format_orgs(row.past_organizations, row.current_organization) }}</span></div>
                      </div>
                    </details>
                  </td>
                </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      {% else %}
        <div class="panel-body">
          <div class="message">No analysis run yet.</div>
        </div>
      {% endif %}
    </section>
  </main>

  <footer>
    <span>Service endpoint: /api/analyze</span>
  </footer>

  <script>
    const statusEl = document.getElementById("connectionStatus");
    const statusDot = statusEl.querySelector(".status-dot");
    const statusText = statusEl.querySelector("span:last-child");
    const hostInput = document.getElementById("ollamaHost");
    const portInput = document.getElementById("ollamaPort");
    const modelInput = document.getElementById("modelName");
    const modelOptions = document.getElementById("modelOptions");
    const submitButton = document.getElementById("submitButton");
    const form = document.getElementById("screeningForm");

    function setStatus(text, state) {
      statusText.textContent = text;
      statusDot.style.background = state === "ok" ? "#22c55e" : state === "bad" ? "#ef4444" : "#8b9994";
    }

    function endpointParams() {
      return new URLSearchParams({ host: hostInput.value, port: portInput.value });
    }

    document.getElementById("refreshModels").addEventListener("click", async () => {
      setStatus("Refreshing", "idle");
      try {
        const response = await fetch(`/api/models?${endpointParams().toString()}`);
        const payload = await response.json();
        modelOptions.innerHTML = "";
        for (const name of payload.models || []) {
          const option = document.createElement("option");
          option.value = name;
          modelOptions.appendChild(option);
        }
        if (payload.models && payload.models.length && !payload.models.includes(modelInput.value)) {
          modelInput.value = payload.models[0];
        }
        setStatus(payload.models.length ? `${payload.models.length} model(s)` : "No models", payload.models.length ? "ok" : "bad");
      } catch (error) {
        setStatus("Unreachable", "bad");
      }
    });

    document.getElementById("testConnection").addEventListener("click", async () => {
      setStatus("Testing", "idle");
      try {
        const response = await fetch("/api/test-connection", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            host: hostInput.value,
            port: portInput.value,
            model: modelInput.value
          })
        });
        const payload = await response.json();
        setStatus(payload.message || "Checked", payload.ok ? "ok" : "bad");
      } catch (error) {
        setStatus("Unreachable", "bad");
      }
    });

    form.addEventListener("submit", () => {
      submitButton.disabled = true;
      submitButton.textContent = "Analysing";
      setStatus("Analysing", "idle");
    });
  </script>
</body>
</html>
"""


def extract_pdf(path: str) -> str:
    try:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as exc:  # pragma: no cover - depends on document parser internals
        return f"[PDF read error: {exc}]"


def extract_docx(path: str) -> str:
    try:
        from docx import Document

        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as exc:  # pragma: no cover - depends on document parser internals
        return f"[DOCX read error: {exc}]"


def extract_pptx(path: str) -> str:
    try:
        from pptx import Presentation

        prs = Presentation(path)
        parts: list[str] = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    parts.append(shape.text)
        return "\n".join(parts)
    except Exception as exc:  # pragma: no cover - depends on document parser internals
        return f"[PPTX read error: {exc}]"


def extract_txt(path: str) -> str:
    try:
        return Path(path).read_text(encoding="utf-8", errors="replace")
    except Exception as exc:  # pragma: no cover - depends on filesystem errors
        return f"[TXT read error: {exc}]"


def extract_text(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return extract_pdf(path)
    if ext in {".doc", ".docx"}:
        return extract_docx(path)
    if ext in {".ppt", ".pptx"}:
        return extract_pptx(path)
    if ext == ".txt":
        return extract_txt(path)
    return f"[Unsupported file type: {ext}]"


def _strip_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```[a-z]*\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def _ollama_url(host: str, port: str, path: str) -> str:
    host = host.strip().rstrip("/")
    if not host.startswith(("http://", "https://")):
        host = f"http://{host}"
    return f"{host}:{port}{path}"


def ollama_generate(host: str, port: str, model: str, prompt: str) -> str:
    url = _ollama_url(host, port, "/api/generate")
    payload = json.dumps(
        {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "format": "json",
            "options": {"temperature": 0.2},
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        url, data=payload, headers={"Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=OLLAMA_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Model server HTTP {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Cannot reach model server at {url}: {exc.reason}") from exc
    return data.get("response", "")


def ollama_list_models(host: str, port: str) -> list[str]:
    url = _ollama_url(host, port, "/api/tags")
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return [m.get("name", "") for m in data.get("models", []) if m.get("name")]
    except Exception:
        return []


def analyse_resume(path: str, jd_text: str, host: str, port: str, model: str) -> dict[str, Any]:
    resume_text = extract_text(path)
    prompt = ANALYSIS_PROMPT.format(jd=jd_text, resume=resume_text)
    raw = ollama_generate(host, port, model, prompt)
    parsed = json.loads(_strip_fences(raw))
    if not isinstance(parsed, dict):
        raise RuntimeError("Model returned JSON, but not a JSON object.")
    return parsed


def fmt_exp(exp: dict[str, Any] | None) -> str:
    if not exp:
        return "N/A"
    try:
        years = int(exp.get("years", 0))
        months = int(exp.get("months", 0))
    except (TypeError, ValueError):
        return "N/A"
    return f"{years}y {months}m"


def fmt_education(edu: list[Any] | None) -> str:
    if not edu:
        return "-"
    seen: set[str] = set()
    out: list[str] = []
    for item in edu:
        if not isinstance(item, dict):
            continue
        degree = str(item.get("degree", "") or "").strip()
        college = str(item.get("college", "") or "").strip()
        year = item.get("year")
        year_text = ""
        if isinstance(year, (int, float)) and year:
            year_text = str(int(year))
        elif year:
            year_text = str(year).strip()
        parts = [part for part in (degree, college, year_text) if part]
        if not parts:
            continue
        line = ", ".join(parts)
        key = line.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(line)
    return " | ".join(out) if out else "-"


def fmt_orgs(orgs: list[Any] | None, exclude: str = "") -> str:
    if not orgs:
        return "-"
    excluded = (exclude or "").strip().lower()
    seen: set[str] = set()
    out: list[str] = []
    for org in orgs:
        text = str(org or "").strip()
        key = text.lower()
        if not text or key in seen or key == excluded:
            continue
        seen.add(key)
        out.append(text)
    return ", ".join(out) if out else "-"


def _normalise_result(result: dict[str, Any], file_name: str) -> dict[str, Any]:
    score = result.get("score", 0)
    try:
        score = int(score)
    except (TypeError, ValueError):
        score = 0
    score = max(0, min(10, score))
    decision = str(result.get("decision") or ("Select" if score > 6 else "Reject")).strip()
    if decision.lower() not in {"select", "reject"}:
        decision = "Select" if score > 6 else "Reject"
    normalised = dict(result)
    normalised["_file"] = file_name
    normalised["name"] = str(normalised.get("name") or "Unknown").strip() or "Unknown"
    normalised["education"] = normalised.get("education") if isinstance(normalised.get("education"), list) else []
    normalised["current_organization"] = str(normalised.get("current_organization") or "").strip()
    normalised["past_organizations"] = (
        normalised.get("past_organizations")
        if isinstance(normalised.get("past_organizations"), list)
        else []
    )
    normalised["total_experience"] = (
        normalised.get("total_experience")
        if isinstance(normalised.get("total_experience"), dict)
        else {"years": 0, "months": 0}
    )
    normalised["relevant_experience"] = (
        normalised.get("relevant_experience")
        if isinstance(normalised.get("relevant_experience"), dict)
        else {"years": 0, "months": 0}
    )
    normalised["highlights"] = (
        normalised.get("highlights") if isinstance(normalised.get("highlights"), list) else []
    )
    normalised["lowlights"] = (
        normalised.get("lowlights") if isinstance(normalised.get("lowlights"), list) else []
    )
    normalised["score"] = score
    normalised["decision"] = "Select" if decision.lower() == "select" else "Reject"
    return normalised


def _summary(results: list[dict[str, Any]]) -> dict[str, int]:
    selected = sum(1 for row in results if str(row.get("decision", "")).lower() == "select")
    return {
        "total": len(results),
        "selected": selected,
        "rejected": len(results) - selected,
    }


def _view_insights(
    results: list[dict[str, Any]],
    errors: list[dict[str, str]],
) -> dict[str, Any]:
    scores = []
    for row in results:
        try:
            scores.append(int(row.get("score", 0)))
        except (TypeError, ValueError):
            scores.append(0)
    selected = sum(1 for row in results if str(row.get("decision", "")).lower() == "select")
    total = len(results)
    return {
        "total": total,
        "selected": selected,
        "rejected": total - selected,
        "errors": len(errors),
        "average_score": round(sum(scores) / total, 1) if total else 0,
        "top_score": max(scores) if scores else 0,
        "shortlist_rate": round((selected / total) * 100) if total else 0,
        "top_candidates": sorted(
            results,
            key=lambda row: int(row.get("score") or 0),
            reverse=True,
        )[:3],
    }


def _validate_port(port: str) -> str:
    text = str(port or "").strip()
    if not text.isdigit() or not 1 <= int(text) <= 65535:
        raise ValueError("Port must be a number from 1 to 65535.")
    return text


def _validate_file(file: FileStorage | None, label: str) -> None:
    if file is None or not file.filename:
        raise ValueError(f"{label} is required.")
    ext = Path(file.filename).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"{label} has an unsupported file type: {ext or 'none'}.")


def _save_upload(file: FileStorage, target_dir: Path) -> Path:
    _validate_file(file, "Upload")
    safe_name = secure_filename(file.filename or f"upload-{uuid.uuid4().hex}")
    dest = target_dir / f"{uuid.uuid4().hex}-{safe_name}"
    file.save(dest)
    return dest


def _screen_documents(
    jd_path: Path,
    resume_paths: list[Path],
    host: str,
    port: str,
    model: str,
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    jd_text = extract_text(str(jd_path))
    results: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []

    for path in resume_paths:
        try:
            result = analyse_resume(str(path), jd_text, host, port, model)
            results.append(_normalise_result(result, path.name.split("-", 1)[-1]))
        except json.JSONDecodeError as exc:
            errors.append({"file": path.name.split("-", 1)[-1], "error": f"Model returned invalid JSON: {exc}"})
        except Exception as exc:
            errors.append({"file": path.name.split("-", 1)[-1], "error": str(exc)})
    return results, errors


def _save_job(
    results: list[dict[str, Any]],
    errors: list[dict[str, str]],
    config: dict[str, str],
) -> str:
    EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    payload = {
        "job_id": job_id,
        "summary": _summary(results),
        "results": results,
        "errors": errors,
        "config": config,
    }
    (EXPORT_ROOT / f"{job_id}.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return job_id


def _load_job(job_id: str) -> dict[str, Any]:
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        abort(404)
    path = EXPORT_ROOT / f"{job_id}.json"
    if not path.exists():
        abort(404)
    return json.loads(path.read_text(encoding="utf-8"))


def _csv_bytes(results: list[dict[str, Any]]) -> bytes:
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(
        [
            "#",
            "Name",
            "Education",
            "Current Organization",
            "Past Organizations",
            "Total Experience",
            "Relevant Experience",
            "Highlights",
            "Lowlights",
            "Score",
            "Decision",
        ]
    )
    for idx, row in enumerate(results, 1):
        current_org = str(row.get("current_organization") or "").strip()
        writer.writerow(
            [
                idx,
                row.get("name", ""),
                fmt_education(row.get("education")),
                current_org,
                fmt_orgs(row.get("past_organizations"), current_org),
                fmt_exp(row.get("total_experience")),
                fmt_exp(row.get("relevant_experience")),
                " | ".join(str(item) for item in row.get("highlights", [])),
                " | ".join(str(item) for item in row.get("lowlights", [])),
                row.get("score", ""),
                row.get("decision", ""),
            ]
        )
    return output.getvalue().encode("utf-8-sig")


def _excel_bytes(results: list[dict[str, Any]]) -> bytes:
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    except ImportError as exc:
        raise RuntimeError("openpyxl is not installed. Install requirements.txt before exporting Excel.") from exc

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Resume Analysis"

    header_fill = PatternFill("solid", fgColor="17211F")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Side(style="thin", color="DBE3DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    headers = [
        "#",
        "Name",
        "Education",
        "Current Organization",
        "Past Organizations",
        "Total Experience",
        "Relevant Experience",
        "Highlights",
        "Lowlights",
        "Score",
        "Decision",
    ]
    for col, title in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col, value=title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    select_fill = PatternFill("solid", fgColor="DFF7EA")
    reject_fill = PatternFill("solid", fgColor="FEE4E2")
    center_cols = {1, 6, 7, 10, 11}

    for row_idx, row in enumerate(results, 2):
        current_org = str(row.get("current_organization") or "").strip()
        is_select = str(row.get("decision", "")).lower() == "select"
        fill = select_fill if is_select else reject_fill
        values = [
            row_idx - 1,
            row.get("name", ""),
            fmt_education(row.get("education")).replace(" | ", "\n"),
            current_org,
            fmt_orgs(row.get("past_organizations"), current_org).replace(", ", "\n"),
            fmt_exp(row.get("total_experience")),
            fmt_exp(row.get("relevant_experience")),
            "\n".join(f"- {item}" for item in row.get("highlights", [])),
            "\n".join(f"- {item}" for item in row.get("lowlights", [])),
            row.get("score", ""),
            row.get("decision", ""),
        ]
        for col_idx, value in enumerate(values, 1):
            cell = worksheet.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
            cell.border = border
            cell.alignment = center if col_idx in center_cols else left

    for col, width in zip("ABCDEFGHIJK", [5, 24, 38, 24, 32, 16, 16, 42, 32, 9, 13]):
        worksheet.column_dimensions[col].width = width
    worksheet.freeze_panes = "A2"

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _form_defaults() -> dict[str, str]:
    return {
        "ollama_host": DEFAULT_OLLAMA_HOST,
        "ollama_port": DEFAULT_OLLAMA_PORT,
        "model": DEFAULT_MODEL,
    }


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("RESUMES_SCREENER_MAX_UPLOAD_MB", "128")) * 1024 * 1024
    app.jinja_env.globals.update(
        format_exp=fmt_exp,
        format_education=fmt_education,
        format_orgs=fmt_orgs,
    )
    UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
    EXPORT_ROOT.mkdir(parents=True, exist_ok=True)

    @app.get("/health")
    def health() -> Response:
        return jsonify({"status": "ok", "service": APP_NAME})

    @app.get("/api/models")
    def api_models() -> Response:
        host = request.args.get("host", DEFAULT_OLLAMA_HOST)
        port = request.args.get("port", DEFAULT_OLLAMA_PORT)
        try:
            port = _validate_port(port)
        except ValueError as exc:
            return jsonify({"ok": False, "models": [], "error": str(exc)}), 400
        models = ollama_list_models(host, port)
        return jsonify({"ok": bool(models), "models": models})

    @app.post("/api/test-connection")
    def api_test_connection() -> Response:
        data = request.get_json(silent=True) or request.form
        host = str(data.get("host") or DEFAULT_OLLAMA_HOST).strip()
        model = str(data.get("model") or DEFAULT_MODEL).strip()
        try:
            port = _validate_port(str(data.get("port") or DEFAULT_OLLAMA_PORT))
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

        models = ollama_list_models(host, port)
        if not models:
            return jsonify({"ok": False, "message": "Unreachable", "models": []}), 503
        if model not in models:
            return jsonify({"ok": False, "message": "Model missing", "models": models}), 404
        return jsonify({"ok": True, "message": f"Connected: {model}", "models": models})

    @app.route("/", methods=["GET", "POST"])
    def index() -> str | tuple[str, int]:
        form = _form_defaults()
        message = ""
        message_type = ""
        results: list[dict[str, Any]] = []
        errors: list[dict[str, str]] = []
        summary: dict[str, int] | None = None
        job_id: str | None = None

        if request.method == "POST":
            form = {
                "ollama_host": request.form.get("ollama_host", DEFAULT_OLLAMA_HOST).strip(),
                "ollama_port": request.form.get("ollama_port", DEFAULT_OLLAMA_PORT).strip(),
                "model": request.form.get("model", DEFAULT_MODEL).strip(),
            }
            try:
                payload = _handle_analysis_request(form)
                results = payload["results"]
                errors = payload["errors"]
                summary = payload["summary"]
                job_id = payload["job_id"]
                message = "Analysis complete."
                message_type = "ok"
            except ValueError as exc:
                message = str(exc)
                message_type = "error"
                return (
                    render_template(
                        "index.html",
                        app_name=APP_NAME,
                        org_name=ORG_NAME,
                        form=form,
                        known_models=[DEFAULT_MODEL],
                        message=message,
                        message_type=message_type,
                        results=results,
                        errors=errors,
                        summary=summary,
                        insights=_view_insights(results, errors),
                        job_id=job_id,
                    ),
                    400,
                )

        return render_template(
            "index.html",
            app_name=APP_NAME,
            org_name=ORG_NAME,
            form=form,
            known_models=[DEFAULT_MODEL, "llama3", "mistral", "qwen2.5", "phi3"],
            message=message,
            message_type=message_type,
            results=results,
            errors=errors,
            summary=summary,
            insights=_view_insights(results, errors),
            job_id=job_id,
        )

    @app.post("/api/analyze")
    def api_analyze() -> Response:
        form = {
            "ollama_host": request.form.get("ollama_host") or request.form.get("host") or DEFAULT_OLLAMA_HOST,
            "ollama_port": request.form.get("ollama_port") or request.form.get("port") or DEFAULT_OLLAMA_PORT,
            "model": request.form.get("model") or DEFAULT_MODEL,
        }
        try:
            payload = _handle_analysis_request(form)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        return jsonify(
            {
                "ok": not payload["errors"],
                "job_id": payload["job_id"],
                "summary": payload["summary"],
                "results": payload["results"],
                "errors": payload["errors"],
                "downloads": {
                    "csv": url_for("download_csv", job_id=payload["job_id"], _external=True),
                    "excel": url_for("download_excel", job_id=payload["job_id"], _external=True),
                },
            }
        )

    @app.get("/download/<job_id>.csv")
    def download_csv(job_id: str) -> Response:
        payload = _load_job(job_id)
        data = _csv_bytes(payload.get("results", []))
        return send_file(
            io.BytesIO(data),
            mimetype="text/csv",
            as_attachment=True,
            download_name=f"resume_analysis_{job_id}.csv",
        )

    @app.get("/download/<job_id>.xlsx")
    def download_excel(job_id: str) -> Response:
        payload = _load_job(job_id)
        try:
            data = _excel_bytes(payload.get("results", []))
        except RuntimeError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 500
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"resume_analysis_{job_id}.xlsx",
        )

    def _handle_analysis_request(form: dict[str, str]) -> dict[str, Any]:
        host = str(form.get("ollama_host") or DEFAULT_OLLAMA_HOST).strip()
        port = _validate_port(str(form.get("ollama_port") or DEFAULT_OLLAMA_PORT))
        model = str(form.get("model") or DEFAULT_MODEL).strip()
        if not host:
            raise ValueError("Host / IP is required.")
        if not model:
            raise ValueError("Model is required.")

        jd_file = request.files.get("jd_file")
        resume_files = request.files.getlist("resumes")
        _validate_file(jd_file, "Job description")
        if not resume_files or not any(file.filename for file in resume_files):
            raise ValueError("At least one resume file is required.")
        for resume_file in resume_files:
            _validate_file(resume_file, "Resume")

        with tempfile.TemporaryDirectory(prefix="job-", dir=UPLOAD_ROOT) as work_dir_name:
            work_dir = Path(work_dir_name)
            assert jd_file is not None
            jd_path = _save_upload(jd_file, work_dir)
            resume_paths = [_save_upload(file, work_dir) for file in resume_files if file.filename]
            results, errors = _screen_documents(jd_path, resume_paths, host, port, model)

        config = {"host": host, "port": port, "model": model}
        saved_job_id = _save_job(results, errors, config)
        return {
            "job_id": saved_job_id,
            "summary": _summary(results),
            "results": results,
            "errors": errors,
        }

    return app


app = create_app()


def main() -> None:
    host = os.environ.get("RESUMES_SCREENER_HOST", "0.0.0.0")
    port = int(os.environ.get("RESUMES_SCREENER_PORT", "80"))
    threads = int(os.environ.get("RESUMES_SCREENER_THREADS", "8"))
    try:
        from waitress import serve
    except ImportError as exc:
        raise SystemExit("waitress is not installed. Run: python -m pip install -r requirements.txt") from exc
    serve(app, host=host, port=port, threads=threads)


if __name__ == "__main__":
    main()
