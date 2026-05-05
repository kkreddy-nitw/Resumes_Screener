#!/usr/bin/env bash
set -euo pipefail

SERVICE_NAME="${SERVICE_NAME:-resumes-screener}"
HOST="${HOST:-0.0.0.0}"
PORT="${PORT:-80}"
SERVICE_USER="${SERVICE_USER:-root}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

if [[ "${EUID}" -ne 0 ]]; then
  echo "Run this script with sudo."
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
VENV_DIR="${VENV_DIR:-${APP_DIR}/.venv}"

"${PYTHON_BIN}" -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/python" -m pip install --upgrade pip
"${VENV_DIR}/bin/python" -m pip install -r "${APP_DIR}/requirements.txt"

mkdir -p "${APP_DIR}/uploads" "${APP_DIR}/exports" "${APP_DIR}/logs"
if id "${SERVICE_USER}" >/dev/null 2>&1; then
  chown -R "${SERVICE_USER}:${SERVICE_USER}" "${APP_DIR}/uploads" "${APP_DIR}/exports" "${APP_DIR}/logs"
fi

cat > "/etc/systemd/system/${SERVICE_NAME}.service" <<UNIT
[Unit]
Description=Resumes Screener
After=network.target

[Service]
Type=simple
WorkingDirectory=${APP_DIR}
Environment=RESUMES_SCREENER_HOST=${HOST}
Environment=RESUMES_SCREENER_PORT=${PORT}
ExecStart=${VENV_DIR}/bin/python ${APP_DIR}/resumes_screener.py
Restart=always
RestartSec=5
User=${SERVICE_USER}

[Install]
WantedBy=multi-user.target
UNIT

systemctl daemon-reload
systemctl enable --now "${SERVICE_NAME}"

echo "Resumes Screener service installed and started on http://${HOST}:${PORT}"
echo "Use: systemctl status ${SERVICE_NAME}"
