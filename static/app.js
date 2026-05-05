const statusEl = document.getElementById("connectionStatus");
const statusDot = statusEl.querySelector(".status-dot");
const statusText = statusEl.querySelector(".status-text");
const hostInput = document.getElementById("ollamaHost");
const portInput = document.getElementById("ollamaPort");
const modelInput = document.getElementById("modelName");
const modelOptions = document.getElementById("modelOptions");
const submitButton = document.getElementById("submitButton");
const form = document.getElementById("screeningForm");

function setStatus(text, state) {
  statusText.textContent = text;
  statusDot.style.background = state === "ok" ? "#22c55e" : state === "bad" ? "#ef4444" : "#94a3b8";
}

function endpointParams() {
  return new URLSearchParams({ host: hostInput.value, port: portInput.value });
}

document.querySelectorAll("input[type='file'][data-file-label]").forEach((input) => {
  input.addEventListener("change", () => {
    const label = document.getElementById(input.dataset.fileLabel);
    const count = input.files ? input.files.length : 0;
    if (!count) {
      return;
    }
    label.textContent = count === 1 ? input.files[0].name : `${count} files selected`;
  });
});

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
        model: modelInput.value,
      }),
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
