/**
 * Report Automation Tool — Frontend JS
 * Handles file upload, form submission, progress animation, and result display.
 */

"use strict";

/* ─── DOM refs ──────────────────────────────────────────────────────────────── */
const uploadZone      = document.getElementById("upload-zone");
const fileInput       = document.getElementById("file-input");
const fileSelectedBox = document.getElementById("file-selected");
const fileNameEl      = document.getElementById("file-name");
const fileSizeEl      = document.getElementById("file-size");
const removeFileBtn   = document.getElementById("remove-file");

const uploadEmpZone      = document.getElementById("upload-emp-zone");
const empFileInput       = document.getElementById("emp-file-input");
const empFileSelectedBox = document.getElementById("emp-file-selected");
const empFileNameEl      = document.getElementById("emp-file-name");
const empFileSizeEl      = document.getElementById("emp-file-size");
const empRemoveFileBtn   = document.getElementById("emp-remove-file");

const convertBtn      = document.getElementById("btn-convert");
const convertBtnText  = document.getElementById("btn-text");
const progressWrap    = document.getElementById("progress-wrap");
const progressFill    = document.getElementById("progress-fill");
const progressMsg     = document.getElementById("progress-msg");
const statusBar       = document.getElementById("status-bar");
const statusIcon      = document.getElementById("status-icon");
const statusText      = document.getElementById("status-text");
const statsRow        = document.getElementById("stats-row");
const statInput       = document.getElementById("stat-input");
const statFiltered    = document.getElementById("stat-filtered");
const statFinal       = document.getElementById("stat-final");
const downloadBtn     = document.getElementById("btn-download");

/* ─── State ─────────────────────────────────────────────────────────────────── */
let selectedFile = null;
let selectedEmpFile = null;
let progressInterval = null;

function updateConvertBtn() {
  convertBtn.disabled = !(selectedFile && selectedEmpFile);
}

/* ─── File selection helpers ────────────────────────────────────────────────── */
function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function setFile(file) {
  if (!file || !file.name.toLowerCase().endsWith(".xlsx")) {
    showStatus("error", "⚠️", "Please select a valid .xlsx file.");
    return;
  }
  selectedFile = file;
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatBytes(file.size);
  fileSelectedBox.classList.add("visible");
  updateConvertBtn();
  clearStatus();
}

function clearFile() {
  selectedFile = null;
  fileInput.value = "";
  fileSelectedBox.classList.remove("visible");
  updateConvertBtn();
  clearStatus();
  resetResults();
}

function setEmpFile(file) {
  if (!file || !file.name.toLowerCase().endsWith(".xlsx")) {
    showStatus("error", "⚠️", "Please select a valid .xlsx file for emp_info.");
    return;
  }
  selectedEmpFile = file;
  empFileNameEl.textContent = file.name;
  empFileSizeEl.textContent = formatBytes(file.size);
  empFileSelectedBox.classList.add("visible");
  updateConvertBtn();
  clearStatus();
}

function clearEmpFile() {
  selectedEmpFile = null;
  empFileInput.value = "";
  empFileSelectedBox.classList.remove("visible");
  updateConvertBtn();
  clearStatus();
  resetResults();
}

/* ─── Drag & Drop ───────────────────────────────────────────────────────────── */
uploadZone.addEventListener("click", () => fileInput.click());

uploadZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadZone.classList.add("drag-over");
});

uploadZone.addEventListener("dragleave", () => {
  uploadZone.classList.remove("drag-over");
});

uploadZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

removeFileBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  clearFile();
});

// Emp File zone listeners
uploadEmpZone.addEventListener("click", () => empFileInput.click());

uploadEmpZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadEmpZone.classList.add("drag-over");
});

uploadEmpZone.addEventListener("dragleave", () => {
  uploadEmpZone.classList.remove("drag-over");
});

uploadEmpZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadEmpZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) setEmpFile(file);
});

empFileInput.addEventListener("change", () => {
  if (empFileInput.files[0]) setEmpFile(empFileInput.files[0]);
});

empRemoveFileBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  clearEmpFile();
});

/* ─── Status helpers ────────────────────────────────────────────────────────── */
function showStatus(type, icon, msg) {
  statusBar.className = `status-bar visible ${type}`;
  statusIcon.textContent = icon;
  statusText.textContent = msg;
}

function clearStatus() {
  statusBar.classList.remove("visible");
}

/* ─── Progress animation ────────────────────────────────────────────────────── */
const PROGRESS_STAGES = [
  { pct: 5,  msg: "Reading raw Excel file…" },
  { pct: 20, msg: "Loading employee mapping…" },
  { pct: 35, msg: "Mapping team types…" },
  { pct: 48, msg: "Filtering Super Cats…" },
  { pct: 58, msg: "Cleaning null values…" },
  { pct: 68, msg: "Mapping city & paid flags…" },
  { pct: 76, msg: "Classifying Owner/Orphan…" },
  { pct: 85, msg: "Building summary…" },
  { pct: 94, msg: "Writing output Excel…" },
];

function startProgress() {
  progressWrap.classList.add("visible");
  progressFill.style.width = "0%";
  let stageIdx = 0;

  progressInterval = setInterval(() => {
    if (stageIdx >= PROGRESS_STAGES.length) {
      clearInterval(progressInterval);
      return;
    }
    const stage = PROGRESS_STAGES[stageIdx];
    progressFill.style.width = stage.pct + "%";
    progressMsg.textContent = stage.msg;
    stageIdx++;
  }, 3500);
}

function stopProgress(success) {
  clearInterval(progressInterval);
  progressFill.style.width = success ? "100%" : "0%";
  progressMsg.textContent = success ? "Complete!" : "";
  setTimeout(() => {
    progressWrap.classList.remove("visible");
    progressFill.style.width = "0%";
  }, 1200);
}

/* ─── Results display ───────────────────────────────────────────────────────── */
function resetResults() {
  statsRow.classList.remove("visible");
  downloadBtn.classList.remove("visible");
}

function renderStats(stats) {
  statInput.textContent   = (stats.input_rows   || 0).toLocaleString();
  statFiltered.textContent = (stats.after_supercat_filter || 0).toLocaleString();
  statFinal.textContent   = (stats.final_rows   || 0).toLocaleString();
  statsRow.classList.add("visible");
}

/* ─── Main conversion ───────────────────────────────────────────────────────── */
convertBtn.addEventListener("click", async () => {
  if (!selectedFile) {
    showStatus("error", "⚠️", "Please select a file first.");
    return;
  }

  // Reset UI
  resetResults();
  clearStatus();
  convertBtn.disabled = true;
  convertBtn.classList.add("loading");
  convertBtnText.textContent = "Processing…";
  showStatus("info", "⚡", "Uploading and processing your file. This may take up to 60 seconds…");
  startProgress();

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("emp_file", selectedEmpFile);

  try {
    const resp = await fetch("/process", {
      method: "POST",
      body: formData,
    });

    const data = await resp.json();
    stopProgress(resp.ok && data.status === "ok");

    if (!resp.ok || data.status !== "ok") {
      showStatus("error", "❌", data.message || "An unexpected error occurred.");
      return;
    }

    // Success
    showStatus("success", "✅", `Processing complete! ${data.stats.final_rows?.toLocaleString() ?? ""} records written to output.`);
    renderStats(data.stats);
    downloadBtn.dataset.month = data.month || "July";
    downloadBtn.classList.add("visible");

  } catch (err) {
    stopProgress(false);
    showStatus("error", "❌", "Network error: could not reach the server. Is it running?");
    console.error(err);
  } finally {
    convertBtn.disabled = false;
    convertBtn.classList.remove("loading");
    convertBtnText.textContent = "Convert & Generate Report";
  }
});

/* ─── Download ──────────────────────────────────────────────────────────────── */
downloadBtn.addEventListener("click", async () => {
  try {
    const month = downloadBtn.dataset.month || "July";
    const resp = await fetch(`/download?month=${encodeURIComponent(month)}`);
    if (!resp.ok) {
      showStatus("error", "❌", "Download failed. Please process a file first.");
      return;
    }
    // Get filename from Content-Disposition header, fall back to default
    const disposition = resp.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : "supercats_output_report.xlsx";

    // Blob approach: only reliable way to force a named download in Chrome
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    showStatus("error", "❌", "Download failed: " + err.message);
    console.error(err);
  }
});
