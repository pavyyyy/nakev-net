"use strict";

pdfjsLib.GlobalWorkerOptions.workerSrc = "../assets/vendor/pdf.worker.min.js";

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  runButton: document.querySelector("#runButton"),
  clearButton: document.querySelector("#clearButton"),
  status: document.querySelector("#status"),
  previewWrap: document.querySelector("#previewWrap"),
  previewBody: document.querySelector("#previewBody"),
};

let selectedFile = null;
let parsedRows = [];

els.fileInput.addEventListener("change", async () => {
  const file = els.fileInput.files[0] || null;
  if (file) await setFile(file);
});
els.runButton.addEventListener("click", generate);
els.clearButton.addEventListener("click", clear);
setupDropzone(els.dropzone, async (files) => {
  const file = files.find((item) => /\.pdf$/i.test(item.name));
  if (file) await setFile(file);
});

function setupDropzone(element, onFiles) {
  element.addEventListener("dragover", (event) => {
    event.preventDefault();
    element.classList.add("dragover");
  });
  element.addEventListener("dragleave", () => element.classList.remove("dragover"));
  element.addEventListener("drop", async (event) => {
    event.preventDefault();
    element.classList.remove("dragover");
    await onFiles([...event.dataTransfer.files]);
  });
}

async function setFile(file) {
  selectedFile = file;
  parsedRows = [];
  els.runButton.disabled = true;
  els.clearButton.disabled = false;
  setStatus(`Reading ${file.name}...`);

  try {
    const lines = await extractPdfLines(file);
    parsedRows = parseAbbInvoiceRows(lines);
    renderRows(parsedRows);
    els.runButton.disabled = !parsedRows.length;
    setStatus(parsedRows.length ? `Found ${parsedRows.length} row(s).` : "No invoice rows found.", parsedRows.length ? "ok" : "error");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  }
}

function clear() {
  selectedFile = null;
  parsedRows = [];
  els.fileInput.value = "";
  els.runButton.disabled = true;
  els.clearButton.disabled = true;
  els.previewWrap.hidden = true;
  els.previewBody.innerHTML = "";
  setStatus("Ready.");
}

function generate() {
  if (!selectedFile || !parsedRows.length) return;
  const rows = [["Your document", "Date", "Gross amount"]];
  parsedRows.forEach((row) => rows.push([row.document, row.date, row.grossAmount]));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [{ wch: 22 }, { wch: 14 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  downloadBlob(
    new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    replaceExtension(selectedFile.name, ".xlsx"),
  );
  setStatus(`Generated ${replaceExtension(selectedFile.name, ".xlsx")}.`, "ok");
}

async function extractPdfLines(file) {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const lines = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const text = await page.getTextContent();
    lines.push(...textItemsToLines(text.items));
  }
  return lines.map((line) => line.trim()).filter(Boolean);
}

function textItemsToLines(items) {
  const grouped = [];
  items.forEach((item) => {
    const text = String(item.str || "").trim();
    if (!text) return;
    const x = item.transform[4];
    const y = item.transform[5];
    let line = grouped.find((candidate) => Math.abs(candidate.y - y) < 3);
    if (!line) {
      line = { y, parts: [] };
      grouped.push(line);
    }
    line.parts.push({ x, text });
  });
  return grouped
    .sort((a, b) => b.y - a.y)
    .map((line) => line.parts.sort((a, b) => a.x - b.x).map((part) => part.text).join(" "));
}

function parseAbbInvoiceRows(lines) {
  const fullText = lines.join(" ");
  const separator = /_{20,}\s*/g;
  const sections = fullText.includes("____________________")
    ? fullText.split(separator).map((part) => part.trim()).filter(Boolean)
    : lines;

  const rows = parseLikeOriginalPython(sections);
  if (rows.length) return rows;
  return parseWithLineFallback(lines);
}

function parseLikeOriginalPython(sections) {
  const kept = sections.filter((_, index) => index % 2 === 1);
  const tokenRows = kept.map((section) => section.split(/\s+/).filter(Boolean)).filter((tokens) => tokens.length >= 5);
  tokenRows.forEach((tokens, index) => {
    if (index > 0 && tokens.length > 5) tokens.splice(0, Math.min(5, tokens.length - 5));
  });

  return tokenRows.map((tokens) => {
    const last = tokens.slice(-5);
    if (last.length < 5) return null;
    const reordered = [last[2], last[3], last[4], last[0], last[1]];
    const gross = parseEuropeanNumber(reordered[2]);
    if (!Number.isFinite(gross)) return null;
    return { document: reordered[4], date: reordered[0], grossAmount: gross };
  }).filter(Boolean);
}

function parseWithLineFallback(lines) {
  const rows = [];
  const dateRe = /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/;
  lines.forEach((line) => {
    const dateMatch = line.match(dateRe);
    if (!dateMatch) return;
    const tokens = line.split(/\s+/).filter(Boolean);
    const amountToken = [...tokens].reverse().find((token) => /^[+-]?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})$/.test(token));
    const documentToken = tokens.find((token) => /[A-Za-z0-9/-]{4,}/.test(token) && token !== amountToken && token !== dateMatch[0]);
    const gross = parseEuropeanNumber(amountToken);
    if (documentToken && Number.isFinite(gross)) {
      rows.push({ document: documentToken, date: dateMatch[0], grossAmount: gross });
    }
  });
  return rows;
}

function parseEuropeanNumber(value) {
  if (!value) return NaN;
  const cleaned = String(value).replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
  return Number(cleaned);
}

function renderRows(rows) {
  els.previewWrap.hidden = !rows.length;
  els.previewBody.innerHTML = rows.map((row) => (
    `<tr><td>${escapeHtml(row.document)}</td><td>${escapeHtml(row.date)}</td><td>${escapeHtml(row.grossAmount.toFixed(2))}</td></tr>`
  )).join("");
}

function setStatus(message, kind = "") {
  els.status.textContent = message;
  els.status.className = `status ${kind}`.trim();
}

function replaceExtension(name, extension) {
  return safeFileName(String(name).replace(/\.[^.]+$/, "") + extension);
}

function safeFileName(name) {
  return String(name).replace(/[<>:"/\\|?*]/g, "_").trim() || "output.xlsx";
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  }[char]));
}
