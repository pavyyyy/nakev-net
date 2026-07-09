"use strict";

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  runButton: document.querySelector("#runButton"),
  clearButton: document.querySelector("#clearButton"),
  status: document.querySelector("#status"),
  previewWrap: document.querySelector("#previewWrap"),
  previewBody: document.querySelector("#previewBody"),
};

let selectedFiles = [];

els.fileInput.addEventListener("change", () => setFiles([...els.fileInput.files]));
els.runButton.addEventListener("click", generate);
els.clearButton.addEventListener("click", clear);

setupDropzone(els.dropzone, (files) => setFiles(files.filter(isExcelFile)));

function setupDropzone(element, onFiles) {
  element.addEventListener("dragover", (event) => {
    event.preventDefault();
    element.classList.add("dragover");
  });
  element.addEventListener("dragleave", () => element.classList.remove("dragover"));
  element.addEventListener("drop", (event) => {
    event.preventDefault();
    element.classList.remove("dragover");
    onFiles([...event.dataTransfer.files]);
  });
}

function setFiles(files) {
  selectedFiles = files;
  els.runButton.disabled = !selectedFiles.length;
  els.clearButton.disabled = !selectedFiles.length;
  els.previewWrap.hidden = !selectedFiles.length;
  els.previewBody.innerHTML = selectedFiles.map((file) => (
    `<tr><td>${escapeHtml(file.name)}</td><td>Ready</td><td>${escapeHtml(groupFromName(file.name))}</td></tr>`
  )).join("");
  setStatus(selectedFiles.length ? `${selectedFiles.length} file(s) selected.` : "Ready.", "ok");
}

function clear() {
  selectedFiles = [];
  els.fileInput.value = "";
  els.previewBody.innerHTML = "";
  els.previewWrap.hidden = true;
  els.runButton.disabled = true;
  els.clearButton.disabled = true;
  setStatus("Ready.");
}

async function generate() {
  if (!selectedFiles.length) return;
  els.runButton.disabled = true;
  try {
    const sheets = [];
    for (const file of selectedFiles) {
      const result = await convertAccountingFile(file);
      sheets.push(result);
    }

    const workbook = buildCombinedWorkbook(sheets);
    const fileName = buildCombinedFileName(sheets);
    downloadBytes(writeWorkbook(workbook), fileName, XLSX_MIME);

    renderResults(sheets);
    setStatus(`Generated ${fileName}.`, "ok");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    els.runButton.disabled = !selectedFiles.length;
  }
}

function renderResults(outputs) {
  els.previewBody.innerHTML = outputs.map((item) => (
    `<tr><td>${escapeHtml(item.sheetName)}</td><td>${item.rowCount}</td><td>${escapeHtml(item.group)}</td></tr>`
  )).join("");
}

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function convertAccountingFile(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: true,
    cellNF: true,
  });
  const source = workbook.Sheets[workbook.SheetNames[0]];
  if (!source) throw new Error(`${file.name}: no worksheet found.`);

  const range = XLSX.utils.decode_range(source["!ref"] || "A1:A1");
  const group = groupFromName(file.name);
  let firm = cellValue(source, "A11");
  const rows = [["Контрагенти", "Номер на документ", "От дата", "Валута", "Дебит", "Кредит"]];

  for (let r = 0; r <= range.e.r; r += 1) {
    const rowNumber = r + 1;
    const documentNumber = cellValue(source, `B${rowNumber}`);
    const dateValue = readDate(source, `E${rowNumber}`);
    const currency = cellValue(source, `G${rowNumber}`);
    const debit = cellValue(source, `M${rowNumber}`);
    const credit = cellValue(source, `O${rowNumber}`);

    if (documentNumber != null && String(documentNumber).slice(0, 3) === group) {
      firm = documentNumber;
    }

    if (dateValue) {
      rows.push([
        firm || "",
        normalizeDocumentNumber(documentNumber),
        formatDateDots(dateValue),
        currency,
        debit,
        credit,
      ]);
    }
  }

  return { group, sheetName: group, rows, rowCount: rows.length - 1 };
}

function buildCombinedWorkbook(outputs) {
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set();

  outputs.forEach((item) => {
    const sheet = XLSX.utils.aoa_to_sheet(item.rows);
    sheet["!cols"] = [{ wch: 24 }, { wch: 20 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
    const sheetName = uniqueSheetName(item.sheetName, usedNames);
    usedNames.add(sheetName);
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  });

  return workbook;
}

function uniqueSheetName(baseName, usedNames) {
  const cleanBaseName = String(baseName || "Sheet").slice(0, 31) || "Sheet";
  if (!usedNames.has(cleanBaseName)) return cleanBaseName;

  let index = 2;
  while (true) {
    const suffix = `_${index}`;
    const candidate = `${cleanBaseName.slice(0, 31 - suffix.length)}${suffix}`;
    if (!usedNames.has(candidate)) return candidate;
    index += 1;
  }
}

function buildCombinedFileName(outputs) {
  const groups = [...new Set(outputs.map((item) => item.group).filter(Boolean))];
  return safeFileName(`${groups.join("_") || "auto-excel"}.xlsx`);
}

function writeWorkbook(workbook) {
  return XLSX.write(workbook, { bookType: "xlsx", type: "array", cellDates: true });
}

function readDate(sheet, address) {
  const cell = sheet[address];
  if (!cell) return null;
  if (cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) return cell.v;
  if (typeof cell.v === "number" && looksLikeDateFormat(cell.z)) {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (typeof cell.v === "string") {
    const match = cell.v.trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
    if (match) {
      const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
      return new Date(year, Number(match[2]) - 1, Number(match[1]));
    }
  }
  return null;
}

function looksLikeDateFormat(format) {
  return /[dmy]/i.test(String(format || ""));
}

function cellValue(sheet, address) {
  const cell = sheet[address];
  return cell ? cell.v : null;
}

function groupFromName(name) {
  return String(name || "").slice(0, 3);
}

function formatDateDots(date) {
  return [date.getDate(), date.getMonth() + 1, date.getFullYear()]
    .map((part) => String(part).padStart(2, "0"))
    .join(".");
}

function normalizeDocumentNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return Math.trunc(value);
  if (typeof value !== "string") return value;

  const trimmed = value.trim();
  if (!/^\d+$/.test(trimmed)) return value;

  const withoutLeadingZeroes = trimmed.replace(/^0+/, "");
  return withoutLeadingZeroes || "0";
}

function replaceExtension(name, extension) {
  return safeFileName(String(name).replace(/\.[^.]+$/, "") + extension);
}

function isExcelFile(file) {
  return /\.(xls|xlsx)$/i.test(file.name);
}

function setStatus(message, kind = "") {
  els.status.textContent = message;
  els.status.className = `status ${kind}`.trim();
}

function downloadBytes(bytes, fileName, mime) {
  downloadBlob(new Blob([bytes], { type: mime }), fileName);
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

function safeFileName(name) {
  return String(name).replace(/[<>:"/\\|?*]/g, "_").trim() || "output.xlsx";
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
