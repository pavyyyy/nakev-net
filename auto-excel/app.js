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
    const outputs = [];
    for (const file of selectedFiles) {
      const result = await convertAccountingFile(file, { convertDocumentNumber: false, datesAsText: false });
      outputs.push({
        fileName: replaceExtension(file.name, ".xlsx"),
        bytes: writeWorkbook(result.workbook),
        rowCount: result.rowCount,
      });
    }

    if (outputs.length === 1) {
      downloadBytes(outputs[0].bytes, outputs[0].fileName, XLSX_MIME);
    } else {
      const zip = new JSZip();
      outputs.forEach((item) => zip.file(item.fileName, item.bytes));
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, `auto-excel-${timestampForName()}.zip`);
    }

    renderResults(outputs);
    setStatus(`Generated ${outputs.length} workbook(s).`, "ok");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    els.runButton.disabled = !selectedFiles.length;
  }
}

function renderResults(outputs) {
  els.previewBody.innerHTML = outputs.map((item) => (
    `<tr><td>${escapeHtml(item.fileName)}</td><td>${item.rowCount}</td><td>${escapeHtml(groupFromName(item.fileName))}</td></tr>`
  )).join("");
}

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function convertAccountingFile(file, options) {
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
        options.convertDocumentNumber ? maybeNumber(documentNumber) : documentNumber,
        options.datesAsText ? formatDateSlash(dateValue) : dateValue,
        currency,
        debit,
        credit,
      ]);
    }
  }

  const output = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [{ wch: 24 }, { wch: 20 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
  for (let r = 2; r <= rows.length; r += 1) {
    const dateCell = sheet[`C${r}`];
    if (dateCell && dateCell.v instanceof Date) dateCell.z = "dd.mm.yyyy";
  }
  XLSX.utils.book_append_sheet(output, sheet, "Sheet1");
  return { workbook: output, rows, rowCount: rows.length - 1 };
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

function maybeNumber(value) {
  if (typeof value !== "string") return value;
  const normalized = value.trim().replace(",", ".");
  if (!normalized) return value;
  const number = Number(normalized);
  if (!Number.isFinite(number)) return value;
  return Number.isInteger(number) ? Math.trunc(number) : number;
}

function groupFromName(name) {
  return String(name || "").slice(0, 3);
}

function formatDateSlash(date) {
  return [date.getDate(), date.getMonth() + 1, date.getFullYear()]
    .map((part) => String(part).padStart(2, "0"))
    .join("/");
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

function timestampForName() {
  const now = new Date();
  return [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "-",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
  ].join("");
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
