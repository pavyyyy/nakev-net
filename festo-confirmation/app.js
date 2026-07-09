"use strict";

pdfjsLib.GlobalWorkerOptions.workerSrc = "../assets/vendor/pdf.worker.min.js";

const { PDFDocument, StandardFonts, rgb } = PDFLib;

const FROM_ADDRESS_LINES = [
  "EAZ OOD",
  "Treti Mart 42",
  "4225 Perushtitsa",
];

const FOOTER_LINES = [
  "Katya Kusheva",
  "Deputy production manager",
  "T: + 359 32 654 115 | F: + 359 32 654 100",
  "M: +359 888 829 316 | e: k.kusheva@eaz-bg.com | http://www.eaz-bg.com",
];

const DELIVERY_ADDRESS_EXCLUDES = [
  "EAZ OOD",
  "Appedix 1",
  "Appendix 1",
  "Treti Mart 42",
  "4225 Perushtitsa",
];

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  runButton: document.querySelector("#runButton"),
  clearButton: document.querySelector("#clearButton"),
  status: document.querySelector("#status"),
  previewWrap: document.querySelector("#previewWrap"),
  previewBody: document.querySelector("#previewBody"),
  log: document.querySelector("#log"),
};

let selectedItems = [];
const ORDER_LABELS = ["Purchase order", "Pilot order"];

els.fileInput.addEventListener("change", async () => loadFiles([...els.fileInput.files]));
els.runButton.addEventListener("click", generate);
els.clearButton.addEventListener("click", clear);
setupDropzone(els.dropzone, async (files) => loadFiles(files.filter((file) => /\.pdf$/i.test(file.name))));

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

async function loadFiles(files) {
  selectedItems = [];
  els.runButton.disabled = true;
  els.clearButton.disabled = !files.length;
  els.previewBody.innerHTML = "";
  els.previewWrap.hidden = true;
  setLog("");

  if (!files.length) {
    setStatus("Ready.");
    return;
  }

  setStatus(`Reading ${files.length} PDF file(s)...`);
  for (const file of files) {
    try {
      const layout = await extractPdfLayout(file);
      const data = extractConfirmationData(layout);
      selectedItems.push({ file, data });
      log(`${file.name}: parsed ${data.order_type.toLowerCase()} ${data.order_number}`);
    } catch (error) {
      selectedItems.push({ file, error });
      log(`${file.name}: ${error.message || error}`);
    }
  }

  renderPreview();
  const okCount = selectedItems.filter((item) => item.data).length;
  els.runButton.disabled = okCount === 0;
  setStatus(okCount ? `Parsed ${okCount} of ${files.length} file(s).` : "No orders could be parsed.", okCount ? "ok" : "error");
}

function clear() {
  selectedItems = [];
  els.fileInput.value = "";
  els.runButton.disabled = true;
  els.clearButton.disabled = true;
  els.previewBody.innerHTML = "";
  els.previewWrap.hidden = true;
  setLog("");
  setStatus("Ready.");
}

async function generate() {
  const validItems = selectedItems.filter((item) => item.data);
  if (!validItems.length) return;

  els.runButton.disabled = true;
  try {
    const outputs = [];
    const templateBytes = await fetch("PO CONFIRM page template.pdf").then((response) => {
      if (!response.ok) throw new Error("Could not load PDF template.");
      return response.arrayBuffer();
    });

    for (const item of validItems) {
      const bytes = await buildConfirmationPdf(item.data, templateBytes);
      outputs.push({ fileName: outputFileName(item.data), bytes });
    }

    for (const item of outputs) {
      downloadBlob(new Blob([item.bytes], { type: "application/pdf" }), item.fileName);
      await delay(250);
    }

    setStatus(`Generated ${outputs.length} PDF file(s).`, "ok");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    els.runButton.disabled = !selectedItems.some((item) => item.data);
  }
}

async function extractPdfLayout(file) {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const text = await page.getTextContent();
    pages.push(buildPageLayout(text.items, pageNumber));
  }
  return {
    sourceName: file.name,
    pages,
    lines: pages.flatMap((page) => page.lines.map((line) => line.text)),
  };
}

function buildPageLayout(items, pageNumber) {
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
    line.parts.push({ x, text, width: item.width || 0 });
  });

  const lines = grouped
    .sort((a, b) => b.y - a.y)
    .map((line) => {
      const parts = line.parts.sort((a, b) => a.x - b.x);
      return {
        pageNumber,
        y: line.y,
        parts,
        text: normalizeSpaces(parts.map((part) => part.text).join(" ")),
      };
    })
    .filter((line) => line.text);

  return { pageNumber, lines };
}

function extractConfirmationData(layout) {
  const lines = layout.lines;
  const { orderType, orderNumber } = detectOrderInfo(layout);
  if (!orderNumber) throw new Error(`Could not parse ${orderType.toLowerCase()}.`);

  const supplierCurrencyIndex = findIndexContains(lines, "Supplier number Customer number Currency");
  let supplierNumber = "";
  lines.slice(supplierCurrencyIndex + 1, supplierCurrencyIndex + 6).some((candidate) => {
    const match = candidate.match(/^\s*(\d+)\b/);
    if (match) supplierNumber = match[1];
    return Boolean(supplierNumber);
  });
  if (!supplierNumber) throw new Error("Could not parse supplier number.");

  let deliveryAddressLines = [];
  try {
    deliveryAddressLines = extractDeliveryAddressLines(layout.pages[0]);
  } catch {
    deliveryAddressLines = [];
  }

  let orderDate = "";
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].trim().toLowerCase() !== "date") continue;
    const found = lines.slice(i + 1, i + 4).find((candidate) => /^\d{2}\.\d{2}\.\d{4}$/.test(candidate.trim()));
    if (found) {
      orderDate = found.trim();
      break;
    }
  }
  if (!orderDate) throw new Error("Could not parse order date.");

  const partNumberHeader = findIndexContains(lines, "Part number");
  const [itemLineIndex, itemMatch] = findRegex(
    lines,
    /^\s*(.+?)\s+([A-Z0-9][A-Z0-9\-./:]+)\s+([\d.,]+)\s+items\s*$/i,
    partNumberHeader,
  );
  const partName = normalizeSpaces(itemMatch[1]);
  const partNumber = itemMatch[2].trim();
  const orderQty = itemMatch[3].trim();

  let toleranceUnder = "";
  let toleranceOver = "";
  const toleranceSource = lines.find((line) => /underdelivery:/i.test(line) && /overdelivery:/i.test(line));
  if (toleranceSource) {
    const toleranceMatch = toleranceSource.match(/underdelivery:\s*([\d.,]+)\s*%\s*\/\s*overdelivery:\s*([\d.,]+)\s*%/i);
    if (toleranceMatch) {
      toleranceUnder = toleranceMatch[1];
      toleranceOver = toleranceMatch[2];
    }
  }

  const materialStartPatterns = [
    "Material code",
    "Materials designation",
    "Materials number",
    "Materials festo standard",
    "Coating",
    "RoHS compliant",
    "Sensitivity code",
    "The following documents belong to this item:",
    "Drawing no.",
  ];
  let materialStartIndex = null;
  for (let i = itemLineIndex + 1; i < lines.length; i += 1) {
    if (materialStartPatterns.some((pattern) => lines[i].startsWith(pattern))) {
      materialStartIndex = i;
      break;
    }
  }

  const deliverySearchStart = materialStartIndex ?? itemLineIndex;
  const [, deliveryMatch] = findRegex(lines, /Delivery date\s+([\d.]+)/, deliverySearchStart);
  const [, priceMatch] = findRegex(lines, /Price per\s+(\d+)\s+PC\s+([\d.,]+)\s+([\d.,]+)/i, deliverySearchStart);

  const materialLines = [];
  if (materialStartIndex !== null) {
    const deliveryDateIndex = findIndexContains(lines, "Delivery date", materialStartIndex);
    const allowedPrefixes = materialStartPatterns.filter((item) => item !== "The following documents belong to this item:");
    lines.slice(materialStartIndex, deliveryDateIndex).forEach((rawLine) => {
      const clean = normalizeSpaces(rawLine);
      if (clean && allowedPrefixes.some((prefix) => clean.startsWith(prefix))) materialLines.push(clean);
    });
  }

  return {
    order_type: orderType,
    order_number: orderNumber,
    supplier_number: supplierNumber,
    part_number: partNumber,
    part_name: partName,
    delivery_tolerance_under: toleranceUnder,
    delivery_tolerance_over: toleranceOver,
    material_description_lines: materialLines,
    order_date: orderDate,
    delivery_date: deliveryMatch[1],
    order_qty: orderQty,
    price_unit_count: priceMatch[1],
    price_per_piece: priceMatch[2],
    price_total: priceMatch[3],
    delivery_address: deliveryAddressLines.join("\n"),
  };
}

async function buildConfirmationPdf(data, templateBytes) {
  const output = await PDFDocument.create();
  const template = await PDFDocument.load(templateBytes);
  const [templatePage] = await output.copyPages(template, [0]);
  const page = output.addPage(templatePage);
  const width = page.getWidth();
  const height = page.getHeight();

  const font = await output.embedFont(StandardFonts.Helvetica);
  const bold = await output.embedFont(StandardFonts.HelveticaBold);
  const italic = await output.embedFont(StandardFonts.HelveticaOblique);

  const left = mm(28);
  const rightBlock = mm(108);
  let y = height - mm(50);

  drawLine(page, "From:", left, y, 13, font);
  drawLine(page, "Deliver To:", rightBlock, y, 13, font);
  y -= 22;
  drawMultiline(page, FROM_ADDRESS_LINES, left, y, 12.5, font, 16);
  drawMultiline(page, splitLines(data.delivery_address), rightBlock, y, 12.5, font, 16);
  y -= 50;
  drawLine(page, `supplier number: ${data.supplier_number}`, left, y, 12.5, font);

  y -= 40;
  const intro = `We would like to confirm your ${data.order_type.toLowerCase()}:`;
  const poWidth = bold.widthOfTextAtSize(data.order_number, 13.5);
  const introWidth = font.widthOfTextAtSize(intro, 13.5);
  const introX = (width - introWidth - poWidth - 10) / 2;
  drawLine(page, intro, introX, y, 13.5, font);
  page.drawRectangle({ x: introX + introWidth + 8, y: y - 2, width: poWidth + 5, height: 16, color: rgb(1, 0.95, 0) });
  drawLine(page, data.order_number, introX + introWidth + 10, y, 13.5, bold);

  y -= 42;
  drawCentered(page, `P/N ${data.part_number}   -   ${data.part_name}`, y, 12.5, font);

  y -= 48;
  drawLine(page, `Price per ${data.price_unit_count} piece(s): ${data.price_per_piece}`, left, y, 11.5, font);
  drawLine(page, `QTY: ${data.order_qty}`, left + mm(58), y, 11.5, font);
  drawLine(page, `Price total: ${data.price_total} EUR`, left + mm(100), y, 11.5, font);
  y -= 22;
  drawLine(page, `Order Date: ${data.order_date}`, left, y, 11.5, font);
  drawLine(page, `Delivery date: ${data.delivery_date}`, left + mm(100), y, 11.5, font);

  y -= 42;
  if (data.delivery_tolerance_under || data.delivery_tolerance_over) {
    drawLine(page, "Delivery tolerance:", left, y, 11.5, bold);
    y -= 18;
    if (data.delivery_tolerance_under) {
      drawLine(page, `- underdelivery: ${data.delivery_tolerance_under} %`, left, y, 11.2, font);
      y -= 16;
    }
    if (data.delivery_tolerance_over) {
      drawLine(page, `+ overdelivery: ${data.delivery_tolerance_over} %`, left, y, 11.2, font);
      y -= 16;
    }
    y -= 10;
  }

  if (data.material_description_lines.length) {
    drawLine(page, "Material Descriptions:", left, y, 11.5, bold);
    y -= 17;
    for (const line of data.material_description_lines) {
      const wrapped = wrapText(formatMaterialLine(line), font, 11.2, width - left - mm(28));
      wrapped.forEach((part) => {
        drawLine(page, part, left, y, 11.2, line.startsWith("RoHS compliant") ? italic : font);
        y -= 14.5;
      });
    }
  }

  drawMultiline(page, FOOTER_LINES, left, mm(48), 8.7, font, 11);
  return output.save();
}

function drawLine(page, text, x, y, size, font) {
  page.drawText(toPdfText(text), { x, y, size, font, color: rgb(0, 0, 0) });
}

function detectOrderInfo(layout) {
  const lines = layout.lines || [];
  const joinedText = normalizeSpaces(lines.join(" "));
  const lowered = joinedText.toLowerCase();
  const sourceName = String(layout.sourceName || "").toLowerCase();
  let orderType = sourceName.includes("pilot order")
    ? "Pilot order"
    : sourceName.includes("purchase order")
      ? "Purchase order"
      : ORDER_LABELS.find((label) => lowered.includes(label.toLowerCase()));
  if (!orderType && lowered.includes("pilot")) orderType = "Pilot order";
  if (!orderType && lowered.includes("purchase")) orderType = "Purchase order";

  const numberPattern = /\b[A-Z]{2,}\/\d{6,}\b/;
  const orderNumberMatch = joinedText.match(numberPattern);
  let orderNumber = orderNumberMatch ? orderNumberMatch[0] : "";

  if (!orderType) {
    try {
      const found = findOrderHeader(lines);
      orderType = found.label;
      if (!orderNumber) orderNumber = readOrderNumber(lines, found.index, found.label);
    } catch {
      if (sourceName.includes("pilot order")) orderType = "Pilot order";
      else orderType = "Purchase order";
    }
  }

  return { orderType, orderNumber };
}

function drawCentered(page, text, y, size, font) {
  const safeText = toPdfText(text);
  const textWidth = font.widthOfTextAtSize(safeText, size);
  drawLine(page, safeText, (page.getWidth() - textWidth) / 2, y, size, font);
}

function drawMultiline(page, lines, x, y, size, font, leading) {
  lines.filter(Boolean).forEach((line, index) => {
    drawLine(page, line, x, y - index * leading, size, font);
  });
}

function wrapText(text, font, size, maxWidth) {
  const words = toPdfText(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(next, size) <= maxWidth || !current) {
      current = next;
    } else {
      lines.push(current);
      current = word;
    }
  });
  if (current) lines.push(current);
  return lines;
}

function formatMaterialLine(line) {
  return toPdfText(normalizeSpaces(String(line || "").replace(/,$/, "")));
}

function renderPreview() {
  els.previewWrap.hidden = !selectedItems.length;
  els.previewBody.innerHTML = selectedItems.map((item) => {
    if (item.error) {
      return `<tr><td>${escapeHtml(item.file.name)}</td><td colspan="3">${escapeHtml(item.error.message || item.error)}</td></tr>`;
    }
    return `<tr><td>${escapeHtml(item.file.name)}</td><td>${escapeHtml(item.data.order_number)}</td><td>${escapeHtml(item.data.part_number)}</td><td>${escapeHtml(item.data.delivery_date)}</td></tr>`;
  }).join("");
}

function findOrderHeader(lines) {
  for (const label of ORDER_LABELS) {
    try {
      return { label, index: findIndexContains(lines, label) };
    } catch {
      // Try the next supported order label.
    }
  }
  throw new Error(`Could not find text: ${ORDER_LABELS.join(" or ")}`);
}

function readOrderNumber(lines, orderTypeIndex, orderType) {
  const sameLine = normalizeSpaces(lines[orderTypeIndex] || "");
  const nearby = lines
    .slice(Math.max(0, orderTypeIndex), Math.min(lines.length, orderTypeIndex + 8))
    .map((line) => normalizeSpaces(line))
    .filter(Boolean);
  const numberPattern = /\b[A-Z]{2,}\/\d{6,}\b/;

  const sameLineNumber = sameLine.match(numberPattern);
  if (sameLineNumber) return sameLineNumber[0];

  const inlineMatch = sameLine.match(new RegExp(`^${escapeRegex(orderType)}\\s+(.+)$`, "i"));
  if (inlineMatch) {
    const inlineNumber = inlineMatch[1].match(numberPattern);
    if (inlineNumber) return inlineNumber[0];
  }

  for (const candidate of nearby) {
    const match = candidate.match(numberPattern);
    if (match) return match[0];
  }

  for (let i = orderTypeIndex + 1; i < Math.min(lines.length, orderTypeIndex + 5); i += 1) {
    const candidate = normalizeSpaces(lines[i]);
    if (!candidate || ORDER_LABELS.includes(candidate) || candidate.toLowerCase() === "date") continue;
    const candidateNumber = candidate.match(numberPattern);
    if (candidateNumber) return candidateNumber[0];
  }

  return "";
}

function findIndexContains(lines, text, start = 0) {
  for (let i = start; i < lines.length; i += 1) {
    if (lines[i].includes(text)) return i;
  }
  throw new Error(`Could not find text: ${text}`);
}

function findRegex(lines, regex, start = 0) {
  for (let i = start; i < lines.length; i += 1) {
    const match = lines[i].match(regex);
    if (match) return [i, match];
  }
  throw new Error(`Could not find pattern: ${regex}`);
}

function normalizeSpaces(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function escapeRegex(text) {
  return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function splitLines(text) {
  return String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
}

function cleanDeliveryAddressLine(line) {
  let cleaned = normalizeSpaces(line);
  for (const exclude of DELIVERY_ADDRESS_EXCLUDES) {
    const pattern = new RegExp(escapeRegex(exclude), "gi");
    cleaned = cleaned.replace(pattern, " ");
  }
  cleaned = normalizeSpaces(cleaned);
  if (DELIVERY_ADDRESS_EXCLUDES.some((exclude) => cleaned.toLowerCase().includes(exclude.toLowerCase()))) {
    return "";
  }
  return cleaned;
}

function extractDeliveryAddressLines(page) {
  if (!page) return [];
  const labelIndex = page.lines.findIndex((line) => line.text.includes("Please deliver to:"));
  if (labelIndex === -1) throw new Error("Could not find delivery address.");

  const labelLine = page.lines[labelIndex];
  const labelPart = labelLine.parts.find((part) => /Please/i.test(part.text)) || labelLine.parts[0];
  const rightColumnX = labelPart ? labelPart.x - 2 : 250;
  const stopPatterns = ["For all questions", "Supplier number", "Telephone number", "Faxnumber", "Your contact partner"];
  const deliveryLines = [];

  for (let i = labelIndex + 1; i < page.lines.length; i += 1) {
    const line = page.lines[i];
    if (stopPatterns.some((pattern) => line.text.includes(pattern))) break;
    const rightParts = line.parts.filter((part) => part.x >= rightColumnX);
    if (!rightParts.length) continue;
    const cleaned = cleanDeliveryAddressLine(rightParts.map((part) => part.text).join(" "));
    if (cleaned) deliveryLines.push(cleaned);
  }

  return [...new Set(deliveryLines)];
}

function toPdfText(text) {
  return transliterateCyrillic(normalizeSpaces(String(text || "")));
}

function transliterateCyrillic(text) {
  const map = {
    А: "A", а: "a",
    Б: "B", б: "b",
    В: "V", в: "v",
    Г: "G", г: "g",
    Д: "D", д: "d",
    Е: "E", е: "e",
    Ж: "Zh", ж: "zh",
    З: "Z", з: "z",
    И: "I", и: "i",
    Й: "Y", й: "y",
    К: "K", к: "k",
    Л: "L", л: "l",
    М: "M", м: "m",
    Н: "N", н: "n",
    О: "O", о: "o",
    П: "P", п: "p",
    Р: "R", р: "r",
    С: "S", с: "s",
    Т: "T", т: "t",
    У: "U", у: "u",
    Ф: "F", ф: "f",
    Х: "H", х: "h",
    Ц: "Ts", ц: "ts",
    Ч: "Ch", ч: "ch",
    Ш: "Sh", ш: "sh",
    Щ: "Sht", щ: "sht",
    Ъ: "A", ъ: "a",
    Ь: "Y", ь: "y",
    Ю: "Yu", ю: "yu",
    Я: "Ya", я: "ya",
  };

  return Array.from(String(text || ""), (char) => map[char] ?? char).join("");
}

function mm(value) {
  return value * 72 / 25.4;
}

function outputFileName(data) {
  const suffix = String(data.order_number || "confirmation").split("/").pop();
  return safeFileName(`${data.order_type} Confirmation ${suffix}.pdf`);
}

function safeFileName(name) {
  return String(name).replace(/[<>:"/\\|?*]/g, "_").trim() || "confirmation.pdf";
}

function setStatus(message, kind = "") {
  els.status.textContent = message;
  els.status.className = `status ${kind}`.trim();
}

function log(message) {
  els.log.hidden = false;
  els.log.textContent += `${message}\n`;
}

function setLog(message) {
  els.log.textContent = message;
  els.log.hidden = !message;
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

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
