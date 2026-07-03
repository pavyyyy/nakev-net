"use strict";

const MACHINE_NAMES = {
  139: "GW 20",
  140: "GW 20",
  141: "GW 20",
  142: "GW 32",
  143: "GW 42",
  146: "GW 20",
  147: "GW 20",
  148: "GW 20",
  149: "CC 52",
  150: "JF 32",
  151: "JF 32",
  152: "JF 328",
  153: "JF 328",
  154: "JF 42",
  155: "GW 20",
  156: "GW 32",
  157: "JF 32",
  158: "JF 20",
  159: "JF 328",
  160: "TR GT26",
  161: "TR Nano",
};

const appState = {
  files: [],
  programs: [],
  inventory: null,
  lastClickedProgramKey: null,
  sortMode: "machine-program",
  singleProgram: { path1: null, path2: null },
};

const MACHINE_COLORS = ["#174f83", "#1f6f43", "#7a4c12", "#6c3f91", "#8f2f3f", "#156a72", "#5d6518", "#324a9a"];
const STORAGE_LOCATION_RE = /^(?:ш|sh)\s*\d+(?:[.,]\d+)?\s*[-–/]\s*\d+/i;

const els = {
  libraryStatus: document.querySelector("#libraryStatus"),
  folderInput: document.querySelector("#folderInput"),
  path1Input: document.querySelector("#path1Input"),
  path2Input: document.querySelector("#path2Input"),
  inventoryInput: document.querySelector("#inventoryInput"),
  darkModeToggle: document.querySelector("#darkModeToggle"),
  sortSelect: document.querySelector("#sortSelect"),
  folderDrop: document.querySelector("#folderDrop"),
  inventoryDrop: document.querySelector("#inventoryDrop"),
  programSearch: document.querySelector("#programSearch"),
  programList: document.querySelector("#programList"),
  programCount: document.querySelector("#programCount"),
  selectAllButton: document.querySelector("#selectAllButton"),
  clearButton: document.querySelector("#clearButton"),
  generateButton: document.querySelector("#generateButton"),
  summaryNote: document.querySelector("#summaryNote"),
  singleProgramNote: document.querySelector("#singleProgramNote"),
  logOutput: document.querySelector("#logOutput"),
};

window.addEventListener("load", () => {
  if (window.ExcelJS && window.JSZip) {
    els.libraryStatus.textContent = "Готово";
    els.libraryStatus.classList.add("ready");
  } else {
    els.libraryStatus.textContent = "Excel инструментите не се заредиха";
  }
});

els.folderInput.addEventListener("change", async (event) => {
  await loadMachineFiles([...event.target.files]);
  event.target.value = "";
});

els.path1Input.addEventListener("change", async (event) => {
  appState.singleProgram.path1 = event.target.files[0] || null;
  await loadSingleProgramFiles();
  event.target.value = "";
});

els.path2Input.addEventListener("change", async (event) => {
  appState.singleProgram.path2 = event.target.files[0] || null;
  await loadSingleProgramFiles();
  event.target.value = "";
});

els.inventoryInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (file) await loadInventory(file);
});

els.programSearch.addEventListener("input", renderProgramList);
els.sortSelect.addEventListener("change", () => {
  appState.sortMode = els.sortSelect.value;
  renderProgramList();
});
els.darkModeToggle.addEventListener("change", () => {
  document.documentElement.classList.toggle("dark", els.darkModeToggle.checked);
});
els.selectAllButton.addEventListener("click", selectVisiblePrograms);
els.clearButton.addEventListener("click", clearSelectedPrograms);
els.generateButton.addEventListener("click", generateSelectedPrograms);

setupDropzone(els.folderDrop, async (files) => {
  await loadMachineFiles(files);
});

setupDropzone(els.inventoryDrop, async (files) => {
  const workbookFile = files.find((file) => /\.(xlsx|xlsm|xls)$/i.test(file.name));
  if (workbookFile) await loadInventory(workbookFile);
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
    const files = await filesFromDrop(event.dataTransfer);
    await onFiles(files);
  });
}

async function filesFromDrop(dataTransfer) {
  const items = [...dataTransfer.items || []];
  const entries = items.map((item) => item.webkitGetAsEntry && item.webkitGetAsEntry()).filter(Boolean);
  if (!entries.length) return [...dataTransfer.files || []];

  const files = [];
  for (const entry of entries) {
    await collectEntryFiles(entry, "", files);
  }
  return files;
}

function collectEntryFiles(entry, prefix, files) {
  return new Promise((resolve, reject) => {
    if (entry.isFile) {
      entry.file((file) => {
        file.relativePath = `${prefix}${file.name}`;
        files.push(file);
        resolve();
      }, reject);
      return;
    }

    if (!entry.isDirectory) {
      resolve();
      return;
    }

    const reader = entry.createReader();
    const directoryPrefix = `${prefix}${entry.name}/`;
    const readBatch = () => {
      reader.readEntries(async (entries) => {
        if (!entries.length) {
          resolve();
          return;
        }
        for (const child of entries) {
          await collectEntryFiles(child, directoryPrefix, files);
        }
        readBatch();
      }, reject);
    };
    readBatch();
  });
}

async function loadMachineFiles(files) {
  const newEntries = files.map((file) => ({
    file,
    path: normalizePath(file.webkitRelativePath || file.relativePath || file.name),
    name: file.name,
  }));
  await mergeProgramFileEntries(newEntries, `Заредени ${newEntries.length} нови файла`);
}

async function loadSingleProgramFiles() {
  const { path1, path2 } = appState.singleProgram;
  appState.files = appState.files.filter((entry) => !entry.singleProgram);

  const baseName = sanitizePathPart((path1 || path2)?.name || "single-program");
  const newEntries = [];
  if (path1) newEntries.push({ file: path1, path: `single-program/${baseName}/PATH1/${path1.name}`, name: path1.name, singleProgram: true });
  if (path2) newEntries.push({ file: path2, path: `single-program/${baseName}/PATH2/${path2.name}`, name: path2.name, singleProgram: true });

  els.singleProgramNote.textContent = newEntries.length
    ? `Избрани: ${path1 ? `PATH1 ${path1.name}` : "PATH1 няма"} | ${path2 ? `PATH2 ${path2.name}` : "PATH2 няма"}`
    : "Няма избрана единична програма.";

  if (!newEntries.length) {
    appState.programs = await discoverPrograms(appState.files);
    renderProgramList();
    updateReadyState();
    return;
  }

  await mergeProgramFileEntries(newEntries, "Заредена единична програма");
}

async function mergeProgramFileEntries(newEntries, messagePrefix) {
  const merged = new Map(appState.files.map((entry) => [entry.path.toLowerCase(), entry]));
  for (const entry of newEntries) merged.set(entry.path.toLowerCase(), entry);
  appState.files = [...merged.values()];
  appState.programs = await discoverPrograms(appState.files);
  log(`${messagePrefix}, общо ${appState.files.length}. Намерени ${appState.programs.length} програми.`);
  renderProgramList();
  updateReadyState();
}

async function loadInventory(file) {
  log(`Четене на Excel справките: ${file.name}`);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  appState.inventory = new InventoryLookup(workbook);
  log(`Справките са заредени: ${appState.inventory.searchRows.length} реда с артикули, ${appState.inventory.inventoryByCode.size} складови позиции.`);
  updateReadyState();
}

async function discoverPrograms(fileEntries) {
  const byPath = new Map(fileEntries.map((entry) => [entry.path.toLowerCase(), entry]));
  const programs = [];

  for (const entry of fileEntries) {
    if (!/^O\d+$/i.test(entry.name)) continue;
    if (naturalProgramNumber(entry.name) >= 8000) continue;
    const pathNum = detectPathNumber(entry.path);
    if (pathNum !== "1") continue;

    const related = findRelatedProgramEntries(entry, byPath, fileEntries);
    const summary = await readProgramSummary([...related.entries()].sort((a, b) => Number(a[0]) - Number(b[0])).map((item) => item[1].file));
    const machine = inferMachineNumber(entry.path);
    programs.push({
      key: entry.path,
      code: entry.name.toUpperCase(),
      title: summary.title,
      date: summary.date,
      dateValue: summary.dateValue,
      machine,
      machineName: MACHINE_NAMES[machine] || "N/A",
      entries: related,
      selected: false,
    });
  }

  sortPrograms(programs);
  return programs;
}

function findRelatedProgramEntries(selectedEntry, byPath, fileEntries) {
  const selectedPath = selectedEntry.path;
  const paths = new Map([[detectPathNumber(selectedPath), selectedEntry]]);

  for (const wanted of ["1", "2", "3", "4"]) {
    if (paths.has(wanted)) continue;
    const candidatePath = replacePathNumber(selectedPath, wanted);
    const candidate = candidatePath ? byPath.get(candidatePath.toLowerCase()) : null;
    if (candidate) paths.set(wanted, candidate);
  }

  if (paths.has("1") && paths.has("2")) return paths;

  const ancestors = parentPaths(selectedPath);
  for (const ancestor of ancestors) {
    let sawPathFolder = false;
    for (const entry of fileEntries) {
      if (entry.name.toLowerCase() !== selectedEntry.name.toLowerCase()) continue;
      const relative = relativeToAncestor(entry.path, ancestor);
      if (!relative) continue;
      const parts = relative.split("/");
      const pathIndex = parts.findIndex((part) => /^PATH\d+$/i.test(part));
      if (pathIndex === -1 || pathIndex > 1) continue;
      sawPathFolder = true;
      const pathNum = parts[pathIndex].slice(4);
      if (!paths.has(pathNum)) paths.set(pathNum, entry);
    }
    if ((paths.has("1") && paths.has("2")) || sawPathFolder || isMachineRoot(ancestor)) break;
  }

  return paths;
}

async function readProgramSummary(files) {
  const summary = { title: "", date: "", dateValue: 0 };
  for (const file of files) {
    const text = await readTextFile(file);
    const lines = text.split(/\r?\n/).map((value) => value.trim()).filter(Boolean);
    if (!summary.title) {
      const programLine = lines.find((value) => /^O\d+/i.test(value));
      const match = programLine && programLine.match(/^O\d+\s*(?:\((.*?)\))?/i);
      summary.title = sanitizeText(match && match[1] ? match[1] : "");
    }
    if (!summary.date) {
      const date = findDateInText(lines);
      if (date) {
        summary.date = normalizeDate(date);
        summary.dateValue = dateSortValue(summary.date);
      }
    }
    if (summary.title && summary.date) break;
  }
  return summary;
}

function renderProgramList() {
  const query = els.programSearch.value.trim().toLowerCase();
  const sorted = sortPrograms([...appState.programs]);
  const visible = sorted.filter((program) => {
    const haystack = `${program.code} ${program.title} ${program.date} ${program.key} ${program.machine} ${program.machineName}`.toLowerCase();
    return !query || haystack.includes(query);
  });

  els.programCount.textContent = `${visible.length} показани / ${appState.programs.length} намерени`;
  els.summaryNote.textContent = appState.files.length
    ? `${appState.files.length} файла заредени. ${selectedPrograms().length} програми избрани.`
    : "Още няма заредена папка.";

  if (!appState.programs.length) {
    els.programList.className = "program-list empty";
    els.programList.innerHTML = '<div class="empty-state">Зареди папка на машина, за да видиш програмите.</div>';
    updateReadyState();
    return;
  }

  els.programList.className = "program-list";
  els.programList.innerHTML = "";
  const fragment = document.createDocumentFragment();
  for (const program of visible) {
    const row = document.createElement("label");
    row.className = "program-row";
    row.style.setProperty("--machine-color", machineColor(program.machine));
    const paths = [...program.entries.keys()].map((value) => `PATH${value}`).join(" + ");
    row.innerHTML = `
      <input type="checkbox" ${program.selected ? "checked" : ""} />
      <span class="program-code">${escapeHtml(program.code)}</span>
      <span class="program-machine">${escapeHtml(program.machine)} ${escapeHtml(program.machineName)}</span>
      <span class="program-title">${escapeHtml(program.title || program.key)}</span>
      <span class="program-meta">
        <span>${escapeHtml(paths)}${program.date ? ` | ${escapeHtml(program.date)}` : ""}</span>
        <span class="program-path">${escapeHtml(program.key)}</span>
      </span>
    `;
    row.querySelector("input").addEventListener("click", (event) => {
      applyProgramSelection(program, event.target.checked, event.shiftKey, visible);
      updateReadyState();
      renderProgramList();
    });
    fragment.appendChild(row);
  }
  els.programList.appendChild(fragment);
  updateReadyState();
}

function applyProgramSelection(program, checked, shiftKey, visiblePrograms) {
  if (shiftKey && appState.lastClickedProgramKey) {
    const from = visiblePrograms.findIndex((item) => item.key === appState.lastClickedProgramKey);
    const to = visiblePrograms.findIndex((item) => item.key === program.key);
    if (from >= 0 && to >= 0) {
      const start = Math.min(from, to);
      const end = Math.max(from, to);
      for (let index = start; index <= end; index += 1) visiblePrograms[index].selected = checked;
      appState.lastClickedProgramKey = program.key;
      return;
    }
  }
  program.selected = checked;
  appState.lastClickedProgramKey = program.key;
}

function selectVisiblePrograms() {
  const query = els.programSearch.value.trim().toLowerCase();
  for (const program of appState.programs) {
    const haystack = `${program.code} ${program.title} ${program.date} ${program.key} ${program.machine} ${program.machineName}`.toLowerCase();
    if (!query || haystack.includes(query)) program.selected = true;
  }
  renderProgramList();
}

function clearSelectedPrograms() {
  for (const program of appState.programs) program.selected = false;
  renderProgramList();
}

function selectedPrograms() {
  return appState.programs.filter((program) => program.selected);
}

function updateReadyState() {
  els.generateButton.disabled = !(appState.inventory && selectedPrograms().length && window.ExcelJS && window.JSZip);
}

async function generateSelectedPrograms() {
  const programs = selectedPrograms();
  if (!programs.length || !appState.inventory) return;

  els.generateButton.disabled = true;
  try {
    const generated = [];
    for (const program of programs) {
      log(`Генериране ${program.machine} ${program.machineName} ${program.code} ${program.title || ""}`.trim());
      const output = await generateDossierWorkbook(program, appState.inventory);
      generated.push(output);
    }

    if (generated.length === 1) {
      downloadBlob(generated[0].blob, generated[0].fileName);
    } else {
      const zip = new JSZip();
      for (const item of generated) zip.file(item.fileName, item.blob);
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, zipFileName(generated));
    }
    log(`Готово. Генерирани ${generated.length} Excel файла.`);
  } catch (error) {
    console.error(error);
    log(`ГРЕШКА: ${error.message || error}`);
  } finally {
    updateReadyState();
  }
}

async function generateDossierWorkbook(program, inventory) {
  const lines = [];
  for (const [pathNum, entry] of [...program.entries.entries()].sort((a, b) => Number(a[0]) - Number(b[0]))) {
    const text = await readTextFile(entry.file);
    for (const rawLine of text.split(/\r?\n/)) {
      const cleaned = cleanProgramLine(rawLine);
      if (cleaned) lines.push({ text: cleaned, pathNum, sourcePath: entry.path });
    }
  }

  const metadata = parseMetadata(program.entries.get("1")?.path || program.key, lines);
  const tools = parseTools(lines, inventory, metadata);
  const workbook = buildWorkbook(metadata, tools);
  const buffer = await workbook.xlsx.writeBuffer();
  const fileName = safeFileName(`${metadata.part} # ${metadata.programNumber} # ${machineFileLabel(metadata)}.xlsx`);
  return { fileName, metadata, blob: new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }) };
}

class InventoryLookup {
  constructor(workbook) {
    const sheet1 = workbook.getWorksheet("Sheet1") || workbook.worksheets[0];
    const sheet2 = workbook.getWorksheet("Sheet2") || workbook.worksheets[1];
    if (!sheet1 || !sheet2) throw new Error("Excel файлът трябва да има Sheet1 и Sheet2.");

    const sheet1Rows = worksheetValues(sheet1);
    const sheet2Rows = worksheetValues(sheet2);
    const header1 = sheet1Rows[0].map(cellToText);
    const header2 = sheet2Rows[0].map(cellToText);

    this.codeCol = findHeader(header1, "Код", 2);
    this.shortNameCol = findHeader(header1, "Кратко име", 3);
    this.barcodeCol = findHeader(header1, "Баркод", 5);
    const invCodeCol = findHeader(header2, "Код", 1);
    const invLocationCol = detectStorageLocationColumn(sheet2Rows, findHeader(header2, "Места", 2));

    this.searchRows = sheet1Rows;
    this.searchNames = sheet1Rows.map((row) => normalizeSearchName(row[this.shortNameCol]));
    this.searchNumbers = this.searchNames.map(extractNumbers);
    this.inventoryByCode = new Map();

    for (const row of sheet2Rows.slice(1)) {
      const code = cellToText(row[invCodeCol]);
      const location = cellToText(row[invLocationCol]);
      if (code && location && location !== "10101" && !this.inventoryByCode.has(code)) {
        this.inventoryByCode.set(code, splitStorageLocations(location));
      }
    }
  }

  matchTool(toolName) {
    const query = normalizeToolName(toolName);
    if (query.length < 6) return blankMatch();

    const queryNumbers = extractNumbers(query);
    let bestScore = 0;
    let bestIndex = -1;

    for (let index = 0; index < this.searchNames.length; index += 1) {
      const candidate = this.searchNames[index];
      if (!candidate || !numbersMatch(queryNumbers, this.searchNumbers[index])) continue;
      const score = tokenRatio(query, candidate);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    }

    if (bestIndex === -1 || bestScore < 65) return blankMatch(bestScore);

    const row = this.searchRows[bestIndex];
    const erpCode = cellToText(row[this.codeCol]);
    const shortName = cellToText(row[this.shortNameCol]);
    const barcode = cellToText(row[this.barcodeCol]);
    if (shortName.toUpperCase() === "BARCODE") return blankMatch(bestScore);

    return {
      erpCode: erpCode || " ",
      barcode: barcode || " ",
      shortName: shortName || " ",
      storage: this.inventoryByCode.get(erpCode) || " ",
      compare: shortName,
      score: bestScore,
    };
  }
}

function parseMetadata(programPath, lines) {
  const metadata = {
    machineNum: inferMachineNumber(programPath),
    machineName: "N/A",
    client: "N/A",
    programNumber: "N/A",
    part: "N/A",
    rev: "N/A",
    date: "N/A",
    material: "N/A",
    head1Spindles: "N/A",
    head2Spindle: "N/A",
    backupDrawingNumber: "N/A",
  };
  metadata.machineName = MACHINE_NAMES[metadata.machineNum] || "N/A";

  const programLine = lines.find((line) => /^O\d+/i.test(line.text));
  const programMatch = programLine && programLine.text.match(/^(O\d+)\s*(?:\((.*?)\))?/i);
  if (programMatch) {
    metadata.programNumber = programMatch[1].toUpperCase();
    if (programMatch[2]) {
      metadata.part = sanitizeText(programMatch[2]);
      const firstToken = metadata.part.split(" ")[0];
      if (firstToken && !/^\d+$/.test(firstToken)) metadata.client = firstToken;
    }
  }

  const lineCommentsByPath = new Map();
  const commentsByPath = new Map();
  for (const line of lines) {
    if (isOperationHeader(line.text)) continue;
    const comments = [...line.text.matchAll(/\(([^()]*)\)/g)].map((match) => sanitizeText(match[1])).filter(Boolean);
    if (!comments.length) continue;
    const lineComments = lineCommentsByPath.get(line.pathNum) || [];
    const lineIndex = lineComments.length;
    lineComments.push(sanitizeText(comments.join(" ")));
    lineCommentsByPath.set(line.pathNum, lineComments);

    const pathComments = commentsByPath.get(line.pathNum) || [];
    for (const comment of comments) pathComments.push({ lineIndex, comment });
    commentsByPath.set(line.pathNum, pathComments);
  }

  for (const [pathNum, comments] of commentsByPath.entries()) {
    for (const { lineIndex, comment } of comments) {
      const upper = comment.toUpperCase();
      if (isDate(comment)) {
        if (metadata.date === "N/A") metadata.date = normalizeDate(comment);
        continue;
      }
      if (upper.includes("REV") && metadata.rev === "N/A") metadata.rev = comment;
      else if (comment.length > 6 && [...comment].filter((char) => !/\d/.test(char)).length <= 3) metadata.backupDrawingNumber = comment;

      if (upper.startsWith("CANGA") || upper.startsWith("CANGI")) {
        if (pathNum === "1") {
          metadata.head1Spindles = comment;
          const material = findMaterialBeforeCanga((lineCommentsByPath.get(pathNum) || []).slice(0, lineIndex));
          if (material) metadata.material = material;
        } else if (pathNum === "2") {
          metadata.head2Spindle = comment;
        }
      }
    }
  }

  if (
    metadata.backupDrawingNumber !== "N/A" &&
    metadata.part !== "N/A" &&
    digitCount(metadata.backupDrawingNumber) > digitCount(metadata.part)
  ) {
    metadata.part += ` # ${metadata.backupDrawingNumber}`;
  }

  return metadata;
}

function parseTools(lines, inventory, metadata = null) {
  const parsed = [];
  const metadataNames = metadataToolNames(metadata);
  for (const [start, end] of collectOperationRanges(lines)) {
    const header = lines[start];
    const opNum = operationNumber(header.text);
    const opName = operationName(header.text);
    const operation = `SP${header.pathNum}-N${opNum}${opName ? ` ${opName}` : ""}`;
    let toolIndex = -1;
    let toolPosition = "N/A";

    for (let index = start + 1; index < end; index += 1) {
      const position = extractToolPosition(lines[index].text);
      if (position) {
        toolIndex = index;
        toolPosition = position;
        break;
      }
    }
    if (toolIndex === -1) continue;

    const tools = [];
    const comments = [];
    const geometry = [];

    for (let index = start + 1; index < toolIndex; index += 1) {
      const text = lines[index].text;
      const upper = text.toUpperCase();
      if (text.includes("&")) {
        comments.push(sanitizeText(text));
        continue;
      }
      if (text.startsWith("(")) {
        const content = commentText(text);
        const upperContent = content.toUpperCase();
        if (!content) continue;
        let classification = classifyToolText(content);
        if (isMetadataToolName(content, metadataNames)) classification = "comment";
        if (classification === "comment") comments.push(content);
        else if (classification === "tool") tools.push(content);
        continue;
      }
      const inlineComment = commentText(text);
      const classification = classifyToolText(inlineComment || text);
      if (classification === "comment") comments.push(inlineComment || sanitizeText(text));
    }

    let filteredTools = tools.filter((tool) => !isIgnoredToolName(tool) && !isMetadataToolName(tool, metadataNames));

    if (!filteredTools.length) {
      const toolLineComment = commentText(lines[toolIndex].text);
      const codeWithoutComments = sanitizeText(stripComments(lines[toolIndex].text));
      if (
        toolLineComment &&
        toolLineComment !== codeWithoutComments &&
        classifyToolText(toolLineComment) === "tool" &&
        !isIgnoredToolName(toolLineComment) &&
        !isMetadataToolName(toolLineComment, metadataNames)
      ) {
        filteredTools.push(toolLineComment);
      }
    }
    if (!filteredTools.length) filteredTools.push("");

    filteredTools.forEach((tool, index) => {
      const match = inventory.matchTool(tool);
      parsed.push({
        position: toolPosition,
        tool,
        operation: index === 0 ? operation : "",
        geometry: "",
        erpCode: match.erpCode,
        barcode: match.barcode,
        shortName: match.shortName,
        storage: match.storage,
        comment: index === 0 ? comments.join(", ") : "",
        compare: match.compare,
      });
    });
  }

  parsed.sort((a, b) => toolSortNumber(a.position) - toolSortNumber(b.position) || a.tool.localeCompare(b.tool));
  return mergeDuplicateTools(parsed);
}

function buildWorkbook(metadata, tools) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Dossier Detaili Web";
  const machineLabel = machineFileLabel(metadata);
  workbook.title = safeFileName(`${metadata.part} # ${metadata.programNumber}`);
  workbook.subject = `Machine ${metadata.machineNum}`;
  workbook.category = machineLabel;
  workbook.keywords = `machine ${metadata.machineNum}; ${machineLabel}; ${metadata.machineName}; ${metadata.programNumber}`;
  workbook.description = `Machine ${metadata.machineNum} / ${metadata.machineName} / ${metadata.programNumber}`;
  workbook.calcProperties.fullCalcOnLoad = true;
  const workbookDate = parseWorkbookDate(metadata.date);
  if (workbookDate) {
    workbook.created = workbookDate;
    workbook.modified = workbookDate;
  }
  const sheet = workbook.addWorksheet("Sheet", {
    pageSetup: {
      paperSize: 9,
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.25, right: 0.25, top: 0.295, bottom: 0.295, header: 0.3, footer: 0.3 },
      printArea: "A1:G1048576",
    },
  });

  sheet.getCell("A1").value = "Чертеж";
  sheet.getCell("B1").value = metadata.part;
  sheet.getCell("A2").value = "Ревизия";
  sheet.getCell("B2").value = metadata.rev;
  sheet.getCell("A3").value = "Дата";
  sheet.getCell("B3").value = metadata.date;
  sheet.getCell("A4").value = "Материал";
  sheet.getCell("B4").value = metadata.material;
  sheet.getCell("C2").value = "Машина";
  sheet.getCell("D2").value = /^\d+$/.test(metadata.machineNum) ? Number(metadata.machineNum) : metadata.machineNum;
  sheet.getCell("C3").value = "Модел";
  sheet.getCell("D3").value = metadata.machineName;
  sheet.getCell("C4").value = "Програма";
  sheet.getCell("D4").value = metadata.programNumber;

  for (let row = 1; row <= 4; row += 1) {
    for (let col = 1; col <= 4; col += 1) {
      sheet.getCell(row, col).border = metadataBorder();
    }
  }
  ["A1", "A2", "A3", "A4", "C2", "C3", "C4"].forEach((cell) => {
    sheet.getCell(cell).font = { bold: true };
  });
  sheet.getCell("B4").alignment = { wrapText: true };

  const headers = ["Позиция", "Име", "ERP Код", "Баркод", "Barcode", "Склад", "Коментар", "от МАШИНА", "от ERP"];
  headers.forEach((header, index) => {
    const cell = sheet.getCell(6, index + 1);
    cell.value = header;
    if (index < 7) {
      cell.border = thickBorder();
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    }
  });

  const rows = addSpindleRows(metadata, tools);
  let rowIndex = 7;
  let previousNumber = null;
  for (const tool of rows) {
    const currentNumber = positionNumber(tool.position);
    if (previousNumber !== null && currentNumber !== null && previousNumber < 20 && currentNumber >= 20) {
      for (let col = 1; col <= 7; col += 1) {
        const cell = sheet.getCell(rowIndex, col);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } };
        cell.font = { size: 2 };
      }
      sheet.getRow(rowIndex).height = 4;
      rowIndex += 1;
    }

    writeToolRow(sheet, rowIndex, tool);
    if (currentNumber !== null) previousNumber = currentNumber;
    rowIndex += 1;
  }

  stripeToolRows(sheet, 7, rowIndex - 1);
  applyLayout(sheet);
  return workbook;
}

function writeToolRow(sheet, rowIndex, tool) {
  const displayName = tool.compare && tool.compare.length > 6 ? tool.compare : tool.tool;
  const comments = [tool.operation, tool.comment, tool.geometry].filter(Boolean).join("\n").replaceAll("28H0.", "G97M");
  const values = [
    tool.position,
    displayName,
    formatLongCode(tool.erpCode),
    formatLongCode(tool.barcode),
    { formula: `SUBSTITUTE("*"&D${rowIndex}&"*", CHAR(10), "")` },
    tool.storage,
    comments,
    tool.tool,
    tool.compare,
  ];

  values.forEach((value, index) => {
    const cell = sheet.getCell(rowIndex, index + 1);
    cell.value = value;
    if (index < 7) {
      cell.border = thinRowBorder();
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    }
  });
  sheet.getCell(rowIndex, 5).font = { name: "IDAutomationHC39M Free Version", size: 10 };
}

function addSpindleRows(metadata, tools) {
  const spindles = [
    ["PRPOD", metadata.head1Spindles],
    ["SP1", metadata.head1Spindles],
    ["LUN", metadata.head1Spindles],
    ["SP2", metadata.head2Spindle],
  ].map(([position, value]) => ({
    position,
    tool: spindleToolName(value),
    operation: "",
    geometry: "",
    erpCode: " ",
    barcode: " ",
    shortName: " ",
    storage: " ",
    comment: "",
    compare: "",
  }));
  return [...spindles, ...tools];
}

function applyLayout(sheet) {
  const widths = [9.5, 32, 12, 10, 31, 9, 28, 28, 32];
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
}

function stripeToolRows(sheet, startRow, endRow) {
  let alternate = false;
  let previousPosition = null;
  for (let row = startRow; row <= endRow; row += 1) {
    const rowValues = [];
    for (let col = 1; col <= 7; col += 1) rowValues.push(sheet.getCell(row, col).value);
    if (!rowValues.some(Boolean)) continue;
    const position = sheet.getCell(row, 1).value;
    if (position !== previousPosition) {
      alternate = !alternate;
      previousPosition = position;
    }
    const argb = alternate ? "FFD3D3D3" : "FFFFFFFF";
    for (let col = 1; col <= 7; col += 1) {
      sheet.getCell(row, col).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
    }
  }
}

function metadataBorder() {
  return {
    left: { style: "thin" },
    right: { style: "thin" },
    top: { style: "thick" },
    bottom: { style: "thick" },
  };
}

function thickBorder() {
  return {
    left: { style: "thick" },
    right: { style: "thick" },
    top: { style: "thick" },
    bottom: { style: "thick" },
  };
}

function thinRowBorder() {
  return {
    left: { style: "thin" },
    right: { style: "thin" },
    top: { style: "thick" },
    bottom: { style: "thick" },
  };
}

function worksheetValues(sheet) {
  const rows = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    const values = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      values[colNumber - 1] = cell.value && typeof cell.value === "object" && "result" in cell.value ? cell.value.result : cell.value;
    });
    rows.push(values);
  });
  return rows;
}

function collectOperationRanges(lines) {
  const starts = lines.map((line, index) => (isOperationHeader(line.text) ? index : -1)).filter((index) => index >= 0);
  return starts.map((start, index) => [start, index + 1 < starts.length ? starts[index + 1] : lines.length]);
}

function isOperationHeader(line) {
  const match = line.match(/^N(\d+)(?=\s|\(|$)/i);
  if (!match) return false;
  const number = Number(match[1]);
  return ![0, 98, 99, 1000].includes(number) && !line.toUpperCase().includes("ZAHVASHTANE");
}

function operationNumber(line) {
  const match = line.match(/^N(\d+)/i);
  return match ? match[1] : "";
}

function operationName(line) {
  const comments = [...line.matchAll(/\(([^()]*)\)/g)].map((match) => sanitizeText(match[1])).filter(Boolean);
  return comments.join(" ");
}

function cleanProgramLine(rawLine) {
  let line = rawLine.trim();
  if (!line || line.startsWith("/N")) return "";
  if (line.startsWith("/")) line = line.slice(1).trimStart();
  return line;
}

function extractToolPosition(line) {
  const codeOnly = stripComments(line);
  const matches = [...codeOnly.matchAll(/T\s*(\d{1,4})(?!\d)/gi)];
  const positions = matches.map((match) => toolNumberFromDigits(match[1])).filter(Boolean).map((number) => `T${number}`);
  return positions.length ? positions[positions.length - 1] : null;
}

function toolNumberFromDigits(digits) {
  if (!digits || Number(digits) === 0) return null;
  if (digits.length >= 4) return Number(digits.slice(0, 2));
  if (digits.length === 3) {
    if (digits.endsWith("00")) return Number(digits[0]);
    if (digits.startsWith("0")) return Number(digits.slice(1));
    if (digits.endsWith("0")) return Number(digits.slice(0, 2));
    return Number(digits.slice(0, 2));
  }
  return Number(digits);
}

function mergeDuplicateTools(tools) {
  const merged = [];
  const byKey = new Map();
  for (const tool of tools) {
    const key = `${tool.position}\u0000${tool.tool}`;
    if (!byKey.has(key)) {
      byKey.set(key, tool);
      merged.push(tool);
      continue;
    }
    const existing = byKey.get(key);
    for (const value of [tool.operation, tool.comment, tool.geometry]) {
      if (value && !existing.operation.includes(value)) existing.operation = [existing.operation, value].filter(Boolean).join("\n");
    }
  }
  return merged;
}

function findMaterialBeforeCanga(previousComments) {
  for (let index = previousComments.length - 1; index >= 0; index -= 1) {
    const comment = previousComments[index];
    const upper = comment.toUpperCase();
    if (!comment || isDate(comment) || upper.startsWith("HEAD")) continue;
    if (upper.includes("REV") || upper.includes("PROBA") || upper.includes("MOSTRI")) continue;
    if (digitCount(comment) > 8 && letterCount(comment) < 3) continue;
    return comment;
  }
  return "";
}

function normalizePath(path) {
  return path.replaceAll("\\", "/").replace(/^\/+/, "");
}

function detectPathNumber(path) {
  const part = normalizePath(path).split("/").find((value) => /^PATH\d+$/i.test(value));
  return part ? part.slice(4) : "1";
}

function replacePathNumber(path, wanted) {
  const parts = normalizePath(path).split("/");
  const index = parts.findIndex((part) => /^PATH\d+$/i.test(part));
  if (index === -1) return null;
  parts[index] = `PATH${wanted}`;
  return parts.join("/");
}

function parentPaths(path) {
  const parts = normalizePath(path).split("/");
  const parents = [];
  for (let index = parts.length - 2; index >= 0; index -= 1) {
    parents.push(parts.slice(0, index + 1).join("/"));
  }
  return parents;
}

function relativeToAncestor(path, ancestor) {
  const normalized = normalizePath(path);
  if (!ancestor) return normalized;
  const prefix = `${ancestor}/`;
  return normalized.startsWith(prefix) ? normalized.slice(prefix.length) : "";
}

function isMachineRoot(path) {
  const name = normalizePath(path).split("/").pop() || "";
  return /^\d{3}$/.test(name) || /^CNC[_ -]?\d{3}(?:\D|$)/i.test(name);
}

function inferMachineNumber(path) {
  const parts = normalizePath(path).split("/");
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    const cncMatch = parts[index].match(/^CNC[_ -]?(\d{3})(?:\D|$)/i);
    if (cncMatch) return cncMatch[1];
  }
  for (let index = 0; index < parts.length; index += 1) {
    if (/^PATH\d+$/i.test(parts[index]) && index > 0) {
      const parent = parts[index - 1];
      const match = parent.match(/\b(\d{3})\b/);
      if (match) return match[1];
    }
  }
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    if (/^\d{3}$/.test(parts[index])) return parts[index];
  }
  return "N/A";
}

function naturalProgramNumber(code) {
  const match = code.match(/\d+/);
  return match ? Number(match[0]) : Number.MAX_SAFE_INTEGER;
}

function sanitizePathPart(value) {
  return String(value || "single-program").replace(/[<>:"/\\|?*]/g, "_").trim() || "single-program";
}

function sortPrograms(programs) {
  const byProgram = (a, b) => naturalProgramNumber(a.code) - naturalProgramNumber(b.code) || a.code.localeCompare(b.code);
  const byMachine = (a, b) => machineSortNumber(a.machine) - machineSortNumber(b.machine) || a.machineName.localeCompare(b.machineName);
  const sorters = {
    "machine-program": (a, b) => byMachine(a, b) || byProgram(a, b) || a.key.localeCompare(b.key),
    program: (a, b) => byProgram(a, b) || byMachine(a, b) || a.key.localeCompare(b.key),
    name: (a, b) => (a.title || "").localeCompare(b.title || "") || byMachine(a, b) || byProgram(a, b),
    machine: (a, b) => byMachine(a, b) || (a.title || "").localeCompare(b.title || "") || byProgram(a, b),
    date: (a, b) => (b.dateValue || 0) - (a.dateValue || 0) || byMachine(a, b) || byProgram(a, b),
    path: (a, b) => a.key.localeCompare(b.key),
  };
  programs.sort(sorters[appState.sortMode] || sorters["machine-program"]);
  return programs;
}

function machineSortNumber(machine) {
  return /^\d+$/.test(String(machine)) ? Number(machine) : Number.MAX_SAFE_INTEGER;
}

function machineColor(machine) {
  const num = machineSortNumber(machine);
  const index = Number.isFinite(num) ? num % MACHINE_COLORS.length : 0;
  return MACHINE_COLORS[index];
}

async function readTextFile(file) {
  const buffer = await file.arrayBuffer();
  for (const encoding of ["utf-8", "windows-1251", "iso-8859-1"]) {
    try {
      return new TextDecoder(encoding, { fatal: encoding === "utf-8" }).decode(buffer);
    } catch (error) {
      continue;
    }
  }
  return new TextDecoder().decode(buffer);
}

function commentText(line) {
  const comments = [...line.matchAll(/\(([^()]*)\)/g)].map((match) => sanitizeText(match[1]));
  if (comments.length) return comments.join(" ");
  if (line.startsWith("(") && line.endsWith(")")) return sanitizeText(line.slice(1, -1));
  return sanitizeText(line);
}

function sanitizeText(text) {
  return String(text || "").replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
}

function stripComments(line) {
  return line.replace(/\([^()]*\)/g, "");
}

function isDate(text) {
  return /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(String(text).trim());
}

function normalizeDate(text) {
  return String(text).trim().replaceAll("-", ".").replaceAll("/", ".");
}

function findDateInText(lines) {
  for (const line of lines) {
    const comments = [...line.matchAll(/\(([^()]*)\)/g)].map((match) => sanitizeText(match[1]));
    for (const comment of comments) {
      if (isDate(comment)) return comment;
    }
    if (isDate(line)) return line;
  }
  return "";
}

function dateSortValue(value) {
  const normalized = normalizeDate(value);
  const match = normalized.match(/^(\d{1,2})[.](\d{1,2})[.](\d{2,4})$/);
  if (!match) return 0;
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  const month = Number(match[2]);
  const day = Number(match[1]);
  return year * 10000 + month * 100 + day;
}

function digitCount(text) {
  return (String(text).match(/\d/g) || []).length;
}

function letterCount(text) {
  return (String(text).match(/[A-Za-zА-Яа-я]/g) || []).length;
}

function cellToText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object") {
    if ("text" in value) return String(value.text || "").trim();
    if ("result" in value) return cellToText(value.result);
    if ("richText" in value) return value.richText.map((part) => part.text).join("").trim();
  }
  if (typeof value === "number" && Number.isInteger(value)) return String(value);
  return String(value).trim();
}

function findHeader(header, name, fallback) {
  const index = header.indexOf(name);
  return index >= 0 ? index : fallback;
}

function detectStorageLocationColumn(rows, fallback) {
  const scores = new Map();
  for (const row of rows.slice(0, 1000)) {
    row.forEach((value, index) => {
      if (isStorageLocation(cellToText(value))) scores.set(index, (scores.get(index) || 0) + 1);
    });
  }

  let bestIndex = fallback;
  let bestScore = 0;
  for (const [index, score] of scores.entries()) {
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  }
  return bestScore >= 3 ? bestIndex : fallback;
}

function isStorageLocation(value) {
  return STORAGE_LOCATION_RE.test(String(value || "").trim());
}

function normalizeSearchName(text) {
  let cleaned = cellToText(text).replace(/[\u0400-\u04FF]+/g, "");
  cleaned = cleaned.replace(/\b\/?ISCAR\b/gi, "");
  return cleaned.replace(/\s+/g, " ").trim();
}

function normalizeToolName(text) {
  return String(text || "").replaceAll("/", " ").replace(/\s+/g, " ").trim();
}

function extractNumbers(text) {
  return String(text || "").match(/\d+\.?\d*/g) || [];
}

function numbersMatch(searchNumbers, candidateNumbers) {
  if (!searchNumbers.length) return true;
  const matches = searchNumbers.filter((number) => candidateNumbers.includes(number)).length;
  return matches / searchNumbers.length >= 2 / 3;
}

function tokenRatio(first, second) {
  const direct = similarity(String(first).toLowerCase(), String(second).toLowerCase());
  const tokenized = similarity(sortedTokens(first), sortedTokens(second));
  return Math.max(direct, tokenized) * 100;
}

function sortedTokens(text) {
  return String(text || "").toLowerCase().split(/\s+/).filter(Boolean).sort().join(" ");
}

function similarity(first, second) {
  if (first === second) return 1;
  if (!first || !second) return 0;
  const previous = Array(second.length + 1).fill(0);
  const current = Array(second.length + 1).fill(0);
  for (let j = 0; j <= second.length; j += 1) previous[j] = j;
  for (let i = 1; i <= first.length; i += 1) {
    current[0] = i;
    for (let j = 1; j <= second.length; j += 1) {
      const cost = first[i - 1] === second[j - 1] ? 0 : 1;
      current[j] = Math.min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + cost);
    }
    for (let j = 0; j <= second.length; j += 1) previous[j] = current[j];
  }
  const distance = previous[second.length];
  return 1 - distance / Math.max(first.length, second.length);
}

function blankMatch(score = 0) {
  return { erpCode: " ", barcode: " ", shortName: " ", storage: " ", compare: "", score };
}

function isIgnoredToolName(toolName) {
  return /^[MT]\d{1,4}$/i.test(normalizeToolName(toolName));
}

function metadataToolNames(metadata) {
  if (!metadata) return new Set();
  const values = [
    metadata.client,
    metadata.programNumber,
    metadata.part,
    metadata.rev,
    metadata.date,
    metadata.material,
    metadata.head1Spindles,
    metadata.head2Spindle,
    metadata.backupDrawingNumber,
    spindleToolName(metadata.head1Spindles),
    spindleToolName(metadata.head2Spindle),
  ];
  return new Set(values.filter((value) => value && value !== "N/A").map(normalizeForCompare));
}

function normalizeForCompare(value) {
  return normalizeToolName(value).toUpperCase();
}

function isMetadataToolName(toolName, metadataNames) {
  const normalized = normalizeForCompare(toolName);
  return Boolean(normalized && metadataNames.has(normalized));
}

function classifyToolText(text) {
  const value = normalizeToolName(text);
  const upper = value.toUpperCase();
  if (!value) return "ignore";
  if (isDate(value)) return "ignore";
  if (/^[MT]\d{1,4}$/i.test(value)) return "ignore";
  if (/^HEAD\s*\d+$/i.test(value)) return "ignore";
  if (/^G28H-?\d+(?:[.,]\d+)?$/i.test(value)) return "ignore";
  if (/^G0C-?\d+(?:[.,]\d+)?$/i.test(value)) return "ignore";
  if (upper.includes("NEVARTYASHTO")) return "comment";
  if (/\bNOVA\b/i.test(value)) return "comment";
  if (upper.includes("VARTYASHT")) return "comment";
  if (upper.includes("GEOM") || upper.includes("ZNAK")) return "comment";
  if (/^(CANGA|CANGI)\b/i.test(value)) return "comment";
  if (upper.includes("PLAST") || upper.includes("VARHA") || upper.includes("OFSSET")) return "comment";
  if (/^G50[XYZ]/i.test(value) || /\/G50[XYZ]/i.test(value)) return "comment";
  if (/^\d+(?:[,.]\d+)?\s*MM\s+OT\b/i.test(value)) return "comment";
  if (/^(W|VC|S|FZ|L\d+)\s*=/i.test(value)) return "comment";
  if (/\b(W|VC|S|FZ|L\d+)\s*=/i.test(value)) return "comment";
  if (/\bOB\/MIN\b/i.test(value)) return "comment";
  if (/\bS\s*=\s*[-\d,.]+\s*[\/ ]\s*F\s*=/i.test(value)) return "comment";
  return "tool";
}

function splitStorageLocations(value) {
  return String(value).split(",").map((part) => part.trim()).filter(Boolean).join("\n");
}

function spindleToolName(value) {
  if (!value || value === "N/A") return "";
  const parts = String(value).split(/\s+/);
  return parts.length > 1 ? parts.slice(1).join(" ") : String(value);
}

function formatLongCode(value) {
  const text = cellToText(value) || " ";
  if (text.length <= 10) return text;
  const middle = Math.floor(text.length / 2);
  return `${text.slice(0, middle)}\n${text.slice(middle)}`;
}

function toolSortNumber(position) {
  const match = String(position || "").match(/^T(\d+)/);
  return match ? Number(match[1]) : 9999;
}

function positionNumber(position) {
  const match = String(position || "").match(/^T(\d+)/);
  return match ? Number(match[1]) : null;
}

function machineFileLabel(metadata) {
  const machineType = String(metadata.machineName || "").split(/\s+/)[0] || "MACHINE";
  return `${machineType} ${metadata.machineNum}`;
}

function parseWorkbookDate(value) {
  const match = String(value || "").match(/^(\d{1,2})[.](\d{1,2})[.](\d{2,4})$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]) - 1;
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  const date = new Date(year, month, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

function safeFileName(name) {
  return String(name).replace(/[<>:"/\\|?*]/g, "_").trim() || "dossier.xlsx";
}

function timestampForName() {
  const now = new Date();
  return [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "_",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
  ].join("");
}

function zipFileName(generated) {
  const labels = [...new Set(generated.map((item) => machineZipLabel(item.metadata)))];
  const machineLabel = labels.length === 1 ? labels[0] : "MULTI MACHINE";
  return safeFileName(`${machineLabel} # ${timestampForName()}.zip`);
}

function machineZipLabel(metadata) {
  const machineType = String(metadata.machineName || "").split(/\s+/)[0] || "MACHINE";
  return `${metadata.machineNum} ${machineType}`;
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

function escapeHtml(text) {
  return String(text || "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  }[char]));
}

function log(message) {
  const time = new Date().toLocaleTimeString();
  els.logOutput.textContent += `\n[${time}] ${message}`;
  els.logOutput.scrollTop = els.logOutput.scrollHeight;
}
