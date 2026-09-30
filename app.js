// ==============================================================================
// Excel-Prüfer Web Controller (100% Client-Side für Netlify & Browser)
// ==============================================================================

document.addEventListener("DOMContentLoaded", () => {
  // App-Zustand im Arbeitsspeicher
  const state = {
    refWorkbook: null,
    targetWorkbook: null,
    refFileName: "",
    targetFileName: "",
    currentRefSheet: "",
    currentTargetSheet: "",
    mapping: [],
    allResults: [],
    currentFilter: "errors",
    searchQuery: "",
    appliedCorrections: {}, // cellKey ("row_col") -> newVal
    lastTargetBuffer: null,
    staffList: [],
    staffMap: {},
    timesheetData: null
  };

  // DOM Elemente
  const dropRef = document.getElementById("drop-ref");
  const inputRef = document.getElementById("input-ref");
  const refFileInfo = document.getElementById("ref-file-info");
  const refFilename = document.getElementById("ref-filename");
  const selectRefSheet = document.getElementById("select-ref-sheet");
  const refStorageBadge = document.getElementById("ref-storage-badge");
  const refInfoDetail = document.getElementById("ref-info-detail");
  const btnChangeRef = document.getElementById("btn-change-ref");
  const btnClearRef = document.getElementById("btn-clear-ref");

  const dropTgt = document.getElementById("drop-tgt");
  const inputTgt = document.getElementById("input-tgt");
  const tgtFileInfo = document.getElementById("tgt-file-info");
  const tgtFilename = document.getElementById("tgt-filename");
  const selectTgtSheet = document.getElementById("select-tgt-sheet");
  const tgtStatusBadge = document.getElementById("tgt-status-badge");
  const btnChangeTgt = document.getElementById("btn-change-tgt");

  // Mitarbeiter & Personal DOM
  const dropStaff = document.getElementById("drop-staff");
  const inputStaff = document.getElementById("input-staff");
  const staffFileInfo = document.getElementById("staff-file-info");
  const staffInfoTitle = document.getElementById("staff-info-title");
  const staffInfoDetail = document.getElementById("staff-info-detail");
  const staffDropZoneContent = document.getElementById("staff-drop-zone-content");
  const staffStorageBadge = document.getElementById("staff-storage-badge");
  const btnChangeStaff = document.getElementById("btn-change-staff");
  const btnClearStaff = document.getElementById("btn-clear-staff");
  const btnDownloadStaffTemplate = document.getElementById("btn-download-staff-template");

  // Arbeitszeiten / Timesheet DOM
  const timesheetSection = document.getElementById("timesheet-section");
  const selectTsDateCol = document.getElementById("select-ts-date-col");
  const selectTsHoursCol = document.getElementById("select-ts-hours-col");
  const selectTsResourceCol = document.getElementById("select-ts-resource-col");
  const chkTsOnlyStaff = document.getElementById("chk-ts-only-staff");
  const tsStaffStatusBadge = document.getElementById("ts-staff-status-badge");
  const timesheetThead = document.getElementById("timesheet-thead");
  const timesheetTbody = document.getElementById("timesheet-tbody");
  const timesheetTfoot = document.getElementById("timesheet-tfoot");
  const btnExportTimesheet = document.getElementById("btn-export-timesheet");

  const btnLoadDemo = document.getElementById("btn-load-demo");
  const configSection = document.getElementById("config-section");
  const mappingTbody = document.getElementById("mapping-tbody");
  const checkAllCols = document.getElementById("check-all-cols");
  const btnStartCheck = document.getElementById("btn-start-check");

  const resultsSection = document.getElementById("results-section");
  const resultsTbody = document.getElementById("results-tbody");
  const tableSearchInput = document.getElementById("table-search-input");
  const filterPills = document.getElementById("filter-pills");

  const btnGeneralApply = document.getElementById("btn-general-apply");
  const btnFixZahlendreher = document.getElementById("btn-fix-zahlendreher");
  const btnResetFixes = document.getElementById("btn-reset-fixes");
  const btnExportExcel = document.getElementById("btn-export-excel");
  const btnExportReport = document.getElementById("btn-export-report");
  const toast = document.getElementById("toast");

  // Ansichten-Umschalter (Tabs)
  const tabBtnDiagnostics = document.getElementById("tab-btn-diagnostics");
  const tabBtnFulltable = document.getElementById("tab-btn-fulltable");
  const viewDiagnostics = document.getElementById("view-diagnostics");
  const viewFulltable = document.getElementById("view-fulltable");
  const tabBadgeErrors = document.getElementById("tab-badge-errors");
  const tabBadgeTotalRows = document.getElementById("tab-badge-total-rows");

  // Gesamte Excel-Datei Filter & Tabelle
  const filterFullStaff = document.getElementById("filter-full-staff");
  const filterFullDate = document.getElementById("filter-full-date");
  const filterFullStatus = document.getElementById("filter-full-status");
  const filterFullSearch = document.getElementById("filter-full-search");
  const fulltableCountsText = document.getElementById("fulltable-counts-text");
  const btnResetFullFilters = document.getElementById("btn-reset-full-filters");
  const fulltableThead = document.getElementById("fulltable-thead");
  const fulltableTbody = document.getElementById("fulltable-tbody");

  // --- IndexedDB Speicher für die Referenzdatei ---
  const IDB_NAME = "ExcelPrueferStorage";
  const IDB_VERSION = 1;
  const IDB_STORE = "app_data";
  const IDB_KEY_REF = "saved_reference_file";

  function openAppDB() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        reject(new Error("IndexedDB wird von diesem Browser nicht unterstützt."));
        return;
      }
      const req = window.indexedDB.open(IDB_NAME, IDB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(IDB_STORE)) {
          db.createObjectStore(IDB_STORE);
        }
      };
      req.onsuccess = (e) => resolve(e.target.result);
      req.onerror = (e) => reject(e.target.error || new Error("Konnte IndexedDB nicht öffnen"));
    });
  }

  async function saveRefFileToStorage(fileName, buffer, selectedSheet) {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const data = {
          fileName: fileName,
          buffer: buffer,
          selectedSheet: selectedSheet || "",
          savedAt: new Date().toISOString()
        };
        const req = store.put(data, IDB_KEY_REF);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Speichern der Referenzdatei in IndexedDB:", err);
      return false;
    }
  }

  async function updateRefSheetInStorage(selectedSheet) {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const getReq = store.get(IDB_KEY_REF);
        getReq.onsuccess = () => {
          if (getReq.result) {
            const item = getReq.result;
            item.selectedSheet = selectedSheet;
            store.put(item, IDB_KEY_REF);
          }
          resolve(true);
        };
        getReq.onerror = () => resolve(false);
      });
    } catch (err) {
      console.warn("Konnte gewähltes Arbeitsblatt nicht im Speicher aktualisieren:", err);
    }
  }

  async function loadRefFileFromStorage() {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readonly");
        const store = tx.objectStore(IDB_STORE);
        const req = store.get(IDB_KEY_REF);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Laden der Referenzdatei aus IndexedDB:", err);
      return null;
    }
  }

  async function deleteRefFileFromStorage() {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const req = store.delete(IDB_KEY_REF);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Löschen der Referenzdatei aus IndexedDB:", err);
      return false;
    }
  }

  async function restoreSavedReferenceFile() {
    try {
      const saved = await loadRefFileFromStorage();
      if (saved && saved.buffer) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(saved.buffer);
        state.refWorkbook = wb;
        state.refFileName = saved.fileName || "Referenz_Stammdaten.xlsx";

        if (refFilename) refFilename.textContent = state.refFileName;
        populateSheetSelect(selectRefSheet, wb.worksheets);

        if (saved.selectedSheet && wb.worksheets.some(ws => ws.name === saved.selectedSheet)) {
          selectRefSheet.value = saved.selectedSheet;
        }
        state.currentRefSheet = selectRefSheet.value;

        if (refFileInfo) refFileInfo.classList.remove("hidden");
        const dropContent = dropRef ? dropRef.querySelector(".drop-zone-content") : null;
        if (dropContent) dropContent.classList.add("hidden");

        if (refStorageBadge) {
          refStorageBadge.textContent = "Im Browser gespeichert";
          refStorageBadge.className = "badge badge-ok";
        }
        if (refInfoDetail) {
          refInfoDetail.textContent = "💾 Im Browser gespeichert (bleibt dauerhaft erhalten)";
        }

        checkReadyForConfig();
        console.log(`Referenzdatei "${state.refFileName}" aus dem Browser-Speicher wiederhergestellt.`);
      }
    } catch (err) {
      console.warn("Konnte gespeicherte Referenzdatei nicht laden:", err);
    }
  }

  // --- Drag & Drop Einrichten ---
  setupDropZone(dropRef, inputRef, async (file) => {
    try {
      state.refFileName = file.name;
      const arrayBuffer = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      state.refWorkbook = wb;

      refFilename.textContent = file.name;
      populateSheetSelect(selectRefSheet, wb.worksheets);
      state.currentRefSheet = selectRefSheet.value;
      refFileInfo.classList.remove("hidden");
      dropRef.querySelector(".drop-zone-content").classList.add("hidden");

      if (refStorageBadge) {
        refStorageBadge.textContent = "Im Browser gespeichert";
        refStorageBadge.className = "badge badge-ok";
      }
      if (refInfoDetail) {
        refInfoDetail.textContent = "💾 Im Browser gespeichert (bleibt dauerhaft erhalten)";
      }

      await saveRefFileToStorage(file.name, arrayBuffer, state.currentRefSheet);

      checkReadyForConfig();
      showToast(`Referenzdatei "${file.name}" geladen und im Browser gespeichert.`);
    } catch (err) {
      console.error(err);
      alert("Fehler beim Laden der Referenzdatei: " + err.message);
    }
  });

  setupDropZone(dropTgt, inputTgt, async (file) => {
    try {
      state.targetFileName = file.name;
      const arrayBuffer = await file.arrayBuffer();
      state.lastTargetBuffer = arrayBuffer;
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      state.targetWorkbook = wb;

      tgtFilename.textContent = file.name;
      populateSheetSelect(selectTgtSheet, wb.worksheets);
      state.currentTargetSheet = selectTgtSheet.value;
      tgtFileInfo.classList.remove("hidden");
      dropTgt.querySelector(".drop-zone-content").classList.add("hidden");

      if (tgtStatusBadge) {
        tgtStatusBadge.textContent = "Bereit zur Prüfung";
        tgtStatusBadge.className = "badge badge-ok";
        tgtStatusBadge.style.display = "inline-block";
      }

      checkReadyForConfig();
      showToast(`Prüfdatei "${file.name}" geladen.`);
    } catch (err) {
      console.error(err);
      alert("Fehler beim Laden der Prüfdatei: " + err.message);
    }
  });

  // --- Mitarbeiter-Stammdaten: Browser-Speicher & Dropzone ---
  const STORAGE_KEY_STAFF = "excel_pruefer_mitarbeiter_v1";

  function loadStaffFromStorage() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_STAFF);
      if (stored) {
        const list = JSON.parse(stored);
        if (Array.isArray(list) && list.length > 0) {
          applyStaffList(list);
          return true;
        }
      }
    } catch (e) {
      console.warn("Fehler beim Laden der Mitarbeiter aus localStorage:", e);
    }
    updateStaffUI();
    return false;
  }

  // --- Mitarbeiter-Stammdaten: Hilfsfunktionen, Registrierung & Normalisierung ---
  function registerStaffEntry(rawRes, name) {
    if (!rawRes || !name) return;
    const clean = String(rawRes).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
    if (!clean) return;

    // 1. Exakter String & Kleinschreibung
    state.staffMap[clean] = name;
    state.staffMap[clean.toLowerCase()] = name;

    // 2. Ohne float-Endung .0 oder ,0 (Excel)
    const noFloat = clean.replace(/[,.]0+$/, "");
    state.staffMap[noFloat] = name;
    state.staffMap[noFloat.toLowerCase()] = name;

    // 3. Wenn es Ziffern sind (z. B. "45" oder "0045"):
    // Unpadded und alle gängigen Padding-Längen (1- bis 10-stellig) mappen
    if (/^\d+$/.test(noFloat)) {
      const numVal = parseInt(noFloat, 10);
      const unpadded = String(numVal);
      state.staffMap[unpadded] = name;
      for (let len = 1; len <= 10; len++) {
        const padded = unpadded.padStart(len, "0");
        state.staffMap[padded] = name;
      }
    } else {
      const digitsMatch = noFloat.match(/\b\d+\b/);
      if (digitsMatch) {
        const numVal = parseInt(digitsMatch[0], 10);
        const unpadded = String(numVal);
        state.staffMap[unpadded] = name;
        for (let len = 1; len <= 10; len++) {
          const padded = unpadded.padStart(len, "0");
          state.staffMap[padded] = name;
        }
      }
    }
  }

  function getStaffName(rawRes) {
    if (rawRes === null || rawRes === undefined) return null;
    const s = String(rawRes).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
    if (!s) return null;

    // 1. Direkter Treffer
    if (state.staffMap[s]) return state.staffMap[s];

    // 2. Ohne .0 oder ,0 Float-Endung
    const noFloat = s.replace(/[,.]0+$/, "");
    if (state.staffMap[noFloat]) return state.staffMap[noFloat];

    // 3. Numerischer Abgleich (unpadded und padded 1 bis 10 Stellen)
    if (/^\d+$/.test(noFloat)) {
      const numVal = parseInt(noFloat, 10);
      const unpadded = String(numVal);
      if (state.staffMap[unpadded]) return state.staffMap[unpadded];
      for (let len = 1; len <= 10; len++) {
        const p = unpadded.padStart(len, "0");
        if (state.staffMap[p]) return state.staffMap[p];
      }
    }

    // 4. Case-insensitive
    if (state.staffMap[s.toLowerCase()]) return state.staffMap[s.toLowerCase()];
    if (state.staffMap[noFloat.toLowerCase()]) return state.staffMap[noFloat.toLowerCase()];

    // 5. Ziffern-Extraktion falls Präfix (z. B. "Nr. 45" oder "MA-0045")
    const digitsMatch = noFloat.match(/\b\d+\b/);
    if (digitsMatch) {
      const numVal = parseInt(digitsMatch[0], 10);
      const unpadded = String(numVal);
      if (state.staffMap[unpadded]) return state.staffMap[unpadded];
      for (let len = 1; len <= 10; len++) {
        const p = unpadded.padStart(len, "0");
        if (state.staffMap[p]) return state.staffMap[p];
      }
    }

    return null;
  }

  function applyStaffList(list) {
    state.staffList = list;
    state.staffMap = {};
    list.forEach(item => {
      const rawRes = String(item.resource || item.res || "").replace(/\u00a0/g, " ").trim();
      const name = String(item.name || item.mitarbeiter || "").replace(/\u00a0/g, " ").trim();
      if (rawRes && name) {
        registerStaffEntry(rawRes, name);
      }
    });
    updateStaffUI();
  }

  function saveStaffToStorage(list) {
    try {
      localStorage.setItem(STORAGE_KEY_STAFF, JSON.stringify(list));
      applyStaffList(list);
      if (state.timesheetData) {
        renderTimesheetMatrix();
      }
    } catch (e) {
      console.error("Fehler beim Speichern in localStorage:", e);
    }
  }

  function clearStaffStorage() {
    if (!confirm("Möchten Sie die gespeicherten Mitarbeiter wirklich aus dem Browser löschen?")) return;
    try {
      localStorage.removeItem(STORAGE_KEY_STAFF);
    } catch (e) {}
    state.staffList = [];
    state.staffMap = {};
    updateStaffUI();
    if (state.timesheetData) {
      renderTimesheetMatrix();
    }
    showToast("Mitarbeiterliste aus dem Browser gelöscht.");
  }

  function updateStaffUI() {
    if (!staffFileInfo || !staffDropZoneContent) return;
    const count = state.staffList.length;
    if (count > 0) {
      staffDropZoneContent.classList.add("hidden");
      staffFileInfo.classList.remove("hidden");
      if (staffInfoTitle) staffInfoTitle.textContent = `${count} Mitarbeiter aktiv`;
      if (staffInfoDetail) staffInfoDetail.textContent = `Im Browser gespeichert (Ressourcen verknüpft)`;
      if (staffStorageBadge) {
        staffStorageBadge.textContent = `${count} Gespeichert`;
        staffStorageBadge.className = "badge badge-ok";
      }
    } else {
      staffDropZoneContent.classList.remove("hidden");
      staffFileInfo.classList.add("hidden");
      if (staffStorageBadge) {
        staffStorageBadge.textContent = "Optional";
        staffStorageBadge.className = "badge badge-leer-ok";
      }
    }
  }

  if (dropStaff && inputStaff) {
    setupDropZone(dropStaff, inputStaff, async (file) => {
      try {
        showToast(`Lese Mitarbeiterdatei "${file.name}"...`);
        const arrayBuffer = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(arrayBuffer);
        const ws = wb.worksheets.find(s => s && (s.rowCount > 0 || s.actualRowCount > 0)) || wb.worksheets[0];
        if (!ws) {
          alert("Die Arbeitsmappe enthält keine Tabellenblätter.");
          return;
        }

        const parsedList = [];
        const seenResKeys = new Set();

        const addRow = (rawA, rawB, rawC, rowNum) => {
          let resVal = String(rawA || "").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
          resVal = resVal.replace(/[,.]0+$/, "");
          if (!resVal) return;

          let nameVal = String(rawB || "").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();

          // Falls Spalte A und B vertauscht sein sollten (A ist Name, B ist Nummer):
          if (!/^\d+$/.test(resVal) && /^\d+$/.test(nameVal)) {
            const tmp = resVal;
            resVal = nameVal;
            nameVal = tmp;
          }

          if (!nameVal) {
            nameVal = `Mitarbeiter ${resVal}`;
          }

          // Zeile 1 nur überspringen, wenn es EINDEUTIG eine Kopfzeile ist (keine Ziffern in A und typische Headerwörter in A und B):
          if (rowNum === 1) {
            const aLow = resVal.toLowerCase();
            const bLow = nameVal.toLowerCase();
            const isHeadA = ["ressource", "resource", "personal", "mitarbeiter-nr", "personalnummer", "mitarbeiternr", "persnr", "pers-nr", "nummer", "nr"].some(k => aLow === k);
            const isHeadB = ["name", "mitarbeiter", "person", "vorname", "nachname", "bezeichnung"].some(k => bLow === k);
            if (!/^\d+$/.test(resVal) && isHeadA && isHeadB) {
              return; // Kopfzeile überspringen
            }
          }

          const dKey = `${resVal}___${nameVal.toLowerCase()}`;
          if (!seenResKeys.has(dKey)) {
            seenResKeys.add(dKey);
            parsedList.push({
              resource: resVal,
              name: nameVal,
              dept: ""
            });
          }
        };

        // 1. Alle Zeilen via eachRow erfassen
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          let cellA = WebExcelEngine.extractCellValue(row.getCell(1).value, row.getCell(1));
          if (!cellA && row.getCell(1).text) cellA = row.getCell(1).text;
          if (!cellA && row.getCell(1).model && row.getCell(1).model.value) cellA = row.getCell(1).model.value;

          let cellB = WebExcelEngine.extractCellValue(row.getCell(2).value, row.getCell(2));
          if (!cellB && row.getCell(2).text) cellB = row.getCell(2).text;
          if (!cellB && row.getCell(2).model && row.getCell(2).model.value) cellB = row.getCell(2).model.value;

          let cellC = WebExcelEngine.extractCellValue(row.getCell(3).value, row.getCell(3));
          if (!cellC && row.getCell(3).text) cellC = row.getCell(3).text;

          addRow(cellA, cellB, cellC, rowNumber);
        });

        // 2. Zur Sicherheit auch direkt 1 bis rowCount durchlaufen (falls Zeilen in eachRow ausgelassen wurden)
        const totalRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
        for (let r = 1; r <= totalRows; r++) {
          const row = ws.getRow(r);
          if (!row) continue;
          let cellA = WebExcelEngine.extractCellValue(row.getCell(1).value, row.getCell(1));
          if (!cellA && row.getCell(1).text) cellA = row.getCell(1).text;
          if (!cellA && row.getCell(1).model && row.getCell(1).model.value) cellA = row.getCell(1).model.value;

          let cellB = WebExcelEngine.extractCellValue(row.getCell(2).value, row.getCell(2));
          if (!cellB && row.getCell(2).text) cellB = row.getCell(2).text;
          if (!cellB && row.getCell(2).model && row.getCell(2).model.value) cellB = row.getCell(2).model.value;

          let cellC = WebExcelEngine.extractCellValue(row.getCell(3).value, row.getCell(3));
          if (!cellC && row.getCell(3).text) cellC = row.getCell(3).text;

          addRow(cellA, cellB, cellC, r);
        }

        if (parsedList.length === 0) {
          alert("In der Datei wurden keine Mitarbeiter gefunden. Bitte stellen Sie sicher, dass in Spalte A die Nummern und in Spalte B die Namen stehen.");
          return;
        }

        saveStaffToStorage(parsedList);
        showToast(`✅ ${parsedList.length} Mitarbeiter geladen und dauerhaft im Browser gespeichert!`);

        // Falls Ziel-Datei bereits geladen ist: Spaltenauswahl & Arbeitszeiten & Tabellen sofort neu synchronisieren
        if (state.targetWorkbook && state.currentTargetSheet) {
          const wsTarget = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
          if (wsTarget) {
            const tgtHeaders = [];
            const headerRow = wsTarget.getRow(1);
            let colCount = wsTarget.columnCount;
            headerRow.eachCell((cell, colNum) => { if (colNum > colCount) colCount = colNum; });
            for (let c = 1; c <= colCount; c++) {
              const hCell = headerRow.getCell(c);
              const val = WebExcelEngine.extractCellValue(hCell.value, hCell).trim();
              tgtHeaders.push({
                colNum: c,
                letter: getColLetter(c),
                name: val || `Spalte ${getColLetter(c)}`
              });
            }
            populateTimesheetColSelects(tgtHeaders);
          }
          if (typeof renderTimesheetMatrix === "function") renderTimesheetMatrix();
          if (typeof populateFullTableFilters === "function") populateFullTableFilters();
          if (typeof renderFullExcelTable === "function") renderFullExcelTable();
        }
      } catch (err) {
        console.error(err);
        alert("Fehler beim Lesen der Mitarbeiterdatei: " + err.message);
      }
    });
  }

  if (btnChangeStaff && inputStaff) {
    btnChangeStaff.addEventListener("click", (e) => {
      e.stopPropagation();
      inputStaff.click();
    });
  }

  if (btnClearStaff) {
    btnClearStaff.addEventListener("click", (e) => {
      e.stopPropagation();
      clearStaffStorage();
    });
  }

  if (btnChangeRef && inputRef) {
    btnChangeRef.addEventListener("click", (e) => {
      e.stopPropagation();
      inputRef.click();
    });
  }

  if (btnClearRef) {
    btnClearRef.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!confirm("Möchten Sie die gespeicherte Referenzdatei wirklich aus dem Browser löschen?")) return;
      await deleteRefFileFromStorage();
      state.refWorkbook = null;
      state.refFileName = "";
      state.currentRefSheet = "";
      if (refFilename) refFilename.textContent = "";
      if (selectRefSheet) selectRefSheet.innerHTML = "";
      if (refFileInfo) refFileInfo.classList.add("hidden");
      const dropContent = dropRef ? dropRef.querySelector(".drop-zone-content") : null;
      if (dropContent) dropContent.classList.remove("hidden");
      if (refStorageBadge) {
        refStorageBadge.textContent = "Nicht geladen";
        refStorageBadge.className = "badge badge-secondary";
      }
      if (configSection) configSection.classList.add("hidden");
      showToast("Referenzdatei aus dem Browser gelöscht.");
    });
  }

  if (btnChangeTgt && inputTgt) {
    btnChangeTgt.addEventListener("click", (e) => {
      e.stopPropagation();
      inputTgt.click();
    });
  }

  if (btnDownloadStaffTemplate) {
    btnDownloadStaffTemplate.addEventListener("click", async (e) => {
      e.stopPropagation();
      try {
        const buf = await generateStaffTemplateExcel();
        downloadBuffer(buf, "Mitarbeiter_Stammdaten_Vorlage.xlsx");
        showToast("📥 Vorlage für Mitarbeiter-Stammdaten heruntergeladen.");
      } catch (err) {
        console.error(err);
        alert("Fehler beim Erstellen der Vorlage: " + err.message);
      }
    });
  }

  // Gespeicherte Referenzdatei & Mitarbeiter beim Start aus Browser-Speicher laden
  restoreSavedReferenceFile();
  loadStaffFromStorage();

  function setupDropZone(dropZone, fileInput, onFileLoaded) {
    if (!dropZone || !fileInput) return;

    dropZone.addEventListener("click", (e) => {
      // Nicht auslösen, falls ein Button, Select, Label, Link oder das Info-Feld angeklickt wurde
      if (
        e.target.closest("button") ||
        e.target.closest("select") ||
        e.target.closest("option") ||
        e.target.closest("label") ||
        e.target.closest("a") ||
        e.target.closest(".file-info")
      ) {
        return;
      }
      fileInput.click();
    });

    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        onFileLoaded(fileInput.files[0]);
        fileInput.value = "";
      }
    });

    ["dragenter", "dragover"].forEach(evt => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach(evt => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove("dragover");
      });
    });

    dropZone.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("dragover");
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        onFileLoaded(e.dataTransfer.files[0]);
      }
    });
  }

  function populateSheetSelect(selectEl, worksheets) {
    selectEl.innerHTML = "";
    worksheets.forEach(ws => {
      const opt = document.createElement("option");
      opt.value = ws.name;
      opt.textContent = `${ws.name} (${ws.rowCount} Zeilen)`;
      selectEl.appendChild(opt);
    });
  }

  selectRefSheet.addEventListener("change", async (e) => {
    state.currentRefSheet = e.target.value;
    await updateRefSheetInStorage(e.target.value);
    updateMapping();
  });

  selectTgtSheet.addEventListener("change", (e) => {
    state.currentTargetSheet = e.target.value;
    updateMapping();
  });

  function checkReadyForConfig() {
    if (state.refWorkbook && state.targetWorkbook) {
      updateMapping();
      configSection.classList.remove("hidden");
      configSection.scrollIntoView({ behavior: "smooth" });
    }
  }

  // --- Spalten-Mapping ermitteln ---
  function updateMapping() {
    if (!state.refWorkbook || !state.targetWorkbook) return;
    const refWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
    const tgtWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!refWs || !tgtWs) return;

    const refHeaders = [];
    const refHeaderRow = refWs.getRow(1);
    refHeaderRow.eachCell((cell, colNum) => {
      const extracted = WebExcelEngine.extractCellValue(cell.value, cell);
      const val = extracted ? extracted.trim() : `Spalte ${getColLetter(colNum)}`;
      refHeaders.push({ colNum, name: val });
    });

    const tgtHeaders = [];
    const tgtHeaderRow = tgtWs.getRow(1);
    tgtHeaderRow.eachCell((cell, colNum) => {
      const extracted = WebExcelEngine.extractCellValue(cell.value, cell);
      const val = extracted ? extracted.trim() : `Spalte ${getColLetter(colNum)}`;
      tgtHeaders.push({ colNum, letter: getColLetter(colNum), name: val });
    });

    const mapping = [];
    tgtHeaders.forEach(tgt => {
      const tgtClean = tgt.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
      let bestMatch = "";

      for (const ref of refHeaders) {
        const refClean = ref.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
        if (tgtClean === refClean && tgtClean) {
          bestMatch = ref.name;
          break;
        }
      }
      if (!bestMatch) {
        for (const ref of refHeaders) {
          const refClean = ref.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
          if (tgtClean && refClean && (tgtClean.includes(refClean) || refClean.includes(tgtClean))) {
            bestMatch = ref.name;
            break;
          }
        }
      }

      // Spalten B, C und E sind standardmäßig angehakt, alle anderen abwählbar / optional
      const isDefaultCol = ["B", "C", "E"].includes(tgt.letter.toUpperCase()) || [2, 3, 5].includes(tgt.colNum);
      const isRessource = (tgt.letter.toUpperCase() === "C") || (tgt.colNum === 3) || ["ressource", "resource", "resour", "res-nr", "res_nr", "resnr"].some(k => tgt.name.toLowerCase().includes(k));
      const isLeistung = (tgt.letter.toUpperCase() === "E") || (tgt.colNum === 5) || ["leistung", "leist", "leist-nr", "leist_nr", "leistnr"].some(k => tgt.name.toLowerCase().includes(k));

      // Falls kein Match über Header-Namen gefunden wurde, nach gleicher Position/Buchstabe suchen
      if (!bestMatch) {
        for (const ref of refHeaders) {
          if (ref.colNum === tgt.colNum || (ref.letter && ref.letter.toUpperCase() === tgt.letter.toUpperCase())) {
            bestMatch = ref.name;
            break;
          }
        }
      }

      mapping.push({
        target_col_idx: tgt.colNum,
        target_col_letter: tgt.letter,
        target_col_name: tgt.name,
        ref_col_name: bestMatch,
        selected: isDefaultCol,
        allow_empty: isLeistung,
        pad_to_4: isRessource
      });
    });

    state.mapping = mapping;
    renderMappingTable(mapping, refHeaders.map(h => h.name));
    populateTimesheetColSelects(tgtHeaders);
    populateFullTableFilters();
    renderFullExcelTable();
  }

  function renderMappingTable(mapping, refHeaderNames) {
    mappingTbody.innerHTML = mapping.map((m, idx) => {
      const isChecked = m.selected ? "checked" : "";
      let optionsHtml = `<option value="">-- Nicht prüfen --</option>`;
      refHeaderNames.forEach(h => {
        const isSel = (h === m.ref_col_name) ? "selected" : "";
        optionsHtml += `<option value="${escapeHtml(h)}" ${isSel}>${escapeHtml(h)}</option>`;
      });

      const isAllowEmptyChecked = Boolean(m.allow_empty) ? "checked" : "";

      return `
        <tr>
          <td>
            <input type="checkbox" class="col-checkbox" data-idx="${idx}" ${isChecked}>
          </td>
          <td>
            <strong>${escapeHtml(m.target_col_name)}</strong>
            <span class="diag-detail"> (Spalte ${m.target_col_letter})</span>
          </td>
          <td style="color: var(--text-muted); font-size: 1.1rem;">➔</td>
          <td>
            <select class="form-select ref-col-select" data-idx="${idx}">
              ${optionsHtml}
            </select>
          </td>
          <td>
            <label class="custom-control">
              <input type="checkbox" class="allow-empty-checkbox" data-idx="${idx}" ${isAllowEmptyChecked}>
              <span>Leere Zellen erlauben</span>
            </label>
          </td>
        </tr>
      `;
    }).join("");

    // Synchronisiere "Alle auswählen"-Checkbox
    checkAllCols.checked = mapping.length > 0 && mapping.every(m => m.selected);

    // Event Listener für Mapping-Tabelle
    mappingTbody.querySelectorAll(".col-checkbox").forEach(cb => {
      cb.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].selected = e.target.checked;
        checkAllCols.checked = state.mapping.length > 0 && state.mapping.every(m => m.selected);
      });
    });

    mappingTbody.querySelectorAll(".ref-col-select").forEach(sel => {
      sel.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].ref_col_name = e.target.value;
        state.mapping[idx].selected = Boolean(e.target.value);
        const rowCb = mappingTbody.querySelector(`.col-checkbox[data-idx="${idx}"]`);
        if (rowCb) rowCb.checked = state.mapping[idx].selected;
        checkAllCols.checked = state.mapping.length > 0 && state.mapping.every(m => m.selected);
      });
    });

    mappingTbody.querySelectorAll(".allow-empty-checkbox").forEach(cb => {
      cb.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].allow_empty = e.target.checked;
      });
    });
  }

  checkAllCols.addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    state.mapping.forEach(m => { m.selected = isChecked; });
    mappingTbody.querySelectorAll(".col-checkbox").forEach(cb => { cb.checked = isChecked; });
  });

  // --- Datumsprüfung Konfiguration & Status-Badge ---
  const chkDateVal = document.getElementById("chk-date-validation");
  const badgeDateStatus = document.getElementById("badge-date-status");
  const inputDateWeeks = document.getElementById("input-date-weeks");
  const dateControls = document.getElementById("date-check-controls");

  function updateDateBadge() {
    if (!chkDateVal) return;
    if (!chkDateVal.checked) {
      if (badgeDateStatus) {
        badgeDateStatus.textContent = "Deaktiviert";
        badgeDateStatus.className = "badge badge-leer-ok";
      }
      if (dateControls) dateControls.style.opacity = "0.5";
    } else {
      const mode = document.querySelector('input[name="date-mode"]:checked')?.value || "weeks";
      const w = inputDateWeeks ? inputDateWeeks.value : 3;
      if (badgeDateStatus) {
        badgeDateStatus.textContent = mode === "weeks" ? `Aktiv (Letzte ${w} Wochen)` : "Aktiv (Fester Zeitraum)";
        badgeDateStatus.className = "badge badge-ok";
      }
      if (dateControls) dateControls.style.opacity = "1";
    }
  }

  if (chkDateVal) {
    chkDateVal.addEventListener("change", updateDateBadge);
  }
  if (inputDateWeeks) {
    inputDateWeeks.addEventListener("input", updateDateBadge);
  }
  document.querySelectorAll('input[name="date-mode"]').forEach(r => {
    r.addEventListener("change", updateDateBadge);
  });

  // --- Prüfung starten ---
  btnStartCheck.addEventListener("click", () => {
    const activeCols = state.mapping.filter(m => m.selected && m.ref_col_name);
    const dateCheckOn = chkDateVal ? chkDateVal.checked : true;
    if (activeCols.length === 0 && !dateCheckOn) {
      alert("Bitte wählen Sie mindestens eine Spalte für den Abgleich aus.");
      return;
    }

    btnStartCheck.disabled = true;
    btnStartCheck.innerHTML = `<span class="icon">⏳</span> Prüfung läuft...`;

    setTimeout(() => {
      runInspection();
      btnStartCheck.disabled = false;
      btnStartCheck.innerHTML = `<span class="icon">🚀</span> Prüfung starten`;
    }, 50);
  });

  function runInspection() {
    const refWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
    const tgtWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);

    // Datumsoptionen auslesen
    const dateCheckEnabled = chkDateVal ? chkDateVal.checked : true;
    const dateRangeMode = document.querySelector('input[name="date-mode"]:checked')?.value || "weeks";
    const weeksVal = parseInt(document.getElementById("input-date-weeks")?.value, 10) || 3;
    const startVal = document.getElementById("input-date-start")?.value || "";
    const endVal = document.getElementById("input-date-end")?.value || "";

    const today = new Date();
    today.setHours(23, 59, 59, 999);
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);

    let minDate = null;
    let maxDate = today;
    let weeksInfo = `letzte ${weeksVal} Wochen`;

    if (dateRangeMode === "custom" && (startVal || endVal)) {
      minDate = startVal ? WebExcelEngine.parseDateValue(startVal) : new Date(Date.now() - 21 * 86400000);
      maxDate = endVal ? WebExcelEngine.parseDateValue(endVal) : today;
      if (minDate) minDate.setHours(0, 0, 0, 0);
      if (maxDate) maxDate.setHours(23, 59, 59, 999);
      const sDisp = minDate ? WebExcelEngine.formatDate(minDate) : "?";
      const eDisp = maxDate ? WebExcelEngine.formatDate(maxDate) : "?";
      weeksInfo = `Zeitraum von ${sDisp} bis ${eDisp}`;
    } else {
      minDate = new Date(Date.now() - weeksVal * 7 * 86400000);
      minDate.setHours(0, 0, 0, 0);
      weeksInfo = `letzte ${weeksVal} Wochen (ab ${WebExcelEngine.formatDate(minDate)})`;
    }

    // 1. Referenz-Indizes aufbauen
    const colIndices = {};
    state.mapping.forEach(m => {
      if (!m.selected || !m.ref_col_name) return;

      // Finde Spalten-Index in Referenz
      let refColIdx = null;
      refWs.getRow(1).eachCell((cell, colNum) => {
        const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
        if (hVal === m.ref_col_name) {
          refColIdx = colNum;
        }
      });

      if (refColIdx) {
        const vals = [];
        for (let r = 2; r <= refWs.rowCount; r++) {
          const cell = refWs.getRow(r).getCell(refColIdx);
          const norm = WebExcelEngine.normalizeCellValue(cell.value, cell);
          if (norm) {
            vals.push(norm);
            if (m.pad_to_4 || /^\d{1,4}$/.test(norm)) {
              vals.push(WebExcelEngine.padNumber(norm, 4));
              const unpadded = norm.replace(/^0+/, "");
              if (unpadded) vals.push(unpadded);
            }
          }
        }
        colIndices[m.target_col_name] = new WebExcelEngine.ColumnIndex(vals);
      }
    });

    // 2. Prüfdatei analysieren
    const results = [];
    const stats = {
      total_cells: 0,
      ok: 0,
      zahlendreher: 0,
      zifferntausch: 0,
      tippfehler: 0,
      ziffer_zuviel_zuwenig: 0,
      nicht_existent: 0,
      leer: 0,
      datum_warnung: 0,
      total_errors: 0
    };

    state.appliedCorrections = {};

    const colAHeader = tgtWs.getRow(1).getCell(1);
    const colAName = WebExcelEngine.extractCellValue(colAHeader.value, colAHeader).trim() || "Datum";

    for (let r = 2; r <= tgtWs.rowCount; r++) {
      const row = tgtWs.getRow(r);

      // Datumsprüfung Spalte A
      if (dateCheckEnabled) {
        const cellCoordA = `A${r}`;
        const cellKeyA = `${r}_1`;
        const cellA = row.getCell(1);
        const parsedD = WebExcelEngine.parseDateValue(cellA.value, cellA);
        const rawDateStr = WebExcelEngine.extractCellValue(cellA.value, cellA).trim();
        const dispValA = parsedD ? WebExcelEngine.formatDate(parsedD) : (rawDateStr || "");
        stats.total_cells++;

        const todayStr = WebExcelEngine.formatDate(new Date());

        if (cellA.value === null || cellA.value === undefined || rawDateStr === "") {
          stats.total_errors++;
          stats.leer++;
          stats.datum_warnung++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: "",
            current_value: "",
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.0,
            suggestion: todayStr,
            detail: "Datum in Spalte A fehlt komplett",
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (!parsedD) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.0,
            suggestion: todayStr,
            detail: `Ungültiges Datumsformat oder Zahl ('${dispValA}')`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (minDate && parsedD < minDate) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          const daysDiff = Math.round((startOfToday - parsedD) / (86400000));
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.5,
            suggestion: todayStr,
            detail: `Datum ${dispValA} liegt ${daysDiff} Tage zurück (erwartet: ${weeksInfo})`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (maxDate && parsedD > maxDate) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          const daysDiff = Math.round((parsedD - today) / (86400000));
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.5,
            suggestion: todayStr,
            detail: `Datum ${dispValA} liegt in der Zukunft (+${daysDiff} Tage)`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else {
          stats.ok++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "OK",
            status_label: "Gültig",
            badge_class: "badge-ok",
            score: 1.0,
            suggestion: dispValA,
            detail: `Aktuelles Datum (${dispValA})`,
            candidates: [],
            is_corrected: false
          });
        }
      }

      state.mapping.forEach(m => {
        if (!m.selected || !m.ref_col_name) return;
        // Wenn Datumsprüfung aktiv ist, Spalte A nicht nochmals gegen Referenzpool abgleichen
        if (dateCheckEnabled && (m.target_col_idx === 1 || m.target_col_letter === "A")) return;

        const index = colIndices[m.target_col_name];
        if (!index) return;

        const cellCoord = `${m.target_col_letter}${r}`;
        const cellKey = `${r}_${m.target_col_idx}`;
        const cell = row.getCell(m.target_col_idx);
        const origRaw = WebExcelEngine.extractCellValue(cell.value, cell);
        const valStr = WebExcelEngine.normalizeCellValue(cell.value, cell);
        const origText = (origRaw && !origRaw.toLowerCase().includes("object")) ? origRaw.trim() : (valStr || "");
        stats.total_cells++;

        const allowEmpty = Boolean(m.allow_empty);

        if (!valStr) {
          if (allowEmpty) {
            stats.ok++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: "",
              status: "OK_LEER",
              status_label: "Leer (erlaubt)",
              badge_class: "badge-leer-ok",
              score: 1.0,
              suggestion: "",
              detail: "Feld ist leer (optional, kein Fehler)",
              candidates: [],
              is_corrected: false
            });
          } else {
            stats.leer++;
            stats.total_errors++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: "",
              status: "LEER",
              status_label: "Pflichtfeld leer",
              badge_class: "badge-leer",
              score: 0.0,
              suggestion: "",
              detail: "Zelle ist leer, obwohl Pflichtfeld",
              candidates: [],
              is_corrected: false
            });
          }
          return;
        }

        const targetToCheck = m.pad_to_4 ? WebExcelEngine.padNumber(valStr, 4) : valStr;
        const unpaddedTarget = valStr.replace(/^0+/, "");
        const neededPadding = (targetToCheck !== valStr);
        const isOk = index.contains(targetToCheck) || index.contains(valStr) || (unpaddedTarget && index.contains(unpaddedTarget));

        if (isOk) {
          stats.ok++;
          if (m.pad_to_4 && neededPadding) {
            // Automatisch für den 4-stelligen Export mit 0000 vormerken, aber NICHT als Fehler anzeigen:
            state.appliedCorrections[cellKey] = targetToCheck;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: targetToCheck,
              status: "OK",
              status_label: "Gültig",
              badge_class: "badge-ok",
              score: 1.0,
              suggestion: targetToCheck,
              detail: `Gültige Ressource (wird beim Speichern automatisch 4-stellig als '${targetToCheck}' formatiert)`,
              candidates: [],
              is_corrected: false
            });
          } else {
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: "OK",
              status_label: "Gültig",
              badge_class: "badge-ok",
              score: 1.0,
              suggestion: valStr,
              detail: "Exakte Übereinstimmung mit Referenz",
              candidates: [],
              is_corrected: false
            });
          }
        } else {
          stats.total_errors++;
          const candidates = index.findCandidates(targetToCheck);

          if (candidates.length > 0) {
            const best = candidates[0];
            let lbl = "Tippfehler";
            let bclass = "badge-tippfehler";

            if (best.type === "ZAHLENDREHER") {
              stats.zahlendreher++;
              lbl = "Zahlendreher";
              bclass = "badge-zahlendreher";
            } else if (best.type === "ZIFFERENTAUSCH") {
              stats.zifferntausch++;
              lbl = "Zifferntausch";
              bclass = "badge-zifferntausch";
            } else if (best.type === "TIPPFEHLER") {
              stats.tippfehler++;
              lbl = "Tippfehler (1 Ziffer)";
              bclass = "badge-tippfehler";
            } else if (best.type === "ZIFFER_ZUVIEL" || best.type === "ZIFFER_FEHLT") {
              stats.ziffer_zuviel_zuwenig++;
              lbl = "Ziffer zuviel/fehlt";
              bclass = "badge-tippfehler";
            } else if (best.type === "NAECHSTE_ALTERNATIVE") {
              stats.nicht_existent++;
              lbl = "Falscher Wert (Alternative)";
              bclass = "badge-fehler";
            }

            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: best.type,
              status_label: lbl,
              badge_class: bclass,
              score: best.score,
              suggestion: best.value,
              detail: best.detail,
              candidates: candidates,
              is_corrected: false
            });
          } else {
            stats.nicht_existent++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: "NICHT_EXISTENT",
              status_label: "Nicht in Stammdaten",
              badge_class: "badge-fehler",
              score: 0.0,
              suggestion: "",
              detail: "Nummer existiert nicht in der Referenzspalte",
              candidates: [],
              is_corrected: false
            });
          }
        }
      });
    }

    state.allResults = results;
    updateStats(stats);
    updateFilterCounts(results);
    renderResultsTable();
    renderTimesheetMatrix();
    populateFullTableFilters();
    renderFullExcelTable();

    resultsSection.classList.remove("hidden");
    resultsSection.scrollIntoView({ behavior: "smooth" });
    showToast(`Prüfung abgeschlossen: ${stats.total_errors} Abweichungen gefunden.`);
  }

  // --- Statistiken aktualisieren ---
  function updateStats(stats) {
    document.getElementById("stat-total").textContent = stats.total_cells.toLocaleString("de-DE");
    document.getElementById("stat-ok").textContent = stats.ok.toLocaleString("de-DE");
    document.getElementById("stat-zahlendreher").textContent = (stats.zahlendreher + stats.zifferntausch).toLocaleString("de-DE");
    document.getElementById("stat-tippfehler").textContent = (stats.tippfehler + stats.ziffer_zuviel_zuwenig).toLocaleString("de-DE");
    document.getElementById("stat-fehler").textContent = (stats.nicht_existent + stats.leer).toLocaleString("de-DE");
  }

  function updateFilterCounts(results) {
    const isError = (r) => (r.status !== "OK" && r.status !== "OK_LEER");
    const total = results.length;
    const errors = results.filter(isError).length;
    const zd = results.filter(r => r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH").length;
    const tp = results.filter(r => r.status === "TIPPFEHLER" || r.status.startsWith("ZIFFER_")).length;
    const ne = results.filter(r => r.status === "NICHT_EXISTENT" || r.status === "LEER" || r.status === "NAECHSTE_ALTERNATIVE").length;
    const datum = results.filter(r => r.status === "DATUM_WARNUNG").length;
    const corr = results.filter(r => r.is_corrected).length;

    document.getElementById("count-all").textContent = total;
    document.getElementById("count-errors").textContent = errors;
    document.getElementById("count-zd").textContent = zd;
    document.getElementById("count-tp").textContent = tp;
    document.getElementById("count-ne").textContent = ne;
    const countDatumEl = document.getElementById("count-datum");
    if (countDatumEl) countDatumEl.textContent = datum;
    document.getElementById("count-corr").textContent = corr;

    if (tabBadgeErrors) {
      tabBadgeErrors.textContent = `${errors} Abweichungen`;
    }
    if (tabBadgeTotalRows && state.targetWorkbook && state.currentTargetSheet) {
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      const totalDataRows = ws ? Math.max(0, ws.rowCount - 1) : 0;
      tabBadgeTotalRows.textContent = `${totalDataRows} Zeilen`;
    }
  }

  // --- Filter Event Listener ---
  filterPills.querySelectorAll(".pill").forEach(p => {
    p.addEventListener("click", () => {
      filterPills.querySelectorAll(".pill").forEach(el => el.classList.remove("active"));
      p.classList.add("active");
      state.currentFilter = p.dataset.filter;
      renderResultsTable();
    });
  });

  tableSearchInput.addEventListener("input", (e) => {
    state.searchQuery = e.target.value.toLowerCase().trim();
    renderResultsTable();
  });

  // --- Ergebnistabelle Rendern ---
  function renderResultsTable() {
    const query = state.searchQuery;
    const filter = state.currentFilter;

    const filtered = state.allResults.filter(r => {
      const isError = (r.status !== "OK" && r.status !== "OK_LEER");
      if (filter === "errors" && !isError) return false;
      if (filter === "zahlendreher" && r.status !== "ZAHLENDREHER" && r.status !== "ZIFFERENTAUSCH") return false;
      if (filter === "tippfehler" && r.status !== "TIPPFEHLER" && !r.status.startsWith("ZIFFER_")) return false;
      if (filter === "datum" && r.status !== "DATUM_WARNUNG") return false;
      if (filter === "nicht_existent" && r.status !== "NICHT_EXISTENT" && r.status !== "LEER" && r.status !== "NAECHSTE_ALTERNATIVE") return false;
      if (filter === "corrected" && !r.is_corrected) return false;

      if (query) {
        const rowStr = `${r.cell} ${r.col_name} ${r.row} ${r.original_value} ${r.current_value} ${r.status_label} ${r.detail} ${r.suggestion}`.toLowerCase();
        if (!rowStr.includes(query)) return false;
      }
      return true;
    });

    if (filtered.length === 0) {
      resultsTbody.innerHTML = `<tr><td colspan="7" style="text-align: center; padding: 2rem; color: var(--text-muted);">Keine Einträge für die aktuellen Filterkriterien gefunden.</td></tr>`;
      return;
    }

    resultsTbody.innerHTML = filtered.map(r => {
      const isCorrected = r.is_corrected;
      const isZahlendreher = (r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH");
      const isError = (r.status !== "OK" && r.status !== "OK_LEER");

      let rowClass = "";
      if (isCorrected) rowClass = "row-corrected";
      else if (isZahlendreher) rowClass = "row-zahlendreher";
      else if (isError) rowClass = "row-error";

      const strikedClass = isCorrected ? "is-striked" : "";
      const inputCorrClass = isCorrected ? "is-corrected" : "";

      let suggHtml = "-";
      if (r.suggestion) {
        let otherChips = "";
        if (r.candidates && r.candidates.length > 1) {
          const others = r.candidates.filter(c => c.value !== r.suggestion).slice(0, 3);
          if (others.length > 0) {
            otherChips = `
              <div class="alt-chips">
                <span>Weitere:</span>
                ${others.map(c => `
                  <button type="button" class="alt-chip" data-id="${r.id}" data-val="${escapeHtml(c.value)}" title="${escapeHtml(c.detail)}">
                    ${escapeHtml(c.value)}
                  </button>
                `).join("")}
              </div>
            `;
          }
        }

        suggHtml = `
          <div class="alt-chips-container">
            <button class="btn-apply-sugg" data-id="${r.id}" data-val="${escapeHtml(r.suggestion)}" title="Diesen Vorschlag übernehmen">
              <span class="icon">✓</span> <strong>${escapeHtml(r.suggestion)}</strong> übernehmen
            </button>
            ${otherChips}
          </div>
        `;
      }

      const manualHtml = `
        <div class="manual-edit-box">
          <input type="text" class="manual-input ${inputCorrClass}" 
            data-id="${r.id}" 
            value="${escapeHtml(r.current_value)}" 
            placeholder="Nummer eintragen..."
            title="Nummer manuell anpassen und Enter drücken">
          <button class="btn-save-manual" data-id="${r.id}" title="Händisch speichern">💾 Speichern</button>
        </div>
        ${isCorrected ? '<div style="margin-top: 4px;"><span class="badge badge-corrected">✓ Geändert</span></div>' : ''}
      `;

      const displayOrig = (r.original_value && !String(r.original_value).toLowerCase().includes("object")) ? r.original_value : "—";

      return `
        <tr class="${rowClass}">
          <td><span class="cell-coord">${escapeHtml(r.cell)}</span></td>
          <td><strong>${escapeHtml(r.col_name)}</strong></td>
          <td style="color: var(--text-muted);">Zeile ${r.row}</td>
          <td><span class="val-old ${strikedClass}">${escapeHtml(displayOrig)}</span></td>
          <td>
            <span class="badge ${r.badge_class}">${escapeHtml(r.status_label)}</span>
            <div class="diag-detail" style="margin-top: 3px;">${escapeHtml(r.detail)}</div>
          </td>
          <td>${suggHtml}</td>
          <td>${manualHtml}</td>
        </tr>
      `;
    }).join("");

    // Event Listener für Vorschläge
    resultsTbody.querySelectorAll(".btn-apply-sugg").forEach(btn => {
      btn.addEventListener("click", () => {
        applySingleCorrection(btn.dataset.id, btn.dataset.val);
      });
    });

    resultsTbody.querySelectorAll(".alt-chip").forEach(chip => {
      chip.addEventListener("click", () => {
        applySingleCorrection(chip.dataset.id, chip.dataset.val);
      });
    });

    resultsTbody.querySelectorAll(".btn-save-manual").forEach(btn => {
      btn.addEventListener("click", () => {
        const id = btn.dataset.id;
        const input = resultsTbody.querySelector(`.manual-input[data-id="${id}"]`);
        if (input) {
          applySingleCorrection(id, input.value.trim());
        }
      });
    });

    resultsTbody.querySelectorAll(".manual-input").forEach(input => {
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          applySingleCorrection(input.dataset.id, input.value.trim());
          input.blur();
        }
      });
    });
  }

  function applySingleCorrection(cellId, newVal) {
    const targetItem = state.allResults.find(r => r.id === cellId);
    if (!targetItem) return;

    const origVal = targetItem.original_value;
    const colIdx = targetItem.col_idx;
    let count = 0;

    state.allResults.forEach(r => {
      if (r.id === cellId || (targetItem && r.col_idx === colIdx && r.original_value === origVal)) {
        r.current_value = newVal;
        r.is_corrected = (newVal !== r.original_value);
        state.appliedCorrections[r.id] = newVal;
        count++;
      }
    });

    updateFilterCounts(state.allResults);
    renderResultsTable();
    if (state.timesheetData) renderTimesheetMatrix();
    renderFullExcelTable();

    if (count > 1) {
      showToast(`Korrektur "${newVal}" für alle ${count} Zellen mit Wert "${origVal}" übernommen!`);
    } else {
      showToast(`Zelle ${targetItem.cell} aktualisiert auf "${newVal}".`);
    }
  }

  // --- GENERAL-BUTTON: Datei erstellen & Änderungen übernehmen ---
  if (btnGeneralApply) {
    btnGeneralApply.addEventListener("click", async () => {
      if (!state.allResults || state.allResults.length === 0) {
        showToast("Bitte führen Sie zuerst eine Prüfung durch.");
        return;
      }

      btnGeneralApply.disabled = true;
      btnGeneralApply.innerHTML = `<span class="icon">⏳</span> Erstelle korrigierte Excel-Datei...`;
      showToast("General-Button: Erstelle korrigierte Excel-Datei...");

      // 1. Alle im Formular manuell eingegebenen Werte sammeln
      resultsTbody.querySelectorAll(".manual-input").forEach(input => {
        const id = input.dataset.id;
        const val = input.value.trim();
        const item = state.allResults.find(r => r.id === id);
        if (item && val && val !== item.original_value) {
          item.current_value = val;
          item.is_corrected = true;
          state.appliedCorrections[id] = val;
        }
      });

      // 2. Excel-Datei 1:1 klonen und aufbauen
      try {
        const outBuffer = await generateCorrectedExcel();
        const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
        const outFileName = `${baseName}_KORRIGIERT.xlsx`;

        // 3. Browser-Download anstoßen
        downloadBuffer(outBuffer, outFileName);

        btnGeneralApply.disabled = false;
        btnGeneralApply.innerHTML = `<span class="icon">🚀</span> <strong>General-Button: Datei erstellen &amp; Änderungen übernehmen</strong>`;

        // Erfolgsbanner einblenden
        const banner = document.getElementById("download-success-banner");
        const nameEl = document.getElementById("download-file-name");
        if (banner && nameEl) {
          nameEl.textContent = outFileName;
          banner.classList.remove("hidden");
          banner.scrollIntoView({ behavior: "smooth" });
        }

        if (btnExportExcel) btnExportExcel.classList.remove("hidden");
        if (btnExportReport) btnExportReport.classList.remove("hidden");

        renderResultsTable();
        updateFilterCounts(state.allResults);
        if (state.timesheetData) renderTimesheetMatrix();
        renderFullExcelTable();
        showToast(`✅ Datei "${outFileName}" erfolgreich im Download-Ordner gespeichert!`);
      } catch (err) {
        console.error(err);
        btnGeneralApply.disabled = false;
        btnGeneralApply.innerHTML = `<span class="icon">🚀</span> <strong>General-Button: Datei erstellen &amp; Änderungen übernehmen</strong>`;
        alert("Fehler beim Erstellen der Datei: " + err.message);
      }
    });
  }

  // --- Nur Zahlendreher korrigieren ---
  if (btnFixZahlendreher) {
    btnFixZahlendreher.addEventListener("click", () => {
      let count = 0;
      state.allResults.forEach(r => {
        if ((r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH") && r.suggestion) {
          r.current_value = r.suggestion;
          r.is_corrected = true;
          state.appliedCorrections[r.id] = r.suggestion;
          count++;
        }
      });
      updateFilterCounts(state.allResults);
      renderResultsTable();
      if (state.timesheetData) renderTimesheetMatrix();
      renderFullExcelTable();
      showToast(`⚡ ${count} Zahlendreher wurden automatisch korrigiert!`);
    });
  }

  // --- Zurücksetzen ---
  if (btnResetFixes) {
    btnResetFixes.addEventListener("click", () => {
      if (!confirm("Alle Korrekturen rückgängig machen?")) return;
      state.allResults.forEach(r => {
        r.current_value = r.original_value;
        r.is_corrected = false;
      });
      state.appliedCorrections = {};
      updateFilterCounts(state.allResults);
      renderResultsTable();
      if (state.timesheetData) renderTimesheetMatrix();
      renderFullExcelTable();
      showToast("Alle Korrekturen wurden zurückgesetzt.");
    });
  }

  // --- Nochmals herunterladen ---
  const btnDownloadAgain = document.getElementById("btn-download-again");
  if (btnDownloadAgain) {
    btnDownloadAgain.addEventListener("click", async () => {
      const outBuffer = await generateCorrectedExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(outBuffer, `${baseName}_KORRIGIERT.xlsx`);
      showToast("📥 Datei erneut heruntergeladen.");
    });
  }

  if (btnExportExcel) {
    btnExportExcel.addEventListener("click", async () => {
      const outBuffer = await generateCorrectedExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(outBuffer, `${baseName}_KORRIGIERT.xlsx`);
      showToast("📥 Korrigierte Excel wird heruntergeladen...");
    });
  }

  if (btnExportReport) {
    btnExportReport.addEventListener("click", async () => {
      const reportBuffer = await generateReportExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(reportBuffer, `Pruefbericht_${baseName}.xlsx`);
      showToast("📊 Prüfbericht wird heruntergeladen...");
    });
  }

  // --- Excel-Erstellung im 1:1 Original-Layout mit 0000 Formatierung ---
  async function generateCorrectedExcel() {
    // Frisches Arbeitsblatt aus Original-Buffer klonen
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(state.lastTargetBuffer);
    const ws = wb.getWorksheet(state.currentTargetSheet);

    // 1. Angewendete Korrekturen eintragen
    Object.entries(state.appliedCorrections).forEach(([cellKey, newVal]) => {
      const [rStr, cStr] = cellKey.split("_");
      const r = parseInt(rStr, 10);
      const c = parseInt(cStr, 10);
      const cell = ws.getRow(r).getCell(c);

      const valStr = String(newVal).trim();
      if (c === 1) {
        cell.value = valStr;
        cell.numFmt = "DD.MM.YYYY";
      } else if (/^\d+$/.test(valStr) && valStr.startsWith("0") && valStr.length > 1) {
        cell.value = valStr; // Text mit führenden Nullen
      } else if (/^\d+$/.test(valStr)) {
        cell.value = parseInt(valStr, 10);
      } else if (/^-?\d+\.\d+$/.test(valStr)) {
        cell.value = parseFloat(valStr);
      } else {
        cell.value = valStr;
      }
    });

    // 2. Formatierung für alle Spalten vorgeben:
    // Spalte "A" (Spalte 1): Typ Datum "DD.MM.YYYY"
    // Spalte "C" (Spalte 3): Typ "0000" (Ressourcen-Nummern mit führenden Nullen)
    // Alle anderen Spalten: Typ Standard ("General")
    for (let colIdx = 1; colIdx <= ws.columnCount; colIdx++) {
      const col = ws.getColumn(colIdx);
      const isColA = (colIdx === 1 || col.letter === "A");
      const isColC = (colIdx === 3 || col.letter === "C");

      if (isColA) {
        col.numFmt = "DD.MM.YYYY";
        for (let r = 2; r <= ws.rowCount; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          // WICHTIG: Original-Datumsangaben NIEMALS in JavaScript Date umwandeln!
          // (Verhindert jegliche UTC-Zeitzonenverschiebung um -1 Tag)
          // Falls die Originalzelle eine Zahl ist (Excel-Seriennummer), als DD.MM.YYYY formatieren.
          if (typeof cell.value === "number") {
            cell.numFmt = "DD.MM.YYYY";
          }
        }
      } else if (isColC) {
        col.numFmt = "0000";
        for (let r = 2; r <= ws.rowCount; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          cell.numFmt = "0000";
          if (cell.value !== null && cell.value !== undefined) {
            const s = WebExcelEngine.extractCellValue(cell.value, cell).trim();
            if (/^\d+$/.test(s)) {
              cell.value = parseInt(s, 10);
            }
          }
        }
      } else {
        col.numFmt = "General";
        for (let r = 2; r <= ws.rowCount; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          cell.numFmt = "General";
        }
      }
    }

    return await wb.xlsx.writeBuffer();
  }

  // --- Detaillierten Prüfbericht als Excel erstellen ---
  async function generateReportExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Prüfbericht");

    ws.columns = [
      { header: "Zelle", key: "cell", width: 10 },
      { header: "Zeile", key: "row", width: 8 },
      { header: "Spalte", key: "col_name", width: 22 },
      { header: "Originalwert", key: "orig", width: 18 },
      { header: "Aktueller Wert", key: "curr", width: 18 },
      { header: "Status", key: "status_lbl", width: 25 },
      { header: "Vorschlag", key: "sugg", width: 18 },
      { header: "Diagnosedetails", key: "detail", width: 45 },
      { header: "Wurde Korrigiert?", key: "is_corr", width: 18 }
    ];

    const hdrRow = ws.getRow(1);
    hdrRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    hdrRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E293B" }
    };

    state.allResults.forEach(r => {
      ws.addRow({
        cell: r.cell,
        row: r.row,
        col_name: r.col_name,
        orig: r.original_value,
        curr: r.current_value,
        status_lbl: r.status_label,
        sugg: r.suggestion || "—",
        detail: r.detail,
        is_corr: r.is_corrected ? "Ja" : "Nein"
      });
    });

    return await wb.xlsx.writeBuffer();
  }

  function downloadBuffer(buffer, fileName) {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 1000);
  }

  // ==============================================================================
  // Arbeitszeiten-Übersicht & Mitarbeiter-Matrix (Schritt 4)
  // ==============================================================================

  function populateTimesheetColSelects(tgtHeaders) {
    if (!selectTsDateCol || !selectTsHoursCol || !selectTsResourceCol) return;

    selectTsDateCol.innerHTML = "";
    selectTsHoursCol.innerHTML = "";
    selectTsResourceCol.innerHTML = "";

    // 1. Suche nach Datumsspalte
    let bestDateCol = null;
    tgtHeaders.forEach(h => {
      const lower = h.name.toLowerCase();
      if (!bestDateCol && (lower.includes("datum") || lower.includes("date") || lower.includes("tag") || lower.includes("zeitpunkt") || lower.includes("buchung"))) {
        bestDateCol = h.colNum;
      }
    });
    if (!bestDateCol && tgtHeaders.length > 0) {
      bestDateCol = tgtHeaders[0].colNum;
    }

    // 2. Suche nach Stundenspalte / Menge
    let bestHoursCol = null;
    tgtHeaders.forEach(h => {
      const lower = h.name.toLowerCase();
      if (!bestHoursCol && (lower.includes("stunde") || lower.includes("zeit") || lower.includes("dauer") || lower.includes("ist-stunden") || lower.includes("menge") || lower.includes("std") || lower === "h")) {
        bestHoursCol = h.colNum;
      }
    });
    if (!bestHoursCol) {
      const mengeCol = tgtHeaders.find(h => h.name.toLowerCase().includes("menge") || h.letter === "D" || h.letter === "G");
      if (mengeCol) bestHoursCol = mengeCol.colNum;
      else if (tgtHeaders.length >= 4) bestHoursCol = tgtHeaders[3].colNum;
    }

    // 3. Suche nach Ressourcenspalte
    let bestResCol = null;

    // Falls Mitarbeiter geladen sind: Prüfe die ersten Zeilen der Ziel-Tabelle auf die meisten Treffer mit Mitarbeiternummern!
    if (state.staffList && state.staffList.length > 0 && state.targetWorkbook && state.currentTargetSheet) {
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (ws) {
        const colMatchCounts = {};
        const maxSampleRows = Math.min(ws.rowCount, 60);
        for (let r = 2; r <= maxSampleRows; r++) {
          const sampleRow = ws.getRow(r);
          tgtHeaders.forEach(h => {
            const val = WebExcelEngine.extractCellValue(sampleRow.getCell(h.colNum).value, sampleRow.getCell(h.colNum));
            if (val && getStaffName(val)) {
              colMatchCounts[h.colNum] = (colMatchCounts[h.colNum] || 0) + 1;
            }
          });
        }
        let maxMatches = 0;
        let matchedColNum = null;
        for (const [colNumStr, count] of Object.entries(colMatchCounts)) {
          if (count > maxMatches) {
            maxMatches = count;
            matchedColNum = parseInt(colNumStr, 10);
          }
        }
        if (maxMatches >= 1) {
          bestResCol = matchedColNum;
        }
      }
    }

    if (!bestResCol) {
      tgtHeaders.forEach(h => {
        const lower = h.name.toLowerCase();
        if (!bestResCol && (
          lower.includes("ressource") || lower.includes("resource") || lower.includes("pers") ||
          lower.includes("mitarbeiter") || lower.includes("monteur") || lower.includes("techniker") ||
          lower.includes("arbeiter") || lower.includes("personal") || lower.includes("manr") ||
          lower.includes("kollege") || h.letter === "C" || h.colNum === 3
        )) {
          bestResCol = h.colNum;
        }
      });
    }
    if (!bestResCol) {
      const colC = tgtHeaders.find(h => h.letter === "C" || h.colNum === 3);
      if (colC) bestResCol = colC.colNum;
      else if (tgtHeaders.length >= 3) bestResCol = tgtHeaders[2].colNum;
    }

    // Optionen für Datum
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestDateCol) opt.selected = true;
      selectTsDateCol.appendChild(opt);
    });
    const optNoDate = document.createElement("option");
    optNoDate.value = "none";
    optNoDate.textContent = "-- Ohne Datum (Gesamtüberblick) --";
    selectTsDateCol.appendChild(optNoDate);

    // Optionen für Stunden
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestHoursCol) opt.selected = true;
      selectTsHoursCol.appendChild(opt);
    });
    const optFixedOne = document.createElement("option");
    optFixedOne.value = "fixed_one";
    optFixedOne.textContent = "-- Pauschal 1 Std pro Buchung --";
    selectTsHoursCol.appendChild(optFixedOne);

    // Optionen für Ressourcen
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestResCol) opt.selected = true;
      selectTsResourceCol.appendChild(opt);
    });
  }

  // Event Listener für Spaltenauswahl & Filter der Arbeitszeiten
  [selectTsDateCol, selectTsHoursCol, selectTsResourceCol, chkTsOnlyStaff].forEach(sel => {
    if (sel) {
      sel.addEventListener("change", () => {
        if (state.targetWorkbook && state.currentTargetSheet) {
          renderTimesheetMatrix();
          populateFullTableFilters();
          renderFullExcelTable();
        }
      });
    }
  });

  function parseRowDate(cellVal, cell) {
    if (!cellVal && cellVal !== 0) return { isoKey: "ohne_datum", displayDate: "Ohne Datum", weekday: "", label: "Ohne Datum" };

    let d = null;
    if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
      d = cellVal;
    } else if (typeof cellVal === "number" && cellVal > 1000 && cellVal < 100000) {
      // Excel-Seriennummer (25569 Tage zwischen 01.01.1900 und 01.01.1970)
      const ms = Math.round((cellVal - 25569) * 86400 * 1000);
      d = new Date(ms);
    } else {
      let str = "";
      if (typeof cellVal === "object" && cellVal) {
        str = String(cellVal.result || cellVal.text || "").trim();
      } else {
        str = String(cellVal).trim();
      }

      // Format TT.MM.JJJJ oder TT.MM.JJ
      const deMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
      if (deMatch) {
        let day = parseInt(deMatch[1], 10);
        let month = parseInt(deMatch[2], 10) - 1;
        let year = parseInt(deMatch[3], 10);
        if (year < 100) year += 2000;
        d = new Date(year, month, day);
      } else {
        // Format JJJJ-MM-TT
        const isoMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (isoMatch) {
          d = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
        } else {
          const parsed = Date.parse(str);
          if (!isNaN(parsed)) {
            d = new Date(parsed);
          }
        }
      }
    }

    if (!d || isNaN(d.getTime())) {
      const clean = String(cellVal).trim() || "Ohne Datum";
      return { isoKey: clean, displayDate: clean, weekday: "", label: clean };
    }

    const pad2 = (n) => String(n).padStart(2, "0");
    const yyyy = d.getFullYear();
    const mm = pad2(d.getMonth() + 1);
    const dd = pad2(d.getDate());
    const isoKey = `${yyyy}-${mm}-${dd}`;
    const displayDate = `${dd}.${mm}.${yyyy}`;
    const WEEKDAYS = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
    const weekday = WEEKDAYS[d.getDay()] || "";
    const label = weekday ? `${weekday}, ${dd}.${mm}.` : displayDate;

    return { isoKey, displayDate, weekday, label };
  }

  function parseHours(cellVal) {
    if (cellVal === null || cellVal === undefined || cellVal === "") return 0;
    if (typeof cellVal === "number") return isNaN(cellVal) ? 0 : cellVal;

    let str = "";
    if (typeof cellVal === "object") {
      str = String(cellVal.result || cellVal.text || "").trim();
    } else {
      str = String(cellVal).trim();
    }
    if (!str) return 0;

    // Zeitformat z.B. "08:30" oder "8:15"
    const timeMatch = str.match(/^(\d{1,2}):(\d{2})$/);
    if (timeMatch) {
      const h = parseInt(timeMatch[1], 10);
      const m = parseInt(timeMatch[2], 10);
      return h + (m / 60);
    }

    // Komma durch Punkt ersetzen und Zahlen filtern
    str = str.replace(/[^0-9,.-]/g, "").replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  function formatHoursDe(hours) {
    if (hours === 0) return "0,0";
    return Number(hours).toLocaleString("de-DE", { minimumFractionDigits: 1, maximumFractionDigits: 2 });
  }

  function renderTimesheetMatrix() {
    if (!timesheetSection || !state.targetWorkbook || !state.currentTargetSheet) {
      if (timesheetSection) timesheetSection.classList.add("hidden");
      return;
    }

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) {
      timesheetSection.classList.add("hidden");
      return;
    }

    const dateColVal = selectTsDateCol?.value;
    const isNoDate = (dateColVal === "none");
    const dateColIdx = isNoDate ? null : parseInt(dateColVal, 10);

    const hoursColVal = selectTsHoursCol?.value;
    const isFixedHours = (hoursColVal === "fixed_one");
    const hoursColIdx = isFixedHours ? null : parseInt(hoursColVal, 10);

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;

    const onlyStaff = chkTsOnlyStaff ? chkTsOnlyStaff.checked : true;
    const hasStaffList = state.staffList && state.staffList.length > 0;

    let skippedNonStaffRows = 0;
    const skippedNonStaffResources = new Set();

    // Maps für Aggregation
    const dateMap = new Map();
    const empMap = new Map();

    let grandTotalHours = 0;
    let grandTotalEntries = 0;

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);

      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;

      // 1. Ressource ermitteln (inklusive etwaiger Benutzer-Korrekturen)
      const cellKey = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined)
        ? String(state.appliedCorrections[cellKey]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");

      if (!rawRes) continue; // Zeilen ohne Ressourcennummer überspringen

      // Mitarbeiter-Name aus Stammdaten ermitteln
      const empNameFromMap = getStaffName(rawRes);
      const isKnownWorker = Boolean(empNameFromMap);

      // WICHTIG: Wenn Mitarbeiterdatei hinterlegt ist und Filter aktiv ist,
      // ausschließlich echte Arbeiter anzeigen! Fremd-/Maschinen-Ressourcen überspringen.
      if (hasStaffList && onlyStaff && !isKnownWorker) {
        skippedNonStaffRows++;
        skippedNonStaffResources.add(rawRes);
        continue;
      }

      const paddedRes = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
      const displayRes = paddedRes;
      const empName = empNameFromMap || `Ressource ${displayRes}`;
      const hasCustomName = Boolean(empNameFromMap);

      // 2. Datum ermitteln
      let dateObj = null;
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        dateObj = parseRowDate(cell.value, cell);
      } else {
        dateObj = { isoKey: "gesamt", displayDate: "Gesamtzeit", weekday: "", label: "Gesamtzeit" };
      }

      // 3. Stunden ermitteln
      let hours = 0;
      if (isFixedHours) {
        hours = 1;
      } else if (hoursColIdx) {
        const cell = row.getCell(hoursColIdx);
        hours = parseHours(cell.value);
      } else {
        hours = 1;
      }

      // Datum aggregieren
      if (!dateMap.has(dateObj.isoKey)) {
        dateMap.set(dateObj.isoKey, {
          isoKey: dateObj.isoKey,
          displayDate: dateObj.displayDate,
          weekday: dateObj.weekday,
          label: dateObj.label,
          totalHours: 0
        });
      }
      dateMap.get(dateObj.isoKey).totalHours += hours;

      // Mitarbeiter aggregieren (Eindeutiger Schlüssel: empName falls bekannt, sonst displayRes)
      const empKey = isKnownWorker ? empName : displayRes;
      if (!empMap.has(empKey)) {
        empMap.set(empKey, {
          resource: displayRes,
          name: empName,
          hasCustomName: hasCustomName,
          hoursByDate: {},
          totalHours: 0,
          totalEntries: 0
        });
      }
      const empRecord = empMap.get(empKey);
      empRecord.hoursByDate[dateObj.isoKey] = (empRecord.hoursByDate[dateObj.isoKey] || 0) + hours;
      empRecord.totalHours += hours;
      empRecord.totalEntries += 1;

      grandTotalHours += hours;
      grandTotalEntries += 1;
    }

    // Sortierung der Kalendertage
    const sortedDates = Array.from(dateMap.values()).sort((a, b) => a.isoKey.localeCompare(b.isoKey));

    // Sortierung der Mitarbeiter nach Ressourcennummer aufsteigend
    const sortedEmployees = Array.from(empMap.values()).sort((a, b) => {
      const numA = parseInt(a.resource, 10);
      const numB = parseInt(b.resource, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.resource.localeCompare(b.resource);
    });

    state.timesheetData = {
      sortedDates,
      employees: sortedEmployees,
      grandTotalHours,
      grandTotalEntries
    };

    // Status-Badge aktualisieren
    if (tsStaffStatusBadge) {
      if (hasStaffList) {
        const workerCount = sortedEmployees.length;
        const totalStaffInList = state.staffList.length;
        if (onlyStaff) {
          tsStaffStatusBadge.className = "badge badge-ok";
          let badgeHtml = `👷 ${workerCount} von ${totalStaffInList} Arbeiter erfasst`;
          if (skippedNonStaffRows > 0) {
            const resListStr = Array.from(skippedNonStaffResources).sort().join(", ");
            badgeHtml += ` • <span title="Ausgeblendete Ressourcen: ${escapeHtml(resListStr)}">🚫 ${skippedNonStaffRows} Buchung(en) Fremd-Ressourcen ausgeblendet (${skippedNonStaffResources.size} Nr.: ${escapeHtml(resListStr)})</span>`;
          }
          tsStaffStatusBadge.innerHTML = badgeHtml;
        } else {
          tsStaffStatusBadge.className = "badge badge-info";
          tsStaffStatusBadge.innerHTML = `⚠️ Alle ${sortedEmployees.length} Ressourcen aktiv (inkl. Fremd/Maschinen)`;
        }
      } else {
        tsStaffStatusBadge.className = "badge badge-leer-ok";
        tsStaffStatusBadge.innerHTML = `⚠️ Keine Mitarbeiterdatei geladen (alle Ressourcen werden angezeigt)`;
      }
    }

    if (sortedEmployees.length === 0) {
      timesheetThead.innerHTML = `<tr><th class="ts-col-emp">Mitarbeiter / Ressource</th><th>Status</th></tr>`;
      let emptyMsg = "Keine Buchungen in der ausgewählten Ressourcenspalte gefunden.";
      if (hasStaffList && onlyStaff && skippedNonStaffRows > 0) {
        const resListStr = Array.from(skippedNonStaffResources).sort().join(", ");
        emptyMsg = `Es wurden keine Arbeiter aus der Mitarbeiterdatei gefunden (${skippedNonStaffRows} Buchungen anderer Ressourcen [${resListStr}] ausgeblendet).`;
      }
      timesheetTbody.innerHTML = `<tr><td colspan="2" style="text-align: center; padding: 1.5rem; color: var(--text-muted);">${escapeHtml(emptyMsg)}</td></tr>`;
      timesheetTfoot.innerHTML = "";
      timesheetSection.classList.remove("hidden");
      return;
    }

    // THEAD
    timesheetThead.innerHTML = `
      <tr>
        <th class="ts-col-emp">Mitarbeiter / Ressource</th>
        ${sortedDates.map(d => `<th title="${escapeHtml(d.displayDate)}">${escapeHtml(d.label)}</th>`).join("")}
        <th class="ts-col-total">Gesamtstunden</th>
        <th>Buchungen</th>
      </tr>
    `;

    // TBODY
    timesheetTbody.innerHTML = sortedEmployees.map(emp => {
      const dateCells = sortedDates.map(d => {
        const h = emp.hoursByDate[d.isoKey] || 0;
        if (h > 0) {
          return `<td><span class="ts-hour-badge ts-hour-clickable" data-res="${escapeHtml(emp.resource)}" data-date="${escapeHtml(d.isoKey)}" data-display-date="${escapeHtml(d.displayDate)}" title="🔍 Klicken, um alle Zeilen von ${escapeHtml(emp.name)} am ${escapeHtml(d.displayDate)} (${formatHoursDe(h)} Std) in der Gesamttabelle anzuzeigen &amp; zu bearbeiten">${formatHoursDe(h)}</span></td>`;
        } else {
          return `<td><span class="ts-hour-empty">-</span></td>`;
        }
      }).join("");

      return `
        <tr>
          <td class="ts-cell-emp">
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 0.5rem;">
              <span class="ts-emp-clickable" data-res="${escapeHtml(emp.resource)}" style="font-weight: 600; color: #1e293b;" title="🔍 Klicken, um alle Zeilen von ${escapeHtml(emp.name)} in der Gesamttabelle anzuzeigen">${escapeHtml(emp.name)}</span>
              <span class="badge ${emp.hasCustomName ? 'badge-ok' : 'badge-leer-ok'}">Nr. ${escapeHtml(emp.resource)}</span>
            </div>
          </td>
          ${dateCells}
          <td class="ts-cell-total"><span class="ts-hour-badge ts-hour-total">${formatHoursDe(emp.totalHours)} h</span></td>
          <td>${emp.totalEntries}</td>
        </tr>
      `;
    }).join("");

    // TFOOT
    timesheetTfoot.innerHTML = `
      <tr>
        <td class="ts-cell-emp"><strong>Gesamtsumme Tag:</strong></td>
        ${sortedDates.map(d => `<td><strong>${formatHoursDe(d.totalHours)} h</strong></td>`).join("")}
        <td class="ts-cell-total"><span class="ts-grand-total">${formatHoursDe(grandTotalHours)} h</span></td>
        <td><strong>${grandTotalEntries}</strong></td>
      </tr>
    `;

    timesheetSection.classList.remove("hidden");

    // Event Listener für interaktiven Klick-Sprung in die Gesamttabelle
    timesheetTbody.querySelectorAll(".ts-hour-clickable").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        drillDownToFullTable(el.dataset.res, el.dataset.date);
      });
    });

    timesheetTbody.querySelectorAll(".ts-emp-clickable").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        drillDownToFullTable(el.dataset.res, null);
      });
    });
  }

  async function generateTimesheetExcel() {
    if (!state.timesheetData || state.timesheetData.employees.length === 0) {
      alert("Es liegen noch keine Daten für die Arbeitszeiten-Übersicht vor.");
      return null;
    }

    const { sortedDates, employees, grandTotalHours, grandTotalEntries } = state.timesheetData;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Arbeitszeiten_Uebersicht");

    const cols = [
      { header: "Mitarbeiter", key: "name", width: 28 },
      { header: "Ressource", key: "res", width: 14 }
    ];

    sortedDates.forEach((d, i) => {
      cols.push({
        header: d.label || d.displayDate,
        key: `d_${i}`,
        width: 14
      });
    });

    cols.push({ header: "Gesamt (Std)", key: "total", width: 16 });
    cols.push({ header: "Buchungen", key: "entries", width: 12 });

    ws.columns = cols;

    // Header styling
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell, colNum) => {
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0F766E" } // Teal header
      };
      cell.alignment = { vertical: "middle", horizontal: (colNum <= 1 ? "left" : "center"), wrapText: true };
      cell.border = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "medium", color: { argb: "FF0D5F58" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
      };
    });

    // Data rows
    employees.forEach(emp => {
      const rowValues = [emp.name, emp.resource];
      sortedDates.forEach(d => {
        const h = emp.hoursByDate[d.isoKey] || 0;
        rowValues.push(h > 0 ? h : null);
      });
      rowValues.push(emp.totalHours);
      rowValues.push(emp.totalEntries);

      const row = ws.addRow(rowValues);
      row.height = 22;

      // Spalte B (Ressource) als 0000 formatieren
      const resCell = row.getCell(2);
      resCell.numFmt = "0000";
      resCell.alignment = { horizontal: "center", vertical: "middle" };

      // Datumsspalten formatieren
      for (let c = 3; c < 3 + sortedDates.length; c++) {
        const cell = row.getCell(c);
        cell.numFmt = '#,##0.0 "h"';
        cell.alignment = { horizontal: "right", vertical: "middle" };
      }

      // Gesamtstunden
      const totalCell = row.getCell(3 + sortedDates.length);
      totalCell.numFmt = '#,##0.0 "h"';
      totalCell.font = { bold: true };
      totalCell.alignment = { horizontal: "right", vertical: "middle" };
      totalCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0FDFA" }
      };

      // Buchungen
      const entriesCell = row.getCell(4 + sortedDates.length);
      entriesCell.numFmt = '0';
      entriesCell.alignment = { horizontal: "center", vertical: "middle" };

      row.eachCell(cell => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } }
        };
      });
    });

    // Summenzeile im Footer
    const footValues = ["Gesamtsumme Tag", ""];
    sortedDates.forEach(d => {
      footValues.push(d.totalHours);
    });
    footValues.push(grandTotalHours);
    footValues.push(grandTotalEntries);

    const footRow = ws.addRow(footValues);
    footRow.height = 26;
    footRow.eachCell((cell, colNum) => {
      cell.font = { bold: true, color: { argb: "FF0F172A" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE2E8F0" }
      };
      cell.border = {
        top: { style: "double", color: { argb: "FF475569" } },
        bottom: { style: "medium", color: { argb: "FF475569" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
      };

      if (colNum >= 3 && colNum <= 3 + sortedDates.length) {
        cell.numFmt = '#,##0.0 "h"';
        cell.alignment = { horizontal: "right", vertical: "middle" };
      } else if (colNum === 4 + sortedDates.length) {
        cell.numFmt = '0';
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      }
    });

    // Grand Total hervorheben
    const grandCell = footRow.getCell(3 + sortedDates.length);
    grandCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFCCFBF1" }
    };
    grandCell.font = { bold: true, color: { argb: "FF0F766E" }, size: 11 };

    return await wb.xlsx.writeBuffer();
  }

  if (btnExportTimesheet) {
    btnExportTimesheet.addEventListener("click", async () => {
      try {
        const buf = await generateTimesheetExcel();
        if (!buf) return;
        const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
        downloadBuffer(buf, `Arbeitszeiten_Uebersicht_${baseName}.xlsx`);
        showToast("✅ Arbeitszeiten-Matrix erfolgreich als Excel exportiert!");
      } catch (err) {
        console.error("Fehler beim Exportieren der Arbeitszeiten:", err);
        alert("Fehler beim Erstellen der Arbeitszeiten-Excel: " + err.message);
      }
    });
  }

  async function generateStaffTemplateExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Mitarbeiter");

    ws.columns = [
      { header: "Ressourcennummer", key: "res", width: 20 },
      { header: "Nachname", key: "nachname", width: 22 },
      { header: "Vorname", key: "vorname", width: 20 },
      { header: "Abteilung", key: "abt", width: 24 },
      { header: "Bemerkung", key: "bem", width: 28 }
    ];

    const headerRow = ws.getRow(1);
    headerRow.height = 26;
    headerRow.eachCell(cell => {
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0284C7" } // Blue header
      };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      cell.border = {
        bottom: { style: "medium", color: { argb: "FF0369A1" } }
      };
    });

    const sampleData = [
      ["0045", "Mustermann", "Max", "Montage", "Leitender Monteur"],
      ["0120", "Schmidt", "Anna", "Kundendienst", "Technikerin"],
      ["0300", "Weber", "Michael", "Logistik", "Lagerleitung"],
      ["0400", "Fischer", "Sarah", "Projektleitung", "Bauleiterin"],
      ["0500", "Becker", "Thomas", "Qualitätssicherung", "Prüftechniker"],
      ["0600", "Wagner", "Julia", "Service", "Außendienst"],
      ["0700", "Hoffmann", "Stefan", "Instandhaltung", "Elektroniker"]
    ];

    sampleData.forEach(rowArr => {
      const row = ws.addRow(rowArr);
      row.height = 20;
      const resCell = row.getCell(1);
      resCell.numFmt = "0000";
      resCell.alignment = { horizontal: "center", vertical: "middle" };
    });

    return await wb.xlsx.writeBuffer();
  }

  // ==============================================================================
  // Gesamte Excel-Datei: Vollständige Tabellenansicht mit Filtern & Live-Editing
  // ==============================================================================

  function activateTab(tabName) {
    if (tabName === "fulltable") {
      if (tabBtnDiagnostics) tabBtnDiagnostics.classList.remove("active");
      if (tabBtnFulltable) tabBtnFulltable.classList.add("active");
      if (viewDiagnostics) viewDiagnostics.classList.add("hidden");
      if (viewFulltable) viewFulltable.classList.remove("hidden");
      renderFullExcelTable();
    } else {
      if (tabBtnFulltable) tabBtnFulltable.classList.remove("active");
      if (tabBtnDiagnostics) tabBtnDiagnostics.classList.add("active");
      if (viewFulltable) viewFulltable.classList.add("hidden");
      if (viewDiagnostics) viewDiagnostics.classList.remove("hidden");
    }
  }

  if (tabBtnDiagnostics) {
    tabBtnDiagnostics.addEventListener("click", () => activateTab("diagnostics"));
  }
  if (tabBtnFulltable) {
    tabBtnFulltable.addEventListener("click", () => activateTab("fulltable"));
  }

  function populateFullTableFilters() {
    if (!filterFullStaff || !filterFullDate || !state.targetWorkbook || !state.currentTargetSheet) return;

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return;

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
    const dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;

    const staffCounts = new Map();
    const dateCounts = new Map();

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;

      // Ressource
      const cellKeyRes = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");

      if (rawRes) {
        const key = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
        staffCounts.set(key, (staffCounts.get(key) || 0) + 1);
      }

      // Datum
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        const dObj = parseRowDate(cell.value, cell);
        if (dObj && dObj.isoKey) {
          if (!dateCounts.has(dObj.isoKey)) {
            dateCounts.set(dObj.isoKey, { label: dObj.label || dObj.displayDate, count: 0 });
          }
          dateCounts.get(dObj.isoKey).count += 1;
        }
      }
    }

    const curStaffVal = filterFullStaff.value;
    filterFullStaff.innerHTML = "";

    const hasStaffList = state.staffList && state.staffList.length > 0;
    const workerKeys = [];
    const nonWorkerKeys = [];

    Array.from(staffCounts.keys()).sort().forEach(res => {
      const isKnownWorker = Boolean(getStaffName(res));
      if (hasStaffList) {
        if (isKnownWorker) workerKeys.push(res);
        else nonWorkerKeys.push(res);
      } else {
        workerKeys.push(res);
      }
    });

    const optAll = document.createElement("option");
    optAll.value = "all";
    optAll.textContent = `Alle Ressourcen (${staffCounts.size})`;
    filterFullStaff.appendChild(optAll);

    if (hasStaffList && nonWorkerKeys.length > 0) {
      const optAllWorkers = document.createElement("option");
      optAllWorkers.value = "all_staff";
      optAllWorkers.textContent = `👷 Nur echte Arbeiter (${workerKeys.length})`;
      filterFullStaff.appendChild(optAllWorkers);

      const optAllNonWorkers = document.createElement("option");
      optAllNonWorkers.value = "all_non_staff";
      optAllNonWorkers.textContent = `⚙️ Nur Fremd-/Maschinen-Ressourcen (${nonWorkerKeys.length})`;
      filterFullStaff.appendChild(optAllNonWorkers);

      const groupWorkers = document.createElement("optgroup");
      groupWorkers.label = "👷 Arbeiter (aus Stammdaten)";
      workerKeys.forEach(res => {
        const name = getStaffName(res) || `Arbeiter ${res}`;
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `${name} (Nr. ${res}) [${count} Zeilen]`;
        if (res === curStaffVal) opt.selected = true;
        groupWorkers.appendChild(opt);
      });
      filterFullStaff.appendChild(groupWorkers);

      const groupNonWorkers = document.createElement("optgroup");
      groupNonWorkers.label = "⚙️ Sonstige Ressourcen (Maschinen / Fremd)";
      nonWorkerKeys.forEach(res => {
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `Ressource ${res} [${count} Zeilen] (Kein Mitarbeiter)`;
        if (res === curStaffVal) opt.selected = true;
        groupNonWorkers.appendChild(opt);
      });
      filterFullStaff.appendChild(groupNonWorkers);
    } else {
      workerKeys.forEach(res => {
        const name = getStaffName(res) || `Ressource ${res}`;
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `${name} (Nr. ${res}) [${count} Zeilen]`;
        if (res === curStaffVal) opt.selected = true;
        filterFullStaff.appendChild(opt);
      });
    }

    if (curStaffVal && (curStaffVal === "all" || curStaffVal === "all_staff" || curStaffVal === "all_non_staff" || staffCounts.has(curStaffVal))) {
      filterFullStaff.value = curStaffVal;
    }

    const curDateVal = filterFullDate.value;
    filterFullDate.innerHTML = `<option value="all">Alle Tage (${dateCounts.size})</option>`;
    Array.from(dateCounts.keys()).sort().forEach(isoKey => {
      const dInfo = dateCounts.get(isoKey);
      const opt = document.createElement("option");
      opt.value = isoKey;
      opt.textContent = `${dInfo.label} [${dInfo.count} Zeilen]`;
      if (isoKey === curDateVal) opt.selected = true;
      filterFullDate.appendChild(opt);
    });
  }

  function renderFullExcelTable() {
    if (!fulltableThead || !fulltableTbody || !state.targetWorkbook || !state.currentTargetSheet) return;

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return;

    const headerRow = ws.getRow(1);
    let colCount = ws.columnCount;
    headerRow.eachCell((cell, colNum) => {
      if (colNum > colCount) colCount = colNum;
    });

    // Kopfzeile erstellen
    let theadHtml = `<tr><th class="col-sticky-row"># Zeile</th>`;
    for (let c = 1; c <= colCount; c++) {
      const headerCell = headerRow.getCell(c);
      const colName = WebExcelEngine.extractCellValue(headerCell.value, headerCell) || `Spalte ${getColLetter(c)}`;
      theadHtml += `<th><span style="color:var(--text-muted); font-size:0.75rem;">${getColLetter(c)}:</span> ${escapeHtml(colName)}</th>`;
    }
    theadHtml += `</tr>`;
    fulltableThead.innerHTML = theadHtml;

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
    const dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;

    const staffFilter = filterFullStaff ? filterFullStaff.value : "all";
    const dateFilter = filterFullDate ? filterFullDate.value : "all";
    const statusFilter = filterFullStatus ? filterFullStatus.value : "all";
    const searchQuery = filterFullSearch ? filterFullSearch.value.trim().toLowerCase() : "";

    let totalDataRows = 0;
    let matchingRowsCount = 0;
    let tbodyHtml = "";

    const resultMap = new Map();
    if (state.allResults) {
      state.allResults.forEach(r => {
        resultMap.set(`${r.row}_${r.col_idx}`, r);
      });
    }

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;
      totalDataRows++;

      // Mitarbeiter-Filter prüfen
      const cellKeyRes = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");
      const paddedRes = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
      const unpaddedRes = rawRes.replace(/^0+/, "");
      const empName = getStaffName(rawRes);
      const isKnownWorker = Boolean(empName);

      if (staffFilter === "all_staff") {
        if (!isKnownWorker) continue;
      } else if (staffFilter === "all_non_staff") {
        if (isKnownWorker) continue;
      } else if (staffFilter !== "all") {
        const filterName = getStaffName(staffFilter);
        const matchesStaff = Boolean(empName && filterName && empName === filterName);
        const matchesKey = (rawRes === staffFilter || paddedRes === staffFilter || unpaddedRes === staffFilter);
        if (!matchesStaff && !matchesKey) {
          continue;
        }
      }

      // Datums-Filter prüfen
      let rowDateIso = "";
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        const dObj = parseRowDate(cell.value, cell);
        rowDateIso = dObj.isoKey;
      }
      if (dateFilter !== "all" && rowDateIso !== dateFilter) {
        continue;
      }

      // Status- & Such-Prüfung
      let rowHasError = false;
      let rowHasCorrection = false;
      let rowTextAcc = "";

      for (let c = 1; c <= colCount; c++) {
        const cellKey = `${r}_${c}`;
        if (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined) {
          rowHasCorrection = true;
        }
        const errObj = resultMap.get(cellKey);
        if (errObj && errObj.status !== "OK" && errObj.status !== "OK_LEER") {
          rowHasError = true;
        }
        const val = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined)
          ? String(state.appliedCorrections[cellKey])
          : WebExcelEngine.extractCellValue(row.getCell(c).value, row.getCell(c));
        rowTextAcc += " " + String(val).toLowerCase();
      }

      const empNameForSearch = empName || "";
      if (empNameForSearch) rowTextAcc += " " + empNameForSearch.toLowerCase();

      if (statusFilter === "errors" && !rowHasError) continue;
      if (statusFilter === "corrected" && !rowHasCorrection) continue;
      if (searchQuery && !rowTextAcc.includes(searchQuery)) continue;

      matchingRowsCount++;

      let rowIndicator = "";
      if (rowHasError) rowIndicator = `<span title="Enthält Prüffehler" style="color:#ef4444; font-size:0.8rem;">⚠️</span>`;
      else if (rowHasCorrection) rowIndicator = `<span title="Manuell korrigiert" style="color:#10b981; font-size:0.8rem;">✏️</span>`;

      tbodyHtml += `<tr>`;
      tbodyHtml += `<td class="col-sticky-row">${r} ${rowIndicator}</td>`;

      for (let c = 1; c <= colCount; c++) {
        const cellKey = `${r}_${c}`;
        const cell = row.getCell(c);
        const isCorrected = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined);
        const currentVal = isCorrected ? state.appliedCorrections[cellKey] : WebExcelEngine.extractCellValue(cell.value, cell);

        const errObj = resultMap.get(cellKey);
        const isError = errObj && errObj.status !== "OK" && errObj.status !== "OK_LEER" && !isCorrected;

        let cellClass = "full-cell-editable";
        let badgeHtml = "";

        if (isCorrected) {
          cellClass += " cell-corrected";
          badgeHtml = `<span class="cell-diag-badge badge-corr" title="Manuell korrigiert">✓</span>`;
        } else if (isError) {
          cellClass += " cell-error";
          badgeHtml = `<span class="cell-diag-badge badge-err" title="${escapeHtml(errObj.status_label)}: ${escapeHtml(errObj.detail)}">⚠️</span>`;
        }

        let displayStr = String(currentVal ?? "");
        if (typeof currentVal === "number" && !isNaN(currentVal) && !Number.isInteger(currentVal)) {
          displayStr = currentVal.toLocaleString("de-DE", { minimumFractionDigits: 1, maximumFractionDigits: 2 });
        }

        tbodyHtml += `
          <td class="${cellClass}" data-row="${r}" data-col="${c}" data-val="${escapeHtml(currentVal ?? '')}" title="Klick zum Ändern (Zelle ${getColLetter(c)}${r})">
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 4px;">
              <span>${escapeHtml(displayStr)}</span>
              ${badgeHtml}
            </div>
          </td>
        `;
      }
      tbodyHtml += `</tr>`;
    }

    if (matchingRowsCount === 0) {
      tbodyHtml = `<tr><td colspan="${colCount + 1}" style="text-align: center; padding: 2rem; color: var(--text-muted);">Keine Zeilen für die aktuellen Filterkriterien gefunden. Klicken Sie auf "Filter zurücksetzen".</td></tr>`;
    }

    fulltableTbody.innerHTML = tbodyHtml;

    if (fulltableCountsText) {
      fulltableCountsText.innerHTML = `Zeige <strong>${matchingRowsCount}</strong> von <strong>${totalDataRows}</strong> Zeilen`;
    }
    if (tabBadgeTotalRows) {
      tabBadgeTotalRows.textContent = `${totalDataRows} Zeilen`;
    }

    // Inline-Editor für alle Tabellenzellen aktivieren
    fulltableTbody.querySelectorAll(".full-cell-editable").forEach(td => {
      td.addEventListener("click", () => {
        makeCellEditable(td);
      });
    });
  }

  function makeCellEditable(td) {
    if (td.classList.contains("full-cell-editing")) return;
    const rowIdx = parseInt(td.dataset.row, 10);
    const colIdx = parseInt(td.dataset.col, 10);
    const origVal = td.dataset.val ?? "";

    td.classList.add("full-cell-editing");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-inline-input";
    input.value = origVal;
    td.innerHTML = "";
    td.appendChild(input);
    input.focus();
    input.select();

    let isSaved = false;
    function finishEdit() {
      if (isSaved) return;
      isSaved = true;
      const newVal = input.value.trim();
      if (newVal !== origVal) {
        saveCellCorrection(rowIdx, colIdx, newVal);
      } else {
        renderFullExcelTable();
      }
    }

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        finishEdit();
      } else if (e.key === "Escape") {
        isSaved = true;
        renderFullExcelTable();
      }
    });

    input.addEventListener("blur", () => {
      finishEdit();
    });
  }

  function saveCellCorrection(rowIdx, colIdx, newVal) {
    const cellKey = `${rowIdx}_${colIdx}`;
    state.appliedCorrections[cellKey] = newVal;

    const ws = state.targetWorkbook ? state.targetWorkbook.getWorksheet(state.currentTargetSheet) : null;
    if (ws) {
      const cell = ws.getRow(rowIdx).getCell(colIdx);
      const valStr = String(newVal).trim();
      if (colIdx === 1) {
        cell.value = valStr;
        cell.numFmt = "DD.MM.YYYY";
      } else if (/^\d+$/.test(valStr) && valStr.startsWith("0") && valStr.length > 1) {
        cell.value = valStr;
      } else if (/^\d+$/.test(valStr)) {
        cell.value = parseInt(valStr, 10);
      } else if (/^-?\d+\.\d+$/.test(valStr.replace(",", "."))) {
        cell.value = parseFloat(valStr.replace(",", "."));
      } else {
        cell.value = valStr;
      }
    }

    const resItem = state.allResults.find(item => item.id === cellKey);
    if (resItem) {
      resItem.current_value = newVal;
      resItem.is_corrected = (newVal !== resItem.original_value);
    }

    updateFilterCounts(state.allResults);
    renderResultsTable();

    // Arbeitszeiten-Matrix sofort neu berechnen
    renderTimesheetMatrix();

    // Gesamttabelle und Filter neu aktualisieren
    populateFullTableFilters();
    renderFullExcelTable();

    showToast(`✅ Zelle ${getColLetter(colIdx)}${rowIdx} aktualisiert auf "${newVal}". Arbeitszeiten & Excel synchronisiert!`);
  }

  function drillDownToFullTable(resource, dateIso) {
    activateTab("fulltable");
    if (filterFullStaff && resource) {
      filterFullStaff.value = resource;
    }
    if (filterFullDate && dateIso) {
      filterFullDate.value = dateIso;
    } else if (filterFullDate) {
      filterFullDate.value = "all";
    }
    if (filterFullStatus) filterFullStatus.value = "all";
    if (filterFullSearch) filterFullSearch.value = "";

    renderFullExcelTable();

    const fullTableEl = document.getElementById("view-fulltable");
    if (fullTableEl) {
      fullTableEl.scrollIntoView({ behavior: "smooth" });
    }

    const empName = getStaffName(resource) || `Ressource ${resource}`;
    showToast(`🔍 Zeige alle Zeilen für ${empName} ${dateIso ? 'am ' + dateIso : ''}. Klicken Sie auf eine Zelle zum Ändern.`);
  }

  // Filter Event Listener für Gesamttabelle
  [filterFullStaff, filterFullDate, filterFullStatus].forEach(sel => {
    if (sel) {
      sel.addEventListener("change", () => {
        renderFullExcelTable();
      });
    }
  });

  if (filterFullSearch) {
    filterFullSearch.addEventListener("input", () => {
      renderFullExcelTable();
    });
  }

  if (btnResetFullFilters) {
    btnResetFullFilters.addEventListener("click", () => {
      if (filterFullStaff) filterFullStaff.value = "all";
      if (filterFullDate) filterFullDate.value = "all";
      if (filterFullStatus) filterFullStatus.value = "all";
      if (filterFullSearch) filterFullSearch.value = "";
      renderFullExcelTable();
      showToast("Filter zurückgesetzt: Alle Zeilen werden angezeigt.");
    });
  }

  // --- Demo-Dateien direkt im Speicher laden ---
  btnLoadDemo.addEventListener("click", async () => {
    btnLoadDemo.disabled = true;
    btnLoadDemo.innerHTML = `<span class="icon">⏳</span> Erstelle Demo...`;

    // 0. Falls noch keine Mitarbeiter gespeichert sind, Demo-Mitarbeiter im Browser speichern
    if (state.staffList.length === 0) {
      const demoStaff = [
        { resource: "0045", name: "Max Mustermann", dept: "Montage" },
        { resource: "0120", name: "Anna Schmidt", dept: "Kundendienst" },
        { resource: "0300", name: "Michael Weber", dept: "Logistik" },
        { resource: "0400", name: "Sarah Fischer", dept: "Projektleitung" },
        { resource: "0500", name: "Thomas Becker", dept: "Qualitätssicherung" },
        { resource: "0600", name: "Julia Wagner", dept: "Service" },
        { resource: "0700", name: "Stefan Hoffmann", dept: "Instandhaltung" }
      ];
      saveStaffToStorage(demoStaff);
    }

    // 1. Referenz-Arbeitsmappe
    const refWb = new ExcelJS.Workbook();
    const refWs = refWb.addWorksheet("Stammdaten");
    refWs.columns = [
      { header: "Artikelnummer", key: "art", width: 16 },
      { header: "EAN_Code", key: "ean", width: 18 },
      { header: "Kundennummer", key: "knd", width: 16 },
      { header: "PLZ", key: "plz", width: 10 },
      { header: "Leistungsnummer", key: "leist", width: 18 },
      { header: "Resourcen Nummer", key: "res", width: 18 },
      { header: "Bezeichnung", key: "bez", width: 30 }
    ];

    const refRows = [
      ["10021", "4012345000101", "KND-80410", "10115", "12",   "0045", "Laptop Pro 15 Zoll"],
      ["10022", "4012345000102", "KND-80420", "20095", "45",   "0120", "Kabellose Maus Optical"],
      ["10023", "4012345000103", "KND-80430", "30159", "78",   "0300", "Mechanische Tastatur RGB"],
      ["10024", "4012345000104", "KND-80440", "40213", "88",   "0400", "USB-C Dockingstation"],
      ["10025", "4012345000105", "KND-80450", "50667", "1100", "0500", "Ultra-HD Monitor 27 Zoll"],
      ["10026", "4012345000106", "KND-80460", "60311", "1200", "0600", "Noise-Cancelling Headset"],
      ["10027", "4012345000107", "KND-80470", "70173", "1300", "0700", "Externe NVMe SSD 1TB"]
    ];

    refRows.forEach(r => refWs.addRow(r));
    const refBuf = await refWb.xlsx.writeBuffer();
    state.refWorkbook = refWb;
    state.refFileName = "Beispiel_Referenz_Stammdaten.xlsx";
    refFilename.textContent = state.refFileName;
    populateSheetSelect(selectRefSheet, refWb.worksheets);
    state.currentRefSheet = selectRefSheet.value;
    refFileInfo.classList.remove("hidden");
    dropRef.querySelector(".drop-zone-content").classList.add("hidden");

    // Falls noch keine Referenzdatei im Browser gespeichert ist, Demo-Referenz speichern
    const storedRef = await loadRefFileFromStorage();
    if (!storedRef) {
      await saveRefFileToStorage(state.refFileName, refBuf, state.currentRefSheet);
      if (refStorageBadge) {
        refStorageBadge.textContent = "Im Browser gespeichert";
        refStorageBadge.className = "badge badge-ok";
      }
      if (refInfoDetail) {
        refInfoDetail.textContent = "💾 Im Browser gespeichert (dauerhaft erhalten)";
      }
    } else {
      if (refStorageBadge) {
        refStorageBadge.textContent = "Demo aktiv";
        refStorageBadge.className = "badge badge-warning";
      }
      if (refInfoDetail) {
        refInfoDetail.textContent = "Demo-Daten aktiv (gespeicherte Datei bleibt erhalten)";
      }
    }

    // 2. Prüfdatei-Arbeitsmappe (mit Datum, Menge/Stunden und Ressourcen)
    const tgtWb = new ExcelJS.Workbook();
    const tgtWs = tgtWb.addWorksheet("Report");
    tgtWs.columns = [
      { header: "Datum", key: "dat", width: 14 },
      { header: "Artikelnummer", key: "art", width: 16 },
      { header: "Resourcen Nummer", key: "res", width: 18 },
      { header: "Menge", key: "mng", width: 10 },
      { header: "Leistung", key: "leist", width: 16 },
      { header: "EAN_Code", key: "ean", width: 18 }
    ];

    const tgtRows = [
      ["15.09.2026", "10021", "0045", 8.0, "",     "4012345000101"], // Leistung leer -> OK_LEER; 8 Std für Max Mustermann
      ["15.09.2026", "10022", "45",   4.5, "45",   "4012345000102"], // Res '45' -> wird '0045'; Leistung '45' OK; 4.5 Std für Max
      ["16.09.2026", "10023", "120",  7.5, "78",   "4012345000103"], // Res '120' -> wird '0120'; Leistung '78' OK; 7.5 Std für Anna Schmidt
      ["16.09.2026", "10042", "0400", 8.0, "54",   "4012345000104"], // Artikel Zahlendreher (10042); Leistung Zahlendreher ('54'); 8 Std Sarah Fischer
      ["17.09.2026", "10025", "0500", 6.0, "1100", "4012345000150"], // EAN Zahlendreher; 6 Std Thomas Becker
      ["17.09.2026", "99999", "0700", 8.5, "",     "4012345000107"], // 99999 Nicht existent; 8.5 Std Stefan Hoffmann
      ["17.09.2026", "10026", "9900", 4.0, "1200", "4012345000106"]  // Fremd-Ressource 9900 (Maschine/Bagger - kein Mitarbeiter)
    ];

    tgtRows.forEach(r => tgtWs.addRow(r));
    const tgtBuf = await tgtWb.xlsx.writeBuffer();
    state.targetWorkbook = tgtWb;
    state.lastTargetBuffer = tgtBuf;
    state.targetFileName = "Beispiel_Zu_Pruefen.xlsx";
    tgtFilename.textContent = state.targetFileName;
    populateSheetSelect(selectTgtSheet, tgtWb.worksheets);
    state.currentTargetSheet = selectTgtSheet.value;
    tgtFileInfo.classList.remove("hidden");
    dropTgt.querySelector(".drop-zone-content").classList.add("hidden");

    if (tgtStatusBadge) {
      tgtStatusBadge.textContent = "Bereit zur Prüfung";
      tgtStatusBadge.className = "badge badge-ok";
      tgtStatusBadge.style.display = "inline-block";
    }

    btnLoadDemo.disabled = false;
    btnLoadDemo.innerHTML = `<span class="icon">✨</span> Demo-Dateien laden`;

    checkReadyForConfig();
    showToast("Demo-Dateien & Mitarbeiter geladen! Klicken Sie auf 'Prüfung starten'.");
  });

  function getColLetter(colIdx) {
    let temp = "";
    let letter = "";
    while (colIdx > 0) {
      temp = (colIdx - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIdx = (colIdx - temp - 1) / 26;
    }
    return letter;
  }

  let toastTimer = null;
  function showToast(msg) {
    if (toastTimer) clearTimeout(toastTimer);
    toast.textContent = msg;
    toast.classList.remove("hidden");
    toastTimer = setTimeout(() => {
      toast.classList.add("hidden");
    }, 3500);
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
});
