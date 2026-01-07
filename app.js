(() => {
  "use strict";

  const STORAGE_KEY = "multithreader.v1";
  const DEFAULT_ROWS = 21;
  const DEFAULT_COLS = 10;
  const MIN_SPLIT = 0.2;
  const MAX_SPLIT = 0.8;
  const MIN_COL_WIDTH = 60;
  const MIN_ROW_HEIGHT = 32;
  const UNDO_LIMIT = 10;
  const HELP_TEXT = [
    "MULTITHREADER HELP",
    "",
    "Views: All (A), Scheme (S), Data (D), Swap (X), Zen (Z).",
    "Threads: New/Rename/Delete, drag tabs to reorder, double-click or F2 to rename.",
    "Break timer: set HH:MM, quick 30/60/90, snooze 10m.",
    "Find/Replace: press F to toggle, target Scheme/Data/Both.",
    "Scheme: plain text only, paste strips formatting, Timestamp inserts time.",
    "Data: Find/Replace in the panel, formulas SUM/AVG/MIN/MAX/COUNT.",
    "Copy Scheme: button beside the status chip.",
    "Data: formulas SUM/AVG/MIN/MAX/COUNT, sort by column, shade, align, resize rows/cols.",
    "Selection: drag to select multiple cells; shade/align apply to selection.",
    "Copy CSV: button beside the status chip.",
    "Export: TXT, CSV, ZIP. Print: split or separate pages.",
    "Import: drag CSV onto Data or use Import CSV.",
    "Cache: autosaves locally. Optional cache file writes when supported (secure context).",
    "Backup/Restore: save a portable JSON and load it on another machine."
  ].join("\n");

  const app = document.getElementById("app");
  const tabsEl = document.getElementById("threadTabs");
  const newThreadBtn = document.getElementById("newThread");
  const renameThreadBtn = document.getElementById("renameThread");
  const deleteThreadBtn = document.getElementById("deleteThread");
  const breakTimerBtn = document.getElementById("breakTimerBtn");
  const breakTimerPanel = document.getElementById("breakTimerPanel");
  const breakHoursInput = document.getElementById("breakHours");
  const breakMinutesInput = document.getElementById("breakMinutes");
  const breakStartBtn = document.getElementById("breakStart");
  const breakPauseBtn = document.getElementById("breakPause");
  const breakResetBtn = document.getElementById("breakReset");
  const breakRemainingEl = document.getElementById("breakRemaining");
  const breakAlarmModal = document.getElementById("breakAlarmModal");
  const breakSnoozeBtn = document.getElementById("breakSnooze");
  const breakDismissBtn = document.getElementById("breakDismiss");

  const viewButtons = Array.from(document.querySelectorAll(".view-controls .pill[data-view]"));
  const swapBtn = document.getElementById("swapPanels");

  const exportTxtBtn = document.getElementById("exportTxt");
  const exportCsvBtn = document.getElementById("exportCsv");
  const exportBothBtn = document.getElementById("exportBoth");
  const exportZipBtn = document.getElementById("exportZip");

  const printModeSelect = document.getElementById("printMode");
  const printBtn = document.getElementById("printPdf");
  const pageSettingsToggle = document.getElementById("pageSettingsToggle");
  const pageSettingsPanel = document.getElementById("pageSettingsPanel");
  const importCsvBtn = document.getElementById("importCsv");
  const exportBundleBtn = document.getElementById("exportBundle");
  const openDiagnosticsBtn = document.getElementById("openDiagnostics");
  const toggleZenBtn = document.getElementById("toggleZen");
  const csvFileInput = document.getElementById("csvFileInput");
  const backupFileInput = document.getElementById("backupFileInput");
  const openHelpBtn = document.getElementById("openHelp");
  const setCacheLocationBtn = document.getElementById("setCacheLocation");
  const viewCacheBtn = document.getElementById("viewCache");
  const backupCacheBtn = document.getElementById("backupCache");
  const restoreBackupBtn = document.getElementById("restoreBackup");
  const clearCacheBtn = document.getElementById("clearCache");
  const cacheLocationLabel = document.getElementById("cacheLocationLabel");
  const cacheLocationText = document.getElementById("cacheLocationText");
  const cacheLocationLink = document.getElementById("cacheLocationLink");

  const schemeSettingsToggle = document.getElementById("schemeSettingsToggle");
  const schemePopover = document.getElementById("schemePopover");
  const dataOptionsToggle = document.getElementById("dataOptionsToggle");
  const dataOptionsPopover = document.getElementById("dataOptionsPopover");

  const journalEl = document.getElementById("journal");
  const journalFontEl = document.getElementById("journalFont");
  const journalSizeEl = document.getElementById("journalSize");
  const journalColorEl = document.getElementById("journalColor");
  const journalUndoBtn = document.getElementById("journalUndo");
  const journalFindBtn = document.getElementById("journalFind");
  const journalTimestampBtn = document.getElementById("journalTimestamp");
  const journalStatus = document.getElementById("journalStatus");
  const clearSchemeBtn = document.getElementById("clearScheme");

  const wordCountEl = document.getElementById("wordCount");
  const charCountEl = document.getElementById("charCount");
  const tokenCountEl = document.getElementById("tokenCount");

  const findPanel = document.getElementById("findPanel");
  const findHeader = document.getElementById("findHeader");
  const findInput = document.getElementById("findInput");
  const replaceInput = document.getElementById("replaceInput");
  const matchCaseInput = document.getElementById("matchCase");
  const findNextBtn = document.getElementById("findNext");
  const replaceOneBtn = document.getElementById("replaceOne");
  const replaceAllBtn = document.getElementById("replaceAll");
  const closeFindBtn = document.getElementById("closeFind");
  const findTargetButtons = Array.from(document.querySelectorAll("[data-find-target]"));

  const sheetFontEl = document.getElementById("sheetFont");
  const sheetSizeEl = document.getElementById("sheetSize");
  const sheetColorEl = document.getElementById("sheetColor");
  const copyCsvBtn = document.getElementById("copyCsv");
  const clearSheetBtn = document.getElementById("clearSheet");
  const copySchemeBtn = document.getElementById("copyScheme");
  const sheetUndoBtn = document.getElementById("sheetUndo");
  const sheetFindBtn = document.getElementById("sheetFind");
  const addRowBtn = document.getElementById("addRow");
  const addColBtn = document.getElementById("addCol");
  const removeRowBtn = document.getElementById("removeRow");
  const removeColBtn = document.getElementById("removeCol");
  const sortColumnEl = document.getElementById("sortColumn");
  const sortOrderEl = document.getElementById("sortOrder");
  const sortApplyBtn = document.getElementById("sortApply");
  const shadeColorEl = document.getElementById("shadeColor");
  const applyShadeBtn = document.getElementById("applyShade");
  const clearShadeBtn = document.getElementById("clearShade");
  const alignLeftBtn = document.getElementById("alignLeft");
  const alignCenterBtn = document.getElementById("alignCenter");
  const alignRightBtn = document.getElementById("alignRight");
  const sheetStatus = document.getElementById("sheetStatus");

  const cellLabelEl = document.getElementById("cellLabel");
  const formulaInput = document.getElementById("formulaInput");
  const sheetTable = document.getElementById("sheetTable");
  const panelsEl = document.querySelector(".panels");
  const dividerEl = document.getElementById("panelDivider");
  const journalPanel = document.querySelector(".journal-panel");
  const sheetPanel = document.querySelector(".sheet-panel");

  const renameModal = document.getElementById("renameModal");
  const renameInput = document.getElementById("renameInput");
  const renameColorInput = document.getElementById("renameColor");
  const clearRenameColorBtn = document.getElementById("clearRenameColor");
  const sheetContextMenu = document.getElementById("sheetContextMenu");
  const cellContextMenu = document.getElementById("cellContextMenu");
  const cellEditMenu = document.getElementById("cellEditMenu");
  const cellAlignMenu = document.getElementById("cellAlignMenu");
  const cellOptionsMenu = document.getElementById("cellOptionsMenu");
  const cellFormulaMenu = document.getElementById("cellFormulaMenu");
  const cellTextColorInput = document.getElementById("cellTextColor");
  const clearCellTextColorBtn = document.getElementById("clearCellTextColor");
  const cellBgColorInput = document.getElementById("cellBgColor");
  const clearCellBgColorBtn = document.getElementById("clearCellBgColor");
  const schemeContextMenu = document.getElementById("schemeContextMenu");
  const schemeTextColorInput = document.getElementById("schemeTextColor");
  const clearSchemeTextColorBtn = document.getElementById("clearSchemeTextColor");
  const tabsContextMenu = document.getElementById("tabsContextMenu");
  const cancelRenameBtn = document.getElementById("cancelRename");
  const confirmRenameBtn = document.getElementById("confirmRename");
  const cacheViewModal = document.getElementById("cacheViewModal");
  const cacheViewContent = document.getElementById("cacheViewContent");
  const closeCacheViewBtn = document.getElementById("closeCacheView");
  const diagnosticsModal = document.getElementById("diagnosticsModal");
  const diagnosticsContent = document.getElementById("diagnosticsContent");
  const closeDiagnosticsBtn = document.getElementById("closeDiagnostics");

  const journalUndoStacks = new Map();
  const journalLast = new Map();
  const sheetUndoStacks = new Map();

  const CACHE_HANDLE_DB = "multithreader-cache";
  const CACHE_HANDLE_STORE = "handles";
  const CACHE_HANDLE_KEY = "cacheFile";
  const titleEl = document.querySelector(".title");
  const taglineEl = document.querySelector(".tagline");

  let state = loadState();
  let selectedCell = { row: 0, col: 0 };
  let saveTimer = null;
  let sheetInputs = [];
  let sheetInputList = [];
  let sheetEditSession = null;
  let dragState = null;
  let resizeActive = false;
  let columnResize = null;
  let rowResize = null;
  let selectionRange = null;
  let selectionAnchor = null;
  let selectionActive = false;
  let ignoreFocusSelection = false;
  let cellEditMode = false;
  let renameTargetId = null;
  let renameColorTouched = false;
  let renameColorCleared = false;
  let cacheHandle = null;
  let cacheWriteInFlight = false;
  let cacheWriteQueued = false;
  let cacheHandlePromise = null;
  let breakTickTimer = null;
  let breakReminderTimer = null;
  let focusStartAt = null;
  let accumulatedFocusMs = 0;
  let reminderStage = 0;
  let reminderPaused = false;
  let breakAlarmActive = false;
  let contextMenuTarget = null;
  let cellContextTarget = null;
  let cellContextPosition = null;
  let activeCellSubmenu = null;
  let activeCellSubmenuTrigger = null;
  let schemeContextActive = false;
  let tabsContextTargetId = null;
  let lastSheetCopy = null;
  let sheetFindPointer = null;
  let sheetFindLastQuery = "";
  let lastActivePanel = "journal";
  let findTarget = "journal";
  let findDragState = null;

  if (!state) {
    state = buildDefaultState();
  }

  normalizeState(state);
  renderTabs();
  loadActiveThread();
  bindEvents();
  updateCacheLabel();
  updateFindTargetButtons();
  initCacheHandle();
  initBreakTimer();
  syncTitleTaglineWidth();
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(syncTitleTaglineWidth);
  }

  function buildDefaultState() {
    const first = createThread("Idea 1");
    return {
      activeId: first.id,
      threads: [first],
      ui: {
        view: "split",
        swap: false,
        splitRatio: 0.5,
        zen: false,
        cacheFileName: "",
        breakTimer: {
          lastSetSeconds: 1800,
          remainingSeconds: 0,
          running: false,
          endAt: 0
        }
      }
    };
  }

  function createThread(name) {
    return {
      id: `t-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      name,
      color: "",
      isHelp: false,
      journal: "",
      settings: {
        journalFont: "space",
        journalSize: 16,
        journalColor: "#d9dde7",
        sheetFont: "space",
        sheetSize: 14,
        sheetColor: "#d9dde7"
      },
      sheet: makeSheet(DEFAULT_ROWS, DEFAULT_COLS)
    };
  }

  function makeSheet(rows, cols) {
    const cells = [];
    for (let r = 0; r < rows; r += 1) {
      const row = [];
      for (let c = 0; c < cols; c += 1) {
        row.push({ value: "", bg: "", align: "left", fg: "", wrap: "truncate" });
      }
      cells.push(row);
    }
    const colWidths = Array.from({ length: cols }, () => null);
    const rowHeights = Array.from({ length: rows }, () => null);
    return { rows, cols, cells, colWidths, rowHeights };
  }

  function normalizeSheet(sheet) {
    if (!sheet) return;
    if (!Array.isArray(sheet.cells)) sheet.cells = [];
    if (!sheet.rows || !sheet.cols) {
      sheet.rows = sheet.cells.length;
      sheet.cols = sheet.cells[0] ? sheet.cells[0].length : 0;
    }
    sheet.cells.forEach((row) => {
      row.forEach((cell) => {
        if (!cell || typeof cell !== "object") return;
        if (typeof cell.value !== "string") cell.value = "";
        if (typeof cell.bg !== "string") cell.bg = "";
        if (typeof cell.align !== "string") cell.align = "left";
        if (typeof cell.fg !== "string") cell.fg = "";
        if (typeof cell.wrap !== "string") cell.wrap = "truncate";
      });
    });
    if (sheet.rows < DEFAULT_ROWS) {
      for (let r = sheet.rows; r < DEFAULT_ROWS; r += 1) {
        const row = [];
        for (let c = 0; c < sheet.cols; c += 1) {
          row.push({ value: "", bg: "", align: "left", fg: "", wrap: "truncate" });
        }
        sheet.cells.push(row);
      }
      sheet.rows = DEFAULT_ROWS;
    }
    if (!Array.isArray(sheet.colWidths)) {
      sheet.colWidths = Array.from({ length: sheet.cols }, () => null);
    }
    if (!Array.isArray(sheet.rowHeights)) {
      sheet.rowHeights = Array.from({ length: sheet.rows }, () => null);
    }
    while (sheet.colWidths.length < sheet.cols) sheet.colWidths.push(null);
    if (sheet.colWidths.length > sheet.cols) sheet.colWidths = sheet.colWidths.slice(0, sheet.cols);
    while (sheet.rowHeights.length < sheet.rows) sheet.rowHeights.push(null);
    if (sheet.rowHeights.length > sheet.rows) sheet.rowHeights = sheet.rowHeights.slice(0, sheet.rows);
  }

  function normalizeState(data) {
    if (!data.ui) data.ui = { view: "split", swap: false, splitRatio: 0.5, zen: false, cacheFileName: "" };
    if (typeof data.ui.splitRatio !== "number") data.ui.splitRatio = 0.5;
    if (typeof data.ui.zen !== "boolean") data.ui.zen = false;
    if (typeof data.ui.cacheFileName !== "string") data.ui.cacheFileName = "";
    if (!data.ui.breakTimer) {
      data.ui.breakTimer = { lastSetSeconds: 1800, remainingSeconds: 0, running: false, endAt: 0 };
    }
    if (typeof data.ui.breakTimer.lastSetSeconds !== "number") data.ui.breakTimer.lastSetSeconds = 1800;
    if (typeof data.ui.breakTimer.remainingSeconds !== "number") data.ui.breakTimer.remainingSeconds = 0;
    if (typeof data.ui.breakTimer.running !== "boolean") data.ui.breakTimer.running = false;
    if (typeof data.ui.breakTimer.endAt !== "number") data.ui.breakTimer.endAt = 0;
    if (!Array.isArray(data.threads) || data.threads.length === 0) {
      const first = createThread("Idea 1");
      data.threads = [first];
      data.activeId = first.id;
    }
    data.threads.forEach((thread, index) => {
      if (!thread.id) thread.id = `t-${index}-${Date.now().toString(36)}`;
      if (!thread.name) thread.name = `Idea ${index + 1}`;
      if (typeof thread.color !== "string") thread.color = "";
      const normalizedColor = normalizeHexColor(thread.color);
      thread.color = normalizedColor || "";
      if (typeof thread.isHelp !== "boolean") thread.isHelp = false;
      if (!thread.settings) {
        thread.settings = {
          journalFont: "space",
          journalSize: 16,
          journalColor: "#d9dde7",
          sheetFont: "space",
          sheetSize: 14,
          sheetColor: "#d9dde7"
        };
      }
      if (!thread.sheet || !thread.sheet.cells) {
        thread.sheet = makeSheet(DEFAULT_ROWS, DEFAULT_COLS);
      } else {
        if (!thread.sheet.rows || !thread.sheet.cols) {
          thread.sheet.rows = thread.sheet.cells.length;
          thread.sheet.cols = thread.sheet.cells[0] ? thread.sheet.cells[0].length : 0;
        }
      }
      normalizeSheet(thread.sheet);
    });
    if (!data.activeId || !data.threads.find((t) => t.id === data.activeId)) {
      data.activeId = data.threads[0].id;
    }
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (err) {
      return null;
    }
  }

  function scheduleSave() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      } catch (err) {
        journalStatus.textContent = "Autosave failed";
      }
      queueCacheWrite();
    }, 350);
  }

  function getActiveThread() {
    return state.threads.find((thread) => thread.id === state.activeId);
  }

  function renderTabs() {
    tabsEl.innerHTML = "";
    state.threads.forEach((thread) => {
      const btn = document.createElement("button");
      btn.className = "tab" + (thread.id === state.activeId ? " active" : "");
      btn.dataset.threadId = thread.id;
      const label = document.createElement("span");
      label.className = "tab-label";
      label.textContent = thread.name;
      btn.appendChild(label);
      btn.style.removeProperty("--tab-dot-color");
      if (thread.color) {
        btn.style.setProperty("--tab-dot-color", thread.color);
      }
      btn.draggable = true;
      btn.addEventListener("click", () => {
        setActiveThread(thread.id);
      });
      btn.addEventListener("dblclick", (event) => {
        event.preventDefault();
        event.stopPropagation();
        renameThread(thread);
      });
      btn.addEventListener("keydown", (event) => {
        if (event.key === "F2") {
          event.preventDefault();
          renameThread(thread);
        }
      });
      btn.addEventListener("dragstart", (event) => handleTabDragStart(event, thread.id));
      btn.addEventListener("dragover", (event) => handleTabDragOver(event, thread.id));
      btn.addEventListener("dragleave", (event) => handleTabDragLeave(event));
      btn.addEventListener("drop", (event) => handleTabDrop(event, thread.id));
      btn.addEventListener("dragend", (event) => handleTabDragEnd(event));
      tabsEl.appendChild(btn);
    });
  }

  function createAndActivateThread() {
    const name = `Idea ${state.threads.length + 1}`;
    const thread = createThread(name);
    state.threads.push(thread);
    setActiveThread(thread.id);
  }

  function setActiveThread(id) {
    state.activeId = id;
    renderTabs();
    loadActiveThread();
    scheduleSave();
  }

  function loadActiveThread() {
    const thread = getActiveThread();
    if (!thread) return;
    journalEl.value = thread.journal || "";
    applySettings(thread);
    updateStats();
    normalizeSheet(thread.sheet);
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    journalLast.set(thread.id, journalEl.value);
    sheetFindPointer = null;
    sheetFindLastQuery = "";
    updateViewState();
  }

  function updateViewState() {
    app.classList.remove("view-split", "view-journal", "view-sheet");
    app.classList.add(`view-${state.ui.view}`);
    app.classList.toggle("swap", !!state.ui.swap);
    viewButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.view === state.ui.view);
    });
    document.body.classList.toggle("zen", !!state.ui.zen);
    closePanelPopovers();
    closePageSettings();
    applySplitRatio();
  }

  function setView(view) {
    if (!view) return;
    state.ui.view = view;
    if (view === "sheet") lastActivePanel = "sheet";
    if (view === "journal") lastActivePanel = "journal";
    updateViewState();
    scheduleSave();
  }

  function applySplitRatio() {
    if (!journalPanel || !sheetPanel) return;
    if (state.ui.view !== "split") {
      journalPanel.style.flex = "";
      journalPanel.style.flexBasis = "";
      sheetPanel.style.flex = "";
      sheetPanel.style.flexBasis = "";
      return;
    }
    const ratio = clamp(state.ui.splitRatio ?? 0.5, MIN_SPLIT, MAX_SPLIT);
    state.ui.splitRatio = ratio;
    const leftPanel = state.ui.swap ? sheetPanel : journalPanel;
    const rightPanel = state.ui.swap ? journalPanel : sheetPanel;
    leftPanel.style.flex = `0 0 ${ratio * 100}%`;
    rightPanel.style.flex = `0 0 ${(1 - ratio) * 100}%`;
  }

  function handleDividerPointerDown(event) {
    if (state.ui.view !== "split") return;
    resizeActive = true;
    document.body.classList.add("resizing");
    if (event.target && event.target.setPointerCapture) {
      event.target.setPointerCapture(event.pointerId);
    }
  }

  function handleDividerPointerMove(event) {
    if (!resizeActive || !panelsEl) return;
    const rect = panelsEl.getBoundingClientRect();
    const ratio = clamp((event.clientX - rect.left) / rect.width, MIN_SPLIT, MAX_SPLIT);
    state.ui.splitRatio = ratio;
    applySplitRatio();
  }

  function handleDividerPointerUp(event) {
    if (!resizeActive) return;
    resizeActive = false;
    document.body.classList.remove("resizing");
    if (event.target && event.target.releasePointerCapture) {
      try {
        event.target.releasePointerCapture(event.pointerId);
      } catch (err) {
        // ignore
      }
    }
    scheduleSave();
  }

  function handlePointerMove(event) {
    handleDividerPointerMove(event);
    handleSheetResizeMove(event);
    handleFindDragMove(event);
  }

  function handlePointerUp(event) {
    handleDividerPointerUp(event);
    handleSheetResizeEnd(event);
    handleFindDragEnd(event);
  }

  function handleFindDragStart(event) {
    if (!findPanel || !findHeader) return;
    if (event.button !== 0) return;
    if (event.target.closest("button") || event.target.closest("input")) return;
    const rect = findPanel.getBoundingClientRect();
    findDragState = {
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top,
      pointerId: event.pointerId
    };
    findPanel.classList.add("dragging");
    if (findHeader.setPointerCapture) {
      try {
        findHeader.setPointerCapture(event.pointerId);
      } catch (err) {
        // ignore
      }
    }
  }

  function handleFindDragMove(event) {
    if (!findDragState || !findPanel) return;
    const padding = 8;
    const rect = findPanel.getBoundingClientRect();
    const maxLeft = window.innerWidth - rect.width - padding;
    const maxTop = window.innerHeight - rect.height - padding;
    const left = clamp(event.clientX - findDragState.offsetX, padding, Math.max(padding, maxLeft));
    const top = clamp(event.clientY - findDragState.offsetY, padding, Math.max(padding, maxTop));
    findPanel.style.left = `${left}px`;
    findPanel.style.top = `${top}px`;
    findPanel.style.right = "auto";
    findPanel.style.bottom = "auto";
  }

  function handleFindDragEnd(event) {
    if (!findDragState || !findPanel || !findHeader) return;
    if (findHeader.releasePointerCapture) {
      try {
        findHeader.releasePointerCapture(findDragState.pointerId);
      } catch (err) {
        // ignore
      }
    }
    findDragState = null;
    findPanel.classList.remove("dragging");
  }

  function handleColumnResizeStart(event) {
    event.preventDefault();
    event.stopPropagation();
    const col = Number(event.currentTarget.dataset.col);
    if (Number.isNaN(col)) return;
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const header = event.currentTarget.parentElement;
    const currentWidth = thread.sheet.colWidths[col];
    const startWidth = Number.isFinite(currentWidth)
      ? currentWidth
      : (header ? header.getBoundingClientRect().width : 120);
    columnResize = {
      col,
      startX: event.clientX,
      startWidth,
      target: event.currentTarget,
      pointerId: event.pointerId
    };
    document.body.classList.add("resizing");
    if (event.currentTarget.setPointerCapture) {
      event.currentTarget.setPointerCapture(event.pointerId);
    }
  }

  function handleRowResizeStart(event) {
    event.preventDefault();
    event.stopPropagation();
    const row = Number(event.currentTarget.dataset.row);
    if (Number.isNaN(row)) return;
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const rowEl = sheetTable.querySelector(`tbody tr[data-row="${row}"]`);
    const currentHeight = thread.sheet.rowHeights[row];
    const startHeight = Number.isFinite(currentHeight)
      ? currentHeight
      : (rowEl ? rowEl.getBoundingClientRect().height : 32);
    rowResize = {
      row,
      startY: event.clientY,
      startHeight,
      target: event.currentTarget,
      pointerId: event.pointerId
    };
    document.body.classList.add("resizing-row");
    if (event.currentTarget.setPointerCapture) {
      event.currentTarget.setPointerCapture(event.pointerId);
    }
  }

  function handleSheetResizeMove(event) {
    if (columnResize) {
      event.preventDefault();
      const width = Math.max(MIN_COL_WIDTH, columnResize.startWidth + (event.clientX - columnResize.startX));
      setColumnWidth(columnResize.col, Math.round(width));
    }
    if (rowResize) {
      event.preventDefault();
      const height = Math.max(MIN_ROW_HEIGHT, rowResize.startHeight + (event.clientY - rowResize.startY));
      setRowHeight(rowResize.row, Math.round(height));
    }
  }

  function handleSheetResizeEnd(event) {
    let didResize = false;
    if (columnResize) {
      if (columnResize.target && columnResize.target.releasePointerCapture) {
        try {
          columnResize.target.releasePointerCapture(columnResize.pointerId);
        } catch (err) {
          // ignore
        }
      }
      columnResize = null;
      document.body.classList.remove("resizing");
      didResize = true;
    }
    if (rowResize) {
      if (rowResize.target && rowResize.target.releasePointerCapture) {
        try {
          rowResize.target.releasePointerCapture(rowResize.pointerId);
        } catch (err) {
          // ignore
        }
      }
      rowResize = null;
      document.body.classList.remove("resizing-row");
      didResize = true;
    }
    if (didResize) scheduleSave();
  }

  function setColumnWidth(col, width) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    thread.sheet.colWidths[col] = width;
    const colEl = sheetTable.querySelector(`col[data-col="${col}"]`);
    if (colEl) colEl.style.width = `${width}px`;
  }

  function setRowHeight(rowIndex, height) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    thread.sheet.rowHeights[rowIndex] = height;
    const row = sheetTable.querySelector(`tbody tr[data-row="${rowIndex}"]`);
    if (!row) return;
    row.style.height = `${height}px`;
    const rowHeader = row.querySelector("th.sticky-col");
    if (rowHeader) rowHeader.style.height = `${height}px`;
    row.querySelectorAll("td").forEach((cell) => {
      cell.style.height = `${height}px`;
    });
  }

  function setZen(enabled) {
    state.ui.zen = enabled;
    document.body.classList.toggle("zen", enabled);
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    scheduleSave();
  }

  function toggleZen() {
    setZen(!state.ui.zen);
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function handleTabDragStart(event, threadId) {
    dragState = { id: threadId, dropId: null, before: true };
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      try {
        event.dataTransfer.setData("text/plain", threadId);
      } catch (err) {
        // ignore
      }
    }
    event.currentTarget.classList.add("dragging");
  }

  function handleTabDragOver(event, threadId) {
    if (!dragState || dragState.id === threadId) return;
    event.preventDefault();
    const rect = event.currentTarget.getBoundingClientRect();
    const before = event.clientX < rect.left + rect.width / 2;
    dragState.dropId = threadId;
    dragState.before = before;
    updateTabDropIndicator(event.currentTarget, before);
  }

  function handleTabDragLeave(event) {
    event.currentTarget.classList.remove("drop-before", "drop-after");
  }

  function handleTabDrop(event, threadId) {
    if (!dragState || dragState.id === threadId) return;
    event.preventDefault();
    event.stopPropagation();
    const rect = event.currentTarget.getBoundingClientRect();
    const before = event.clientX < rect.left + rect.width / 2;
    reorderThreads(dragState.id, threadId, before);
    dragState = null;
    renderTabs();
    scheduleSave();
  }

  function handleTabDragEnd(event) {
    event.currentTarget.classList.remove("dragging");
    clearTabDropIndicators();
    dragState = null;
  }

  function handleTabsDragOver(event) {
    if (!dragState) return;
    event.preventDefault();
    const target = event.target.closest(".tab");
    if (!target) {
      clearTabDropIndicators();
    }
  }

  function handleTabsDrop(event) {
    if (!dragState) return;
    event.preventDefault();
    const target = event.target.closest(".tab");
    if (target) return;
    reorderThreads(dragState.id, null, false);
    dragState = null;
    renderTabs();
    scheduleSave();
  }

  function handleTabsContextMenu(event) {
    if (!tabsContextMenu) return;
    const tab = event.target.closest(".tab");
    event.preventDefault();
    event.stopPropagation();
    tabsContextTargetId = tab ? tab.dataset.threadId : state.activeId;
    openTabsContextMenu(event.clientX, event.clientY);
  }

  function updateTabDropIndicator(target, before) {
    clearTabDropIndicators();
    target.classList.add(before ? "drop-before" : "drop-after");
  }

  function clearTabDropIndicators() {
    tabsEl.querySelectorAll(".tab.drop-before, .tab.drop-after").forEach((tab) => {
      tab.classList.remove("drop-before", "drop-after");
    });
  }

  function reorderThreads(dragId, targetId, before) {
    const fromIndex = state.threads.findIndex((thread) => thread.id === dragId);
    if (fromIndex < 0) return;
    const [item] = state.threads.splice(fromIndex, 1);
    if (!targetId) {
      state.threads.push(item);
      return;
    }
    const targetIndex = state.threads.findIndex((thread) => thread.id === targetId);
    if (targetIndex < 0) {
      state.threads.push(item);
      return;
    }
    const insertIndex = before ? targetIndex : targetIndex + 1;
    state.threads.splice(insertIndex, 0, item);
  }

  function applySettings(thread) {
    const settings = thread.settings;
    journalFontEl.value = settings.journalFont;
    journalSizeEl.value = settings.journalSize;
    journalColorEl.value = settings.journalColor;
    sheetFontEl.value = settings.sheetFont;
    sheetSizeEl.value = settings.sheetSize;
    sheetColorEl.value = settings.sheetColor;

    journalEl.style.fontFamily = settings.journalFont === "mono" ? "var(--font-mono)" : "var(--font-sans)";
    journalEl.style.fontSize = `${settings.journalSize}px`;
    journalEl.style.color = settings.journalColor;

    sheetTable.style.fontFamily = settings.sheetFont === "mono" ? "var(--font-mono)" : "var(--font-sans)";
    sheetTable.style.fontSize = `${settings.sheetSize}px`;
    sheetTable.style.color = settings.sheetColor;
    formulaInput.style.fontFamily = settings.sheetFont === "mono" ? "var(--font-mono)" : "var(--font-sans)";
    formulaInput.style.color = settings.sheetColor;
    formulaInput.style.fontSize = `${settings.sheetSize}px`;
  }

  function updateStats() {
    const text = journalEl.value;
    const trimmed = text.trim();
    const words = trimmed ? trimmed.split(/\s+/).length : 0;
    wordCountEl.textContent = String(words);
    charCountEl.textContent = String(text.length);
    tokenCountEl.textContent = String(Math.ceil(text.length / 4));
  }

  function bindEvents() {
    tabsEl.addEventListener("dragover", handleTabsDragOver);
    tabsEl.addEventListener("drop", handleTabsDrop);
    tabsEl.addEventListener("contextmenu", handleTabsContextMenu);
    document.addEventListener("keydown", handleGlobalKeydown);
    document.addEventListener("click", () => {
      closePageSettings();
      closePanelPopovers();
      closeBreakTimerPanel();
      closeSheetContextMenu();
      closeCellContextMenu();
      closeCellSubmenus();
      closeSchemeContextMenu();
      closeTabsContextMenu();
      closeCacheViewModal();
      closeDiagnosticsModal();
    });

    if (pageSettingsToggle) {
      pageSettingsToggle.addEventListener("click", (event) => {
        event.stopPropagation();
        togglePageSettings();
      });
    }
    if (pageSettingsPanel) {
      pageSettingsPanel.addEventListener("click", (event) => {
        event.stopPropagation();
      });
    }
    if (schemeSettingsToggle) {
      schemeSettingsToggle.addEventListener("click", (event) => {
        event.stopPropagation();
        togglePopover(schemePopover, schemeSettingsToggle);
      });
    }
    if (schemePopover) {
      schemePopover.addEventListener("click", (event) => {
        event.stopPropagation();
      });
    }
    if (dataOptionsToggle) {
      dataOptionsToggle.addEventListener("click", (event) => {
        event.stopPropagation();
        togglePopover(dataOptionsPopover, dataOptionsToggle);
      });
    }
    if (dataOptionsPopover) {
      dataOptionsPopover.addEventListener("click", (event) => {
        event.stopPropagation();
      });
    }
    if (importCsvBtn) {
      importCsvBtn.addEventListener("click", () => {
        openCsvFileDialog();
        closePageSettings();
      });
    }
    if (exportBundleBtn) {
      exportBundleBtn.addEventListener("click", () => {
        exportBundle();
        closePageSettings();
      });
    }
    if (openDiagnosticsBtn) {
      openDiagnosticsBtn.addEventListener("click", () => {
        openDiagnosticsModal();
        closePageSettings();
      });
    }
    if (toggleZenBtn) {
      toggleZenBtn.addEventListener("click", () => {
        toggleZen();
        closePageSettings();
      });
    }
    if (openHelpBtn) {
      openHelpBtn.addEventListener("click", () => {
        openHelpThread();
        closePageSettings();
      });
    }
    if (setCacheLocationBtn) {
      setCacheLocationBtn.addEventListener("click", () => {
        selectCacheLocation();
        closePageSettings();
      });
    }
    if (viewCacheBtn) {
      viewCacheBtn.addEventListener("click", async (event) => {
        event.stopPropagation();
        const ok = await ensureCacheReadPermission();
        openCacheViewModal(ok);
        closePageSettings();
      });
    }
    if (backupCacheBtn) {
      backupCacheBtn.addEventListener("click", () => {
        backupState();
        closePageSettings();
      });
    }
    if (restoreBackupBtn) {
      restoreBackupBtn.addEventListener("click", () => {
        restoreBackup();
        closePageSettings();
      });
    }
    if (clearCacheBtn) {
      clearCacheBtn.addEventListener("click", () => {
        clearCache();
        closePageSettings();
      });
    }
    if (cacheLocationLink) {
      cacheLocationLink.addEventListener("click", (event) => {
        event.stopPropagation();
        openCacheLocationLink();
        closePageSettings();
      });
    }
    if (csvFileInput) {
      csvFileInput.addEventListener("change", handleCsvFileInput);
    }
    if (backupFileInput) {
      backupFileInput.addEventListener("change", handleBackupFileInput);
    }

    if (dividerEl) {
      dividerEl.addEventListener("pointerdown", handleDividerPointerDown);
    }
    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);
    window.addEventListener("resize", applySplitRatio);
    window.addEventListener("resize", syncTitleTaglineWidth);
    document.addEventListener("visibilitychange", handleVisibilityChange);

    if (sheetPanel) {
      sheetPanel.addEventListener("dragover", handleSheetDragOver);
      sheetPanel.addEventListener("dragleave", handleSheetDragLeave);
      sheetPanel.addEventListener("drop", handleSheetDrop);
    }
    if (sheetTable) {
      sheetTable.addEventListener("mousedown", handleSheetMouseDown);
      sheetTable.addEventListener("mouseover", handleSheetMouseOver);
    }
    if (sheetContextMenu) {
      sheetContextMenu.addEventListener("click", (event) => {
        event.stopPropagation();
        const item = event.target.closest(".context-item");
        if (!item || item.disabled) return;
        handleSheetContextAction(item.dataset.action);
      });
      sheetContextMenu.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
    }
    if (cellContextMenu) {
      cellContextMenu.addEventListener("click", (event) => {
        event.stopPropagation();
        const item = event.target.closest(".context-item");
        if (!item || item.disabled) return;
        const submenuId = item.dataset.submenu;
        if (submenuId) {
          openCellSubmenu(item);
          return;
        }
        if (item.dataset.action) {
          handleCellContextAction(item.dataset.action);
        }
      });
      cellContextMenu.addEventListener("pointerover", (event) => {
        const item = event.target.closest(".context-item.has-submenu");
        if (!item || item.disabled) return;
        openCellSubmenu(item);
      });
      cellContextMenu.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
    }
    [cellEditMenu, cellAlignMenu, cellOptionsMenu, cellFormulaMenu].forEach((menu) => {
      if (!menu) return;
      menu.addEventListener("click", (event) => {
        event.stopPropagation();
        const item = event.target.closest(".context-item");
        if (!item || item.disabled) return;
        if (item.dataset.action) {
          handleCellContextAction(item.dataset.action);
          return;
        }
        if (item.dataset.formula) {
          addFormulaAtSelection(item.dataset.formula);
          closeCellContextMenu();
        }
      });
      menu.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
    });
    if (schemeContextMenu) {
      schemeContextMenu.addEventListener("click", (event) => {
        event.stopPropagation();
        const item = event.target.closest(".context-item");
        if (!item || item.disabled) return;
        handleSchemeContextAction(item.dataset.action);
      });
      schemeContextMenu.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
    }
    if (tabsContextMenu) {
      tabsContextMenu.addEventListener("click", (event) => {
        event.stopPropagation();
        const item = event.target.closest(".context-item");
        if (!item || item.disabled) return;
        handleTabsContextAction(item.dataset.action);
      });
      tabsContextMenu.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
    }
    document.addEventListener("mouseup", handleSheetMouseUp);

    newThreadBtn.addEventListener("click", () => {
      createAndActivateThread();
    });

    renameThreadBtn.addEventListener("click", () => {
      const thread = getActiveThread();
      renameThread(thread);
    });

    deleteThreadBtn.addEventListener("click", () => {
      if (state.threads.length <= 1) {
        alert("At least one thread is required.");
        return;
      }
      const thread = getActiveThread();
      if (!thread) return;
      const ok = confirm(`Delete ${thread.name}?`);
      if (!ok) return;
      state.threads = state.threads.filter((t) => t.id !== thread.id);
      state.activeId = state.threads[0].id;
      renderTabs();
      loadActiveThread();
      scheduleSave();
    });

    if (breakTimerBtn) {
      breakTimerBtn.addEventListener("click", (event) => {
        event.stopPropagation();
        toggleBreakTimerPanel();
      });
    }
    if (breakTimerPanel) {
      breakTimerPanel.addEventListener("click", (event) => {
        event.stopPropagation();
      });
      breakTimerPanel.querySelectorAll("[data-minutes]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const minutes = Number(btn.dataset.minutes);
          if (!Number.isFinite(minutes)) return;
          applyBreakPreset(minutes);
        });
      });
    }
    if (breakHoursInput) {
      breakHoursInput.addEventListener("input", handleBreakInputChange);
      breakHoursInput.addEventListener("change", handleBreakInputChange);
    }
    if (breakMinutesInput) {
      breakMinutesInput.addEventListener("input", handleBreakInputChange);
      breakMinutesInput.addEventListener("change", handleBreakInputChange);
    }
    if (breakStartBtn) {
      breakStartBtn.addEventListener("click", () => startBreakFromInputs());
    }
    if (breakPauseBtn) {
      breakPauseBtn.addEventListener("click", () => pauseBreakTimer());
    }
    if (breakResetBtn) {
      breakResetBtn.addEventListener("click", () => resetBreakTimer());
    }
    if (breakAlarmModal) {
      breakAlarmModal.addEventListener("click", (event) => {
        if (event.target === breakAlarmModal) {
          dismissBreakAlarm();
        }
      });
    }
    if (breakSnoozeBtn) {
      breakSnoozeBtn.addEventListener("click", () => snoozeBreakTimer());
    }
    if (breakDismissBtn) {
      breakDismissBtn.addEventListener("click", () => dismissBreakAlarm());
    }

    if (renameModal) {
      renameModal.addEventListener("click", (event) => {
        if (event.target === renameModal) {
          closeRenameModal();
        }
      });
    }
    if (cancelRenameBtn) {
      cancelRenameBtn.addEventListener("click", () => closeRenameModal());
    }
    if (confirmRenameBtn) {
      confirmRenameBtn.addEventListener("click", () => commitRename());
    }
    if (cacheViewModal) {
      cacheViewModal.addEventListener("click", (event) => {
        if (event.target === cacheViewModal) {
          closeCacheViewModal();
        }
        event.stopPropagation();
      });
    }
    if (closeCacheViewBtn) {
      closeCacheViewBtn.addEventListener("click", () => closeCacheViewModal());
    }
    if (diagnosticsModal) {
      diagnosticsModal.addEventListener("click", (event) => {
        if (event.target === diagnosticsModal) {
          closeDiagnosticsModal();
        }
        event.stopPropagation();
      });
    }
    if (closeDiagnosticsBtn) {
      closeDiagnosticsBtn.addEventListener("click", () => closeDiagnosticsModal());
    }
    if (renameInput) {
      renameInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          commitRename();
        }
        if (event.key === "Escape") {
          event.preventDefault();
          closeRenameModal();
        }
      });
    }
    if (cellTextColorInput) {
      cellTextColorInput.addEventListener("input", () => {
        applyTextColor(cellTextColorInput.value);
      });
      cellTextColorInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          applyTextColor(cellTextColorInput.value);
        }
      });
    }
    if (clearCellTextColorBtn) {
      clearCellTextColorBtn.addEventListener("click", () => {
        applyTextColor("");
      });
    }
    if (cellBgColorInput) {
      cellBgColorInput.addEventListener("input", () => {
        applyShade(cellBgColorInput.value);
      });
      cellBgColorInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          applyShade(cellBgColorInput.value);
        }
      });
    }
    if (clearCellBgColorBtn) {
      clearCellBgColorBtn.addEventListener("click", () => {
        applyShade("");
      });
    }
    if (schemeTextColorInput) {
      schemeTextColorInput.addEventListener("input", () => {
        updateSettings("journalColor", schemeTextColorInput.value);
      });
      schemeTextColorInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          updateSettings("journalColor", schemeTextColorInput.value);
        }
      });
    }
    if (clearSchemeTextColorBtn) {
      clearSchemeTextColorBtn.addEventListener("click", () => {
        updateSettings("journalColor", "#d9dde7");
      });
    }
    if (renameColorInput) {
      renameColorInput.addEventListener("input", () => {
        renameColorTouched = true;
        renameColorCleared = false;
        renameColorInput.classList.remove("is-empty");
      });
    }
    if (clearRenameColorBtn) {
      clearRenameColorBtn.addEventListener("click", () => {
        renameColorTouched = true;
        renameColorCleared = true;
        if (renameColorInput) {
          renameColorInput.classList.add("is-empty");
          renameColorInput.value = "#4fd1c5";
        }
      });
    }
    viewButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        setView(btn.dataset.view);
        closePageSettings();
      });
    });

    swapBtn.addEventListener("click", () => {
      state.ui.swap = !state.ui.swap;
      updateViewState();
      scheduleSave();
      closePageSettings();
    });

    exportTxtBtn.addEventListener("click", () => {
      exportText();
      closePageSettings();
    });
    exportCsvBtn.addEventListener("click", () => {
      exportCsv();
      closePageSettings();
    });
    exportBothBtn.addEventListener("click", () => {
      exportText();
      exportCsv();
      closePageSettings();
    });
    exportZipBtn.addEventListener("click", () => {
      exportZip();
      closePageSettings();
    });

    printBtn.addEventListener("click", () => {
      const mode = printModeSelect.value;
      document.body.dataset.printMode = mode;
      window.print();
      closePageSettings();
    });

    window.addEventListener("afterprint", () => {
      delete document.body.dataset.printMode;
    });

    journalEl.addEventListener("input", () => {
      const thread = getActiveThread();
      if (!thread) return;
      lastActivePanel = "journal";
      thread.journal = journalEl.value;
      updateStats();
      scheduleJournalUndo(thread.id, journalEl.value);
      scheduleSave();
    });
    journalEl.addEventListener("focus", () => {
      lastActivePanel = "journal";
    });
    journalEl.addEventListener("click", () => {
      lastActivePanel = "journal";
    });

    journalEl.addEventListener("paste", (event) => {
      event.preventDefault();
      const text = (event.clipboardData || window.clipboardData).getData("text");
      insertAtCursor(journalEl, text);
      journalEl.dispatchEvent(new Event("input"));
    });

    journalEl.addEventListener("contextmenu", (event) => {
      openSchemeContextMenu(event);
    });

    journalUndoBtn.addEventListener("click", () => {
      undoJournal();
    });

    journalFindBtn.addEventListener("click", () => {
      openFindPanel("journal");
    });

    journalTimestampBtn.addEventListener("click", () => {
      const stamp = formatTimestamp(new Date());
      insertAtCursor(journalEl, stamp);
      journalEl.dispatchEvent(new Event("input"));
    });
    if (findInput) {
      findInput.addEventListener("input", () => {
        sheetFindPointer = null;
        sheetFindLastQuery = findInput.value || "";
      });
    }
    if (findNextBtn) {
      findNextBtn.addEventListener("click", () => findNextGlobal());
    }
    if (replaceOneBtn) {
      replaceOneBtn.addEventListener("click", () => replaceCurrentGlobal());
    }
    if (replaceAllBtn) {
      replaceAllBtn.addEventListener("click", () => replaceAllGlobal());
    }
    if (closeFindBtn) {
      closeFindBtn.addEventListener("click", () => closeFindPanel());
    }
    if (findTargetButtons.length) {
      findTargetButtons.forEach((btn) => {
        btn.addEventListener("click", () => {
          setFindTarget(btn.dataset.findTarget);
        });
      });
    }
    if (findHeader) {
      findHeader.addEventListener("pointerdown", handleFindDragStart);
    }
    if (copySchemeBtn) {
      copySchemeBtn.addEventListener("click", () => copySchemeToClipboard());
    }
    if (clearSchemeBtn) {
      clearSchemeBtn.addEventListener("click", () => clearSchemeText());
    }

    journalFontEl.addEventListener("change", () => updateSettings("journalFont", journalFontEl.value));
    journalSizeEl.addEventListener("input", () => updateSettings("journalSize", Number(journalSizeEl.value)));
    journalColorEl.addEventListener("input", () => updateSettings("journalColor", journalColorEl.value));

    sheetFontEl.addEventListener("change", () => updateSettings("sheetFont", sheetFontEl.value));
    sheetSizeEl.addEventListener("input", () => updateSettings("sheetSize", Number(sheetSizeEl.value)));
    sheetColorEl.addEventListener("input", () => updateSettings("sheetColor", sheetColorEl.value));
    if (copyCsvBtn) {
      copyCsvBtn.addEventListener("click", () => copyCsvToClipboard());
    }
    if (clearSheetBtn) {
      clearSheetBtn.addEventListener("click", () => clearSheetData());
    }
    if (sheetFindBtn) {
      sheetFindBtn.addEventListener("click", () => openFindPanel("sheet"));
    }

    sheetUndoBtn.addEventListener("click", () => {
      undoSheet();
    });

    addRowBtn.addEventListener("click", () => addSheetRow());
    addColBtn.addEventListener("click", () => addSheetCol());
    if (removeRowBtn) {
      removeRowBtn.addEventListener("click", () => removeSheetRow(selectedCell.row));
    }
    if (removeColBtn) {
      removeColBtn.addEventListener("click", () => removeSheetCol(selectedCell.col));
    }

    sortApplyBtn.addEventListener("click", () => {
      const thread = getActiveThread();
      if (!thread) return;
      normalizeSheet(thread.sheet);
      const colIndex = Number(sortColumnEl.value);
      if (Number.isNaN(colIndex)) return;
      pushSheetUndo(thread.id, thread.sheet);
      const order = sortOrderEl.value;
      const rows = thread.sheet.cells.map((row, index) => ({
        row,
        height: thread.sheet.rowHeights ? thread.sheet.rowHeights[index] : null
      }));
      rows.sort((aEntry, bEntry) => {
        const aVal = getSortableValue(aEntry.row[colIndex]);
        const bVal = getSortableValue(bEntry.row[colIndex]);
        if (aVal.type === "number" && bVal.type === "number") {
          return order === "asc" ? aVal.value - bVal.value : bVal.value - aVal.value;
        }
        return order === "asc" ? aVal.value.localeCompare(bVal.value) : bVal.value.localeCompare(aVal.value);
      });
      thread.sheet.cells = rows.map((entry) => entry.row);
      if (thread.sheet.rowHeights) {
        thread.sheet.rowHeights = rows.map((entry) => entry.height ?? null);
      }
      buildSheetTable(thread.sheet);
      populateSortColumns(thread.sheet.cols);
      sheetStatus.textContent = `Sorted by Column ${indexToCol(colIndex)} (${order === "asc" ? "Ascending" : "Descending"})`;
      scheduleSave();
    });

    applyShadeBtn.addEventListener("click", () => applyShade(shadeColorEl.value));
    clearShadeBtn.addEventListener("click", () => applyShade(""));
    if (alignLeftBtn) {
      alignLeftBtn.addEventListener("click", () => applyAlignment("left"));
    }
    if (alignCenterBtn) {
      alignCenterBtn.addEventListener("click", () => applyAlignment("center"));
    }
    if (alignRightBtn) {
      alignRightBtn.addEventListener("click", () => applyAlignment("right"));
    }

    formulaInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        const thread = getActiveThread();
        if (!thread) return;
        commitCellValue(thread, selectedCell.row, selectedCell.col, formulaInput.value, true);
        moveSelectionBy(1, 0, event.shiftKey);
      }
    });

    formulaInput.addEventListener("focus", () => {
      lastActivePanel = "sheet";
      startSheetEditSession(selectedCell.row, selectedCell.col);
    });

    formulaInput.addEventListener("blur", () => {
      const thread = getActiveThread();
      if (!thread) return;
      commitCellValue(thread, selectedCell.row, selectedCell.col, formulaInput.value, true);
    });
  }

  function updateSettings(key, value) {
    const thread = getActiveThread();
    if (!thread) return;
    thread.settings[key] = value;
    applySettings(thread);
    scheduleSave();
  }

  function renameThread(thread) {
    if (!thread) return;
    openRenameModal(thread);
  }

  function openRenameModal(thread) {
    if (!renameModal || !renameInput) return;
    closePageSettings();
    closePanelPopovers();
    renameTargetId = thread.id;
    renameInput.value = thread.name;
    renameColorTouched = false;
    renameColorCleared = false;
    if (renameColorInput) {
      const normalized = normalizeHexColor(thread.color);
      if (normalized) {
        renameColorInput.value = normalized;
        renameColorInput.classList.remove("is-empty");
      } else {
        renameColorInput.value = "#4fd1c5";
        renameColorInput.classList.add("is-empty");
      }
    }
    renameModal.classList.add("open");
    renameModal.setAttribute("aria-hidden", "false");
    renameInput.focus();
    renameInput.select();
  }

  function closeRenameModal() {
    if (!renameModal) return;
    renameModal.classList.remove("open");
    renameModal.setAttribute("aria-hidden", "true");
    renameTargetId = null;
    renameColorTouched = false;
    renameColorCleared = false;
  }

  function commitRename() {
    if (!renameTargetId || !renameInput) return;
    const thread = state.threads.find((item) => item.id === renameTargetId);
    if (!thread) {
      closeRenameModal();
      return;
    }
    const trimmed = renameInput.value.trim();
    if (!trimmed) {
      renameInput.focus();
      return;
    }
    thread.name = trimmed;
    if (renameColorInput && renameColorTouched) {
      if (renameColorCleared) {
        thread.color = "";
      } else {
        const normalized = normalizeHexColor(renameColorInput.value);
        thread.color = normalized || "";
      }
    }
    renderTabs();
    scheduleSave();
    closeRenameModal();
  }

  function togglePopover(popover, button) {
    if (!popover || !button) return;
    const isOpen = popover.classList.contains("open");
    closePanelPopovers();
    closePageSettings();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    if (!isOpen) {
      popover.classList.add("open");
      button.classList.add("active");
    }
  }

  function closePanelPopovers() {
    [schemePopover, dataOptionsPopover].forEach((popover) => {
      if (popover) popover.classList.remove("open");
    });
    [schemeSettingsToggle, dataOptionsToggle].forEach((btn) => {
      if (btn) btn.classList.remove("active");
    });
  }

  function openFindPanel(target) {
    if (!findPanel) return;
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    findPanel.classList.add("open");
    findPanel.setAttribute("aria-hidden", "false");
    if (target) {
      setFindTarget(target);
    }
    if (findInput) findInput.focus();
  }

  function closeFindPanel() {
    if (!findPanel) return;
    findPanel.classList.remove("open");
    findPanel.setAttribute("aria-hidden", "true");
  }

  function toggleFindPanelForContext() {
    if (!findPanel) return;
    const isOpen = findPanel.classList.contains("open");
    if (isOpen) {
      closeFindPanel();
    } else {
      openFindPanelForContext();
    }
  }

  function setFindTarget(target) {
    const next = target || "journal";
    findTarget = next;
    if (next === "sheet") lastActivePanel = "sheet";
    if (next === "journal") lastActivePanel = "journal";
    updateFindTargetButtons();
  }

  function updateFindTargetButtons() {
    if (!findTargetButtons.length) return;
    findTargetButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.findTarget === findTarget);
    });
  }

  function openFindPanelForContext() {
    const active = document.activeElement;
    let target = "journal";
    if (sheetPanel && (sheetPanel.contains(active) || active === formulaInput)) {
      target = "sheet";
    } else if (journalPanel && journalPanel.contains(active)) {
      target = "journal";
    } else if (state.ui.view === "sheet" || lastActivePanel === "sheet") {
      target = "sheet";
    }
    openFindPanel(target);
  }

  function handleGlobalKeydown(event) {
    const key = event.key.toLowerCase();
    const targetIsCell = isCellInput(event.target);
    const editingCell = targetIsCell && cellEditMode;
    const multiSelection = selectionIsMultiple();
    const isUndo = (event.ctrlKey || event.metaKey) && key === "z" && !event.shiftKey;
    const isFindKey = key === "f" && !event.shiftKey;
    const inEntryField = isEditableTarget(event.target) && !(targetIsCell && !cellEditMode);
    if ((event.ctrlKey || event.metaKey) && key === "o") {
      event.preventDefault();
      openCsvFileDialog();
      return;
    }
    if (isFindKey && (event.ctrlKey || event.metaKey || (!event.altKey && !event.ctrlKey && !event.metaKey))) {
      if (!inEntryField) {
        event.preventDefault();
        toggleFindPanelForContext();
        return;
      }
    }
    if (event.key === "Escape") {
      closePageSettings();
      closePanelPopovers();
      closeBreakTimerPanel();
      closeSheetContextMenu();
      closeCellContextMenu();
      closeCellSubmenus();
      closeSchemeContextMenu();
      closeTabsContextMenu();
      closeRenameModal();
      closeCacheViewModal();
      closeDiagnosticsModal();
      closeFindPanel();
      dismissBreakAlarm();
      if (state.ui.zen) setZen(false);
      return;
    }
    if (renameModal && renameModal.classList.contains("open")) return;
    if (isUndo) {
      if (event.target === journalEl) {
        event.preventDefault();
        undoJournal();
        return;
      }
      if (targetIsCell || event.target === formulaInput) {
        event.preventDefault();
        undoSheet();
        return;
      }
    }
    if (targetIsCell) {
      if ((event.ctrlKey || event.metaKey) && key === "c") {
        if (!editingCell) {
          event.preventDefault();
          copySelectionToClipboard();
        }
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "x") {
        if (!editingCell) {
          event.preventDefault();
          cutSelectionToClipboard();
        }
        return;
      }
      if (key === "delete" || key === "backspace") {
        if (!editingCell || multiSelection) {
          event.preventDefault();
          clearSelectedCells();
        }
        return;
      }
      if (event.altKey || event.ctrlKey || event.metaKey) return;
      if (editingCell) return;
      return;
    }
    if (isEditableTarget(event.target)) return;
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    if (key === "delete" || key === "backspace") {
      event.preventDefault();
      clearSelectedCells();
      return;
    }
    if (key === "n") {
      createAndActivateThread();
      return;
    }
    if (key === "a") setView("split");
    if (key === "s") setView("journal");
    if (key === "d") setView("sheet");
    if (key === "x") {
      state.ui.swap = !state.ui.swap;
      updateViewState();
      scheduleSave();
    }
    if (key === "z") toggleZen();
  }

  function isEditableTarget(target) {
    if (!target) return false;
    const tag = target.tagName;
    return tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable;
  }

  function isCellInput(target) {
    return !!(target && target.classList && target.classList.contains("cell-input"));
  }

  function togglePageSettings() {
    if (!pageSettingsPanel) return;
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    pageSettingsPanel.classList.toggle("open");
    if (pageSettingsToggle) {
      pageSettingsToggle.classList.toggle("active", pageSettingsPanel.classList.contains("open"));
    }
  }

  function closePageSettings() {
    if (!pageSettingsPanel) return;
    pageSettingsPanel.classList.remove("open");
    if (pageSettingsToggle) pageSettingsToggle.classList.remove("active");
  }

  function updateCacheLabel() {
    if (!cacheLocationLabel) return;
    const name = state.ui.cacheFileName;
    const label = name ? `Cache: ${name}` : "Cache: Local storage";
    if (cacheLocationText) {
      cacheLocationText.textContent = label;
    } else {
      cacheLocationLabel.textContent = label;
    }
    if (cacheLocationLink) {
      cacheLocationLink.style.display = name ? "inline-flex" : "none";
    }
  }

  function syncTitleTaglineWidth() {
    if (!titleEl || !taglineEl) return;
    const text = (taglineEl.textContent || "").trim();
    if (text.length < 2) return;
    taglineEl.style.transform = "";
    taglineEl.style.transformOrigin = "left center";
    taglineEl.style.letterSpacing = "0px";
    const titleWidth = titleEl.getBoundingClientRect().width;
    const baseWidth = taglineEl.getBoundingClientRect().width;
    if (!titleWidth || !baseWidth) return;
    const spacing = (titleWidth - baseWidth) / (text.length - 1);
    if (spacing > 0) {
      taglineEl.style.letterSpacing = `${spacing}px`;
    }
    const adjustedWidth = taglineEl.getBoundingClientRect().width;
    if (!adjustedWidth) return;
    const ratio = titleWidth / adjustedWidth;
    if (Math.abs(1 - ratio) > 0.002) {
      taglineEl.style.transform = `scaleX(${ratio})`;
    }
  }

  async function initCacheHandle() {
    if (!("indexedDB" in window)) return;
    try {
      const handle = await getCachedHandle();
      if (!handle) return;
      cacheHandle = handle;
      let didUpdate = false;
      if (!state.ui.cacheFileName && cacheHandle && cacheHandle.name) {
        state.ui.cacheFileName = cacheHandle.name;
        didUpdate = true;
      }
      updateCacheLabel();
      if (didUpdate) scheduleSave();
      const permission = await queryHandlePermission(cacheHandle, "readwrite");
      if (permission === "granted") {
        queueCacheWrite();
      }
    } catch (err) {
      // ignore cache handle restore errors
    }
  }

  function openCacheDb() {
    if (cacheHandlePromise) return cacheHandlePromise;
    cacheHandlePromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(CACHE_HANDLE_DB, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(CACHE_HANDLE_STORE)) {
          db.createObjectStore(CACHE_HANDLE_STORE);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    return cacheHandlePromise;
  }

  async function getCachedHandle() {
    const db = await openCacheDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(CACHE_HANDLE_STORE, "readonly");
      const store = tx.objectStore(CACHE_HANDLE_STORE);
      const req = store.get(CACHE_HANDLE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  async function setCachedHandle(handle) {
    const db = await openCacheDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(CACHE_HANDLE_STORE, "readwrite");
      const store = tx.objectStore(CACHE_HANDLE_STORE);
      store.put(handle, CACHE_HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function clearCachedHandle() {
    const db = await openCacheDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(CACHE_HANDLE_STORE, "readwrite");
      const store = tx.objectStore(CACHE_HANDLE_STORE);
      store.delete(CACHE_HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function queryHandlePermission(handle, mode) {
    if (!handle || !handle.queryPermission) return "granted";
    try {
      return await handle.queryPermission({ mode });
    } catch (err) {
      return "prompt";
    }
  }

  async function requestHandlePermission(handle, mode) {
    if (!handle || !handle.requestPermission) return "granted";
    try {
      return await handle.requestPermission({ mode });
    } catch (err) {
      return "denied";
    }
  }

  async function ensureCacheReadPermission() {
    if (!cacheHandle) return true;
    const permission = await queryHandlePermission(cacheHandle, "read");
    if (permission === "granted") return true;
    const requested = await requestHandlePermission(cacheHandle, "read");
    return requested === "granted";
  }

  function openCacheViewModal(permissionGranted = true) {
    if (!cacheViewModal || !cacheViewContent) return;
    cacheViewModal.classList.add("open");
    cacheViewModal.setAttribute("aria-hidden", "false");
    if (cacheHandle && !permissionGranted) {
      cacheViewContent.textContent = "Cache file permission required.";
      return;
    }
    cacheViewContent.textContent = "Loading...";
    loadCacheViewContent();
  }

  function closeCacheViewModal() {
    if (!cacheViewModal) return;
    cacheViewModal.classList.remove("open");
    cacheViewModal.setAttribute("aria-hidden", "true");
  }

  function openDiagnosticsModal() {
    if (!diagnosticsModal || !diagnosticsContent) return;
    diagnosticsModal.classList.add("open");
    diagnosticsModal.setAttribute("aria-hidden", "false");
    diagnosticsContent.textContent = buildDiagnosticsReport();
  }

  function closeDiagnosticsModal() {
    if (!diagnosticsModal) return;
    diagnosticsModal.classList.remove("open");
    diagnosticsModal.setAttribute("aria-hidden", "true");
  }

  function buildDiagnosticsReport() {
    const threadCount = state.threads.length;
    const totalCells = state.threads.reduce((sum, thread) => {
      if (!thread.sheet) return sum;
      return sum + thread.sheet.rows * thread.sheet.cols;
    }, 0);
    const lines = [];
    lines.push(`Threads: ${threadCount}`);
    lines.push(`Total cells: ${totalCells}`);
    lines.push(`Active timers: breakTick=${breakTickTimer ? "on" : "off"}, breakReminder=${breakReminderTimer ? "on" : "off"}`);
    lines.push("Undo stacks:");
    state.threads.forEach((thread) => {
      const journalCount = (journalUndoStacks.get(thread.id) || []).length;
      const sheetCount = (sheetUndoStacks.get(thread.id) || []).length;
      lines.push(`- ${thread.name}: journal=${journalCount}, sheet=${sheetCount}`);
    });
    return lines.join("\n");
  }

  async function loadCacheViewContent() {
    if (!cacheViewContent) return;
    let text = "";
    if (cacheHandle) {
      const permission = await queryHandlePermission(cacheHandle, "read");
      if (permission !== "granted") {
        cacheViewContent.textContent = "Cache file permission required.";
        return;
      }
      try {
        const file = await cacheHandle.getFile();
        text = await file.text();
      } catch (err) {
        text = "Failed to read cache file.";
      }
    } else {
      try {
        text = JSON.stringify(state, null, 2);
      } catch (err) {
        text = "Failed to read local cache.";
      }
    }
    cacheViewContent.textContent = text || "Cache empty.";
  }

  async function openCacheLocationLink() {
    if (!cacheHandle) {
      openCacheViewModal();
      return;
    }
    if (window.showOpenFilePicker && window.isSecureContext) {
      try {
        const [handle] = await window.showOpenFilePicker({
          startIn: cacheHandle,
          types: [{ description: "JSON", accept: { "application/json": [".json"] } }],
          excludeAcceptAllOption: true,
          multiple: false
        });
        if (!handle) return;
        cacheHandle = handle;
        state.ui.cacheFileName = handle.name || state.ui.cacheFileName || "multithreader-cache.json";
        await setCachedHandle(handle);
        await requestHandlePermission(handle, "readwrite");
        updateCacheLabel();
        scheduleSave();
        const ok = await ensureCacheReadPermission();
        openCacheViewModal(ok);
      } catch (err) {
        // ignore user cancel
      }
      return;
    }
    const permission = await queryHandlePermission(cacheHandle, "read");
    if (permission !== "granted") {
      const requested = await requestHandlePermission(cacheHandle, "read");
      if (requested !== "granted") {
        openCacheViewModal();
        return;
      }
    }
    try {
      const file = await cacheHandle.getFile();
      const url = URL.createObjectURL(file);
      window.open(url, "_blank", "noopener");
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      openCacheViewModal();
    }
  }

  function openTabsContextMenu(x, y) {
    if (!tabsContextMenu) return;
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeCacheViewModal();
    updateTabsContextMenu();
    tabsContextMenu.classList.add("open");
    tabsContextMenu.setAttribute("aria-hidden", "false");
    positionTabsContextMenu(x, y);
  }

  function positionTabsContextMenu(x, y) {
    if (!tabsContextMenu) return;
    const padding = 8;
    tabsContextMenu.style.left = `${x}px`;
    tabsContextMenu.style.top = `${y}px`;
    const rect = tabsContextMenu.getBoundingClientRect();
    const clampedX = Math.min(window.innerWidth - rect.width - padding, Math.max(padding, x));
    const clampedY = Math.min(window.innerHeight - rect.height - padding, Math.max(padding, y));
    tabsContextMenu.style.left = `${clampedX}px`;
    tabsContextMenu.style.top = `${clampedY}px`;
  }

  function updateTabsContextMenu() {
    if (!tabsContextMenu) return;
    const deleteItem = tabsContextMenu.querySelector("[data-action=\"delete-thread\"]");
    if (deleteItem) {
      deleteItem.disabled = state.threads.length <= 1;
    }
  }

  function closeTabsContextMenu() {
    if (!tabsContextMenu) return;
    tabsContextMenu.classList.remove("open");
    tabsContextMenu.setAttribute("aria-hidden", "true");
    tabsContextTargetId = null;
  }

  function handleTabsContextAction(action) {
    if (!action) return;
    if (action === "add-thread") {
      createAndActivateThread();
      closeTabsContextMenu();
      return;
    }
    if (action === "delete-thread") {
      if (state.threads.length <= 1) {
        alert("At least one thread is required.");
        closeTabsContextMenu();
        return;
      }
      const targetId = tabsContextTargetId || state.activeId;
      const thread = state.threads.find((item) => item.id === targetId);
      if (!thread) {
        closeTabsContextMenu();
        return;
      }
      const ok = confirm(`Delete ${thread.name}?`);
      if (!ok) {
        closeTabsContextMenu();
        return;
      }
      state.threads = state.threads.filter((t) => t.id !== thread.id);
      if (state.activeId === thread.id) {
        state.activeId = state.threads[0].id;
      }
      renderTabs();
      loadActiveThread();
      scheduleSave();
      closeTabsContextMenu();
    }
  }

  function initBreakTimer() {
    if (!breakTimerBtn || !breakRemainingEl || !state.ui.breakTimer) return;
    const timer = state.ui.breakTimer;
    if (timer.running && !timer.endAt) {
      const fallback = timer.remainingSeconds || timer.lastSetSeconds || 0;
      timer.endAt = Date.now() + fallback * 1000;
    }
    const remaining = getBreakRemainingSeconds();
    if (timer.running && remaining <= 0) {
      timer.running = false;
      timer.remainingSeconds = 0;
      timer.endAt = 0;
      scheduleSave();
      triggerBreakAlarm();
    } else if (timer.running) {
      timer.remainingSeconds = remaining;
      startBreakTicker();
    }
    setBreakInputs(timer.lastSetSeconds);
    updateBreakDisplay();
    initBreakReminder();
  }

  function toggleBreakTimerPanel() {
    if (!breakTimerPanel) return;
    const isOpen = breakTimerPanel.classList.contains("open");
    closePageSettings();
    closePanelPopovers();
    closeSheetContextMenu();
    closeCellSubmenus();
    if (isOpen) {
      closeBreakTimerPanel();
      return;
    }
    breakTimerPanel.classList.add("open");
  }

  function closeBreakTimerPanel() {
    if (!breakTimerPanel) return;
    breakTimerPanel.classList.remove("open");
  }

  function handleBreakInputChange() {
    if (!state.ui.breakTimer) return;
    const seconds = getBreakInputSeconds();
    if (seconds <= 0) return;
    state.ui.breakTimer.lastSetSeconds = seconds;
    if (!state.ui.breakTimer.running) {
      state.ui.breakTimer.remainingSeconds = seconds;
      updateBreakDisplay();
    }
    scheduleSave();
  }

  function applyBreakPreset(minutes) {
    const seconds = Math.max(0, Math.floor(minutes * 60));
    setBreakInputs(seconds);
    startBreakTimer(seconds);
  }

  function startBreakFromInputs() {
    const seconds = getBreakInputSeconds();
    if (seconds <= 0) return;
    startBreakTimer(seconds);
  }

  function setBreakInputs(totalSeconds) {
    if (!breakHoursInput || !breakMinutesInput) return;
    const clamped = Math.max(0, Math.floor(totalSeconds || 0));
    const hours = Math.floor(clamped / 3600);
    const minutes = Math.floor((clamped % 3600) / 60);
    breakHoursInput.value = String(hours);
    breakMinutesInput.value = String(minutes);
  }

  function getBreakInputSeconds() {
    if (!breakHoursInput || !breakMinutesInput) return 0;
    const hours = clamp(parseInt(breakHoursInput.value, 10) || 0, 0, 23);
    const minutes = clamp(parseInt(breakMinutesInput.value, 10) || 0, 0, 59);
    breakHoursInput.value = String(hours);
    breakMinutesInput.value = String(minutes);
    return hours * 3600 + minutes * 60;
  }

  function startBreakTimer(seconds) {
    if (!state.ui.breakTimer || seconds <= 0) return;
    const timer = state.ui.breakTimer;
    timer.lastSetSeconds = seconds;
    timer.remainingSeconds = seconds;
    timer.running = true;
    timer.endAt = Date.now() + seconds * 1000;
    startBreakTicker();
    updateBreakDisplay();
    setReminderPaused(true, true);
    scheduleSave();
  }

  function pauseBreakTimer() {
    if (!state.ui.breakTimer) return;
    const timer = state.ui.breakTimer;
    if (!timer.running) return;
    timer.remainingSeconds = getBreakRemainingSeconds();
    timer.running = false;
    timer.endAt = 0;
    stopBreakTicker();
    updateBreakDisplay();
    setReminderPaused(false, true);
    scheduleSave();
  }

  function resetBreakTimer() {
    if (!state.ui.breakTimer) return;
    const timer = state.ui.breakTimer;
    timer.running = false;
    timer.endAt = 0;
    timer.remainingSeconds = timer.lastSetSeconds;
    stopBreakTicker();
    updateBreakDisplay();
    setReminderPaused(false, true);
    scheduleSave();
  }

  function startBreakTicker() {
    stopBreakTicker();
    breakTickTimer = setInterval(() => {
      if (!state.ui.breakTimer) return;
      const timer = state.ui.breakTimer;
      if (!timer.running) return;
      const remaining = getBreakRemainingSeconds();
      if (remaining <= 0) {
        timer.running = false;
        timer.remainingSeconds = 0;
        timer.endAt = 0;
        stopBreakTicker();
        updateBreakDisplay();
        scheduleSave();
        triggerBreakAlarm();
        return;
      }
      timer.remainingSeconds = remaining;
      updateBreakDisplay();
    }, 1000);
  }

  function stopBreakTicker() {
    if (!breakTickTimer) return;
    clearInterval(breakTickTimer);
    breakTickTimer = null;
  }

  function getBreakRemainingSeconds() {
    if (!state.ui.breakTimer) return 0;
    const timer = state.ui.breakTimer;
    if (!timer.running || !timer.endAt) return timer.remainingSeconds || timer.lastSetSeconds || 0;
    return Math.max(0, Math.ceil((timer.endAt - Date.now()) / 1000));
  }

  function updateBreakDisplay() {
    if (!breakRemainingEl || !breakTimerBtn || !state.ui.breakTimer) return;
    const timer = state.ui.breakTimer;
    const seconds = timer.running ? getBreakRemainingSeconds() : (timer.remainingSeconds || timer.lastSetSeconds || 0);
    breakRemainingEl.textContent = formatDuration(seconds);
    breakTimerBtn.classList.toggle("running", timer.running);
  }

  function formatDuration(totalSeconds) {
    const clamped = Math.max(0, Math.floor(totalSeconds || 0));
    const hours = Math.floor(clamped / 3600);
    const minutes = Math.floor((clamped % 3600) / 60);
    const seconds = clamped % 60;
    const hh = String(hours).padStart(2, "0");
    const mm = String(minutes).padStart(2, "0");
    const ss = String(seconds).padStart(2, "0");
    return `${hh}:${mm}:${ss}`;
  }

  function initBreakReminder() {
    stopBreakReminder();
    resetBreakReminderState();
    reminderPaused = !!(state.ui.breakTimer && state.ui.breakTimer.running);
    if (!reminderPaused && document.visibilityState === "visible") {
      focusStartAt = Date.now();
    }
    startBreakReminder();
  }

  function startBreakReminder() {
    if (!breakTimerBtn) return;
    if (breakReminderTimer) {
      clearInterval(breakReminderTimer);
    }
    breakReminderTimer = setInterval(() => {
      updateBreakReminder();
    }, 60 * 1000);
    updateBreakReminder(true);
  }

  function stopBreakReminder() {
    if (!breakReminderTimer) return;
    clearInterval(breakReminderTimer);
    breakReminderTimer = null;
  }

  function resetBreakReminderState() {
    reminderStage = 0;
    accumulatedFocusMs = 0;
    focusStartAt = null;
    updateBreakReminderClasses(0);
  }

  function updateBreakReminder(forcePulse = false) {
    if (!breakTimerBtn || reminderPaused) return;
    const minutes = Math.floor(getFocusedElapsedMs() / 60000);
    let stage = 0;
    if (minutes >= 120) stage = 120;
    else if (minutes >= 90) stage = 90;
    else if (minutes >= 60) stage = 60;
    else if (minutes >= 30) stage = 30;
    if (stage !== reminderStage || forcePulse) {
      reminderStage = stage;
      updateBreakReminderClasses(stage);
      if (stage === 30 || stage === 60 || stage === 90) {
        pulseBreakButton();
      }
    }
  }

  function updateBreakReminderClasses(stage) {
    if (!breakTimerBtn) return;
    breakTimerBtn.classList.remove("reminder-30", "reminder-60", "reminder-90", "reminder-120");
    if (stage) {
      breakTimerBtn.classList.add(`reminder-${stage}`);
    }
  }

  function getFocusedElapsedMs() {
    let total = accumulatedFocusMs;
    if (focusStartAt) {
      total += Date.now() - focusStartAt;
    }
    return total;
  }

  function setReminderPaused(paused, reset = false) {
    if (paused) {
      if (focusStartAt) {
        accumulatedFocusMs += Date.now() - focusStartAt;
        focusStartAt = null;
      }
      reminderPaused = true;
      return;
    }
    reminderPaused = false;
    if (reset) {
      resetBreakReminderState();
    }
    if (document.visibilityState === "visible") {
      focusStartAt = Date.now();
    }
    updateBreakReminder(true);
  }

  function handleVisibilityChange() {
    if (reminderPaused) return;
    if (document.visibilityState === "hidden") {
      if (focusStartAt) {
        accumulatedFocusMs += Date.now() - focusStartAt;
        focusStartAt = null;
      }
      return;
    }
    if (!focusStartAt) {
      focusStartAt = Date.now();
    }
    updateBreakReminder(true);
  }

  function pulseBreakButton() {
    if (!breakTimerBtn) return;
    breakTimerBtn.classList.add("pulse");
    setTimeout(() => {
      if (breakTimerBtn) breakTimerBtn.classList.remove("pulse");
    }, 2400);
  }

  function triggerBreakAlarm() {
    if (breakAlarmActive) return;
    breakAlarmActive = true;
    if (breakAlarmModal) {
      breakAlarmModal.classList.add("open");
      breakAlarmModal.setAttribute("aria-hidden", "false");
    }
    pulseBreakButton();
    playBreakTone();
    setReminderPaused(false, true);
  }

  function dismissBreakAlarm() {
    if (!breakAlarmActive) return;
    breakAlarmActive = false;
    if (breakAlarmModal) {
      breakAlarmModal.classList.remove("open");
      breakAlarmModal.setAttribute("aria-hidden", "true");
    }
  }

  function snoozeBreakTimer() {
    dismissBreakAlarm();
    startBreakTimer(600);
  }

  function playBreakTone() {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const context = new AudioCtx();
      const oscillator = context.createOscillator();
      const gain = context.createGain();
      oscillator.type = "sine";
      oscillator.frequency.value = 432;
      gain.gain.value = 0.05;
      oscillator.connect(gain);
      gain.connect(context.destination);
      oscillator.start();
      oscillator.stop(context.currentTime + 0.6);
      oscillator.onended = () => {
        context.close();
      };
    } catch (err) {
      // ignore audio errors
    }
  }

  async function selectCacheLocation() {
    if (!window.showSaveFilePicker || !window.isSecureContext) {
      alert("Set cache location requires a secure context (use https or http://localhost).");
      return;
    }
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: "multithreader-cache.json",
        types: [{ description: "JSON", accept: { "application/json": [".json"] } }]
      });
      cacheHandle = handle;
      state.ui.cacheFileName = handle.name || "multithreader-cache.json";
      await setCachedHandle(handle);
      await requestHandlePermission(handle, "readwrite");
      updateCacheLabel();
      scheduleSave();
      queueCacheWrite();
    } catch (err) {
      // ignore user cancel
    }
  }

  function buildBackupPayload() {
    return state;
  }

  async function backupState() {
    const payload = buildBackupPayload();
    const content = JSON.stringify(payload, null, 2);
    const name = `multithreader-backup-${safeTimestamp()}.json`;
    if (window.showSaveFilePicker && window.isSecureContext) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: name,
          types: [{ description: "JSON", accept: { "application/json": [".json"] } }]
        });
        const writable = await handle.createWritable();
        await writable.write(content);
        await writable.close();
        if (journalStatus) journalStatus.textContent = "Backup saved";
        if (sheetStatus) sheetStatus.textContent = "Backup saved";
      } catch (err) {
        // ignore user cancel
      }
      return;
    }
    downloadFile(name, content, "application/json");
    if (journalStatus) journalStatus.textContent = "Backup saved";
    if (sheetStatus) sheetStatus.textContent = "Backup saved";
  }

  function handleBackupFileInput(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    restoreBackupFromFile(file, null);
    event.target.value = "";
  }

  async function restoreBackup() {
    if (window.showOpenFilePicker && window.isSecureContext) {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [{ description: "JSON", accept: { "application/json": [".json"] } }],
          excludeAcceptAllOption: true,
          multiple: false
        });
        if (!handle) return;
        const file = await handle.getFile();
        await restoreBackupFromFile(file, handle);
      } catch (err) {
        // ignore user cancel
      }
      return;
    }
    if (!backupFileInput) {
      alert("Restore backup is unavailable in this browser.");
      return;
    }
    backupFileInput.value = "";
    backupFileInput.click();
  }

  async function restoreBackupFromFile(file, handle) {
    if (!file) return;
    let text = "";
    try {
      text = await file.text();
    } catch (err) {
      alert("Failed to read backup file.");
      return;
    }
    let data = null;
    try {
      data = JSON.parse(text);
    } catch (err) {
      alert("Invalid backup file.");
      return;
    }
    const nextState = extractStateFromBackup(data);
    if (!nextState) {
      alert("Invalid backup file.");
      return;
    }
    const ok = confirm("Restore backup? This will replace your current data.");
    if (!ok) return;
    await applyRestoredState(nextState, handle);
    if (journalStatus) journalStatus.textContent = "Backup restored";
    if (sheetStatus) sheetStatus.textContent = "Backup restored";
  }

  function extractStateFromBackup(data) {
    if (!data || typeof data !== "object") return null;
    if (data.state && typeof data.state === "object") {
      return data.state;
    }
    if (Array.isArray(data.threads)) {
      return data;
    }
    return null;
  }

  async function applyRestoredState(nextState, handle) {
    state = nextState;
    normalizeState(state);
    journalUndoStacks.clear();
    journalLast.clear();
    sheetUndoStacks.clear();
    if (handle) {
      cacheHandle = handle;
      state.ui.cacheFileName = handle.name || state.ui.cacheFileName || "multithreader-cache.json";
      try {
        await setCachedHandle(handle);
        await requestHandlePermission(handle, "readwrite");
      } catch (err) {
        // ignore permission errors
      }
    }
    renderTabs();
    loadActiveThread();
    updateCacheLabel();
    stopBreakTicker();
    breakAlarmActive = false;
    if (breakAlarmModal) {
      breakAlarmModal.classList.remove("open");
      breakAlarmModal.setAttribute("aria-hidden", "true");
    }
    initBreakTimer();
    scheduleSave();
  }

  function clearCache() {
    const ok = confirm("Clear cache? Are you sure?");
    if (!ok) return;
    localStorage.removeItem(STORAGE_KEY);
    journalUndoStacks.clear();
    journalLast.clear();
    sheetUndoStacks.clear();
    cacheHandle = null;
    clearCachedHandle().catch(() => {});
    state = buildDefaultState();
    normalizeState(state);
    renderTabs();
    loadActiveThread();
    updateCacheLabel();
    stopBreakTicker();
    breakAlarmActive = false;
    if (breakAlarmModal) {
      breakAlarmModal.classList.remove("open");
      breakAlarmModal.setAttribute("aria-hidden", "true");
    }
    initBreakTimer();
    scheduleSave();
  }

  function queueCacheWrite() {
    if (!cacheHandle) return;
    if (cacheWriteInFlight) {
      cacheWriteQueued = true;
      return;
    }
    cacheWriteInFlight = true;
    writeCacheFile().finally(() => {
      cacheWriteInFlight = false;
      if (cacheWriteQueued) {
        cacheWriteQueued = false;
        queueCacheWrite();
      }
    });
  }

  async function writeCacheFile() {
    if (!cacheHandle) return;
    try {
      const permission = await queryHandlePermission(cacheHandle, "readwrite");
      if (permission !== "granted") return;
      const writable = await cacheHandle.createWritable();
      await writable.write(JSON.stringify(state, null, 2));
      await writable.close();
    } catch (err) {
      // keep cached handle; user can re-grant permission via View/Set
    }
  }

  function openHelpThread() {
    const existing = state.threads.find((thread) => thread.isHelp || thread.name.toLowerCase() === "help");
    if (existing) {
      if (!existing.journal) existing.journal = HELP_TEXT;
      setActiveThread(existing.id);
      return;
    }
    const thread = createThread("Help");
    thread.isHelp = true;
    thread.color = "#ff4d4d";
    thread.journal = HELP_TEXT;
    state.threads.push(thread);
    setActiveThread(thread.id);
    scheduleSave();
  }

  function openCsvFileDialog() {
    if (!csvFileInput) return;
    csvFileInput.value = "";
    csvFileInput.click();
  }

  function handleCsvFileInput(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    importCsvFile(file);
    csvFileInput.value = "";
  }

  function handleSheetDragOver(event) {
    if (!hasCsvFile(event.dataTransfer)) return;
    event.preventDefault();
    sheetPanel.classList.add("drag-over");
  }

  function handleSheetDragLeave() {
    sheetPanel.classList.remove("drag-over");
  }

  function handleSheetDrop(event) {
    event.preventDefault();
    sheetPanel.classList.remove("drag-over");
    const file = getCsvFileFromDataTransfer(event.dataTransfer);
    if (file) importCsvFile(file);
  }

  function hasCsvFile(dataTransfer) {
    if (!dataTransfer || !dataTransfer.items) return false;
    return Array.from(dataTransfer.items).some((item) => item.kind === "file" && isCsvName(item.getAsFile()?.name));
  }

  function getCsvFileFromDataTransfer(dataTransfer) {
    if (!dataTransfer || !dataTransfer.files || dataTransfer.files.length === 0) return null;
    return Array.from(dataTransfer.files).find((file) => isCsvName(file.name)) || null;
  }

  function isCsvName(name) {
    if (!name) return false;
    return name.toLowerCase().endsWith(".csv");
  }

  function importCsvFile(file) {
    if (!file) return;
    file
      .text()
      .then((text) => {
        handleCsvText(text);
      })
      .catch(() => {
        alert("Failed to read CSV file.");
      });
  }

  function handleCsvText(text) {
    const thread = getActiveThread();
    if (!thread) return;
    const grid = parseCsv(text);
    if (!grid.length) return;
    if (sheetHasData(thread.sheet)) {
      const ok = confirm("Replace current sheet with imported CSV?");
      if (!ok) return;
    }
    setSheetFromGrid(thread, grid);
  }

  function sheetHasData(sheet) {
    return sheet.cells.some((row) => row.some((cell) => cell.value || cell.bg || cell.fg));
  }

  function scheduleJournalUndo(threadId, currentValue) {
    const lastValue = journalLast.get(threadId);
    if (lastValue !== undefined && lastValue !== currentValue) {
      pushJournalUndo(threadId, lastValue);
    }
    journalLast.set(threadId, currentValue);
  }

  function pushJournalUndo(threadId, snapshot) {
    const stack = journalUndoStacks.get(threadId) || [];
    if (stack.length && stack[stack.length - 1] === snapshot) {
      return;
    }
    stack.push(snapshot);
    while (stack.length > UNDO_LIMIT) stack.shift();
    journalUndoStacks.set(threadId, stack);
  }

  function insertAtCursor(textarea, text) {
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const value = textarea.value;
    textarea.value = value.slice(0, start) + text + value.slice(end);
    textarea.selectionStart = textarea.selectionEnd = start + text.length;
    textarea.focus();
  }

  function formatTimestamp(date) {
    const pad = (value) => String(value).padStart(2, "0");
    const yyyy = date.getFullYear();
    const mm = pad(date.getMonth() + 1);
    const dd = pad(date.getDate());
    const hh = pad(date.getHours());
    const min = pad(date.getMinutes());
    return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
  }

  function findNext() {
    const query = findInput.value;
    if (!query) return false;
    const text = journalEl.value;
    const startPos = journalEl.selectionEnd;
    const haystack = matchCaseInput.checked ? text : text.toLowerCase();
    const needle = matchCaseInput.checked ? query : query.toLowerCase();
    let index = haystack.indexOf(needle, startPos);
    if (index === -1 && startPos > 0) {
      index = haystack.indexOf(needle, 0);
    }
    if (index === -1) return false;
    journalEl.focus();
    journalEl.setSelectionRange(index, index + query.length);
    return true;
  }

  function replaceCurrent() {
    const query = findInput.value;
    if (!query) return false;
    const selection = journalEl.value.slice(journalEl.selectionStart, journalEl.selectionEnd);
    const isMatch = matchCaseInput.checked ? selection === query : selection.toLowerCase() === query.toLowerCase();
    if (!isMatch) {
      return findNext();
    }
    insertAtCursor(journalEl, replaceInput.value);
    journalEl.dispatchEvent(new Event("input"));
    return true;
  }

  function replaceAll() {
    const query = findInput.value;
    if (!query) return 0;
    const replaceValue = replaceInput.value;
    const flags = matchCaseInput.checked ? "g" : "gi";
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(escaped, flags);
    const matches = journalEl.value.match(regex);
    const count = matches ? matches.length : 0;
    journalEl.value = journalEl.value.replace(regex, replaceValue);
    journalEl.dispatchEvent(new Event("input"));
    return count;
  }

  function findNextGlobal() {
    const target = findTarget || "journal";
    if (target === "journal") return findNext();
    if (target === "sheet") return findNextSheet();
    if (target === "both") {
      if (findNext()) return true;
      return findNextSheet();
    }
    return false;
  }

  function replaceCurrentGlobal() {
    const target = findTarget || "journal";
    if (target === "journal") return replaceCurrent();
    if (target === "sheet") return replaceSheetCurrent();
    if (target === "both") {
      if (lastActivePanel === "sheet") {
        if (replaceSheetCurrent()) return true;
        return replaceCurrent();
      }
      if (replaceCurrent()) return true;
      return replaceSheetCurrent();
    }
    return false;
  }

  function replaceAllGlobal() {
    const target = findTarget || "journal";
    let total = 0;
    if (target === "journal" || target === "both") {
      const count = replaceAll();
      if (count && journalStatus) {
        journalStatus.textContent = `Replaced ${count}`;
      }
      total += count;
    }
    if (target === "sheet" || target === "both") {
      total += replaceSheetAll();
    }
    if (target === "both") {
      if (total) {
        sheetStatus.textContent = `Replaced ${total}`;
        journalStatus.textContent = `Replaced ${total}`;
      }
    }
    return total;
  }

  function findNextSheet() {
    const query = findInput ? findInput.value : "";
    if (!query) return false;
    const thread = getActiveThread();
    if (!thread) return false;
    normalizeSheet(thread.sheet);
    const matchCase = matchCaseInput && matchCaseInput.checked;
    const needle = matchCase ? query : query.toLowerCase();
    const rows = thread.sheet.rows;
    const cols = thread.sheet.cols;
    if (!rows || !cols) return false;
    const total = rows * cols;
    let startIndex = 0;
    if (sheetFindLastQuery === query && sheetFindPointer) {
      startIndex = sheetFindPointer.row * cols + sheetFindPointer.col + 1;
    } else {
      startIndex = selectedCell.row * cols + selectedCell.col;
    }
    sheetFindLastQuery = query;
    for (let i = 0; i < total; i += 1) {
      const index = (startIndex + i) % total;
      const row = Math.floor(index / cols);
      const col = index % cols;
      const cell = thread.sheet.cells[row][col];
      const value = cell && typeof cell.value === "string" ? cell.value : "";
      const haystack = matchCase ? value : value.toLowerCase();
      if (haystack.includes(needle)) {
        sheetFindPointer = { row, col };
        focusCell(row, col);
        sheetStatus.textContent = `Match ${indexToCol(col)}${row + 1}`;
        return true;
      }
    }
    sheetFindPointer = null;
    sheetStatus.textContent = "No match";
    return false;
  }

  function replaceSheetCurrent() {
    const query = findInput ? findInput.value : "";
    if (!query) return false;
    const thread = getActiveThread();
    if (!thread) return false;
    normalizeSheet(thread.sheet);
    const matchCase = matchCaseInput && matchCaseInput.checked;
    let target = sheetFindPointer;
    if (!target || sheetFindLastQuery !== query) {
      const value = thread.sheet.cells[selectedCell.row][selectedCell.col].value || "";
      const haystack = matchCase ? value : value.toLowerCase();
      const needle = matchCase ? query : query.toLowerCase();
      if (haystack.includes(needle)) {
        target = { row: selectedCell.row, col: selectedCell.col };
      }
    }
    if (!target) {
      return findNextSheet();
    }
    const cell = thread.sheet.cells[target.row][target.col];
    const before = cell.value || "";
    const nextValue = replaceCellValue(before, query, replaceInput ? replaceInput.value : "", matchCase, false);
    if (before === nextValue) {
      sheetStatus.textContent = "No match";
      return false;
    }
    pushSheetUndo(thread.id, thread.sheet);
    cell.value = nextValue;
    recalcSheet();
    focusCell(target.row, target.col);
    sheetFindPointer = { row: target.row, col: target.col };
    sheetFindLastQuery = query;
    sheetStatus.textContent = "Replaced";
    scheduleSave();
    return true;
  }

  function replaceSheetAll() {
    const query = findInput ? findInput.value : "";
    if (!query) return 0;
    const thread = getActiveThread();
    if (!thread) return 0;
    normalizeSheet(thread.sheet);
    const matchCase = matchCaseInput && matchCaseInput.checked;
    const replacement = replaceInput ? replaceInput.value : "";
    let replaced = 0;
    let pushedUndo = false;
    for (let r = 0; r < thread.sheet.rows; r += 1) {
      for (let c = 0; c < thread.sheet.cols; c += 1) {
        const cell = thread.sheet.cells[r][c];
        const before = cell.value || "";
        const after = replaceCellValue(before, query, replacement, matchCase, true);
        if (before !== after) {
          if (!pushedUndo) {
            pushSheetUndo(thread.id, thread.sheet);
            pushedUndo = true;
          }
          cell.value = after;
          replaced += 1;
        }
      }
    }
    if (!replaced) {
      sheetStatus.textContent = "No match";
      return 0;
    }
    recalcSheet();
    sheetFindPointer = null;
    sheetFindLastQuery = query;
    sheetStatus.textContent = `Replaced ${replaced}`;
    scheduleSave();
    return replaced;
  }

  function replaceCellValue(value, query, replacement, matchCase, replaceAll) {
    if (!query) return value;
    const flags = matchCase ? "g" : "gi";
    const safe = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(safe, replaceAll ? flags : flags.replace("g", ""));
    return String(value ?? "").replace(regex, replacement);
  }

  function clearSchemeText() {
    const thread = getActiveThread();
    if (!thread) return;
    const ok = confirm("Clear scheme text? This cannot be undone.");
    if (!ok) return;
    pushJournalUndo(thread.id, journalEl.value);
    journalEl.value = "";
    thread.journal = "";
    journalLast.set(thread.id, "");
    updateStats();
    journalStatus.textContent = "Cleared";
    scheduleSave();
  }

  function buildSheetTable(sheet) {
    normalizeSheet(sheet);
    sheetTable.innerHTML = "";
    sheetInputs = [];
    sheetInputList = [];

    const colgroup = document.createElement("colgroup");
    const cornerCol = document.createElement("col");
    colgroup.appendChild(cornerCol);
    for (let c = 0; c < sheet.cols; c += 1) {
      const colEl = document.createElement("col");
      colEl.dataset.col = String(c);
      const width = sheet.colWidths && sheet.colWidths[c];
      if (Number.isFinite(width)) {
        colEl.style.width = `${width}px`;
      }
      colgroup.appendChild(colEl);
    }
    sheetTable.appendChild(colgroup);

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "sticky-col corner-cell";
    corner.textContent = "";
    corner.addEventListener("click", () => {
      selectAllCells();
    });
    headRow.appendChild(corner);
    for (let c = 0; c < sheet.cols; c += 1) {
      const th = document.createElement("th");
      th.classList.add("col-header");
      th.textContent = indexToCol(c);
      th.dataset.col = String(c);
      th.addEventListener("click", (event) => {
        if (event.target.closest(".col-resizer") || event.target.closest(".add-col")) return;
        if (event.shiftKey) {
          selectColumnRange(c);
        } else {
          selectColumn(c);
        }
      });
      th.addEventListener("contextmenu", (event) => {
        if (event.target.closest(".col-resizer") || event.target.closest(".add-col")) return;
        openSheetContextMenu(event, { type: "col", index: c });
      });
      const resizer = document.createElement("div");
      resizer.className = "col-resizer";
      resizer.dataset.col = String(c);
      resizer.addEventListener("pointerdown", handleColumnResizeStart);
      th.appendChild(resizer);
      if (c === sheet.cols - 1) {
        const addCol = document.createElement("button");
        addCol.type = "button";
        addCol.className = "add-pill add-col";
        addCol.textContent = "+";
        addCol.title = "Add column";
        addCol.setAttribute("aria-label", "Add column");
        addCol.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          addSheetCol();
        });
        th.appendChild(addCol);
      }
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    sheetTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let r = 0; r < sheet.rows; r += 1) {
      const row = document.createElement("tr");
      row.dataset.row = String(r);
      const rowHeight = sheet.rowHeights && sheet.rowHeights[r];
      if (Number.isFinite(rowHeight)) {
        row.style.height = `${rowHeight}px`;
      }
      const rowHeader = document.createElement("th");
      rowHeader.className = "sticky-col row-header";
      rowHeader.addEventListener("click", (event) => {
        if (event.target.closest(".row-resizer") || event.target.closest(".add-row")) return;
        if (event.shiftKey) {
          selectRowRange(r);
        } else {
          selectRow(r);
        }
      });
      rowHeader.addEventListener("contextmenu", (event) => {
        if (event.target.closest(".row-resizer") || event.target.closest(".add-row")) return;
        openSheetContextMenu(event, { type: "row", index: r });
      });
      const rowLabel = document.createElement("span");
      rowLabel.className = "row-label";
      rowLabel.textContent = String(r + 1);
      rowHeader.appendChild(rowLabel);
      if (Number.isFinite(rowHeight)) {
        rowHeader.style.height = `${rowHeight}px`;
      }
      const rowResizer = document.createElement("div");
      rowResizer.className = "row-resizer";
      rowResizer.dataset.row = String(r);
      rowResizer.addEventListener("pointerdown", handleRowResizeStart);
      rowHeader.appendChild(rowResizer);
      if (r === sheet.rows - 1) {
        rowHeader.classList.add("has-add");
        const addRow = document.createElement("button");
        addRow.type = "button";
        addRow.className = "add-pill add-row";
        addRow.textContent = "+";
        addRow.title = "Add row";
        addRow.setAttribute("aria-label", "Add row");
        addRow.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          addSheetRow();
        });
        rowHeader.appendChild(addRow);
      }
      row.appendChild(rowHeader);

      const inputRow = [];
      for (let c = 0; c < sheet.cols; c += 1) {
        const td = document.createElement("td");
        if (Number.isFinite(rowHeight)) {
          td.style.height = `${rowHeight}px`;
        }
        const input = document.createElement("textarea");
        input.className = "cell-input";
        input.rows = 1;
        input.wrap = "off";
        input.spellcheck = false;
        input.dataset.row = String(r);
        input.dataset.col = String(c);
        input.addEventListener("focus", () => {
          cellEditMode = false;
          startSheetEditSession(r, c);
          focusCell(r, c);
          input.value = sheet.cells[r][c].value || "";
        });
        input.addEventListener("blur", () => {
          cellEditMode = false;
          const thread = getActiveThread();
          if (!thread) return;
          commitCellValue(thread, r, c, input.value, true);
        });
        input.addEventListener("input", () => {
          cellEditMode = true;
          formulaInput.value = input.value;
        });
        input.addEventListener("dblclick", () => {
          cellEditMode = true;
          input.select();
        });
        input.addEventListener("keydown", (event) => {
          if (event.key === "F2") {
            cellEditMode = true;
            input.select();
            return;
          }
          if (event.key === "Enter" && !event.altKey && !event.shiftKey) {
            event.preventDefault();
            const thread = getActiveThread();
            if (!thread) return;
            commitCellValue(thread, r, c, input.value, true);
            moveSelectionBy(1, 0, event.shiftKey);
            return;
          }
          if (!cellEditMode && !event.altKey && !event.ctrlKey && !event.metaKey) {
            if (event.key === "ArrowUp") {
              event.preventDefault();
              moveSelectionBy(-1, 0, event.shiftKey);
              return;
            }
            if (event.key === "ArrowDown") {
              event.preventDefault();
              moveSelectionBy(1, 0, event.shiftKey);
              return;
            }
            if (event.key === "ArrowLeft") {
              event.preventDefault();
              moveSelectionBy(0, -1, event.shiftKey);
              return;
            }
            if (event.key === "ArrowRight") {
              event.preventDefault();
              moveSelectionBy(0, 1, event.shiftKey);
              return;
            }
          }
          if (event.key === "Enter" && (event.altKey || event.shiftKey)) {
            cellEditMode = true;
            return;
          }
          if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
            cellEditMode = true;
          }
        });
        input.addEventListener("contextmenu", (event) => {
          openCellContextMenu(event, { row: r, col: c });
        });
        input.addEventListener("paste", (event) => {
          event.preventDefault();
          handleCellPaste(event, r, c);
        });
        td.appendChild(input);
        row.appendChild(td);
        inputRow.push(input);
        sheetInputList.push(input);
      }
      tbody.appendChild(row);
      sheetInputs.push(inputRow);
    }
    sheetTable.appendChild(tbody);

    recalcSheet();
    const row = Math.min(selectedCell.row, sheet.rows - 1);
    const col = Math.min(selectedCell.col, sheet.cols - 1);
    focusCell(Math.max(row, 0), Math.max(col, 0));
  }

  function handleCellPaste(event, row, col) {
    const thread = getActiveThread();
    if (!thread) return;
    const text = (event.clipboardData || window.clipboardData).getData("text");
    const forceGrid = event.altKey || isLastSheetCopy(text);
    const grid = parseClipboardGrid(text, forceGrid);
    if (grid.length === 1 && grid[0].length === 1) {
      commitCellValue(thread, row, col, grid[0][0], true);
      const target = sheetInputs[row] ? sheetInputs[row][col] : null;
      if (target) target.value = grid[0][0];
      formulaInput.value = grid[0][0];
      return;
    }
    applyGridToSheet(thread, row, col, grid);
    const target = sheetInputs[row] ? sheetInputs[row][col] : null;
    if (target) target.value = thread.sheet.cells[row][col].value;
    formulaInput.value = thread.sheet.cells[row][col].value || "";
  }

  function isLastSheetCopy(text) {
    return !!(lastSheetCopy && lastSheetCopy.isMulti && lastSheetCopy.text === text);
  }

  function parseClipboardGrid(text, forceGrid = false) {
    const normalized = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const hasTab = normalized.includes("\t");
    const hasNewline = normalized.includes("\n");
    const hasComma = normalized.includes(",");

    if (forceGrid) {
      if (hasTab) {
        return normalized.split("\n").map((line) => line.split("\t"));
      }
      if (hasComma) {
        return parseCsv(normalized);
      }
      if (hasNewline) {
        return normalized.split("\n").map((line) => [line]);
      }
      return [[normalized]];
    }

    if (hasTab) {
      return normalized.split("\n").map((line) => line.split("\t"));
    }
    if (hasComma && hasNewline) {
      return parseCsv(normalized);
    }
    return [[normalized]];
  }

function parseCsv(text) {
    const rows = [];
    let row = [];
    let value = "";
    let inQuotes = false;
    const normalized = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");

    for (let i = 0; i < normalized.length; i += 1) {
      const char = normalized[i];
      if (inQuotes) {
        if (char === "\"") {
          if (normalized[i + 1] === "\"") {
            value += "\"";
            i += 1;
          } else {
            inQuotes = false;
          }
        } else {
          value += char;
        }
      } else if (char === "\"") {
        inQuotes = true;
      } else if (char === ",") {
        row.push(value);
        value = "";
      } else if (char === "\n") {
        row.push(value);
        rows.push(row);
        row = [];
        value = "";
      } else {
        value += char;
      }
    }

    row.push(value);
    if (row.length > 1 || row[0] !== "" || normalized.trim() !== "") {
      rows.push(row);
    }

    while (rows.length && rows[rows.length - 1].every((cell) => cell === "")) {
      rows.pop();
    }
    return rows;
  }

  function applyGridToSheet(thread, startRow, startCol, grid) {
    if (!grid.length) return;
    const maxCols = Math.max(...grid.map((row) => row.length));
    if (!maxCols) return;
    pushSheetUndo(thread.id, thread.sheet);
    const resized = ensureSheetSize(thread.sheet, startRow + grid.length, startCol + maxCols);
    for (let r = 0; r < grid.length; r += 1) {
      for (let c = 0; c < grid[r].length; c += 1) {
        thread.sheet.cells[startRow + r][startCol + c].value = grid[r][c];
      }
    }
    if (resized) {
      buildSheetTable(thread.sheet);
      populateSortColumns(thread.sheet.cols);
      focusCell(startRow, startCol);
    } else {
      recalcSheet();
    }
    scheduleSave();
  }

  function ensureSheetSize(sheet, rowsNeeded, colsNeeded) {
    let resized = false;
    normalizeSheet(sheet);
    while (sheet.rows < rowsNeeded) {
      const row = [];
      for (let c = 0; c < sheet.cols; c += 1) {
        row.push({ value: "", bg: "", align: "left", fg: "", wrap: "truncate" });
      }
      sheet.cells.push(row);
      sheet.rows += 1;
      if (sheet.rowHeights) sheet.rowHeights.push(null);
      resized = true;
    }
    while (sheet.cols < colsNeeded) {
      sheet.cells.forEach((row) => row.push({ value: "", bg: "", align: "left", fg: "", wrap: "truncate" }));
      sheet.cols += 1;
      if (sheet.colWidths) sheet.colWidths.push(null);
      resized = true;
    }
    return resized;
  }

  function setSheetFromGrid(thread, grid) {
    const rows = Math.max(grid.length, 1);
    const cols = Math.max(...grid.map((row) => row.length), 1);
    pushSheetUndo(thread.id, thread.sheet);
    const sheet = makeSheet(rows, cols);
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        sheet.cells[r][c].value = (grid[r] && grid[r][c] !== undefined) ? grid[r][c] : "";
      }
    }
    thread.sheet = sheet;
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    scheduleSave();
  }

  function startSheetEditSession(row, col) {
    const thread = getActiveThread();
    if (!thread) return;
    sheetEditSession = {
      threadId: thread.id,
      row,
      col,
      pushed: false
    };
  }

  function focusCell(row, col) {
    cellEditMode = false;
    selectedCell = { row, col };
    lastActivePanel = "sheet";
    sheetInputList.forEach((input) => input.classList.remove("selected"));
    const target = sheetInputs[row] ? sheetInputs[row][col] : null;
    if (target) {
      target.classList.add("selected");
      if (document.activeElement !== target) {
        target.focus({ preventScroll: true });
      }
    }
    cellLabelEl.textContent = `${indexToCol(col)}${row + 1}`;
    if (sortColumnEl) {
      sortColumnEl.value = String(col);
    }
    if (!ignoreFocusSelection) {
      setSelectionRange(row, col, row, col);
      selectionAnchor = { row, col };
    }
    ignoreFocusSelection = false;
    const thread = getActiveThread();
    if (thread) {
      const cell = thread.sheet.cells[row][col];
      formulaInput.value = cell.value || "";
    }
  }

  function moveSelectionBy(deltaRow, deltaCol, extendSelection) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const nextRow = clamp(selectedCell.row + deltaRow, 0, thread.sheet.rows - 1);
    const nextCol = clamp(selectedCell.col + deltaCol, 0, thread.sheet.cols - 1);
    if (extendSelection) {
      const anchor = selectionAnchor || { row: selectedCell.row, col: selectedCell.col };
      ignoreFocusSelection = true;
      selectionAnchor = anchor;
      setSelectionRange(anchor.row, anchor.col, nextRow, nextCol);
    }
    focusCell(nextRow, nextCol);
  }

  function setSelectionRange(startRow, startCol, endRow, endCol) {
    selectionRange = {
      start: { row: startRow, col: startCol },
      end: { row: endRow, col: endCol }
    };
    updateSelectionClasses();
  }

  function getSelectionBounds() {
    if (!selectionRange) return null;
    const rowStart = Math.min(selectionRange.start.row, selectionRange.end.row);
    const rowEnd = Math.max(selectionRange.start.row, selectionRange.end.row);
    const colStart = Math.min(selectionRange.start.col, selectionRange.end.col);
    const colEnd = Math.max(selectionRange.start.col, selectionRange.end.col);
    return { rowStart, rowEnd, colStart, colEnd };
  }

  function selectionIsMultiple() {
    const bounds = getSelectionBounds();
    if (!bounds) return false;
    return bounds.rowStart !== bounds.rowEnd || bounds.colStart !== bounds.colEnd;
  }

  function updateSelectionClasses() {
    const bounds = getSelectionBounds();
    sheetInputList.forEach((input) => {
      const row = Number(input.dataset.row);
      const col = Number(input.dataset.col);
      const inRange = bounds
        ? row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd
        : false;
      input.classList.toggle("range", inRange);
    });
  }

  function handleSheetMouseDown(event) {
    const target = event.target.closest(".cell-input");
    if (!target || event.button !== 0) return;
    lastActivePanel = "sheet";
    const row = Number(target.dataset.row);
    const col = Number(target.dataset.col);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    cellEditMode = false;
    selectionActive = true;
    if (event.shiftKey && selectionAnchor) {
      ignoreFocusSelection = true;
      setSelectionRange(selectionAnchor.row, selectionAnchor.col, row, col);
    } else {
      selectionAnchor = { row, col };
      setSelectionRange(row, col, row, col);
    }
  }

  function handleSheetMouseOver(event) {
    if (!selectionActive) return;
    if (!(event.buttons & 1)) {
      selectionActive = false;
      return;
    }
    const target = event.target.closest(".cell-input");
    if (!target || !selectionAnchor) return;
    const row = Number(target.dataset.row);
    const col = Number(target.dataset.col);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    setSelectionRange(selectionAnchor.row, selectionAnchor.col, row, col);
  }

  function handleSheetMouseUp() {
    selectionActive = false;
  }

  function forEachSelectedCell(callback) {
    const bounds = getSelectionBounds();
    if (!bounds) {
      callback(selectedCell.row, selectedCell.col);
      return;
    }
    for (let r = bounds.rowStart; r <= bounds.rowEnd; r += 1) {
      for (let c = bounds.colStart; c <= bounds.colEnd; c += 1) {
        callback(r, c);
      }
    }
  }

  function commitCellValue(thread, row, col, raw, pushUndo) {
    const cell = thread.sheet.cells[row][col];
    const sessionMatches =
      sheetEditSession
      && sheetEditSession.threadId === thread.id
      && sheetEditSession.row === row
      && sheetEditSession.col === col;
    if (raw === cell.value) {
      if (sessionMatches) {
        sheetEditSession = null;
      }
      recalcSheet();
      return;
    }
    if (pushUndo) {
      if (sessionMatches) {
        if (!sheetEditSession.pushed) {
          pushSheetUndo(thread.id, thread.sheet);
          sheetEditSession.pushed = true;
        }
      } else {
        pushSheetUndo(thread.id, thread.sheet);
      }
    }
    cell.value = raw;
    recalcSheet();
    scheduleSave();
    sheetEditSession = null;
  }

  function pushSheetUndo(threadId, sheet) {
    const stack = sheetUndoStacks.get(threadId) || [];
    stack.push(cloneSheet(sheet));
    while (stack.length > UNDO_LIMIT) stack.shift();
    sheetUndoStacks.set(threadId, stack);
  }

  function cloneSheet(sheet) {
    return {
      rows: sheet.rows,
      cols: sheet.cols,
      cells: sheet.cells.map((row) => row.map((cell) => ({
        value: cell.value,
        bg: cell.bg,
        align: cell.align || "left",
        fg: cell.fg || "",
        wrap: cell.wrap || "truncate"
      }))),
      colWidths: Array.isArray(sheet.colWidths) ? [...sheet.colWidths] : [],
      rowHeights: Array.isArray(sheet.rowHeights) ? [...sheet.rowHeights] : []
    };
  }

  function recalcSheet() {
    const thread = getActiveThread();
    if (!thread) return;
    const sheet = thread.sheet;
    for (let r = 0; r < sheet.rows; r += 1) {
      const rowInputs = sheetInputs[r];
      if (!rowInputs) continue;
      for (let c = 0; c < sheet.cols; c += 1) {
        const input = rowInputs[c];
        if (!input) continue;
        const cell = sheet.cells[r][c];
        input.style.background = cell.bg || "";
        input.style.textAlign = cell.align || "left";
        input.style.color = cell.fg || "";
        input.style.whiteSpace = cell.wrap === "wrap" ? "pre-wrap" : "nowrap";
        input.style.textOverflow = cell.wrap === "wrap" ? "clip" : "ellipsis";
        input.wrap = cell.wrap === "wrap" ? "soft" : "off";
        if (document.activeElement === input) {
          continue;
        }
        const display = getCellDisplay(sheet, r, c);
        input.value = display;
      }
    }
  }

  function getCellDisplay(sheet, row, col) {
    const cell = sheet.cells[row][col];
    const raw = (cell.value || "").trim();
    if (!raw) return "";
    if (raw.startsWith("=")) {
      const result = evaluateFormula(sheet, raw.slice(1), new Set());
      if (Number.isNaN(result)) return "#ERR";
      return formatNumber(result);
    }
    return raw;
  }

  function evaluateFormula(sheet, expression, visited) {
    let expr = expression.trim();
    if (!expr) return NaN;
    const seen = visited || new Set();

    let guard = 0;
    while (guard < 25) {
      const match = expr.match(/\b(SUM|AVG|MIN|MAX|COUNT)\(([^()]*)\)/i);
      if (!match) break;
      expr = expr.replace(/\b(SUM|AVG|MIN|MAX|COUNT)\(([^()]*)\)/gi, (full, fnName, args) => {
        const values = parseFunctionArgs(sheet, args, seen);
        if (!values.length) return "0";
        const nonEmpty = values.filter((value) => value !== null && value !== undefined && value !== "");
        const numbers = nonEmpty
          .map((value) => Number(value))
          .filter((value) => !Number.isNaN(value));
        if (!numbers.length) return "0";
        const upper = fnName.toUpperCase();
        if (upper === "SUM") return String(numbers.reduce((a, b) => a + b, 0));
        if (upper === "AVG") return String(numbers.reduce((a, b) => a + b, 0) / numbers.length);
        if (upper === "MIN") return String(Math.min(...numbers));
        if (upper === "MAX") return String(Math.max(...numbers));
        if (upper === "COUNT") return String(nonEmpty.length);
        return "0";
      });
      guard += 1;
    }

    expr = expr.replace(/\b([A-Z]+)(\d+)\b/g, (match, colLetters, rowDigits) => {
      const ref = parseCellRef(colLetters, rowDigits);
      if (!ref) return "0";
      const value = getCellNumber(sheet, ref.row, ref.col, seen);
      return String(value);
    });

    const safe = expr.replace(/\s+/g, "");
    if (!safe) return NaN;
    if (!/^[0-9+\-*/().]+$/.test(safe)) return NaN;

    try {
      const result = Function(`"use strict"; return (${safe});`)();
      if (!Number.isFinite(result)) return NaN;
      return result;
    } catch (err) {
      return NaN;
    }
  }

  function parseFunctionArgs(sheet, args, visited) {
    const seen = visited || new Set();
    return args
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean)
      .flatMap((part) => {
        if (part.includes(":")) {
          const cells = expandRange(sheet, part, seen);
          return cells.map((cell) => cell.value);
        }
        if (/^[A-Z]+\d+$/.test(part)) {
          const letterMatch = part.match(/^[A-Z]+/);
          const digitMatch = part.match(/\d+/);
          if (!letterMatch || !digitMatch) return [];
          const ref = parseCellRef(letterMatch[0], digitMatch[0]);
          if (!ref) return [];
          return [getCellValueForFunction(sheet, ref.row, ref.col, seen)];
        }
        const numeric = Number(part);
        if (!Number.isNaN(numeric)) return [numeric];
        return [];
      });
  }

  function expandRange(sheet, rangeText, visited) {
    const seen = visited || new Set();
    const [start, end] = rangeText.split(":");
    if (!start || !end) return [];
    const startLetters = start.match(/^[A-Z]+/);
    const startDigits = start.match(/\d+/);
    const endLetters = end.match(/^[A-Z]+/);
    const endDigits = end.match(/\d+/);
    if (!startLetters || !startDigits || !endLetters || !endDigits) return [];
    const startRef = parseCellRef(startLetters[0], startDigits[0]);
    const endRef = parseCellRef(endLetters[0], endDigits[0]);
    if (!startRef || !endRef) return [];
    const rowStart = Math.min(startRef.row, endRef.row);
    const rowEnd = Math.max(startRef.row, endRef.row);
    const colStart = Math.min(startRef.col, endRef.col);
    const colEnd = Math.max(startRef.col, endRef.col);
    const cells = [];
    for (let r = rowStart; r <= rowEnd; r += 1) {
      for (let c = colStart; c <= colEnd; c += 1) {
        if (sheet.cells[r] && sheet.cells[r][c]) {
          cells.push({ value: getCellValueForFunction(sheet, r, c, seen) });
        }
      }
    }
    return cells;
  }

  function getCellValueForFunction(sheet, row, col, visited) {
    if (!sheet.cells[row] || !sheet.cells[row][col]) return null;
    const key = `${row}:${col}`;
    const seen = visited || new Set();
    if (seen.has(key)) return null;
    seen.add(key);
    const raw = (sheet.cells[row][col].value || "").trim();
    if (!raw) {
      seen.delete(key);
      return null;
    }
    if (raw.startsWith("=")) {
      const result = evaluateFormula(sheet, raw.slice(1), seen);
      seen.delete(key);
      return Number.isNaN(result) ? null : result;
    }
    const numeric = Number(raw);
    seen.delete(key);
    if (!Number.isNaN(numeric)) return numeric;
    return raw;
  }

  function getCellNumber(sheet, row, col, visited) {
    if (!sheet.cells[row] || !sheet.cells[row][col]) return 0;
    const key = `${row}:${col}`;
    if (visited.has(key)) return 0;
    visited.add(key);
    const raw = (sheet.cells[row][col].value || "").trim();
    if (raw.startsWith("=")) {
      const result = evaluateFormula(sheet, raw.slice(1), visited);
      visited.delete(key);
      return Number.isNaN(result) ? 0 : result;
    }
    const numeric = Number(raw);
    visited.delete(key);
    return Number.isNaN(numeric) ? 0 : numeric;
  }

  function parseCellRef(colLetters, rowDigits) {
    const col = colToIndex(colLetters);
    const row = Number(rowDigits) - 1;
    if (Number.isNaN(row) || row < 0 || col < 0) return null;
    return { row, col };
  }

  function colToIndex(letters) {
    let result = 0;
    const text = letters.toUpperCase();
    for (let i = 0; i < text.length; i += 1) {
      result *= 26;
      result += text.charCodeAt(i) - 64;
    }
    return result - 1;
  }

  function indexToCol(index) {
    let result = "";
    let value = index + 1;
    while (value > 0) {
      const mod = (value - 1) % 26;
      result = String.fromCharCode(65 + mod) + result;
      value = Math.floor((value - 1) / 26);
    }
    return result;
  }

  function formatNumber(value) {
    if (Number.isInteger(value)) return String(value);
    return value.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  }

  function populateSortColumns(cols) {
    if (!sortColumnEl) return;
    sortColumnEl.innerHTML = "";
    for (let c = 0; c < cols; c += 1) {
      const option = document.createElement("option");
      option.value = String(c);
      option.textContent = `Column ${indexToCol(c)}`;
      sortColumnEl.appendChild(option);
    }
    const selected = Math.max(0, Math.min(selectedCell.col, cols - 1));
    sortColumnEl.value = String(selected);
  }

  function getSortableValue(cell) {
    const raw = (cell.value || "").trim();
    if (!raw) return { type: "text", value: "" };
    if (raw.startsWith("=")) {
      const thread = getActiveThread();
      const result = thread ? evaluateFormula(thread.sheet, raw.slice(1), new Set()) : NaN;
      if (!Number.isNaN(result)) return { type: "number", value: result };
      return { type: "text", value: "" };
    }
    const numeric = Number(raw);
    if (!Number.isNaN(numeric)) return { type: "number", value: numeric };
    return { type: "text", value: raw.toLowerCase() };
  }

  function applyShade(color) {
    const thread = getActiveThread();
    if (!thread) return;
    const normalized = normalizeHexColor(color);
    const nextColor = normalized || "";
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    forEachSelectedCell((row, col) => {
      const cell = thread.sheet.cells[row] && thread.sheet.cells[row][col];
      if (cell) cell.bg = nextColor;
    });
    recalcSheet();
    updateCellContextMenu();
    scheduleSave();
  }

  function applyTextColor(color) {
    const thread = getActiveThread();
    if (!thread) return;
    const normalized = normalizeHexColor(color);
    const nextColor = normalized || "";
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    forEachSelectedCell((row, col) => {
      const cell = thread.sheet.cells[row] && thread.sheet.cells[row][col];
      if (cell) cell.fg = nextColor;
    });
    recalcSheet();
    updateCellContextMenu();
    scheduleSave();
  }

  function applyWrapMode(mode) {
    const thread = getActiveThread();
    if (!thread) return;
    const nextMode = mode === "wrap" ? "wrap" : "truncate";
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    forEachSelectedCell((row, col) => {
      const cell = thread.sheet.cells[row] && thread.sheet.cells[row][col];
      if (cell) cell.wrap = nextMode;
    });
    recalcSheet();
    sheetStatus.textContent = nextMode === "wrap" ? "Wrap" : "Truncate";
    scheduleSave();
  }

  function applyAlignment(align) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    forEachSelectedCell((row, col) => {
      const cell = thread.sheet.cells[row] && thread.sheet.cells[row][col];
      if (cell) cell.align = align;
    });
    recalcSheet();
    sheetStatus.textContent = `Align ${align}`;
    scheduleSave();
  }

  function undoJournal() {
    const thread = getActiveThread();
    if (!thread) return;
    const stack = journalUndoStacks.get(thread.id) || [];
    if (!stack.length) {
      journalStatus.textContent = "Undo buffer empty";
      return;
    }
    const snapshot = stack.pop();
    journalUndoStacks.set(thread.id, stack);
    journalEl.value = snapshot;
    thread.journal = snapshot;
    journalLast.set(thread.id, snapshot);
    updateStats();
    scheduleSave();
  }

  function undoSheet() {
    const thread = getActiveThread();
    if (!thread) return;
    const stack = sheetUndoStacks.get(thread.id) || [];
    if (!stack.length) {
      sheetStatus.textContent = "Undo buffer empty";
      return;
    }
    const snapshot = stack.pop();
    thread.sheet = snapshot;
    normalizeSheet(thread.sheet);
    sheetUndoStacks.set(thread.id, stack);
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    scheduleSave();
  }

  function clearSelectedCells() {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    forEachSelectedCell((row, col) => {
      const cell = thread.sheet.cells[row] && thread.sheet.cells[row][col];
      if (cell) cell.value = "";
      const input = sheetInputs[row] ? sheetInputs[row][col] : null;
      if (input) input.value = "";
    });
    recalcSheet();
    if (thread.sheet.cells[selectedCell.row] && thread.sheet.cells[selectedCell.row][selectedCell.col]) {
      formulaInput.value = thread.sheet.cells[selectedCell.row][selectedCell.col].value || "";
    }
    sheetStatus.textContent = "Cleared";
    scheduleSave();
  }

  function clearSheetData() {
    const thread = getActiveThread();
    if (!thread) return;
    const ok = confirm("Clear the entire sheet? This cannot be undone.");
    if (!ok) return;
    pushSheetUndo(thread.id, thread.sheet);
    normalizeSheet(thread.sheet);
    for (let r = 0; r < thread.sheet.rows; r += 1) {
      for (let c = 0; c < thread.sheet.cols; c += 1) {
        const cell = thread.sheet.cells[r][c];
        cell.value = "";
        cell.bg = "";
        cell.align = "left";
        cell.fg = "";
        cell.wrap = "truncate";
      }
    }
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    sheetStatus.textContent = "Cleared";
    scheduleSave();
  }

  function selectRow(row) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const lastCol = Math.max(0, thread.sheet.cols - 1);
    ignoreFocusSelection = true;
    focusCell(row, 0);
    setSelectionRange(row, 0, row, lastCol);
    selectionAnchor = { row, col: 0 };
  }

  function selectRowRange(row) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const anchorRow = selectionAnchor ? selectionAnchor.row : selectedCell.row;
    const lastCol = Math.max(0, thread.sheet.cols - 1);
    ignoreFocusSelection = true;
    focusCell(row, 0);
    setSelectionRange(anchorRow, 0, row, lastCol);
    selectionAnchor = { row: anchorRow, col: 0 };
  }


  function selectColumn(col) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const lastRow = Math.max(0, thread.sheet.rows - 1);
    ignoreFocusSelection = true;
    focusCell(0, col);
    setSelectionRange(0, col, lastRow, col);
    selectionAnchor = { row: 0, col };
  }

  function selectColumnRange(col) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const anchorCol = selectionAnchor ? selectionAnchor.col : selectedCell.col;
    const lastRow = Math.max(0, thread.sheet.rows - 1);
    ignoreFocusSelection = true;
    focusCell(0, col);
    setSelectionRange(0, anchorCol, lastRow, col);
    selectionAnchor = { row: 0, col: anchorCol };
  }


  function selectAllCells() {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    const lastRow = Math.max(0, thread.sheet.rows - 1);
    const lastCol = Math.max(0, thread.sheet.cols - 1);
    ignoreFocusSelection = true;
    focusCell(0, 0);
    setSelectionRange(0, 0, lastRow, lastCol);
    selectionAnchor = { row: 0, col: 0 };
  }

  function insertSheetRow(insertIndex) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    const clamped = clamp(insertIndex, 0, thread.sheet.rows);
    const row = [];
    for (let c = 0; c < thread.sheet.cols; c += 1) {
      row.push({ value: "", bg: "", align: "left", fg: "", wrap: "truncate" });
    }
    thread.sheet.cells.splice(clamped, 0, row);
    thread.sheet.rows += 1;
    if (thread.sheet.rowHeights) thread.sheet.rowHeights.splice(clamped, 0, null);
    selectedCell = { row: clamped, col: Math.min(selectedCell.col, thread.sheet.cols - 1) };
    buildSheetTable(thread.sheet);
    scheduleSave();
  }

  function insertSheetCol(insertIndex) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    pushSheetUndo(thread.id, thread.sheet);
    const clamped = clamp(insertIndex, 0, thread.sheet.cols);
    thread.sheet.cells.forEach((row) => row.splice(clamped, 0, { value: "", bg: "", align: "left", fg: "", wrap: "truncate" }));
    thread.sheet.cols += 1;
    if (thread.sheet.colWidths) thread.sheet.colWidths.splice(clamped, 0, null);
    selectedCell = { row: Math.min(selectedCell.row, thread.sheet.rows - 1), col: clamped };
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    scheduleSave();
  }

  function removeSheetRow(rowIndex) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    if (thread.sheet.rows <= 1) {
      sheetStatus.textContent = "Minimum rows reached";
      return;
    }
    const clamped = clamp(rowIndex, 0, thread.sheet.rows - 1);
    pushSheetUndo(thread.id, thread.sheet);
    thread.sheet.cells.splice(clamped, 1);
    thread.sheet.rows -= 1;
    if (thread.sheet.rowHeights) thread.sheet.rowHeights.splice(clamped, 1);
    if (selectedCell.row >= clamped) {
      selectedCell = { row: Math.max(0, selectedCell.row - 1), col: selectedCell.col };
    }
    buildSheetTable(thread.sheet);
    scheduleSave();
  }

  function removeSheetCol(colIndex) {
    const thread = getActiveThread();
    if (!thread) return;
    normalizeSheet(thread.sheet);
    if (thread.sheet.cols <= 1) {
      sheetStatus.textContent = "Minimum columns reached";
      return;
    }
    const clamped = clamp(colIndex, 0, thread.sheet.cols - 1);
    pushSheetUndo(thread.id, thread.sheet);
    thread.sheet.cells.forEach((row) => row.splice(clamped, 1));
    thread.sheet.cols -= 1;
    if (thread.sheet.colWidths) thread.sheet.colWidths.splice(clamped, 1);
    if (selectedCell.col >= clamped) {
      selectedCell = { row: selectedCell.row, col: Math.max(0, selectedCell.col - 1) };
    }
    buildSheetTable(thread.sheet);
    populateSortColumns(thread.sheet.cols);
    scheduleSave();
  }

  function openSheetContextMenu(event, target) {
    if (!sheetContextMenu || !target) return;
    event.preventDefault();
    event.stopPropagation();
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeCellSubmenus();
    contextMenuTarget = target;
    if (target.type === "row") {
      selectRow(target.index);
    } else if (target.type === "col") {
      selectColumn(target.index);
    }
    updateSheetContextMenu(target.type);
    sheetContextMenu.classList.add("open");
    sheetContextMenu.setAttribute("aria-hidden", "false");
    positionSheetContextMenu(event.clientX, event.clientY);
  }

  function positionContextMenu(menu, x, y) {
    if (!menu) return;
    const padding = 8;
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    const rect = menu.getBoundingClientRect();
    const clampedX = Math.min(window.innerWidth - rect.width - padding, Math.max(padding, x));
    const clampedY = Math.min(window.innerHeight - rect.height - padding, Math.max(padding, y));
    menu.style.left = `${clampedX}px`;
    menu.style.top = `${clampedY}px`;
  }

  function positionSheetContextMenu(x, y) {
    positionContextMenu(sheetContextMenu, x, y);
  }

  function updateSheetContextMenu(type) {
    if (!sheetContextMenu) return;
    const items = sheetContextMenu.querySelectorAll(".context-item");
    items.forEach((item) => {
      const action = item.dataset.action || "";
      const isRow = action.includes("row");
      const isCol = action.includes("col");
      item.classList.toggle("hidden", type === "row" ? !isRow : !isCol);
      item.disabled = false;
    });
    const thread = getActiveThread();
    if (!thread) return;
    if (type === "row" && thread.sheet.rows <= 1) {
      const remove = sheetContextMenu.querySelector("[data-action=\"remove-row\"]");
      if (remove) remove.disabled = true;
    }
    if (type === "col" && thread.sheet.cols <= 1) {
      const remove = sheetContextMenu.querySelector("[data-action=\"remove-col\"]");
      if (remove) remove.disabled = true;
    }
  }

  function closeSheetContextMenu() {
    if (!sheetContextMenu) return;
    sheetContextMenu.classList.remove("open");
    sheetContextMenu.setAttribute("aria-hidden", "true");
    contextMenuTarget = null;
  }

  function handleSheetContextAction(action) {
    if (!contextMenuTarget || !action) return;
    const { type, index } = contextMenuTarget;
    if (type === "row") {
      if (action === "insert-row-above") insertSheetRow(index);
      if (action === "insert-row-below") insertSheetRow(index + 1);
      if (action === "remove-row") removeSheetRow(index);
    }
    if (type === "col") {
      if (action === "insert-col-left") insertSheetCol(index);
      if (action === "insert-col-right") insertSheetCol(index + 1);
      if (action === "remove-col") removeSheetCol(index);
    }
    closeSheetContextMenu();
  }

  function isCellInSelection(row, col) {
    const bounds = getSelectionBounds();
    if (!bounds) {
      return row === selectedCell.row && col === selectedCell.col;
    }
    return row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd;
  }

  function closeCellSubmenus() {
    [cellEditMenu, cellAlignMenu, cellOptionsMenu, cellFormulaMenu].forEach((menu) => {
      if (!menu) return;
      menu.classList.remove("open");
      menu.setAttribute("aria-hidden", "true");
    });
    if (activeCellSubmenuTrigger) {
      activeCellSubmenuTrigger.classList.remove("submenu-open");
    }
    activeCellSubmenu = null;
    activeCellSubmenuTrigger = null;
  }

  function openCellSubmenu(trigger) {
    if (!trigger || !trigger.dataset.submenu) return;
    const submenu = document.getElementById(trigger.dataset.submenu);
    if (!submenu) return;
    if (activeCellSubmenu === submenu && submenu.classList.contains("open")) return;
    closeCellSubmenus();
    submenu.classList.add("open");
    submenu.setAttribute("aria-hidden", "false");
    activeCellSubmenu = submenu;
    activeCellSubmenuTrigger = trigger;
    trigger.classList.add("submenu-open");
    positionCellSubmenu(submenu, trigger.getBoundingClientRect());
  }

  function positionCellSubmenu(menu, anchorRect) {
    if (!menu || !anchorRect) return;
    const padding = 8;
    let x = anchorRect.right + 6;
    let y = anchorRect.top;
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    const rect = menu.getBoundingClientRect();
    if (rect.right > window.innerWidth - padding) {
      x = anchorRect.left - rect.width - 6;
    }
    if (rect.left < padding) {
      x = padding;
    }
    if (rect.bottom > window.innerHeight - padding) {
      y = Math.max(padding, window.innerHeight - rect.height - padding);
    }
    if (y < padding) {
      y = padding;
    }
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
  }

  function openCellContextMenu(event, target) {
    if (!cellContextMenu || !target) return;
    event.preventDefault();
    event.stopPropagation();
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    closeDiagnosticsModal();
    closeSchemeContextMenu();
    cellContextTarget = target;
    cellContextPosition = { x: event.clientX, y: event.clientY };
    const inSelection = isCellInSelection(target.row, target.col);
    if (inSelection) {
      ignoreFocusSelection = true;
    }
    focusCell(target.row, target.col);
    updateCellContextMenu();
    cellContextMenu.classList.add("open");
    cellContextMenu.setAttribute("aria-hidden", "false");
    positionContextMenu(cellContextMenu, event.clientX, event.clientY);
  }

  function updateCellContextMenu() {
    if (!cellContextMenu) return;
    const thread = getActiveThread();
    if (!thread) return;
    const cell = thread.sheet.cells[selectedCell.row] && thread.sheet.cells[selectedCell.row][selectedCell.col];
    if (!cell) return;
    if (cellTextColorInput) {
      const fg = normalizeHexColor(cell.fg);
      cellTextColorInput.value = fg || thread.settings.sheetColor;
      cellTextColorInput.classList.toggle("is-empty", !fg);
    }
    if (cellBgColorInput) {
      const bg = normalizeHexColor(cell.bg);
      const fallback = shadeColorEl ? shadeColorEl.value : "#1a2431";
      cellBgColorInput.value = bg || fallback;
      cellBgColorInput.classList.toggle("is-empty", !bg);
    }
    const wrapBtn = cellOptionsMenu ? cellOptionsMenu.querySelector("[data-action=\"wrap\"]") : null;
    const truncateBtn = cellOptionsMenu ? cellOptionsMenu.querySelector("[data-action=\"truncate\"]") : null;
    const isWrap = cell.wrap === "wrap";
    if (wrapBtn) wrapBtn.classList.toggle("active", isWrap);
    if (truncateBtn) truncateBtn.classList.toggle("active", !isWrap);
  }

  function closeCellContextMenu() {
    if (!cellContextMenu) return;
    cellContextMenu.classList.remove("open");
    cellContextMenu.setAttribute("aria-hidden", "true");
    cellContextTarget = null;
    cellContextPosition = null;
    closeCellSubmenus();
  }

  function handleCellContextAction(action) {
    if (!action) return;
    if (action === "copy") copySelectionToClipboard();
    if (action === "cut") cutSelectionToClipboard();
    if (action === "paste") pasteSelectionFromClipboard();
    if (action === "align-left") applyAlignment("left");
    if (action === "align-center") applyAlignment("center");
    if (action === "align-right") applyAlignment("right");
    if (action === "wrap") applyWrapMode("wrap");
    if (action === "truncate") applyWrapMode("truncate");
    closeCellContextMenu();
  }

  function openSchemeContextMenu(event) {
    if (!schemeContextMenu || !journalEl) return;
    event.preventDefault();
    event.stopPropagation();
    closePageSettings();
    closePanelPopovers();
    closeBreakTimerPanel();
    closeSheetContextMenu();
    closeCellContextMenu();
    closeCellSubmenus();
    closeTabsContextMenu();
    closeCacheViewModal();
    closeDiagnosticsModal();
    schemeContextActive = true;
    updateSchemeContextMenu();
    schemeContextMenu.classList.add("open");
    schemeContextMenu.setAttribute("aria-hidden", "false");
    positionContextMenu(schemeContextMenu, event.clientX, event.clientY);
  }

  function updateSchemeContextMenu() {
    if (!schemeContextMenu || !schemeTextColorInput) return;
    const color = normalizeHexColor(journalColorEl.value) || "#d9dde7";
    schemeTextColorInput.value = color;
  }

  function closeSchemeContextMenu() {
    if (!schemeContextMenu) return;
    schemeContextMenu.classList.remove("open");
    schemeContextMenu.setAttribute("aria-hidden", "true");
    schemeContextActive = false;
  }

  function handleSchemeContextAction(action) {
    if (!action) return;
    if (action === "copy") copyJournalSelection();
    if (action === "cut") cutJournalSelection();
    if (action === "paste") pasteTextToJournal();
    closeSchemeContextMenu();
  }

  async function pasteSelectionFromClipboard() {
    const thread = getActiveThread();
    if (!thread) return;
    if (!navigator.clipboard || !window.isSecureContext) {
      alert("Paste is unavailable in this browser context.");
      return;
    }
    let text = "";
    try {
      text = await navigator.clipboard.readText();
    } catch (err) {
      alert("Failed to read clipboard.");
      return;
    }
    const forceGrid = isLastSheetCopy(text);
    const grid = parseClipboardGrid(text, forceGrid);
    if (!grid.length) return;
    if (grid.length === 1 && grid[0].length === 1) {
      commitCellValue(thread, selectedCell.row, selectedCell.col, grid[0][0], true);
      const target = sheetInputs[selectedCell.row] ? sheetInputs[selectedCell.row][selectedCell.col] : null;
      if (target) target.value = grid[0][0];
      formulaInput.value = grid[0][0];
      return;
    }
    applyGridToSheet(thread, selectedCell.row, selectedCell.col, grid);
  }

  function getTextareaSelection(textarea) {
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    return {
      start,
      end,
      text: textarea.value.slice(start, end)
    };
  }

  function replaceTextareaSelection(textarea, replacement) {
    const { start, end } = getTextareaSelection(textarea);
    const value = textarea.value;
    textarea.value = value.slice(0, start) + replacement + value.slice(end);
    textarea.selectionStart = textarea.selectionEnd = start + replacement.length;
  }

  function copyJournalSelection() {
    if (!journalEl) return;
    const selection = getTextareaSelection(journalEl);
    if (!selection.text) return;
    copyTextToClipboard(selection.text, (success) => {
      journalStatus.textContent = success ? "Copied" : "Copy failed";
    });
  }

  function cutJournalSelection() {
    if (!journalEl) return;
    const selection = getTextareaSelection(journalEl);
    if (!selection.text) return;
    copyTextToClipboard(selection.text, (success) => {
      journalStatus.textContent = success ? "Cut" : "Copy failed";
    });
    replaceTextareaSelection(journalEl, "");
    journalEl.dispatchEvent(new Event("input"));
  }

  async function pasteTextToJournal() {
    if (!journalEl) return;
    if (!navigator.clipboard || !window.isSecureContext) {
      alert("Paste is unavailable in this browser context.");
      return;
    }
    let text = "";
    try {
      text = await navigator.clipboard.readText();
    } catch (err) {
      alert("Failed to read clipboard.");
      return;
    }
    replaceTextareaSelection(journalEl, text);
    journalEl.dispatchEvent(new Event("input"));
  }

  function getFormulaSelectionRange() {
    const bounds = getSelectionBounds();
    if (!bounds) return "";
    if (bounds.rowStart === bounds.rowEnd && bounds.colStart === bounds.colEnd) return "";
    const start = `${indexToCol(bounds.colStart)}${bounds.rowStart + 1}`;
    const end = `${indexToCol(bounds.colEnd)}${bounds.rowEnd + 1}`;
    if (start === end) return "";
    return `${start}:${end}`;
  }


  function addFormulaAtSelection(formulaName) {
    const thread = getActiveThread();
    if (!thread) return;
    const name = String(formulaName || "").trim().toUpperCase();
    if (!name) return;
    const range = getFormulaSelectionRange();
    const formula = range ? `=${name}(${range})` : `=${name}()`;
    commitCellValue(thread, selectedCell.row, selectedCell.col, formula, true);
    focusCell(selectedCell.row, selectedCell.col);
    formulaInput.focus();
    const cursor = formula.length - 1;
    formulaInput.setSelectionRange(cursor, cursor);
  }

  function addSheetRow() {
    const thread = getActiveThread();
    if (!thread) return;
    insertSheetRow(thread.sheet.rows);
  }

  function addSheetCol() {
    const thread = getActiveThread();
    if (!thread) return;
    insertSheetCol(thread.sheet.cols);
  }

  function exportBundle() {
    const bundle = {
      version: 1,
      exportedAt: new Date().toISOString(),
      state
    };
    const name = `multithreader-bundle-${safeTimestamp()}.json`;
    downloadFile(name, JSON.stringify(bundle, null, 2), "application/json");
  }

  function exportText() {
    const thread = getActiveThread();
    if (!thread) return;
    const name = safeFileName(thread.name) + ".txt";
    downloadFile(name, thread.journal, "text/plain");
  }

  function buildCsv(sheet) {
    const rows = [];
    for (let r = 0; r < sheet.rows; r += 1) {
      const row = [];
      for (let c = 0; c < sheet.cols; c += 1) {
        row.push(csvEscape(getCellDisplay(sheet, r, c)));
      }
      rows.push(row.join(","));
    }
    return rows.join("\n");
  }

  function exportCsv() {
    const thread = getActiveThread();
    if (!thread) return;
    const csv = buildCsv(thread.sheet);
    const name = safeFileName(thread.name) + ".csv";
    downloadFile(name, csv, "text/csv");
  }

  function copyCsvToClipboard() {
    const thread = getActiveThread();
    if (!thread) return;
    const csv = buildCsv(thread.sheet);
    copyTextToClipboard(csv, (success) => {
      sheetStatus.textContent = success ? "CSV copied" : "Copy failed";
    });
  }

  function copySchemeToClipboard() {
    const thread = getActiveThread();
    if (!thread) return;
    const text = thread.journal || "";
    copyTextToClipboard(text, (success) => {
      journalStatus.textContent = success ? "Copied" : "Copy failed";
    });
  }

  function buildSelectionText(thread) {
    if (!thread) return "";
    const bounds = getSelectionBounds();
    if (!bounds) {
      return getCellDisplay(thread.sheet, selectedCell.row, selectedCell.col);
    }
    const lines = [];
    for (let r = bounds.rowStart; r <= bounds.rowEnd; r += 1) {
      const row = [];
      for (let c = bounds.colStart; c <= bounds.colEnd; c += 1) {
        row.push(getCellDisplay(thread.sheet, r, c));
      }
      lines.push(row.join("\t"));
    }
    return lines.join("\n");
  }

  function copySelectionToClipboard() {
    const thread = getActiveThread();
    if (!thread) return;
    const text = buildSelectionText(thread);
    lastSheetCopy = { text, isMulti: selectionIsMultiple() };
    copyTextToClipboard(text, (success) => {
      sheetStatus.textContent = success ? "Copied" : "Copy failed";
    });
  }

  function cutSelectionToClipboard() {
    const thread = getActiveThread();
    if (!thread) return;
    const text = buildSelectionText(thread);
    lastSheetCopy = { text, isMulti: selectionIsMultiple() };
    copyTextToClipboard(text, (success) => {
      sheetStatus.textContent = success ? "Cut" : "Copy failed";
    });
    clearSelectedCells();
  }

  function exportZip() {
    const thread = getActiveThread();
    if (!thread) return;
    if (!window.JSZip) {
      alert("ZIP export is unavailable. Use TXT/CSV export instead.");
      return;
    }
    const zip = new window.JSZip();
    zip.file(safeFileName(thread.name) + ".txt", thread.journal || "");
    zip.file(safeFileName(thread.name) + ".csv", buildCsv(thread.sheet));
    zip.generateAsync({ type: "blob" }).then((blob) => {
      downloadBlob(safeFileName(thread.name) + ".zip", blob);
    });
  }

  function downloadFile(name, content, type) {
    const blob = new Blob([content], { type });
    downloadBlob(name, blob);
  }

  function downloadBlob(name, blob) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 300);
  }

  function copyTextToClipboard(text, onDone) {
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(
        () => onDone(true),
        () => onDone(fallbackCopyText(text))
      );
      return;
    }
    onDone(fallbackCopyText(text));
  }

  function fallbackCopyText(text) {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    textarea.style.left = "-9999px";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (err) {
      ok = false;
    }
    textarea.remove();
    return ok;
  }

  function normalizeHexColor(input) {
    const value = String(input ?? "").trim();
    if (!value) return null;
    const normalized = value.startsWith("#") ? value : `#${value}`;
    if (!/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(normalized)) return null;
    if (normalized.length === 4) {
      return `#${normalized[1]}${normalized[1]}${normalized[2]}${normalized[2]}${normalized[3]}${normalized[3]}`.toLowerCase();
    }
    return normalized.toLowerCase();
  }

  function csvEscape(value) {
    const text = String(value ?? "");
    if (/[",\n]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  }

  function safeFileName(name) {
    return name.trim().replace(/[^a-z0-9_-]+/gi, "-").replace(/-+/g, "-").toLowerCase() || "thread";
  }

  function safeTimestamp() {
    return new Date().toISOString().replace(/[:.]/g, "-");
  }
})();
