"use strict";

const STORAGE_KEY = "midp_web_v2";
const MAX_FIELDS = 11;
const MIN_MATRIX_ROWS = 12;
const SAVE_DEBOUNCE_MS = 220;
const SEARCH_DEBOUNCE_MS = 140;
const DRAFT_META_START = "__meta_startDate";
const DRAFT_META_END = "__meta_endDate";
const DRAFT_META_BASE_UNITS = "__meta_baseUnits";
const DRAFT_META_REAL_PROGRESS = "__meta_realProgress";
const TRACKING_MAX_NOTE = 120;
const BI_MAX_WIDGETS = 12;
const BI_DETAIL_ROW_LIMIT = 500;
const BI_WIDGET_TABLE_LIMIT = 12;
const BI_CANVAS_MIN_WIDTH = 280;
const BI_CANVAS_MIN_HEIGHT = 220;
const BI_CANVAS_SURFACE_HEIGHT = 780;
const BI_CANVAS_SURFACE_MIN_WIDTH = 1200;
const BI_CANVAS_SURFACE_MIN_EDIT_WIDTH = 760;
const BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT = 420;
const BI_CANVAS_SURFACE_MAX_EDIT_WIDTH = 6000;
const BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT = 4000;
const BI_CANVAS_AUTO_EXPAND_MARGIN = 120;
const BI_CANVAS_AUTO_EXPAND_STEP = 240;
const QUICKSIGHT_CANVAS_MIN_WIDTH = 1200;
const QUICKSIGHT_CANVAS_MIN_HEIGHT = 760;
const QUICKSIGHT_CANVAS_DEFAULT_WIDTH = 1800;
const QUICKSIGHT_CANVAS_DEFAULT_HEIGHT = 1100;
const QUICKSIGHT_CANVAS_ZOOM_DEFAULT = 100;
const QUICKSIGHT_CANVAS_ZOOM_MIN = 25;
const QUICKSIGHT_CANVAS_ZOOM_MAX = 300;
const QUICKSIGHT_CANVAS_ZOOM_STEP = 10;
const QUICKSIGHT_VISUAL_MIN_WIDTH = 320;
const QUICKSIGHT_VISUAL_MIN_HEIGHT = 220;
const QUICKSIGHT_VISUAL_DEFAULT_WIDTH = 420;
const QUICKSIGHT_VISUAL_DEFAULT_HEIGHT = 280;
const QUICKSIGHT_VISUAL_GAP_X = 24;
const QUICKSIGHT_VISUAL_GAP_Y = 24;
const QUICKSIGHT_CANVAS_PADDING = 20;
const QUICKSIGHT_DRILL_PATH_SLOTS = 3;
const BI_CANVAS_ZOOM_DEFAULT = 100;
const BI_CANVAS_ZOOM_MIN = 25;
const BI_CANVAS_ZOOM_MAX = 300;
const BI_CANVAS_ZOOM_STEP = 10;
const BI_CANVAS_ZOOM_PRESETS = [25, 50, 75, 100, 200];
const BI_COLOR_ROWS_LIMIT = 40;
const BI_EXPORT_WIDGET_SCALE = 2;
const BI_EXPORT_BOARD_SCALE = 2;
const BI_CUSTOM_CONTENT_TYPES = new Set(["table", "scorecard", "pivot", "sankey", "timeline"]);
const BI_CHART_TYPES = new Set([
  "bar",
  "bar_horizontal",
  "bar_stacked",
  "line",
  "timeseries",
  "donut",
  "area",
  "combo",
  "pie",
  "treemap",
  "funnel",
  "gauge",
  "table",
  "pivot",
  "scorecard",
  "waterfall",
  "radar",
  "pareto",
  "scatter",
  "bubble",
  "bullet",
  "boxplot",
  "candlestick",
  "sankey",
  "timeline"
]);
const BI_BAR_FAMILY_TYPES = new Set(["bar", "bar_horizontal", "bar_stacked"]);
const BI_HORIZONTAL_BAR_TYPES = new Set(["bar_horizontal"]);
const BI_STACKED_BAR_TYPES = new Set(["bar_stacked"]);
const QUICKSIGHT_ECHARTS_REQUIRED_TYPES = new Set(["bar_horizontal", "bar_stacked"]);
const BI_ALLOWED_FONT_FAMILIES = new Set([
  "Segoe UI",
  "Arial",
  "Tahoma",
  "Verdana",
  "Trebuchet MS",
  "Georgia",
  "Times New Roman",
  "Courier New"
]);
const BI_FIXED_FIELD_ALIASES = new Set([
  "proyecto",
  "creador",
  "fase",
  "sector",
  "nivel",
  "tipo",
  "disciplina",
  "sistema",
  "paquete",
  "numero",
  "escala"
]);

const DEFAULT_FIELD_TEMPLATES = [
  {
    label: "Proyecto",
    includeInCode: true,
    locked: true,
    rows: [
      { name: "Edificio de investigacion multidisciplinaria", code: "EIMI" }
    ]
  },
  {
    label: "Creador",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "ARTADI", code: "DS" },
      { name: "AITEC", code: "IN" }
    ]
  },
  {
    label: "Fase",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "Diseno", code: "I00" },
      { name: "Intervencion", code: "H00" },
      { name: "As-Built", code: "ZZZ" }
    ]
  },
  {
    label: "Sector",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "Pabellon I", code: "N01" },
      { name: "Pabellon H", code: "N02" },
      { name: "Varios sectores", code: "N03" }
    ]
  },
  {
    label: "Nivel",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "Nivel 1", code: "PLN" },
      { name: "Nivel 2", code: "MD" },
      { name: "Nivel 3", code: "MC" }
    ]
  },
  {
    label: "Tipo",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "Planos", code: "ARQ" },
      { name: "Memoria descriptiva", code: "EST" },
      { name: "Memoria de calculo", code: "ELE" }
    ]
  },
  {
    label: "Disciplina",
    includeInCode: true,
    locked: true,
    rows: [
      { name: "Arquitectura", code: "ALM" },
      { name: "Estructuras", code: "ALUM" },
      { name: "Electricas", code: "TMC" }
    ]
  },
  {
    label: "Sistema",
    includeInCode: true,
    locked: true,
    rows: [
      { name: "Alimentadores", code: "001" },
      { name: "Alumbrado", code: "002" },
      { name: "Tomacorrientes", code: "003" }
    ]
  },
  {
    label: "Paquete",
    includeInCode: false,
    locked: true,
    rows: [
      { name: "Paquete 01", code: "PK01" },
      { name: "Paquete 02", code: "PK02" },
      { name: "Paquete 03", code: "PK03" }
    ]
  },
  {
    label: "Numero",
    includeInCode: true,
    locked: false,
    rows: [
      { name: "001", code: "001" },
      { name: "002", code: "002" },
      { name: "003", code: "003" }
    ]
  },
  {
    label: "Escala",
    includeInCode: false,
    locked: false,
    rows: [
      { name: "1/50", code: "1/50" },
      { name: "1/100", code: "1/100" },
      { name: "1/200", code: "1/200" }
    ]
  }
];

const storageAdapter = createStorageAdapter();
let state = loadState();
let activeTab = "fields";
let fieldsEditMode = false;
let midpEditMode = false;
let packageEditMode = false;
let reviewFlowEditMode = false;
let reviewControlsEditMode = false;
let chooserLocked = false;
let drawerCloseTimer = null;
let reviewMilestoneDrawerCloseTimer = null;
let trackingCloseTimer = null;
let quickSightDrillThroughCloseTimer = null;
let trackingPanelTargetId = "";
let trackingPanelTargetType = "";
let pendingSaveTimer = null;
let pendingSearchTimer = null;
let currentSearchQuery = "";
let quickSightDrillThroughQuery = "";
let draggedNomenclatureFieldId = "";
let matrixSelectionAnchor = null;
let matrixSelectionRange = null;
let deliverableSelectionAnchor = null;
let deliverableSelectionRange = null;
let biRailMode = "data";
let biDataPanelOpen = false;
let biInspectorPanelOpen = false;
let biOtherPanelOpen = false;
const BI_STUDIO_PANEL_MIN_WIDTH = 290;
const BI_STUDIO_PANEL_MAX_WIDTH = 620;
const BI_STUDIO_PANEL_DEFAULT_WIDTH = 360;
let biStudioLeftWidth = BI_STUDIO_PANEL_DEFAULT_WIDTH;
let biStudioRightWidth = BI_STUDIO_PANEL_DEFAULT_WIDTH;
let biStudioPanelResize = null;
let biFilterFocusTimer = null;
let biVisualScopeMode = "widget";
let selectedBiWidgetId = "";
let selectedQuickSightVisualId = "";
let quickSightVisualInteraction = null;
let quickSightSuppressClickUntil = 0;
let quickSightWheelZoomRafId = 0;
let quickSightPendingWheelZoomStep = 0;
let quickSightPresentationMode = false;
let quickSightExpandedVisualId = "";
let quickSightRolePickerOpen = "";
let quickSightRolePickerFocusRole = "";
const QUICKSIGHT_HISTORY_LIMIT = 48;
const quickSightHistoryByProject = Object.create(null);
const quickSightPanelState = {
  data: true,
  visuals: true,
  props: true
};
let quickSightEditMenuOpen = false;
let quickSightZoomMenuOpen = false;
let quickSightPropsTab = "visual";
const quickSightPropsSectionState = {
  visual: Object.create(null),
  interactions: Object.create(null)
};
let biPendingImageWidgetId = "";
let biCommandPaletteOpen = false;
let biCommandPaletteSelection = 0;
let biCommandVisibleIds = [];
let biPendingRenderRafId = 0;
const biWidgetSnapshotsByProject = {};
let biCanvasInteraction = null;
let biWheelZoomRafId = 0;
let biPendingWheelZoomStep = 0;
const pendingRealAdvanceLogs = [];
const pendingPackageRealAdvanceLogs = [];
const pendingReviewControlRealAdvanceLogs = [];
const draftValuesByProject = {};
const biRuntimeStatsByProject = {};
const BI_STUDIO_UI_PREFS_KEY = "midp_bi_studio_ui_v1";
const SIDEBAR_COLLAPSED_PREF_KEY = "midp_sidebar_collapsed_v1";
let sidebarCollapsed = false;
const els = {};

document.addEventListener("DOMContentLoaded", async () => {
  bindElements();
  await hydrateStorageState();
  restoreSidebarCollapsedPreference();
  applyBiControlHints();
  wireEvents();
  restoreBiStudioUiPreferences();
  if (!biDataPanelOpen && !biInspectorPanelOpen && !biOtherPanelOpen) {
    applyBiRailMode("data", { announce: false, forceOpen: true });
  } else {
    refreshBiStudioPanelsUi();
  }
  ensureActiveProject();
  renderAll();

  window.addEventListener("beforeunload", () => {
    flushPendingSave();
  });
  window.addEventListener("resize", () => {
    updateBiStudioOverlayOffset();
  });
  document.addEventListener("pointerup", () => {
    if (activeTab === "bi") {
      updateBiWorkspaceInsets();
    }
  });
  document.addEventListener("scroll", () => {
    if (activeTab === "bi") {
      updateBiStudioOverlayOffset();
    }
  }, true);
  document.addEventListener("pointermove", (event) => {
    handleBiStudioPanelResizeMove(event);
  });
  document.addEventListener("pointerup", (event) => {
    if (biStudioPanelResize && event.pointerId === biStudioPanelResize.pointerId) {
      stopBiStudioPanelResize();
    }
  });
  document.addEventListener("pointercancel", (event) => {
    if (biStudioPanelResize && event.pointerId === biStudioPanelResize.pointerId) {
      stopBiStudioPanelResize();
    }
  });
  document.addEventListener("pointermove", (event) => {
    handleQuickSightVisualDragMove(event);
  });
  document.addEventListener("pointerup", (event) => {
    finishQuickSightVisualDrag(event, false);
  });
  document.addEventListener("pointercancel", (event) => {
    finishQuickSightVisualDrag(event, true);
  });

  if (state.projects.length > 1) {
    openProjectChooser(true);
  }
});

function normalizeBiStudioPanelWidth(value, fallback = BI_STUDIO_PANEL_DEFAULT_WIDTH) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return fallback;
  }
  return Math.min(BI_STUDIO_PANEL_MAX_WIDTH, Math.max(BI_STUDIO_PANEL_MIN_WIDTH, Math.round(numeric)));
}

function applyBiStudioPanelWidths() {
  if (!(els.biStudioShell instanceof HTMLElement)) {
    return;
  }
  const normalizedLeft = normalizeBiStudioPanelWidth(biStudioLeftWidth);
  const normalizedRight = normalizeBiStudioPanelWidth(biStudioRightWidth);
  biStudioLeftWidth = normalizedLeft;
  biStudioRightWidth = normalizedRight;
  els.biStudioShell.style.setProperty("--bi-left-panel-width", `${normalizedLeft}px`);
  els.biStudioShell.style.setProperty("--bi-right-panel-width", `${normalizedRight}px`);
}

function stopBiStudioPanelResize() {
  if (!biStudioPanelResize) {
    return;
  }
  biStudioPanelResize = null;
  document.body.classList.remove("bi-resizing-panels");
  saveBiStudioUiPreferences();
  updateBiStudioOverlayOffset();
}

function startBiStudioPanelResize(mode, event) {
  if (window.innerWidth <= 1180) {
    return;
  }
  if (!(event instanceof PointerEvent)) {
    return;
  }
  const normalizedMode = mode === "left" ? "left" : "middle";
  biStudioPanelResize = {
    mode: normalizedMode,
    pointerId: event.pointerId,
    startX: event.clientX,
    leftWidth: normalizeBiStudioPanelWidth(biStudioLeftWidth),
    rightWidth: normalizeBiStudioPanelWidth(biStudioRightWidth)
  };
  document.body.classList.add("bi-resizing-panels");
  event.preventDefault();
}

function handleBiStudioPanelResizeMove(event) {
  if (!biStudioPanelResize) {
    return;
  }
  if (!(event instanceof PointerEvent)) {
    return;
  }
  if (event.pointerId !== biStudioPanelResize.pointerId) {
    return;
  }
  const deltaX = event.clientX - biStudioPanelResize.startX;
  if (biStudioPanelResize.mode === "left") {
    biStudioLeftWidth = normalizeBiStudioPanelWidth(biStudioPanelResize.leftWidth - deltaX);
  } else {
    biStudioRightWidth = normalizeBiStudioPanelWidth(biStudioPanelResize.rightWidth - deltaX);
  }
  applyBiStudioPanelWidths();
  updateBiWorkspaceInsets();
}

function getMidpStorageApi() {
  if (typeof window === "undefined" || !window.MIDPStorage || typeof window.MIDPStorage !== "object") {
    return null;
  }
  return window.MIDPStorage;
}

function createStorageAdapter() {
  const api = getMidpStorageApi();
  if (api && typeof api.createHybridStorageAdapter === "function") {
    return api.createHybridStorageAdapter();
  }
  if (api && typeof api.createLocalStorageAdapter === "function") {
    return api.createLocalStorageAdapter();
  }
  return Object.freeze({
    kind: "unavailable",
    isAvailable() {
      return false;
    },
    getString(key, fallback = "") {
      void key;
      return fallback;
    },
    setString(key, value) {
      void key;
      void value;
      return false;
    },
    getJson(key, fallback = null) {
      void key;
      return fallback;
    },
    setJson(key, value) {
      void key;
      void value;
      return false;
    },
    remove(key) {
      void key;
      return false;
    },
    async hydrateString(key, fallback = "") {
      void key;
      return fallback;
    },
    async hydrateJson(key, fallback = null) {
      void key;
      return fallback;
    }
  });
}

async function hydrateStorageState() {
  if (!storageAdapter || typeof storageAdapter.hydrateJson !== "function" || typeof storageAdapter.hydrateString !== "function") {
    return;
  }

  const hydratedState = await storageAdapter.hydrateJson(STORAGE_KEY, null);
  if (hydratedState) {
    state = normalizeState(hydratedState);
  }

  await storageAdapter.hydrateJson(BI_STUDIO_UI_PREFS_KEY, null);
  await storageAdapter.hydrateString(SIDEBAR_COLLAPSED_PREF_KEY, "");
}

function restoreBiStudioUiPreferences() {
  const parsed = storageAdapter.getJson(BI_STUDIO_UI_PREFS_KEY, null);
  if (parsed) {
    biRailMode = normalizeBiRailMode(parsed?.mode || "data");
    biDataPanelOpen = !!parsed?.dataOpen;
    biInspectorPanelOpen = !!parsed?.inspectorOpen;
    biOtherPanelOpen = !!parsed?.otherOpen;
    biStudioLeftWidth = normalizeBiStudioPanelWidth(parsed?.leftWidth, BI_STUDIO_PANEL_DEFAULT_WIDTH);
    biStudioRightWidth = normalizeBiStudioPanelWidth(parsed?.rightWidth, BI_STUDIO_PANEL_DEFAULT_WIDTH);
  } else {
    biStudioLeftWidth = BI_STUDIO_PANEL_DEFAULT_WIDTH;
    biStudioRightWidth = BI_STUDIO_PANEL_DEFAULT_WIDTH;
  }
  applyBiStudioPanelWidths();
}

function saveBiStudioUiPreferences() {
  storageAdapter.setJson(BI_STUDIO_UI_PREFS_KEY, {
    mode: normalizeBiRailMode(biRailMode),
    dataOpen: !!biDataPanelOpen,
    inspectorOpen: !!biInspectorPanelOpen,
    otherOpen: !!biOtherPanelOpen,
    leftWidth: normalizeBiStudioPanelWidth(biStudioLeftWidth),
    rightWidth: normalizeBiStudioPanelWidth(biStudioRightWidth)
  });
}

function applySidebarCollapsedState(collapsed, persist = true) {
  sidebarCollapsed = !!collapsed;
  if (els.appLayout instanceof HTMLElement) {
    els.appLayout.classList.toggle("sidebar-collapsed", sidebarCollapsed);
  }
  if (els.sidebarToggleButton instanceof HTMLElement) {
    els.sidebarToggleButton.setAttribute("title", sidebarCollapsed ? "Mostrar panel izquierdo" : "Ocultar panel izquierdo");
    els.sidebarToggleButton.setAttribute("aria-label", sidebarCollapsed ? "mostrar sidebar" : "ocultar sidebar");
  }
  if (persist) {
    storageAdapter.setString(SIDEBAR_COLLAPSED_PREF_KEY, sidebarCollapsed ? "1" : "0");
  }
}

function restoreSidebarCollapsedPreference() {
  const saved = storageAdapter.getString(SIDEBAR_COLLAPSED_PREF_KEY, "");
  applySidebarCollapsedState(saved === "1", false);
}

function bindElements() {
  els.appLayout = document.querySelector(".app-layout");
  els.sidebarToggleButton = document.getElementById("sidebarToggleButton");
  els.projectTitle = document.getElementById("projectTitle");
  els.projectSubtitle = document.getElementById("projectSubtitle");
  els.currentViewLabel = document.getElementById("currentViewLabel");
  els.syncButton = document.getElementById("syncButton");
  els.currentUserBadge = document.getElementById("currentUserBadge");
  els.syncIndicator = document.getElementById("syncIndicator");
  els.globalSearchInput = document.getElementById("globalSearchInput");
  els.exportBackupButton = document.getElementById("exportBackupButton");
  els.importBackupButton = document.getElementById("importBackupButton");
  els.importBackupInput = document.getElementById("importBackupInput");
  els.projectTitleInput = document.getElementById("projectTitleInput");
  els.fieldSeparatorInput = document.getElementById("fieldSeparatorInput");
  els.progressControlModeSelect = document.getElementById("progressControlModeSelect");
  els.toggleFieldsEditButton = document.getElementById("toggleFieldsEditButton");
  els.fieldCounter = document.getElementById("fieldCounter");
  els.addFieldButton = document.getElementById("addFieldButton");
  els.addFieldRowButton = document.getElementById("addFieldRowButton");
  els.fieldMatrixHead = document.getElementById("fieldMatrixHead");
  els.fieldMatrixBody = document.getElementById("fieldMatrixBody");
  els.fieldsTableBody = document.getElementById("fieldsTableBody");
  els.nomenclatureOrderHint = document.getElementById("nomenclatureOrderHint");
  els.nomenclatureDragArea = document.getElementById("nomenclatureDragArea");
  els.nomenclaturePreviewText = document.getElementById("nomenclaturePreviewText");
  els.openAddDeliverableButton = document.getElementById("openAddDeliverableButton");
  els.addBlankDeliverableButton = document.getElementById("addBlankDeliverableButton");
  els.toggleMidpEditButton = document.getElementById("toggleMidpEditButton");
  els.toggleMidpAdvanceDetailsButton = document.getElementById("toggleMidpAdvanceDetailsButton");
  els.deliverableDrawerOverlay = document.getElementById("deliverableDrawerOverlay");
  els.closeDeliverableDrawerButton = document.getElementById("closeDeliverableDrawerButton");
  els.deliverableInputs = document.getElementById("deliverableInputs");
  els.drawerStartDateInput = document.getElementById("drawerStartDateInput");
  els.drawerEndDateInput = document.getElementById("drawerEndDateInput");
  els.drawerBaseUnitsInput = document.getElementById("drawerBaseUnitsInput");
  els.drawerRealProgressInput = document.getElementById("drawerRealProgressInput");
  els.saveDeliverableButton = document.getElementById("saveDeliverableButton");
  els.recomputeCodesButton = document.getElementById("recomputeCodesButton");
  els.togglePackageEditButton = document.getElementById("togglePackageEditButton");
  els.deliverablesMetaText = document.getElementById("deliverablesMetaText");
  els.deliverablesHeader = document.getElementById("deliverablesHeader");
  els.deliverablesBody = document.getElementById("deliverablesBody");
  els.syncPackagesButton = document.getElementById("syncPackagesButton");
  els.packageControlsMetaText = document.getElementById("packageControlsMetaText");
  els.packageControlsHeader = document.getElementById("packageControlsHeader");
  els.packageControlsBody = document.getElementById("packageControlsBody");
  els.toggleReviewControlsEditButton = document.getElementById("toggleReviewControlsEditButton");
  els.syncReviewControlsButton = document.getElementById("syncReviewControlsButton");
  els.reviewControlsMetaText = document.getElementById("reviewControlsMetaText");
  els.reviewControlsHeader = document.getElementById("reviewControlsHeader");
  els.reviewControlsBody = document.getElementById("reviewControlsBody");
  els.qsAddVisualButton = document.getElementById("qsAddVisualButton");
  els.qsSettingsButton = document.getElementById("qsSettingsButton");
  els.qsPublishButton = document.getElementById("qsPublishButton");
  els.qsUndoButton = document.getElementById("qsUndoButton");
  els.qsRedoButton = document.getElementById("qsRedoButton");
  els.qsSheetsMenuButton = document.getElementById("qsSheetsMenuButton");
  els.qsPresentationModeButton = document.getElementById("qsPresentationModeButton");
  els.qsRefreshButton = document.getElementById("qsRefreshButton");
  els.qsToggleDataButton = document.getElementById("qsToggleDataButton");
  els.qsToggleVisualsButton = document.getElementById("qsToggleVisualsButton");
  els.qsTogglePropsButton = document.getElementById("qsTogglePropsButton");
  els.qsEditMenuWrap = document.getElementById("qsEditMenuWrap");
  els.qsEditMenuButton = document.getElementById("qsEditMenuButton");
  els.qsEditMenu = document.getElementById("qsEditMenu");
  els.qsCanvasWidthInput = document.getElementById("qsCanvasWidthInput");
  els.qsCanvasHeightInput = document.getElementById("qsCanvasHeightInput");
  els.qsCanvasZoomSelect = document.getElementById("qsCanvasZoomSelect");
  els.qsApplyCanvasSizeButton = document.getElementById("qsApplyCanvasSizeButton");
  els.qsResetCanvasSizeButton = document.getElementById("qsResetCanvasSizeButton");
  els.qsZoomMenuWrap = document.getElementById("qsZoomMenuWrap");
  els.qsCanvasZoomMenuButton = document.getElementById("qsCanvasZoomMenuButton");
  els.qsCanvasZoomMenuLabel = document.getElementById("qsCanvasZoomMenuLabel");
  els.qsCanvasZoomMenu = document.getElementById("qsCanvasZoomMenu");
  els.qsZoomFitAllButton = document.getElementById("qsZoomFitAllButton");
  els.qsZoomFitWidthButton = document.getElementById("qsZoomFitWidthButton");
  els.qsZoomDefaultButton = document.getElementById("qsZoomDefaultButton");
  els.qsZoomFitAllLabel = document.getElementById("qsZoomFitAllLabel");
  els.qsZoomFitWidthLabel = document.getElementById("qsZoomFitWidthLabel");
  els.qsDataSummaryText = document.getElementById("qsDataSummaryText");
  els.qsSourceSelect = document.getElementById("qsSourceSelect");
  els.qsFieldsSearchInput = document.getElementById("qsFieldsSearchInput");
  els.qsFieldsList = document.getElementById("qsFieldsList");
  els.qsRoleDimensionList = document.getElementById("qsRoleDimensionList");
  els.qsRoleMetricList = document.getElementById("qsRoleMetricList");
  els.qsRoleOptionalMetricBlock = document.getElementById("qsRoleOptionalMetricBlock");
  els.qsRoleOptionalMetricList = document.getElementById("qsRoleOptionalMetricList");
  els.qsAddMetricGhostButton = document.getElementById("qsAddMetricGhostButton");
  els.qsRoleBreakdownList = document.getElementById("qsRoleBreakdownList");
  els.qsAddBreakdownGhostButton = document.getElementById("qsAddBreakdownGhostButton");
  els.qsChartTypeSelect = document.getElementById("qsChartTypeSelect");
  els.qsChartTypeGrid = document.getElementById("qsChartTypeGrid");
  els.qsGroupBySelect = document.getElementById("qsGroupBySelect");
  els.qsBreakdownSelect = document.getElementById("qsBreakdownSelect");
  els.qsMetricSelect = document.getElementById("qsMetricSelect");
  els.qsOptionalMetricSelect = document.getElementById("qsOptionalMetricSelect");
  els.qsTopNInput = document.getElementById("qsTopNInput");
  els.qsDeleteVisualButton = document.getElementById("qsDeleteVisualButton");
  els.qsSheetTitle = document.getElementById("qsSheetTitle");
  els.qsSheetAddButton = document.getElementById("qsSheetAddButton");
  els.qsCanvasMetaText = document.getElementById("qsCanvasMetaText");
  els.qsCanvasBoard = document.getElementById("qsCanvasBoard");
  els.qsPropertiesBody = document.getElementById("qsPropertiesBody");
  els.qsDrillThroughOverlay = document.getElementById("qsDrillThroughOverlay");
  els.qsDrillThroughTitle = document.getElementById("qsDrillThroughTitle");
  els.qsDrillThroughSubtitle = document.getElementById("qsDrillThroughSubtitle");
  els.qsDrillThroughMeta = document.getElementById("qsDrillThroughMeta");
  els.qsDrillThroughKpis = document.getElementById("qsDrillThroughKpis");
  els.qsDrillThroughSearchInput = document.getElementById("qsDrillThroughSearchInput");
  els.qsDrillThroughExportButton = document.getElementById("qsDrillThroughExportButton");
  els.qsDrillThroughHeader = document.getElementById("qsDrillThroughHeader");
  els.qsDrillThroughBody = document.getElementById("qsDrillThroughBody");
  els.closeQsDrillThroughButton = document.getElementById("closeQsDrillThroughButton");
  els.qsShell = document.getElementById("qsShell");
  els.qsFrame = document.querySelector("#tab-quicksight .qs-frame");
  els.qsDataPanel = document.getElementById("qsDataPanel");
  els.qsVisualsPanel = document.getElementById("qsVisualsPanel");
  els.qsPropsPanel = document.getElementById("qsPropsPanel");
  els.biRefreshButton = document.getElementById("biRefreshButton");
  els.biOpenCommandPaletteButton = document.getElementById("biOpenCommandPaletteButton");
  els.biExportBoardPngButton = document.getElementById("biExportBoardPngButton");
  els.biExportRowsButton = document.getElementById("biExportRowsButton");
  els.biAddTextWidgetButton = document.getElementById("biAddTextWidgetButton");
  els.biAddImageWidgetButton = document.getElementById("biAddImageWidgetButton");
  els.biAutoLayoutButton = document.getElementById("biAutoLayoutButton");
  els.biGenerateDashboardButton = document.getElementById("biGenerateDashboardButton");
  els.biImageUploadInput = document.getElementById("biImageUploadInput");
  els.biStartDateInput = document.getElementById("biStartDateInput");
  els.biEndDateInput = document.getElementById("biEndDateInput");
  els.biTextFilterInput = document.getElementById("biTextFilterInput");
  els.biGlobalSourceSelect = document.getElementById("biGlobalSourceSelect");
  els.biCrossFilterScopeSelect = document.getElementById("biCrossFilterScopeSelect");
  els.biInvalidOnlyCheckbox = document.getElementById("biInvalidOnlyCheckbox");
  els.biKpiGrid = document.getElementById("biKpiGrid");
  els.biWidgetsMetaText = document.getElementById("biWidgetsMetaText");
  els.biWidgetNameInput = document.getElementById("biWidgetNameInput");
  els.biWidgetSourceSelect = document.getElementById("biWidgetSourceSelect");
  els.biWidgetGroupBySelect = document.getElementById("biWidgetGroupBySelect");
  els.biWidgetMetricSelect = document.getElementById("biWidgetMetricSelect");
  els.biDataRolesSection = document.getElementById("biDataRolesSection");
  els.biWidgetBreakdownSelect = document.getElementById("biWidgetBreakdownSelect");
  els.biWidgetOptionalMetricSelect = document.getElementById("biWidgetOptionalMetricSelect");
  els.biWidgetDateDimensionSelect = document.getElementById("biWidgetDateDimensionSelect");
  els.biWidgetChartTypeSelect = document.getElementById("biWidgetChartTypeSelect");
  els.biChartTypeButtons = document.getElementById("biChartTypeButtons");
  els.biSelectedDimensionChip = document.getElementById("biSelectedDimensionChip");
  els.biSelectedMetricChip = document.getElementById("biSelectedMetricChip");
  els.biWidgetTopNInput = document.getElementById("biWidgetTopNInput");
  els.biWidgetSortModeSelect = document.getElementById("biWidgetSortModeSelect");
  els.biUpdateWidgetButton = document.getElementById("biUpdateWidgetButton");
  els.biAddWidgetButton = document.getElementById("biAddWidgetButton");
  els.biCatalogSourceSelect = document.getElementById("biCatalogSourceSelect");
  els.biCatalogSearchInput = document.getElementById("biCatalogSearchInput");
  els.biCatalogSummaryText = document.getElementById("biCatalogSummaryText");
  els.biCatalogFieldsList = document.getElementById("biCatalogFieldsList");
  els.biCrossFilterText = document.getElementById("biCrossFilterText");
  els.biClearCrossFilterButton = document.getElementById("biClearCrossFilterButton");
  els.biInsightsGrid = document.getElementById("biInsightsGrid");
  els.biCommandPaletteOverlay = document.getElementById("biCommandPaletteOverlay");
  els.biCloseCommandPaletteButton = document.getElementById("biCloseCommandPaletteButton");
  els.biCommandPaletteInput = document.getElementById("biCommandPaletteInput");
  els.biCommandPaletteList = document.getElementById("biCommandPaletteList");
  els.biStudioShell = document.getElementById("biStudioShell");
  els.biStudioLeftResizeHandle = document.getElementById("biStudioLeftResizeHandle");
  els.biStudioMiddleResizeHandle = document.getElementById("biStudioMiddleResizeHandle");
  els.biStudioLeft = document.getElementById("biStudioLeft");
  els.biStudioRight = document.getElementById("biStudioRight");
  els.biStudioOtherPanel = document.getElementById("biStudioOtherPanel");
  els.biStudioTitle = document.getElementById("biStudioTitle");
  els.biHideLeftPanelButton = document.getElementById("biHideLeftPanelButton");
  els.biHideDataPanelButton = document.getElementById("biHideDataPanelButton");
  els.biSourceSection = document.getElementById("biSourceSection");
  els.biDimensionSection = document.getElementById("biDimensionSection");
  els.biMetricSection = document.getElementById("biMetricSection");
  els.biChartSection = document.getElementById("biChartSection");
  els.biChartVisualSettingsSection = document.getElementById("biChartVisualSettingsSection");
  els.biChartConfigIntro = document.getElementById("biChartConfigIntro");
  els.biChartConfigGroups = document.getElementById("biChartConfigGroups");
  els.biElementConfigSection = document.getElementById("biElementConfigSection");
  els.biElementConfigTitle = document.getElementById("biElementConfigTitle");
  els.biElementConfigBody = document.getElementById("biElementConfigBody");
  els.biElementConfigHint = document.getElementById("biElementConfigHint");
  els.biBoardSection = document.getElementById("biBoardSection");
  els.biColorSection = document.getElementById("biColorSection");
  els.biFilterSection = document.getElementById("biFilterSection");
  els.biPerformanceSection = document.getElementById("biPerformanceSection");
  els.biSettingsSection = document.getElementById("biSettingsSection");
  els.biBuilderSection = document.getElementById("biBuilderSection");
  els.biCanvasPresetSelect = document.getElementById("biCanvasPresetSelect");
  els.biCanvasWidthInput = document.getElementById("biCanvasWidthInput");
  els.biCanvasHeightInput = document.getElementById("biCanvasHeightInput");
  els.biCanvasZoomSelect = document.getElementById("biCanvasZoomSelect");
  els.biZoomMenuWrap = document.getElementById("biZoomMenuWrap");
  els.biCanvasZoomMenuButton = document.getElementById("biCanvasZoomMenuButton");
  els.biCanvasZoomMenu = document.getElementById("biCanvasZoomMenu");
  els.biZoomFitAllButton = document.getElementById("biZoomFitAllButton");
  els.biZoomFitWidthButton = document.getElementById("biZoomFitWidthButton");
  els.biZoomDefaultButton = document.getElementById("biZoomDefaultButton");
  els.biZoomFitAllLabel = document.getElementById("biZoomFitAllLabel");
  els.biZoomFitWidthLabel = document.getElementById("biZoomFitWidthLabel");
  els.biApplyCanvasSizeButton = document.getElementById("biApplyCanvasSizeButton");
  els.biResetCanvasSizeButton = document.getElementById("biResetCanvasSizeButton");
  els.biShowCanvasGridCheckbox = document.getElementById("biShowCanvasGridCheckbox");
  els.biSnapToGridCheckbox = document.getElementById("biSnapToGridCheckbox");
  els.biGridSnapSizeInput = document.getElementById("biGridSnapSizeInput");
  els.biPerformanceModeSelect = document.getElementById("biPerformanceModeSelect");
  els.biOptimizeNowButton = document.getElementById("biOptimizeNowButton");
  els.biResetPerformanceButton = document.getElementById("biResetPerformanceButton");
  els.biPerfRenderMsText = document.getElementById("biPerfRenderMsText");
  els.biPerfWidgetsText = document.getElementById("biPerfWidgetsText");
  els.biPerfPointsText = document.getElementById("biPerfPointsText");
  els.biPerfRowsText = document.getElementById("biPerfRowsText");
  els.biPerfSummaryText = document.getElementById("biPerfSummaryText");
  els.biColorSourceSelect = document.getElementById("biColorSourceSelect");
  els.biColorGroupBySelect = document.getElementById("biColorGroupBySelect");
  els.biColorLegendList = document.getElementById("biColorLegendList");
  els.biColorSummaryText = document.getElementById("biColorSummaryText");
  els.biAutoColorButton = document.getElementById("biAutoColorButton");
  els.biClearColorButton = document.getElementById("biClearColorButton");
  els.biShowLegendCheckbox = document.getElementById("biShowLegendCheckbox");
  els.biShowGridCheckbox = document.getElementById("biShowGridCheckbox");
  els.biShowAxisLabelsCheckbox = document.getElementById("biShowAxisLabelsCheckbox");
  els.biShowDataLabelsCheckbox = document.getElementById("biShowDataLabelsCheckbox");
  els.biAxisXLabelInput = document.getElementById("biAxisXLabelInput");
  els.biAxisYLabelInput = document.getElementById("biAxisYLabelInput");
  els.biLabelMaxCharsInput = document.getElementById("biLabelMaxCharsInput");
  els.biValueDecimalsInput = document.getElementById("biValueDecimalsInput");
  els.biFontSizeInput = document.getElementById("biFontSizeInput");
  els.biFontFamilyTitleSelect = document.getElementById("biFontFamilyTitleSelect");
  els.biFontFamilyAxisSelect = document.getElementById("biFontFamilyAxisSelect");
  els.biFontFamilyLabelSelect = document.getElementById("biFontFamilyLabelSelect");
  els.biFontFamilyTooltipSelect = document.getElementById("biFontFamilyTooltipSelect");
  els.biFontSizeTitleInput = document.getElementById("biFontSizeTitleInput");
  els.biFontSizeAxisInput = document.getElementById("biFontSizeAxisInput");
  els.biFontSizeLabelsInput = document.getElementById("biFontSizeLabelsInput");
  els.biFontSizeTooltipInput = document.getElementById("biFontSizeTooltipInput");
  els.biLabelColorModeSelect = document.getElementById("biLabelColorModeSelect");
  els.biLabelColorInput = document.getElementById("biLabelColorInput");
  els.biLineWidthInput = document.getElementById("biLineWidthInput");
  els.biSeriesLineStyleSelect = document.getElementById("biSeriesLineStyleSelect");
  els.biSeriesMarkerStyleSelect = document.getElementById("biSeriesMarkerStyleSelect");
  els.biSeriesMarkerSizeInput = document.getElementById("biSeriesMarkerSizeInput");
  els.biAreaOpacityInput = document.getElementById("biAreaOpacityInput");
  els.biSmoothLinesCheckbox = document.getElementById("biSmoothLinesCheckbox");
  els.biResetVisualSettingsButton = document.getElementById("biResetVisualSettingsButton");
  els.biApplyVisualScopeButton = document.getElementById("biApplyVisualScopeButton");
  els.biResetVisualScopeButton = document.getElementById("biResetVisualScopeButton");
  els.biUiModeSelect = document.getElementById("biUiModeSelect");
  els.biVisualScopeSelect = document.getElementById("biVisualScopeSelect");
  els.biVisualScopeMeta = document.getElementById("biVisualScopeMeta");
  els.biWidgetShowTitleCheckbox = document.getElementById("biWidgetShowTitleCheckbox");
  els.biWidgetShowBorderCheckbox = document.getElementById("biWidgetShowBorderCheckbox");
  els.biWidgetShowBackgroundCheckbox = document.getElementById("biWidgetShowBackgroundCheckbox");
  els.biWidgetChromeMeta = document.getElementById("biWidgetChromeMeta");
  els.biVisualPresetSelect = document.getElementById("biVisualPresetSelect");
  els.biApplyVisualPresetButton = document.getElementById("biApplyVisualPresetButton");
  els.biSavePresetNameInput = document.getElementById("biSavePresetNameInput");
  els.biSaveVisualPresetButton = document.getElementById("biSaveVisualPresetButton");
  els.biNumberFormatSelect = document.getElementById("biNumberFormatSelect");
  els.biValuePrefixInput = document.getElementById("biValuePrefixInput");
  els.biValueSuffixInput = document.getElementById("biValueSuffixInput");
  els.biLegendPositionSelect = document.getElementById("biLegendPositionSelect");
  els.biLegendMaxItemsInput = document.getElementById("biLegendMaxItemsInput");
  els.biAxisScaleSelect = document.getElementById("biAxisScaleSelect");
  els.biAxisMinInput = document.getElementById("biAxisMinInput");
  els.biAxisMaxInput = document.getElementById("biAxisMaxInput");
  els.biShowTargetLineCheckbox = document.getElementById("biShowTargetLineCheckbox");
  els.biTargetLineValueInput = document.getElementById("biTargetLineValueInput");
  els.biTargetLineLabelInput = document.getElementById("biTargetLineLabelInput");
  els.biTargetLineColorInput = document.getElementById("biTargetLineColorInput");
  els.biBarWidthRatioInput = document.getElementById("biBarWidthRatioInput");
  els.biGridOpacityInput = document.getElementById("biGridOpacityInput");
  els.biGridDashInput = document.getElementById("biGridDashInput");
  els.biStackModeSelect = document.getElementById("biStackModeSelect");
  els.biTooltipModeSelect = document.getElementById("biTooltipModeSelect");
  els.biRailButtons = document.querySelectorAll("[data-bi-rail-mode]");
  els.biWidgetsGrid = document.getElementById("biWidgetsGrid");
  els.biRowsMetaText = document.getElementById("biRowsMetaText");
  els.biRowsHeader = document.getElementById("biRowsHeader");
  els.biRowsBody = document.getElementById("biRowsBody");
  els.openAddReviewMilestoneButton = document.getElementById("openAddReviewMilestoneButton");
  els.toggleReviewFlowEditButton = document.getElementById("toggleReviewFlowEditButton");
  els.addBlankReviewMilestoneButton = document.getElementById("addBlankReviewMilestoneButton");
  els.reviewMilestoneNameInput = document.getElementById("reviewMilestoneNameInput");
  els.reviewMilestoneWeightInput = document.getElementById("reviewMilestoneWeightInput");
  els.reviewMilestoneDrawerOverlay = document.getElementById("reviewMilestoneDrawerOverlay");
  els.closeReviewMilestoneDrawerButton = document.getElementById("closeReviewMilestoneDrawerButton");
  els.saveReviewMilestoneButton = document.getElementById("saveReviewMilestoneButton");
  els.reviewMilestonesBody = document.getElementById("reviewMilestonesBody");
  els.reviewMilestonesMetaText = document.getElementById("reviewMilestonesMetaText");
  els.trackingPanelOverlay = document.getElementById("trackingPanelOverlay");
  els.trackingPanelTitle = document.getElementById("trackingPanelTitle");
  els.trackingPanelSubtitle = document.getElementById("trackingPanelSubtitle");
  els.closeTrackingPanelButton = document.getElementById("closeTrackingPanelButton");
  els.addRealAdvanceRowButton = document.getElementById("addRealAdvanceRowButton");
  els.addConsumedHourRowButton = document.getElementById("addConsumedHourRowButton");
  els.realAdvancesBody = document.getElementById("realAdvancesBody");
  els.consumedHoursBody = document.getElementById("consumedHoursBody");
  els.statusMessage = document.getElementById("statusMessage");
  els.switchProjectButton = document.getElementById("switchProjectButton");
  els.projectChooserOverlay = document.getElementById("projectChooserOverlay");
  els.projectChooserList = document.getElementById("projectChooserList");
  els.newProjectNameInput = document.getElementById("newProjectNameInput");
  els.addProjectButton = document.getElementById("addProjectButton");
  els.closeChooserButton = document.getElementById("closeChooserButton");
}

function applyBiControlHints() {
  const hints = {
    biVisualScopeSelect: "Define si el estilo se aplica a todo el proyecto o solo al widget seleccionado.",
    biShowLegendCheckbox: "Muestra u oculta la leyenda del grafico.",
    biShowGridCheckbox: "Activa o desactiva las lineas de rejilla.",
    biShowAxisLabelsCheckbox: "Muestra u oculta textos de ejes y ticks.",
    biShowDataLabelsCheckbox: "Muestra valores sobre cada punto o barra.",
    biLabelColorModeSelect: "Auto: ajusta contraste segun fondo. Manual: usa color fijo.",
    biLabelColorInput: "Color manual para labels cuando el modo es Manual.",
    biSeriesLineStyleSelect: "Estilo de trazo para lineas: continua, punteada o puntos.",
    biSeriesMarkerStyleSelect: "Forma de marcadores de puntos en lineas/areas/combos.",
    biSeriesMarkerSizeInput: "Tamano de marcador en pixeles. 0 oculta marcadores.",
    biStackModeSelect: "Stacking para graficos cartesianos: none, normal o 100%.",
    biFontFamilyTitleSelect: "Fuente para textos de titulo del grafico.",
    biFontFamilyAxisSelect: "Fuente para ejes y ticks.",
    biFontFamilyLabelSelect: "Fuente para labels y leyendas.",
    biFontFamilyTooltipSelect: "Fuente para tooltip interactivo.",
    biFontSizeTitleInput: "Tamano de fuente para titulos.",
    biFontSizeAxisInput: "Tamano de fuente para ejes.",
    biFontSizeLabelsInput: "Tamano de fuente para labels y leyendas.",
    biFontSizeTooltipInput: "Tamano de fuente para tooltip.",
    biAxisScaleSelect: "Escala del eje numerico principal (X o Y segun el grafico): lineal o logaritmica.",
    biAxisMinInput: "Valor minimo manual del eje numerico principal (vacio = Automatico).",
    biAxisMaxInput: "Valor maximo manual del eje numerico principal (vacio = Automatico).",
    biShowTargetLineCheckbox: "Dibuja una linea objetivo horizontal en graficos cartesianos.",
    biTargetLineValueInput: "Valor numerico de la meta para la linea objetivo.",
    biTargetLineLabelInput: "Etiqueta corta mostrada junto a la linea objetivo.",
    biTargetLineColorInput: "Color de la linea objetivo.",
    biWidgetBreakdownSelect: "Dimension secundaria usada en tipos como Pivot, Sankey, Scatter o Bubble.",
    biWidgetOptionalMetricSelect: "Metrica auxiliar para tipos avanzados (ej: eje secundario o burbuja).",
    biWidgetDateDimensionSelect: "Campo de fecha para Time series, Timeline o Candlestick.",
    biWidgetSortModeSelect: "Ordena las categorias del widget por valor o por etiqueta.",
    biUpdateWidgetButton: "Aplica propiedades al widget seleccionado (sin crear uno nuevo).",
    biTooltipModeSelect: "Full muestra detalle completo; Compact solo etiqueta y valor.",
    biCanvasPresetSelect: "Tamanos sugeridos de pizarra para dashboard.",
    biApplyCanvasSizeButton: "Aplica dimensiones personalizadas de la pizarra.",
    biCanvasZoomSelect: "Escala visual de la pizarra (25%, 50%, 75%, 100%, 200%).",
    biCanvasZoomMenuButton: "Abre el menu de zoom (ajustar todo, ancho y porcentajes).",
    biZoomFitAllButton: "Ajusta toda la pizarra a la vista visible.",
    biZoomFitWidthButton: "Ajusta el ancho de la pizarra al area visible.",
    biZoomDefaultButton: "Vuelve al zoom predeterminado 100%.",
    biShowCanvasGridCheckbox: "Muestra u oculta la rejilla de la pizarra.",
    biSnapToGridCheckbox: "Ajusta posicion y tamano de widgets al paso de rejilla.",
    biGridSnapSizeInput: "Define el tamano de celda de la rejilla en pixeles.",
    biPerformanceModeSelect: "Quality prioriza fidelidad, Balanced equilibra, Turbo acelera dashboards grandes.",
    biOptimizeNowButton: "Aplica optimizacion automatica (top N y limites) segun perfil de rendimiento.",
    biResetPerformanceButton: "Restaura el perfil a calidad maxima.",
    biUiModeSelect: "Basico simplifica controles; Avanzado muestra configuracion completa.",
    biOpenCommandPaletteButton: "Abre buscador de comandos rapidos (Ctrl+K).",
    biCrossFilterScopeSelect: "Define si los filtros de grafico aplican a todos o solo misma fuente.",
    biAutoLayoutButton: "Ordena automaticamente widgets en una malla limpia.",
    biGenerateDashboardButton: "Crea un dashboard base con widgets recomendados.",
    biHideLeftPanelButton: "Oculta el panel izquierdo de propiedades.",
    biHideDataPanelButton: "Oculta el panel derecho de datos.",
    sidebarToggleButton: "Oculta o muestra el panel lateral izquierdo de toda la aplicacion."
  };
  Object.entries(hints).forEach(([id, text]) => {
    const node = document.getElementById(id);
    if (!(node instanceof HTMLElement)) {
      return;
    }
    node.setAttribute("title", text);
    node.dataset.biHint = text;
  });
}

function requireFieldsEditMode(message = "Activa Editar para modificar Listas.") {
  if (fieldsEditMode) {
    return true;
  }
  setStatus(message);
  return false;
}

function requireReviewFlowEditMode(message = "Activa Editar para modificar Hitos de Flujo de Revisión.") {
  if (reviewFlowEditMode) {
    return true;
  }
  setStatus(message);
  return false;
}

function wireEvents() {
  els.sidebarToggleButton?.addEventListener("click", () => {
    applySidebarCollapsedState(!sidebarCollapsed, true);
    setStatus(sidebarCollapsed ? "Panel izquierdo oculto." : "Panel izquierdo mostrado.");
  });

  els.syncButton.addEventListener("click", () => {
    flushPendingSave();
    saveState(true);
    renderDeliverablesPanel(getActiveProject());
    renderReviewFlowPanel(getActiveProject());
    renderReviewControlsPanel(getActiveProject());
    renderBiPanel(getActiveProject());
    renderQuickSightPanel(getActiveProject());
    setStatus("Cambios guardados localmente.");
  });

  els.globalSearchInput.addEventListener("input", () => {
    const text = els.globalSearchInput.value || "";
    if (pendingSearchTimer !== null) {
      clearTimeout(pendingSearchTimer);
    }

    pendingSearchTimer = window.setTimeout(() => {
      currentSearchQuery = normalizeLookup(text);
      pendingSearchTimer = null;
      renderFieldsPanel(getActiveProject());
      renderDeliverablesPanel(getActiveProject());
      renderPackageControlsPanel(getActiveProject());
      renderReviewFlowPanel(getActiveProject());
      renderReviewControlsPanel(getActiveProject());
      renderBiPanel(getActiveProject());
      renderQuickSightPanel(getActiveProject());
    }, SEARCH_DEBOUNCE_MS);
  });

  els.exportBackupButton.addEventListener("click", () => {
    try {
      flushPendingSave();
      const serialized = JSON.stringify(state, null, 2);
      const blob = new Blob([serialized], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const stamp = new Date().toISOString().replace(/[:.]/g, "-");
      const a = document.createElement("a");
      a.href = url;
      a.download = `midp_backup_${stamp}.json`;
      a.click();
      URL.revokeObjectURL(url);
      setStatus("Backup exportado.");
    } catch {
      setStatus("No se pudo exportar el backup.");
    }
  });

  els.importBackupButton.addEventListener("click", () => {
    els.importBackupInput.value = "";
    els.importBackupInput.click();
  });

  els.importBackupInput.addEventListener("change", async () => {
    const file = els.importBackupInput.files?.[0];
    if (!file) {
      return;
    }

    try {
      const content = await file.text();
      const parsed = JSON.parse(content);
      state = normalizeState(parsed);
      ensureActiveProject();
      renderAll();
      flushPendingSave();
      saveState();
      setStatus("Backup importado correctamente.");
    } catch {
      setStatus("Backup invalido. Verifica el archivo JSON.");
    }
  });

  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      switchTab(button.dataset.tab || "fields");
    });
  });

  const handleQuickSightConfigChange = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const nextSource = normalizeBiSource(els.qsSourceSelect?.value || project.quickSightConfig.source || "all");
    const nextGroupBy = normalizeBiGroupBy(els.qsGroupBySelect?.value || project.quickSightConfig.groupBy || "disciplina");
    const nextBreakdown = normalizeBiOptionalGroupBy(
      els.qsBreakdownSelect instanceof HTMLSelectElement
        ? els.qsBreakdownSelect.value
        : (project.quickSightConfig.breakdownDimension || "")
    );
    const nextMetric = normalizeBiMetric(els.qsMetricSelect?.value || project.quickSightConfig.metric || "count");
    const nextChartType = normalizeBiChartType(els.qsChartTypeSelect?.value || project.quickSightConfig.chartType || "bar");
    const nextOptionalMetric = normalizeBiOptionalMetric(
      els.qsOptionalMetricSelect instanceof HTMLSelectElement
        ? els.qsOptionalMetricSelect.value
        : (project.quickSightConfig.optionalMetric || "")
    );
    const nextTopN = sanitizeBiTopN(els.qsTopNInput?.value ?? project.quickSightConfig.topN ?? 12);
    const selectedVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
    if (selectedVisual) {
      selectedVisual.source = nextSource;
      selectedVisual.groupBy = nextGroupBy;
      selectedVisual.breakdownDimension = nextBreakdown;
      selectedVisual.metric = nextMetric;
      selectedVisual.chartType = nextChartType;
      selectedVisual.topN = nextTopN;
      syncQuickSightVisualDataRoles(selectedVisual, {
        groupBy: nextGroupBy,
        breakdownDimension: nextBreakdown,
        metric: nextMetric,
        chartType: nextChartType,
        optionalMetric: nextOptionalMetric,
        dateDimension: project.quickSightConfig.dateDimension || ""
      });
      selectedVisual.display = normalizeQuickSightDisplaySettings(selectedVisual.display);
      selectedVisual.visualSettings = applyBiChartTypeVisualDefaults(
        normalizeBiVisualSettings(selectedVisual.visualSettings),
        selectedVisual.chartType
      );
      selectedVisual.chartConfig = normalizeBiChartConfig(selectedVisual.chartConfig, selectedVisual.chartType);
      const safeTypeMap = normalizeBiChartTypeConfigMap(selectedVisual.chartTypeConfig);
      if (!safeTypeMap[selectedVisual.chartType]) {
        safeTypeMap[selectedVisual.chartType] = normalizeBiChartTypeSpecificConfig({}, selectedVisual.chartType);
      }
      selectedVisual.chartTypeConfig = safeTypeMap;
    } else {
      project.quickSightConfig.source = nextSource;
      project.quickSightConfig.groupBy = nextGroupBy;
      project.quickSightConfig.breakdownDimension = nextBreakdown;
      project.quickSightConfig.metric = nextMetric;
      project.quickSightConfig.optionalMetric = nextOptionalMetric;
      project.quickSightConfig.chartType = nextChartType;
      project.quickSightConfig.topN = nextTopN;
    }
    project.quickSightConfig.fieldsSearch = trimOrFallback(els.qsFieldsSearchInput?.value, "").slice(0, 80);
    saveState();
    renderQuickSightPanel(project);
  };

  const normalizeQuickSightRolePickerKey = (value) => {
    const token = trimOrFallback(value, "");
    return new Set(["dimension", "metric", "optionalMetric", "breakdown"]).has(token)
      ? token
      : "";
  };

  const getQuickSightRoleUiLabels = (chartType) => {
    const type = normalizeBiChartType(chartType || "bar");
    if (BI_HORIZONTAL_BAR_TYPES.has(type)) {
      return {
        dimension: "EJE Y",
        metric: "EJE X / VALOR",
        optionalMetric: "EJE X",
        breakdown: "GRUPO/COLOR"
      };
    }
    if (type === "scatter" || type === "bubble") {
      return {
        dimension: "ETIQUETA",
        metric: "EJE Y",
        optionalMetric: "EJE X",
        breakdown: "GRUPO/COLOR"
      };
    }
    if (BI_BAR_FAMILY_TYPES.has(type) || new Set(["line", "timeseries", "area", "combo", "waterfall", "pareto"]).has(type)) {
      return {
        dimension: "EJE X",
        metric: "EJE Y / VALOR",
        optionalMetric: "EJE X",
        breakdown: "GRUPO/COLOR"
      };
    }
    return {
      dimension: "CATEGORIA",
      metric: "VALOR",
      optionalMetric: "EJE X",
      breakdown: "GRUPO/COLOR"
    };
  };

  const getQuickSightRolePickerContext = (project) => {
    ensureQuickSightState(project);
    const selectedVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
    const chartType = normalizeBiChartType(
      els.qsChartTypeSelect?.value
      || selectedVisual?.chartType
      || project.quickSightConfig.chartType
      || "bar"
    );
    const capabilities = getQuickSightChartCapabilities(selectedVisual ? {
      ...selectedVisual,
      chartType
    } : { chartType }, project.quickSightConfig);
    return {
      selectedVisual,
      chartType,
      capabilities
    };
  };

  const closeQuickSightRolePicker = (options = {}) => {
    const rerender = !!options.rerender;
    const hadOpenPicker = !!quickSightRolePickerOpen;
    quickSightRolePickerOpen = "";
    quickSightRolePickerFocusRole = "";
    if (!rerender || !hadOpenPicker || activeTab !== "quicksight") {
      return;
    }
    renderQuickSightPanel(getActiveProject(), { softSelect: true, allowEmptySelection: true });
  };

  const openQuickSightRolePicker = (role) => {
    const nextRole = normalizeQuickSightRolePickerKey(role);
    if (!nextRole || activeTab !== "quicksight") {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    const { capabilities } = getQuickSightRolePickerContext(project);
    if (nextRole === "optionalMetric" && !capabilities.supportsOptionalMetric) {
      setStatus("Este tipo usa una sola medida principal. EJE X adicional se habilita en Scatter y Bubble.");
      return;
    }
    if (nextRole === "breakdown" && !capabilities.supportsBreakdown) {
      setStatus("Este tipo de visual no usa GRUPO/COLOR.");
      return;
    }
    quickSightRolePickerOpen = quickSightRolePickerOpen === nextRole ? "" : nextRole;
    quickSightRolePickerFocusRole = quickSightRolePickerOpen || "";
    renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
  };

  const applyQuickSightRolePickerSelection = (role, rawValue) => {
    const nextRole = normalizeQuickSightRolePickerKey(role);
    const project = getActiveProject();
    if (!project || !nextRole) {
      return;
    }
    const roleLabels = getQuickSightRoleUiLabels(
      els.qsChartTypeSelect?.value
      || project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId)?.chartType
      || project.quickSightConfig.chartType
      || "bar"
    );
    if (nextRole === "dimension" && els.qsGroupBySelect instanceof HTMLSelectElement) {
      const nextGroupBy = normalizeBiGroupBy(rawValue || els.qsGroupBySelect.value);
      if (!nextGroupBy) {
        closeQuickSightRolePicker({ rerender: true });
        return;
      }
      els.qsGroupBySelect.value = nextGroupBy;
      quickSightRolePickerOpen = "";
      quickSightRolePickerFocusRole = "";
      handleQuickSightConfigChange();
      setStatus(`${roleLabels.dimension} actualizado: ${getBiGroupLabel(nextGroupBy, project)}.`);
      return;
    }
    if (nextRole === "metric" && els.qsMetricSelect instanceof HTMLSelectElement) {
      const nextMetric = normalizeBiMetric(rawValue || els.qsMetricSelect.value);
      if (!nextMetric) {
        closeQuickSightRolePicker({ rerender: true });
        return;
      }
      els.qsMetricSelect.value = nextMetric;
      quickSightRolePickerOpen = "";
      quickSightRolePickerFocusRole = "";
      handleQuickSightConfigChange();
      setStatus(`${roleLabels.metric} actualizado: ${getBiMetricLabel(nextMetric)}.`);
      return;
    }
    if (nextRole === "optionalMetric" && els.qsOptionalMetricSelect instanceof HTMLSelectElement) {
      const nextMetric = normalizeBiOptionalMetric(rawValue);
      els.qsOptionalMetricSelect.value = nextMetric;
      quickSightRolePickerOpen = "";
      quickSightRolePickerFocusRole = "";
      handleQuickSightConfigChange();
      setStatus(nextMetric
        ? `${roleLabels.optionalMetric} actualizado: ${getBiMetricLabel(nextMetric)}.`
        : `${roleLabels.optionalMetric} restaurado a Automatico: UP Base total.`);
      return;
    }
    if (nextRole === "breakdown" && els.qsBreakdownSelect instanceof HTMLSelectElement) {
      const nextBreakdown = normalizeBiOptionalGroupBy(rawValue);
      els.qsBreakdownSelect.value = nextBreakdown;
      quickSightRolePickerOpen = "";
      quickSightRolePickerFocusRole = "";
      handleQuickSightConfigChange();
      setStatus(nextBreakdown
        ? `${roleLabels.breakdown} actualizado: ${getBiGroupLabel(nextBreakdown, project)}.`
        : `${roleLabels.breakdown} limpiado.`);
    }
  };

  [els.qsSourceSelect, els.qsGroupBySelect, els.qsBreakdownSelect, els.qsMetricSelect, els.qsOptionalMetricSelect, els.qsChartTypeSelect, els.qsTopNInput, els.qsFieldsSearchInput]
    .forEach((node) => {
      if (!node) {
        return;
      }
      node.addEventListener("input", handleQuickSightConfigChange);
      node.addEventListener("change", handleQuickSightConfigChange);
    });

  els.qsChartTypeGrid?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest("[data-qs-chart-type]");
    if (!(button instanceof HTMLElement)) {
      return;
    }
    const nextType = normalizeBiChartType(button.dataset.qsChartType || "");
    if (els.qsChartTypeSelect instanceof HTMLSelectElement) {
      els.qsChartTypeSelect.value = nextType;
    }
    handleQuickSightConfigChange();
  });

  els.qsRoleDimensionList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("[data-qs-role-picker-select]")) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    openQuickSightRolePicker("dimension");
  });

  els.qsRoleDimensionList?.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement) || target.dataset.qsRolePickerSelect !== "dimension") {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightRolePickerSelection("dimension", target.value);
  });

  els.qsRoleMetricList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("[data-qs-role-picker-select]")) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    openQuickSightRolePicker("metric");
  });

  els.qsRoleMetricList?.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement) || target.dataset.qsRolePickerSelect !== "metric") {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightRolePickerSelection("metric", target.value);
  });

  els.qsRoleOptionalMetricList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("[data-qs-role-picker-select]")) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    openQuickSightRolePicker("optionalMetric");
  });

  els.qsRoleOptionalMetricList?.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement) || target.dataset.qsRolePickerSelect !== "optionalMetric") {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightRolePickerSelection("optionalMetric", target.value);
  });

  els.qsRoleBreakdownList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("[data-qs-role-picker-select]")) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    openQuickSightRolePicker("breakdown");
  });

  els.qsRoleBreakdownList?.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement) || target.dataset.qsRolePickerSelect !== "breakdown") {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightRolePickerSelection("breakdown", target.value);
  });

  els.qsAddMetricGhostButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const selectedVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
    const chartType = normalizeBiChartType(
      els.qsChartTypeSelect?.value
      || selectedVisual?.chartType
      || project.quickSightConfig.chartType
      || "bar"
    );
    const capabilities = getQuickSightChartCapabilities(selectedVisual ? {
      ...selectedVisual,
      chartType
    } : { chartType }, project.quickSightConfig);
    if (!capabilities.supportsOptionalMetric || !(els.qsOptionalMetricSelect instanceof HTMLSelectElement)) {
      els.qsMetricSelect?.focus();
      setStatus("Este tipo usa una sola medida principal. EJE X adicional se habilita en Scatter y Bubble.");
      return;
    }
    openQuickSightRolePicker("optionalMetric");
  });

  els.qsAddBreakdownGhostButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const selectedVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
    const chartType = normalizeBiChartType(
      els.qsChartTypeSelect?.value
      || selectedVisual?.chartType
      || project.quickSightConfig.chartType
      || "bar"
    );
    const capabilities = getQuickSightChartCapabilities(selectedVisual ? {
      ...selectedVisual,
      chartType
    } : { chartType }, project.quickSightConfig);
    if (!capabilities.supportsBreakdown || !(els.qsBreakdownSelect instanceof HTMLSelectElement)) {
      setStatus("Este tipo de visual no usa GRUPO/COLOR.");
      return;
    }
    openQuickSightRolePicker("breakdown");
  });

  const parseQuickSightDragPayload = (dataTransfer) => {
    if (!dataTransfer) {
      return null;
    }
    const raw = dataTransfer.getData("application/x-qs-field")
      || dataTransfer.getData("text/plain")
      || "";
    if (!raw) {
      return null;
    }
    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return null;
      }
      return {
        groupBy: normalizeBiOptionalGroupBy(parsed.groupBy || ""),
        metric: normalizeBiOptionalMetric(parsed.metric || ""),
        name: trimOrFallback(parsed.name, "")
      };
    } catch {
      return null;
    }
  };

  const assignQuickSightRoleFromPayload = (role, payload) => {
    if (!payload || typeof payload !== "object") {
      return;
    }
    if (role === "dimension" && payload.groupBy && els.qsGroupBySelect instanceof HTMLSelectElement) {
      els.qsGroupBySelect.value = payload.groupBy;
      handleQuickSightConfigChange();
      setStatus(`Dimension asignada: ${getBiGroupLabel(payload.groupBy)}.`);
      return;
    }
    if (role === "metric" && payload.metric && els.qsMetricSelect instanceof HTMLSelectElement) {
      els.qsMetricSelect.value = payload.metric;
      handleQuickSightConfigChange();
      setStatus(`Metrica asignada: ${getBiMetricLabel(payload.metric)}.`);
      return;
    }
    if (role === "optionalMetric" && payload.metric && els.qsOptionalMetricSelect instanceof HTMLSelectElement) {
      els.qsOptionalMetricSelect.value = payload.metric;
      handleQuickSightConfigChange();
      setStatus(`EJE X asignado: ${getBiMetricLabel(payload.metric)}.`);
      return;
    }
    if (role === "breakdown" && payload.groupBy && els.qsBreakdownSelect instanceof HTMLSelectElement) {
      els.qsBreakdownSelect.value = payload.groupBy;
      handleQuickSightConfigChange();
      setStatus(`GRUPO/COLOR asignado: ${getBiGroupLabel(payload.groupBy)}.`);
      return;
    }
    setStatus("Ese campo no aplica para ese rol.");
  };

  const bindQuickSightRoleDropZone = (node, role) => {
    if (!(node instanceof HTMLElement)) {
      return;
    }
    node.addEventListener("dragenter", (event) => {
      const payload = parseQuickSightDragPayload(event.dataTransfer);
      if (!payload) {
        return;
      }
      if (
        ((role === "metric" || role === "optionalMetric") && payload.metric)
        || ((role === "dimension" || role === "breakdown") && payload.groupBy)
      ) {
        node.classList.add("is-drop-target");
      }
    });
    node.addEventListener("dragover", (event) => {
      const payload = parseQuickSightDragPayload(event.dataTransfer);
      if (!payload) {
        return;
      }
      const canDrop = (
        ((role === "metric" || role === "optionalMetric") && payload.metric)
        || ((role === "dimension" || role === "breakdown") && payload.groupBy)
      );
      if (!canDrop) {
        return;
      }
      event.preventDefault();
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "copy";
      }
      node.classList.add("is-drop-target");
    });
    node.addEventListener("dragleave", () => {
      node.classList.remove("is-drop-target");
    });
    node.addEventListener("drop", (event) => {
      event.preventDefault();
      node.classList.remove("is-drop-target");
      const payload = parseQuickSightDragPayload(event.dataTransfer);
      assignQuickSightRoleFromPayload(role, payload);
    });
  };

  bindQuickSightRoleDropZone(els.qsRoleDimensionList, "dimension");
  bindQuickSightRoleDropZone(els.qsRoleMetricList, "metric");
  bindQuickSightRoleDropZone(els.qsRoleOptionalMetricList, "optionalMetric");
  bindQuickSightRoleDropZone(els.qsRoleBreakdownList, "breakdown");

  const toggleQuickSightPanel = (panelKey) => {
    quickSightPanelState[panelKey] = !quickSightPanelState[panelKey];
    if (panelKey === "data") {
      applyQuickSightPanelVisibility(quickSightPanelState[panelKey] ? "Panel Datos mostrado." : "Panel Datos oculto.", panelKey);
    } else if (panelKey === "visuals") {
      applyQuickSightPanelVisibility(quickSightPanelState[panelKey] ? "Panel Elementos visuales mostrado." : "Panel Elementos visuales oculto.", panelKey);
    } else if (panelKey === "props") {
      applyQuickSightPanelVisibility(quickSightPanelState[panelKey] ? "Panel Propiedades mostrado." : "Panel Propiedades oculto.", panelKey);
    } else {
      applyQuickSightPanelVisibility("");
    }
  };

  const bindQuickSightToggleButton = (buttonNode, panelKey) => {
    if (!(buttonNode instanceof HTMLElement)) {
      return;
    }
    buttonNode.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      toggleQuickSightPanel(panelKey);
    });
  };
  bindQuickSightToggleButton(els.qsToggleDataButton, "data");
  bindQuickSightToggleButton(els.qsToggleVisualsButton, "visuals");
  bindQuickSightToggleButton(els.qsTogglePropsButton, "props");
  els.qsPresentationModeButton?.addEventListener("click", () => {
    const changed = toggleQuickSightPresentationMode();
    const project = getActiveProject();
    renderQuickSightPanel(project, { softSelect: false, allowEmptySelection: true });
    if (changed) {
      setStatus(quickSightPresentationMode ? "Modo visualizacion QuickSight activado." : "Modo visualizacion QuickSight desactivado.");
    }
  });

  els.qsSettingsButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    quickSightPanelState.props = true;
    quickSightPropsTab = "visual";
    if (!selectedQuickSightVisualId && project.quickSightVisuals.length > 0) {
      selectedQuickSightVisualId = project.quickSightVisuals[0].id;
    }
    applyQuickSightPanelVisibility("Panel Configuracion QuickSight abierto.", "props");
    renderQuickSightPanel(project, { softSelect: false, allowEmptySelection: true });
  });

  els.qsPublishButton?.addEventListener("click", () => {
    exportQuickSightBoardPng(getActiveProject());
  });

  els.qsUndoButton?.addEventListener("click", () => {
    stepQuickSightHistory(getActiveProject(), "undo");
  });

  els.qsRedoButton?.addEventListener("click", () => {
    stepQuickSightHistory(getActiveProject(), "redo");
  });

  const announceSingleSheetMode = () => {
    setStatus("QuickSight opera con una sola hoja por ahora; multi-hoja queda pendiente hasta introducir quickSightSheets[].");
  };
  els.qsSheetsMenuButton?.addEventListener("click", announceSingleSheetMode);
  els.qsSheetAddButton?.addEventListener("click", announceSingleSheetMode);

  els.qsCanvasZoomMenuButton?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (activeTab !== "quicksight") {
      return;
    }
    if (quickSightEditMenuOpen) {
      setQuickSightEditMenuOpen(false);
    }
    setQuickSightZoomMenuOpen(!quickSightZoomMenuOpen);
  });

  els.qsCanvasZoomMenu?.addEventListener("click", (event) => {
    const target = event.target instanceof HTMLElement
      ? event.target.closest("[data-qs-zoom-value], [data-qs-zoom-action]")
      : null;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const project = getActiveProject();
    if (!project) {
      setQuickSightZoomMenuOpen(false);
      return;
    }
    ensureQuickSightState(project);
    const zoomValue = target.dataset.qsZoomValue;
    const action = trimOrFallback(target.dataset.qsZoomAction, "");
    if (zoomValue) {
      setQuickSightCanvasZoom(zoomValue, "Zoom pizarra Quicksigth");
      setQuickSightZoomMenuOpen(false);
      return;
    }
    if (action === "fit-all" || action === "fit-width") {
      const fitMode = action === "fit-width" ? "width" : "all";
      const fitZoom = getQuickSightCanvasFitZoom(project, fitMode);
      const statusLabel = action === "fit-width" ? "Zoom ajustado al ancho" : "Zoom ajustado a vista";
      setQuickSightCanvasZoom(fitZoom, statusLabel);
      setQuickSightZoomMenuOpen(false);
      return;
    }
    if (action === "in" || action === "out") {
      const step = action === "in" ? 1 : -1;
      const nextZoom = getSteppedQuickSightCanvasZoom(project.quickSightConfig.canvasZoom, step);
      setQuickSightCanvasZoom(nextZoom, "Zoom pizarra Quicksigth");
      setQuickSightZoomMenuOpen(false);
    }
  });

  const applyQuickSightCanvasFromEditMenu = (mode = "apply") => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const defaults = createDefaultQuickSightConfig();
    if (mode === "reset") {
      project.quickSightConfig.canvasWidth = defaults.canvasWidth;
      project.quickSightConfig.canvasHeight = defaults.canvasHeight;
      project.quickSightConfig.canvasZoom = defaults.canvasZoom;
    } else {
      project.quickSightConfig.canvasWidth = sanitizeBiCanvasDimension(
        els.qsCanvasWidthInput?.value,
        project.quickSightConfig.canvasWidth,
        QUICKSIGHT_CANVAS_MIN_WIDTH,
        BI_CANVAS_SURFACE_MAX_EDIT_WIDTH
      );
      project.quickSightConfig.canvasHeight = sanitizeBiCanvasDimension(
        els.qsCanvasHeightInput?.value,
        project.quickSightConfig.canvasHeight,
        QUICKSIGHT_CANVAS_MIN_HEIGHT,
        BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
      );
      project.quickSightConfig.canvasZoom = sanitizeQuickSightCanvasZoom(
        els.qsCanvasZoomSelect?.value,
        project.quickSightConfig.canvasZoom
      );
    }
    project.quickSightVisuals = project.quickSightVisuals.map((visual, index) => ({
      ...visual,
      layout: clampQuickSightVisualLayoutToCanvas(
        normalizeQuickSightVisualLayout(visual.layout, index),
        project.quickSightConfig
      )
    }));
    syncQuickSightEditMenuInputs(project);
    updateQuickSightZoomMenuUi(project);
    saveState();
    renderQuickSightPanel(project);
    setStatus(
      `Hoja Quicksigth ${mode === "reset" ? "restablecida" : "actualizada"}: ` +
      `${project.quickSightConfig.canvasWidth} x ${project.quickSightConfig.canvasHeight}px | Zoom ${project.quickSightConfig.canvasZoom}%.`
    );
  };

  els.qsEditMenuButton?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (activeTab !== "quicksight") {
      return;
    }
    const project = getActiveProject();
    if (project) {
      syncQuickSightEditMenuInputs(project);
    }
    if (quickSightZoomMenuOpen) {
      setQuickSightZoomMenuOpen(false);
    }
    setQuickSightEditMenuOpen(!quickSightEditMenuOpen);
  });

  els.qsApplyCanvasSizeButton?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightCanvasFromEditMenu("apply");
  });

  els.qsResetCanvasSizeButton?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    applyQuickSightCanvasFromEditMenu("reset");
  });

  els.qsEditMenu?.addEventListener("click", (event) => {
    event.stopPropagation();
  });

  els.qsCanvasWidthInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    applyQuickSightCanvasFromEditMenu("apply");
  });

  els.qsCanvasHeightInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    applyQuickSightCanvasFromEditMenu("apply");
  });

  els.qsCanvasZoomSelect?.addEventListener("change", () => {
    applyQuickSightCanvasFromEditMenu("apply");
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      let handled = false;
      if (activeTab === "quicksight" && quickSightRolePickerOpen) {
        closeQuickSightRolePicker({ rerender: true });
        setStatus("Selector de rol QuickSight cerrado.");
        handled = true;
      }
      if (activeTab === "quicksight" && quickSightExpandedVisualId) {
        quickSightExpandedVisualId = "";
        renderQuickSightPanel(getActiveProject(), { softSelect: true, allowEmptySelection: true });
        setStatus("Visual QuickSight restaurado al tamano de hoja.");
        handled = true;
      } else if (activeTab === "quicksight" && quickSightPresentationMode) {
        toggleQuickSightPresentationMode(false);
        renderQuickSightPanel(getActiveProject(), { softSelect: true, allowEmptySelection: true });
        setStatus("Modo visualizacion QuickSight desactivado.");
        handled = true;
      }
      if (quickSightEditMenuOpen) {
        setQuickSightEditMenuOpen(false);
        handled = true;
      }
      if (quickSightZoomMenuOpen) {
        setQuickSightZoomMenuOpen(false);
        handled = true;
      }
      if (activeTab === "quicksight" && selectedQuickSightVisualId) {
        selectedQuickSightVisualId = "";
        const project = getActiveProject();
        renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
        setStatus("Visual deseleccionado.");
        handled = true;
      }
      if (handled) {
        event.preventDefault();
        event.stopPropagation();
      }
    }
  });

  document.addEventListener("click", (event) => {
    if (!quickSightEditMenuOpen) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      setQuickSightEditMenuOpen(false);
      return;
    }
    if (target.closest("#qsEditMenuWrap")) {
      return;
    }
    setQuickSightEditMenuOpen(false);
  });

  document.addEventListener("click", (event) => {
    if (!quickSightZoomMenuOpen) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      setQuickSightZoomMenuOpen(false);
      return;
    }
    if (target.closest("#qsZoomMenuWrap")) {
      return;
    }
    setQuickSightZoomMenuOpen(false);
  });

  document.addEventListener("click", (event) => {
    if (!quickSightRolePickerOpen || activeTab !== "quicksight") {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      closeQuickSightRolePicker({ rerender: true });
      return;
    }
    if (target.closest(".qs-role-block")) {
      return;
    }
    closeQuickSightRolePicker({ rerender: true });
  });

  document.addEventListener("click", (event) => {
    if (activeTab !== "quicksight") {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const toggleButton = target.closest("#qsToggleDataButton, #qsToggleVisualsButton, #qsTogglePropsButton");
    if (!(toggleButton instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (toggleButton.id === "qsToggleDataButton") {
      toggleQuickSightPanel("data");
    } else if (toggleButton.id === "qsToggleVisualsButton") {
      toggleQuickSightPanel("visuals");
    } else if (toggleButton.id === "qsTogglePropsButton") {
      toggleQuickSightPanel("props");
    }
  });

  els.qsFieldsList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const item = target.closest(".qs-field-item");
    if (!(item instanceof HTMLElement)) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const groupByValue = trimOrFallback(item.dataset.qsGroupby || "", "");
    const metricValue = trimOrFallback(item.dataset.qsMetric || "", "");
    const legacyKind = item.dataset.qsKind;
    const legacyValue = trimOrFallback(item.dataset.qsValue || "", "");
    const applyAsBreakdown = !!(event.shiftKey || event.altKey);
    if (groupByValue) {
      const nextGroupBy = normalizeBiGroupBy(groupByValue);
      if (applyAsBreakdown) {
        if (els.qsBreakdownSelect) {
          els.qsBreakdownSelect.value = nextGroupBy;
        }
        setStatus(`GRUPO/COLOR seleccionado: ${getBiGroupLabel(nextGroupBy, project)}.`);
      } else {
        if (els.qsGroupBySelect) {
          els.qsGroupBySelect.value = nextGroupBy;
        }
        setStatus(`Dimension seleccionada: ${getBiGroupLabel(nextGroupBy, project)}.`);
      }
    } else if (metricValue) {
      const nextMetric = normalizeBiMetric(metricValue);
      const assignAsOptionalMetric = (
        (event.shiftKey || event.altKey)
        && !els.qsRoleOptionalMetricBlock?.classList.contains("hidden")
        && els.qsOptionalMetricSelect instanceof HTMLSelectElement
      );
      if (assignAsOptionalMetric) {
        els.qsOptionalMetricSelect.value = nextMetric;
        setStatus(`Metrica eje X seleccionada: ${getBiMetricLabel(nextMetric)}.`);
      } else if (els.qsMetricSelect) {
        els.qsMetricSelect.value = nextMetric;
        setStatus(`Metrica seleccionada: ${getBiMetricLabel(nextMetric)}.`);
      }
    } else if (legacyKind === "dimension" && legacyValue) {
      const nextGroupBy = normalizeBiGroupBy(legacyValue);
      if (els.qsGroupBySelect) {
        els.qsGroupBySelect.value = nextGroupBy;
      }
    } else if (legacyKind === "metric" && legacyValue) {
      const nextMetric = normalizeBiMetric(legacyValue);
      if (els.qsMetricSelect) {
        els.qsMetricSelect.value = nextMetric;
      }
    } else {
      setStatus("Campo no vinculable a dimension o metrica.");
      return;
    }
    handleQuickSightConfigChange();
  });

  els.qsFieldsList?.addEventListener("dragstart", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const item = target.closest(".qs-field-item");
    if (!(item instanceof HTMLElement) || !(event.dataTransfer instanceof DataTransfer)) {
      return;
    }
    const groupBy = normalizeBiOptionalGroupBy(item.dataset.qsGroupby || "");
    const metric = normalizeBiOptionalMetric(item.dataset.qsMetric || "");
    if (!groupBy && !metric) {
      event.preventDefault();
      return;
    }
    const payload = {
      groupBy,
      metric,
      name: trimOrFallback(item.dataset.qsFieldName, "")
    };
    const serialized = JSON.stringify(payload);
    event.dataTransfer.effectAllowed = "copy";
    event.dataTransfer.setData("application/x-qs-field", serialized);
    event.dataTransfer.setData("text/plain", serialized);
    item.classList.add("is-dragging");
  });

  els.qsFieldsList?.addEventListener("dragend", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const item = target.closest(".qs-field-item");
    if (!(item instanceof HTMLElement)) {
      return;
    }
    item.classList.remove("is-dragging");
  });

  els.qsAddVisualButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const selectedVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
    const seedConfig = createQuickSightConfigFromVisual(selectedVisual, project.quickSightConfig);
    const visual = createQuickSightVisualFromConfig(seedConfig, project.quickSightVisuals.length);
    project.quickSightVisuals.unshift(visual);
    selectedQuickSightVisualId = visual.id;
    saveState();
    renderQuickSightPanel(project);
    setStatus(`Visual agregado: ${visual.name}.`);
  });

  els.qsDeleteVisualButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    if (!selectedQuickSightVisualId) {
      setStatus("Selecciona un visual para eliminar.");
      return;
    }
    const before = project.quickSightVisuals.length;
    project.quickSightVisuals = project.quickSightVisuals.filter((item) => item.id !== selectedQuickSightVisualId);
    if (project.quickSightVisuals.length === before) {
      setStatus("No se encontr? el visual seleccionado.");
      return;
    }
    if (!project.quickSightVisuals.some((item) => item.id === quickSightExpandedVisualId)) {
      quickSightExpandedVisualId = "";
    }
    selectedQuickSightVisualId = project.quickSightVisuals[0]?.id || "";
    saveState();
    renderQuickSightPanel(project);
    setStatus("Visual eliminado.");
  });

  els.qsRefreshButton?.addEventListener("click", () => {
    renderQuickSightPanel(getActiveProject());
    setStatus("Dechini Quicksigth actualizado.");
  });

  els.qsCanvasBoard?.addEventListener("wheel", (event) => {
    if (activeTab !== "quicksight") {
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    event.preventDefault();
    ensureQuickSightState(project);
    quickSightPendingWheelZoomStep += event.deltaY < 0 ? 1 : -1;
    if (quickSightWheelZoomRafId) {
      return;
    }
    quickSightWheelZoomRafId = window.requestAnimationFrame(() => {
      quickSightWheelZoomRafId = 0;
      const step = quickSightPendingWheelZoomStep;
      quickSightPendingWheelZoomStep = 0;
      if (!step) {
        return;
      }
      project.quickSightConfig.canvasZoom = getSteppedQuickSightCanvasZoom(
        project.quickSightConfig.canvasZoom,
        step > 0 ? 1 : -1
      );
      syncQuickSightEditMenuInputs(project);
      updateQuickSightZoomMenuUi(project);
      saveState();
      renderQuickSightPanel(project);
      setStatus(`Zoom pizarra Quicksigth: ${project.quickSightConfig.canvasZoom}%.`);
    });
  }, { passive: false });

  els.qsCanvasBoard?.addEventListener("pointerdown", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("[data-qs-card-action]")) {
      return;
    }
    if (quickSightExpandedVisualId) {
      return;
    }
    const resizeHandle = target.closest("[data-qs-resize-handle]");
    if (resizeHandle instanceof HTMLElement) {
      const handle = trimOrFallback(resizeHandle.dataset.qsResizeHandle, "").toLowerCase();
      if (!handle) {
        return;
      }
      const visualCard = resizeHandle.closest(".qs-visual-card");
      if (!(visualCard instanceof HTMLElement)) {
        return;
      }
      const visualId = trimOrFallback(visualCard.dataset.qsVisualId, "");
      if (!visualId) {
        return;
      }
      const project = getActiveProject();
      if (!project) {
        return;
      }
      ensureQuickSightState(project);
      const visual = project.quickSightVisuals.find((item) => item.id === visualId);
      if (!visual) {
        return;
      }
      startQuickSightVisualDrag(project, visual, visualCard, event, "resize", handle);
      return;
    }
    const dragHandle = target.closest(".qs-visual-head");
    let visualCard = dragHandle instanceof HTMLElement
      ? dragHandle.closest(".qs-visual-card")
      : null;
    if (!(visualCard instanceof HTMLElement)) {
      const fallbackCard = target.closest(".qs-visual-card");
      if (fallbackCard instanceof HTMLElement && fallbackCard.classList.contains("no-head")) {
        visualCard = fallbackCard;
      }
    }
    if (!(visualCard instanceof HTMLElement)) {
      return;
    }
    const visualId = trimOrFallback(visualCard.dataset.qsVisualId, "");
    if (!visualId) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const visual = project.quickSightVisuals.find((item) => item.id === visualId);
    if (!visual) {
      return;
    }
    startQuickSightVisualDrag(project, visual, visualCard, event);
  });

  els.qsCanvasBoard?.addEventListener("click", (event) => {
    if (Date.now() < quickSightSuppressClickUntil) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const actionButton = target.closest("[data-qs-card-action]");
    if (actionButton instanceof HTMLElement) {
      const action = trimOrFallback(actionButton.dataset.qsCardAction, "");
      const visualId = trimOrFallback(actionButton.dataset.qsVisualId, "");
      const project = getActiveProject();
      if (!project || !visualId) {
        return;
      }
      if (action === "toggle-expand") {
        const changed = toggleQuickSightExpandedVisual(project, visualId);
        if (!changed) {
          return;
        }
        renderQuickSightPanel(project, { softSelect: false, allowEmptySelection: true });
        setStatus(
          quickSightExpandedVisualId
            ? "Visual QuickSight maximizado."
            : "Visual QuickSight restaurado."
        );
      }
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const visualCard = target.closest(".qs-visual-card");
    if (!(visualCard instanceof HTMLElement)) {
      if (selectedQuickSightVisualId) {
        selectedQuickSightVisualId = "";
        renderQuickSightPanel(getActiveProject(), { softSelect: true, allowEmptySelection: true });
        setStatus("Visual deseleccionado.");
      }
      return;
    }
    const visualId = trimOrFallback(visualCard.dataset.qsVisualId, "");
    if (!visualId) {
      return;
    }
    if (visualId === selectedQuickSightVisualId) {
      renderQuickSightPanel(getActiveProject(), { softSelect: true });
      return;
    }
    selectedQuickSightVisualId = visualId;
    renderQuickSightPanel(getActiveProject(), { softSelect: true });
  });

  const handleQuickSightPropertiesChange = (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const prop = trimOrFallback(target.dataset.qsProp, "");
    if (!prop) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureQuickSightState(project);
    const visual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId);
    if (!visual) {
      return;
    }
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) {
      return;
    }
    const rawValue = target instanceof HTMLInputElement && target.type === "checkbox"
      ? target.checked
      : target.value;

    if (prop === "name") {
      visual.name = trimOrFallback(String(rawValue), visual.name).slice(0, 120);
    } else if (prop === "layoutX" || prop === "layoutY" || prop === "layoutW" || prop === "layoutH") {
      const currentLayout = normalizeQuickSightVisualLayout(visual.layout, 0);
      let nextX = currentLayout.x;
      let nextY = currentLayout.y;
      let nextW = currentLayout.w;
      let nextH = currentLayout.h;
      const numeric = Number(rawValue);
      if (prop === "layoutX" && Number.isFinite(numeric)) {
        nextX = Math.max(0, Math.round(numeric));
      }
      if (prop === "layoutY" && Number.isFinite(numeric)) {
        nextY = Math.max(0, Math.round(numeric));
      }
      if (prop === "layoutW" && Number.isFinite(numeric)) {
        nextW = Math.max(QUICKSIGHT_VISUAL_MIN_WIDTH, Math.round(numeric));
      }
      if (prop === "layoutH" && Number.isFinite(numeric)) {
        nextH = Math.max(QUICKSIGHT_VISUAL_MIN_HEIGHT, Math.round(numeric));
      }
      visual.layout = normalizeQuickSightVisualLayout({
        x: nextX,
        y: nextY,
        w: nextW,
        h: nextH
      }, 0);
      visual.layout = clampQuickSightVisualLayoutToCanvas(visual.layout, project.quickSightConfig);
    } else if (prop === "source") {
      visual.source = normalizeBiSource(rawValue || visual.source);
      visual.drillState = createDefaultQuickSightDrillState();
    } else if (prop === "groupBy") {
      visual.groupBy = normalizeBiGroupBy(rawValue || visual.groupBy);
      syncQuickSightVisualDataRoles(visual, project.quickSightConfig);
      visual.drillState = createDefaultQuickSightDrillState();
    } else if (prop === "breakdownDimension") {
      visual.breakdownDimension = normalizeBiOptionalGroupBy(rawValue || "");
      syncQuickSightVisualDataRoles(visual, project.quickSightConfig);
    } else if (prop === "metric") {
      visual.metric = normalizeBiMetric(rawValue || visual.metric);
      syncQuickSightVisualDataRoles(visual, project.quickSightConfig);
    } else if (prop === "optionalMetric") {
      syncQuickSightVisualDataRoles(visual, {
        ...project.quickSightConfig,
        optionalMetric: normalizeBiOptionalMetric(rawValue || "")
      });
    } else if (prop === "chartType") {
      visual.chartType = normalizeBiChartType(rawValue || visual.chartType);
      syncQuickSightVisualDataRoles(visual, project.quickSightConfig);
    } else if (prop === "renderEngine") {
      visual.renderEngine = normalizeQuickSightRenderEngine(rawValue);
    } else if (prop === "topN") {
      visual.topN = sanitizeBiTopN(rawValue || visual.topN);
    } else if (prop === "sortMode") {
      visual.sortMode = normalizeBiSortMode(rawValue || visual.sortMode);
    } else if (prop.startsWith("display:")) {
      const key = prop.slice("display:".length);
      const currentDisplay = normalizeQuickSightDisplaySettings(visual.display);
      visual.display = normalizeQuickSightDisplaySettings({
        ...currentDisplay,
        [key]: rawValue
      });
      visual.subtitle = visual.display.subtitle;
      visual.altText = visual.display.altText;
    } else if (prop.startsWith("visual:")) {
      const key = prop.slice("visual:".length);
      const currentVisual = normalizeBiVisualSettings(visual.visualSettings);
      visual.visualSettings = normalizeBiVisualSettings({
        ...currentVisual,
        [key]: rawValue
      });
    } else if (prop.startsWith("interaction:")) {
      const key = prop.slice("interaction:".length);
      const currentInteractions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
      if (/^drillPath\d+$/.test(key)) {
        const drillIndex = Math.max(0, Number.parseInt(key.slice("drillPath".length), 10) || 0);
        const nextPath = currentInteractions.drillPath.slice(0, QUICKSIGHT_DRILL_PATH_SLOTS);
        nextPath[drillIndex] = rawValue;
        visual.interactions = normalizeQuickSightInteractionSettings({
          ...currentInteractions,
          drillPath: nextPath
        }, visual);
      } else {
        visual.interactions = normalizeQuickSightInteractionSettings({
          ...currentInteractions,
          [key]: rawValue
        }, visual);
      }
      visual.drillState = normalizeQuickSightDrillState(visual.drillState, visual);
    } else if (prop.startsWith("type:")) {
      const key = prop.slice("type:".length);
      const chartType = normalizeBiChartType(visual.chartType || "bar");
      const currentTypeConfig = getBiSpecificChartConfigForWidget(visual, chartType);
      setBiSpecificChartConfigForWidget(visual, chartType, {
        ...currentTypeConfig,
        [key]: rawValue
      });
    }

    ensureQuickSightVisualDrillState(visual);
    const activeDrillThrough = normalizeQuickSightDrillThroughSelection(project.quickSightConfig.drillThroughSelection);
    if (activeDrillThrough && activeDrillThrough.visualId === visual.id && new Set(["source", "groupBy", "breakdownDimension"]).has(prop)) {
      project.quickSightConfig.drillThroughSelection = null;
      quickSightDrillThroughQuery = "";
    }
    visual.chartConfig = normalizeBiChartConfig(visual.chartConfig, visual.chartType);
    const safeTypeMap = normalizeBiChartTypeConfigMap(visual.chartTypeConfig);
    const safeType = normalizeBiChartType(visual.chartType);
    if (!safeTypeMap[safeType]) {
      safeTypeMap[safeType] = normalizeBiChartTypeSpecificConfig({}, safeType);
    }
    visual.chartTypeConfig = safeTypeMap;
    saveState();
    renderQuickSightPanel(project);
  };

  els.qsPropertiesBody?.addEventListener("input", handleQuickSightPropertiesChange);
  els.qsPropertiesBody?.addEventListener("change", handleQuickSightPropertiesChange);
  els.qsPropertiesBody?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const stopSummaryNode = target.closest("[data-qs-stop-summary]");
    if (stopSummaryNode instanceof HTMLElement) {
      event.stopPropagation();
      return;
    }
    const actionButton = target.closest("[data-qs-action]");
    if (actionButton instanceof HTMLElement) {
      const action = trimOrFallback(actionButton.dataset.qsAction, "");
      const project = getActiveProject();
      if (!project) {
        return;
      }
      const visual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;
      if (action === "clear-interaction-filters") {
        const changed = clearQuickSightInteractionFilters(project);
        if (!changed) {
          setStatus("No hay filtros QuickSight activos.");
          return;
        }
        saveState();
        renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
        setStatus("Filtros QuickSight limpiados.");
        return;
      }
      if (action === "drill-up") {
        if (!visual) {
          return;
        }
        const result = rewindQuickSightVisualDrill(project, visual);
        if (!result.changed) {
          setStatus("El visual ya esta en el nivel raiz.");
          return;
        }
        saveState();
        renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
        setStatus(`Drill QuickSight: volvimos a ${getBiGroupLabel(result.currentGroupBy, project)}.`);
        return;
      }
      if (action === "drill-reset") {
        if (!visual) {
          return;
        }
        const changed = resetQuickSightVisualDrill(project, visual);
        if (!changed) {
          setStatus("No hay drill activo en este visual.");
          return;
        }
        saveState();
        renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
        setStatus("Drill QuickSight reiniciado.");
        return;
      }
    }
    const tabButton = target.closest("[data-qs-props-tab]");
    if (!(tabButton instanceof HTMLElement)) {
      return;
    }
    const nextTab = trimOrFallback(tabButton.dataset.qsPropsTab, "");
    if (nextTab !== "visual" && nextTab !== "interactions") {
      return;
    }
    if (quickSightPropsTab === nextTab) {
      return;
    }
    quickSightPropsTab = nextTab;
    const project = getActiveProject();
    renderQuickSightPanel(project || null);
  });

  const handleBiConfigChange = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.startDate = sanitizeDateInput(els.biStartDateInput?.value || "");
    project.biConfig.endDate = sanitizeDateInput(els.biEndDateInput?.value || "");
    project.biConfig.textFilter = trimOrFallback(els.biTextFilterInput?.value, "").slice(0, 120);
    project.biConfig.source = normalizeBiSource(els.biGlobalSourceSelect?.value || "all");
    project.biConfig.crossFilterScope = normalizeBiCrossFilterScope(els.biCrossFilterScopeSelect?.value || "all");
    project.biConfig.invalidOnly = !!els.biInvalidOnlyCheckbox?.checked;
    project.biConfig.uiMode = normalizeBiUiMode(els.biUiModeSelect?.value || project.biConfig.uiMode || "basic");
    saveState();
    renderBiPanel(project);
  };

  [els.biStartDateInput, els.biEndDateInput, els.biTextFilterInput, els.biGlobalSourceSelect, els.biCrossFilterScopeSelect, els.biInvalidOnlyCheckbox, els.biUiModeSelect]
    .forEach((node) => {
      if (!node) {
        return;
      }
      node.addEventListener("input", handleBiConfigChange);
      node.addEventListener("change", handleBiConfigChange);
    });

  const applyBiVisualFromControls = (announceMessage) => {
    const project = getActiveProject();
    if (!project) {
      return false;
    }
    ensureBiState(project);
    const target = resolveBiVisualTarget(project);
    if (target.mode === "widget" && !target.widget) {
      setStatus("Selecciona un widget para aplicar configuracion por widget.");
      return false;
    }
    const nextVisual = collectBiVisualSettingsFromInputs();
    if (target.mode === "widget" && target.widget) {
      target.widget.visualOverride = nextVisual;
    } else {
      project.biConfig.visual = nextVisual;
    }
    syncBiInputs(project.biConfig);
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiWidgets(project, filteredRows);
    saveState();
    if (announceMessage) {
      setStatus(announceMessage);
    }
    return true;
  };

  const handleBiVisualSettingsChange = () => {
    applyBiVisualFromControls("");
  };

  [
    els.biShowLegendCheckbox,
    els.biShowGridCheckbox,
    els.biShowAxisLabelsCheckbox,
    els.biShowDataLabelsCheckbox,
    els.biAxisXLabelInput,
    els.biAxisYLabelInput,
    els.biLabelMaxCharsInput,
    els.biValueDecimalsInput,
    els.biNumberFormatSelect,
    els.biValuePrefixInput,
    els.biValueSuffixInput,
    els.biFontSizeInput,
    els.biFontFamilyTitleSelect,
    els.biFontFamilyAxisSelect,
    els.biFontFamilyLabelSelect,
    els.biFontFamilyTooltipSelect,
    els.biFontSizeTitleInput,
    els.biFontSizeAxisInput,
    els.biFontSizeLabelsInput,
    els.biFontSizeTooltipInput,
    els.biLabelColorModeSelect,
    els.biLabelColorInput,
    els.biLineWidthInput,
    els.biSeriesLineStyleSelect,
    els.biSeriesMarkerStyleSelect,
    els.biSeriesMarkerSizeInput,
    els.biAreaOpacityInput,
    els.biLegendPositionSelect,
    els.biLegendMaxItemsInput,
    els.biAxisScaleSelect,
    els.biAxisMinInput,
    els.biAxisMaxInput,
    els.biShowTargetLineCheckbox,
    els.biTargetLineValueInput,
    els.biTargetLineLabelInput,
    els.biTargetLineColorInput,
    els.biBarWidthRatioInput,
    els.biGridOpacityInput,
    els.biGridDashInput,
    els.biStackModeSelect,
    els.biTooltipModeSelect,
    els.biSmoothLinesCheckbox
  ].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", handleBiVisualSettingsChange);
    node.addEventListener("change", handleBiVisualSettingsChange);
  });

  const applyBiWidgetChromeFromControls = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const widget = getBiSelectedWidget(project);
    if (!widget) {
      setStatus("Selecciona un widget para configurar fondo, borde y titulo.");
      syncBiInputs(project.biConfig);
      return;
    }
    widget.showTitle = !!els.biWidgetShowTitleCheckbox?.checked;
    widget.showBorder = !!els.biWidgetShowBorderCheckbox?.checked;
    widget.showBackground = !!els.biWidgetShowBackgroundCheckbox?.checked;
    saveState();
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiWidgets(project, filteredRows);
    syncBiInputs(project.biConfig);
  };

  [els.biWidgetShowTitleCheckbox, els.biWidgetShowBorderCheckbox, els.biWidgetShowBackgroundCheckbox].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", applyBiWidgetChromeFromControls);
    node.addEventListener("change", applyBiWidgetChromeFromControls);
  });

  [els.biElementConfigBody].forEach((container) => {
    if (!(container instanceof HTMLElement)) {
      return;
    }
    const handler = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      if (!target.dataset.biSpecificKey) {
        return;
      }
      applyBiSpecificConfigChange(target, {
        refreshInspector: event.type === "change"
      });
    };
    container.addEventListener("input", handler);
    container.addEventListener("change", handler);
  });

  els.biVisualScopeSelect?.addEventListener("change", () => {
    biVisualScopeMode = normalizeBiVisualScopeMode(els.biVisualScopeSelect?.value || biVisualScopeMode || "widget");
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    syncBiInputs(project.biConfig);
    setStatus(biVisualScopeMode === "widget"
      ? "Configuracion enfocada en widget seleccionado."
      : "Configuracion global del proyecto activa.");
  });

  els.biApplyVisualScopeButton?.addEventListener("click", () => {
    applyBiVisualFromControls("Configuracion visual aplicada.");
  });

  els.biResetVisualScopeButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const target = resolveBiVisualTarget(project);
    if (target.mode === "widget") {
      if (!target.widget) {
        setStatus("Selecciona un widget para limpiar su configuracion.");
        return;
      }
      target.widget.visualOverride = null;
      saveState();
      syncBiInputs(project.biConfig);
      const filteredRows = queryBiProjectRows(project, project.biConfig);
      renderBiWidgets(project, filteredRows);
      setStatus(`Configuracion del widget "${target.widget.name}" limpiada (usa la global).`);
      return;
    }
    project.biConfig.visual = createDefaultBiVisualSettings();
    saveState();
    syncBiInputs(project.biConfig);
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiWidgets(project, filteredRows);
    setStatus("Configuracion global restablecida.");
  });

  els.biApplyVisualPresetButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const presetId = trimOrFallback(els.biVisualPresetSelect?.value, "");
    const preset = resolveBiVisualPreset(project, presetId);
    if (!preset) {
      setStatus("Preset visual no disponible.");
      return;
    }
    project.biConfig.activeVisualPreset = presetId;
    const target = resolveBiVisualTarget(project);
    if (target.mode === "widget" && !target.widget) {
      setStatus("Selecciona un widget para aplicar el preset.");
      return;
    }
    if (target.mode === "widget" && target.widget) {
      target.widget.visualOverride = normalizeBiVisualSettings(preset);
      setStatus(`Preset aplicado al widget "${target.widget.name}".`);
    } else {
      project.biConfig.visual = normalizeBiVisualSettings(preset);
      setStatus("Preset aplicado a todo el proyecto.");
    }
    saveState();
    syncBiInputs(project.biConfig);
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiWidgets(project, filteredRows);
  });

  els.biSaveVisualPresetButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const rawName = trimOrFallback(els.biSavePresetNameInput?.value, "");
    const presetName = rawName.slice(0, 30);
    if (!presetName) {
      setStatus("Escribe un nombre para guardar el preset.");
      return;
    }
    saveCustomBiVisualPreset(project, presetName, collectBiVisualSettingsFromInputs());
    if (els.biSavePresetNameInput) {
      els.biSavePresetNameInput.value = "";
    }
    refreshBiVisualPresetOptions(project);
    saveState();
    setStatus(`Preset "${presetName}" guardado.`);
  });

  els.biResetVisualSettingsButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const target = resolveBiVisualTarget(project);
    const defaults = createDefaultBiVisualSettings();
    if (target.mode === "widget" && target.widget) {
      target.widget.visualOverride = defaults;
    } else {
      project.biConfig.visual = defaults;
    }
    syncBiInputs(project.biConfig);
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiWidgets(project, filteredRows);
    saveState();
    setStatus(target.mode === "widget" && target.widget
      ? `Estilo del widget "${target.widget.name}" restablecido.`
      : "Configuracion visual restablecida.");
  });

  const handleBiCatalogChange = () => {
    renderBiFieldCatalog(getActiveProject());
  };
  [els.biCatalogSourceSelect, els.biCatalogSearchInput].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", handleBiCatalogChange);
    node.addEventListener("change", handleBiCatalogChange);
  });

  if (els.biRailButtons && els.biRailButtons.length) {
    els.biRailButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        if (!(button instanceof HTMLElement)) {
          return;
        }
        button.blur();
        applyBiRailMode(button.dataset.biRailMode || "data");
      });
    });
  }

  els.biHideLeftPanelButton?.addEventListener("click", () => {
    applyBiRailMode(biRailMode, { forceClose: true });
  });

  els.biHideDataPanelButton?.addEventListener("click", () => {
    applyBiRailMode("data", { forceClose: true });
  });

  [els.biStudioLeftResizeHandle, els.biStudioMiddleResizeHandle].forEach((handle) => {
    if (!(handle instanceof HTMLElement)) {
      return;
    }
    handle.addEventListener("pointerdown", (event) => {
      const mode = trimOrFallback(handle.dataset.biPanelResizer, "middle");
      startBiStudioPanelResize(mode, event);
    });
  });

  els.biOpenCommandPaletteButton?.addEventListener("click", () => {
    toggleBiCommandPalette();
  });

  els.biCloseCommandPaletteButton?.addEventListener("click", () => {
    closeBiCommandPalette();
  });

  els.biCommandPaletteOverlay?.addEventListener("click", (event) => {
    if (event.target === els.biCommandPaletteOverlay) {
      closeBiCommandPalette();
    }
  });

  els.biCommandPaletteInput?.addEventListener("input", () => {
    biCommandPaletteSelection = 0;
    renderBiCommandPalette(getActiveProject());
  });

  els.biCommandPaletteList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const action = target.closest("[data-bi-command-action]");
    if (!(action instanceof HTMLElement)) {
      return;
    }
    const commandId = trimOrFallback(action.dataset.biCommandAction, "");
    if (!commandId) {
      return;
    }
    runBiCommandById(commandId);
  });

  els.biInsightsGrid?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const action = target.closest("[data-bi-insight-action]");
    if (!(action instanceof HTMLElement)) {
      return;
    }
    const insightAction = trimOrFallback(action.dataset.biInsightAction, "");
    const insightValue = trimOrFallback(action.dataset.biInsightValue, "");
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    if (insightAction === "filter_group" && insightValue) {
      const groupBy = normalizeBiGroupBy(action.dataset.biInsightGroupby || "disciplina");
      project.biConfig.crossFilters = [{
        groupBy,
        label: insightValue,
        source: "all"
      }];
      renderBiPanel(project);
      saveState();
      setStatus(`Insight aplicado: ${getBiGroupLabel(groupBy, project)} = ${insightValue}`);
      return;
    }
    if (insightAction === "invalid_dates") {
      project.biConfig.invalidOnly = true;
      if (els.biInvalidOnlyCheckbox) {
        els.biInvalidOnlyCheckbox.checked = true;
      }
      renderBiPanel(project);
      saveState();
      setStatus("Insight aplicado: solo fechas invertidas.");
    }
  });

  els.biCanvasPresetSelect?.addEventListener("change", () => {
    const raw = trimOrFallback(els.biCanvasPresetSelect?.value, "");
    if (!raw) {
      return;
    }
    const [rawWidth, rawHeight] = raw.split("x");
    const width = sanitizeBiCanvasDimension(rawWidth, BI_CANVAS_SURFACE_MIN_WIDTH, BI_CANVAS_SURFACE_MIN_EDIT_WIDTH, BI_CANVAS_SURFACE_MAX_EDIT_WIDTH);
    const height = sanitizeBiCanvasDimension(rawHeight, BI_CANVAS_SURFACE_HEIGHT, BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT, BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT);
    if (els.biCanvasWidthInput) {
      els.biCanvasWidthInput.value = String(width);
    }
    if (els.biCanvasHeightInput) {
      els.biCanvasHeightInput.value = String(height);
    }
  });

  const applyBiCanvasSize = (resetToDefault = false) => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const currentSize = getBiCanvasSizeFromConfig(project.biConfig);
    const widthValue = resetToDefault
      ? BI_CANVAS_SURFACE_MIN_WIDTH
      : (els.biCanvasWidthInput?.value || currentSize.width);
    const heightValue = resetToDefault
      ? BI_CANVAS_SURFACE_HEIGHT
      : (els.biCanvasHeightInput?.value || currentSize.height);
    project.biConfig.canvasWidth = sanitizeBiCanvasDimension(
      widthValue,
      BI_CANVAS_SURFACE_MIN_WIDTH,
      BI_CANVAS_SURFACE_MIN_EDIT_WIDTH,
      BI_CANVAS_SURFACE_MAX_EDIT_WIDTH
    );
    project.biConfig.canvasHeight = sanitizeBiCanvasDimension(
      heightValue,
      BI_CANVAS_SURFACE_HEIGHT,
      BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT,
      BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
    );
    if (resetToDefault && els.biCanvasPresetSelect) {
      els.biCanvasPresetSelect.value = "";
    }
    saveState();
    renderBiPanel(project);
    applyBiRailMode("board", { announce: false, forceOpen: true });
    setStatus(`Pizarra actualizada: ${project.biConfig.canvasWidth} x ${project.biConfig.canvasHeight}px.`);
  };

  els.biApplyCanvasSizeButton?.addEventListener("click", () => {
    applyBiCanvasSize(false);
  });

  els.biResetCanvasSizeButton?.addEventListener("click", () => {
    applyBiCanvasSize(true);
  });

  const applyBiCanvasZoom = (rawZoomValue, statusPrefix = "Zoom de pizarra") => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.canvasZoom = sanitizeBiCanvasZoom(
      rawZoomValue,
      project.biConfig.canvasZoom
    );
    if (els.biCanvasZoomSelect) {
      syncBiZoomSelectControl(els.biCanvasZoomSelect, project.biConfig.canvasZoom);
    }
    updateBiZoomMenuUi(project);
    saveState();
    renderBiPanel(project);
    setStatus(`${statusPrefix}: ${project.biConfig.canvasZoom}%.`);
  };

  [els.biCanvasZoomSelect].forEach((node) => {
    if (!(node instanceof HTMLSelectElement)) {
      return;
    }
    node.addEventListener("input", () => {
      applyBiCanvasZoom(node.value, "Zoom de pizarra");
    });
    node.addEventListener("change", () => {
      applyBiCanvasZoom(node.value, "Zoom de pizarra");
    });
  });

  const closeBiZoomMenu = () => {
    if (!(els.biCanvasZoomMenu instanceof HTMLElement)) {
      return;
    }
    els.biCanvasZoomMenu.classList.add("hidden");
    if (els.biCanvasZoomMenuButton instanceof HTMLButtonElement) {
      els.biCanvasZoomMenuButton.setAttribute("aria-expanded", "false");
    }
  };

  const openBiZoomMenu = () => {
    if (!(els.biCanvasZoomMenu instanceof HTMLElement)) {
      return;
    }
    els.biCanvasZoomMenu.classList.remove("hidden");
    if (els.biCanvasZoomMenuButton instanceof HTMLButtonElement) {
      els.biCanvasZoomMenuButton.setAttribute("aria-expanded", "true");
    }
  };

  const toggleBiZoomMenu = () => {
    if (!(els.biCanvasZoomMenu instanceof HTMLElement)) {
      return;
    }
    if (els.biCanvasZoomMenu.classList.contains("hidden")) {
      openBiZoomMenu();
    } else {
      closeBiZoomMenu();
    }
  };

  if (els.biCanvasZoomMenuButton instanceof HTMLButtonElement) {
    els.biCanvasZoomMenuButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      toggleBiZoomMenu();
    });
  }

  if (els.biCanvasZoomMenu instanceof HTMLElement) {
    els.biCanvasZoomMenu.addEventListener("click", (event) => {
      const target = event.target instanceof HTMLElement
        ? event.target.closest("[data-bi-zoom-value], [data-bi-zoom-action]")
        : null;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      const project = getActiveProject();
      if (!project) {
        closeBiZoomMenu();
        return;
      }
      ensureBiState(project);
      const zoomValue = target.dataset.biZoomValue;
      const action = trimOrFallback(target.dataset.biZoomAction, "");
      if (zoomValue) {
        applyBiCanvasZoom(zoomValue, "Zoom de pizarra");
        closeBiZoomMenu();
        return;
      }
      if (action === "fit-all" || action === "fit-width") {
        const fitMode = action === "fit-width" ? "width" : "all";
        const fitZoom = getBiCanvasFitZoom(project, fitMode);
        const statusLabel = action === "fit-width" ? "Zoom ajustado al ancho" : "Zoom ajustado a vista";
        applyBiCanvasZoom(fitZoom, statusLabel);
        closeBiZoomMenu();
        return;
      }
      if (action === "in" || action === "out") {
        const step = action === "in" ? 1 : -1;
        const nextZoom = getBiSteppedCanvasZoom(project.biConfig.canvasZoom, step);
        applyBiCanvasZoom(nextZoom, "Zoom de pizarra");
        closeBiZoomMenu();
        return;
      }
    });
  }

  document.addEventListener("click", (event) => {
    if (!(els.biCanvasZoomMenu instanceof HTMLElement) || els.biCanvasZoomMenu.classList.contains("hidden")) {
      return;
    }
    const targetNode = event.target instanceof Node ? event.target : null;
    if (!targetNode) {
      closeBiZoomMenu();
      return;
    }
    if (els.biZoomMenuWrap instanceof HTMLElement && els.biZoomMenuWrap.contains(targetNode)) {
      return;
    }
    closeBiZoomMenu();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") {
      return;
    }
    closeBiZoomMenu();
  });

  els.biWidgetsGrid?.addEventListener("wheel", (event) => {
    if (activeTab !== "bi") {
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    event.preventDefault();
    const step = event.deltaY < 0 ? 1 : -1;
    biPendingWheelZoomStep += step;
    if (biWheelZoomRafId) {
      return;
    }
    biWheelZoomRafId = window.requestAnimationFrame(() => {
      biWheelZoomRafId = 0;
      const accumulated = biPendingWheelZoomStep;
      biPendingWheelZoomStep = 0;
      if (!accumulated) {
        return;
      }
      const nextZoom = getBiSteppedCanvasZoom(project.biConfig.canvasZoom, accumulated > 0 ? 1 : -1);
      applyBiCanvasZoom(nextZoom, "Zoom de pizarra");
    });
  }, { passive: false });

  const handleBiBoardGridSettingsChange = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.showCanvasGrid = !!els.biShowCanvasGridCheckbox?.checked;
    project.biConfig.snapToGrid = !!els.biSnapToGridCheckbox?.checked;
    project.biConfig.gridSnapSize = sanitizeBiInteger(els.biGridSnapSizeInput?.value, 12, 4, 120);
    if (els.biGridSnapSizeInput) {
      els.biGridSnapSizeInput.value = String(project.biConfig.gridSnapSize);
    }
    saveState();
    renderBiPanel(project);
    setStatus("Configuracion de pizarra actualizada.");
  };

  [els.biShowCanvasGridCheckbox, els.biSnapToGridCheckbox, els.biGridSnapSizeInput].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", handleBiBoardGridSettingsChange);
    node.addEventListener("change", handleBiBoardGridSettingsChange);
  });

  const handleBiPerformanceModeChange = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.performanceMode = normalizeBiPerformanceMode(els.biPerformanceModeSelect?.value || project.biConfig.performanceMode || "balanced");
    saveState();
    renderBiPanel(project);
    const label = project.biConfig.performanceMode === "turbo"
      ? "Turbo"
      : (project.biConfig.performanceMode === "quality" ? "Calidad máxima" : "Balanceado");
    setStatus(`Perfil de rendimiento aplicado: ${label}.`);
  };

  [els.biPerformanceModeSelect].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", handleBiPerformanceModeChange);
    node.addEventListener("change", handleBiPerformanceModeChange);
  });

  els.biOptimizeNowButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const mode = normalizeBiPerformanceMode(els.biPerformanceModeSelect?.value || project.biConfig.performanceMode || "balanced");
    project.biConfig.performanceMode = mode;
    const result = runBiAutoOptimize(project, mode);
    saveState();
    renderBiPanel(project);
    setStatus(
      `Optimización aplicada (${mode}): ${result.updatedWidgets} widgets ajustados, `
      + `${result.topNChanges} topN, ${result.tableChanges} límites de tabla.`
    );
  });

  els.biResetPerformanceButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.performanceMode = "quality";
    saveState();
    renderBiPanel(project);
    setStatus("Perfil de rendimiento restablecido a Calidad máxima.");
  });

  const handleBiColorScopeChange = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.colorSource = normalizeBiSource(els.biColorSourceSelect?.value || project.biConfig.colorSource || "all");
    project.biConfig.colorGroupBy = normalizeBiGroupBy(els.biColorGroupBySelect?.value || project.biConfig.colorGroupBy || "disciplina");
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiColorPanel(project, filteredRows);
    saveState();
  };

  [els.biColorSourceSelect, els.biColorGroupBySelect].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", handleBiColorScopeChange);
    node.addEventListener("change", handleBiColorScopeChange);
  });

  els.biColorLegendList?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "color") {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const label = trimOrFallback(target.dataset.biColorLabel, "");
    const groupBy = normalizeBiGroupBy(target.dataset.biColorGroupby || project.biConfig.colorGroupBy || "disciplina");
    const key = buildBiColorKey(groupBy, label);
    if (!key) {
      return;
    }
    const colorMap = normalizeBiColorMap(project.biConfig.colorMap);
    colorMap[key] = normalizeBiColorHex(target.value, getBiPaletteColor(0));
    project.biConfig.colorMap = colorMap;
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiColorPanel(project, filteredRows);
    renderBiWidgets(project, filteredRows);
    saveState();
  });

  els.biAutoColorButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const source = normalizeBiSource(els.biColorSourceSelect?.value || project.biConfig.colorSource || "all");
    const groupBy = normalizeBiGroupBy(els.biColorGroupBySelect?.value || project.biConfig.colorGroupBy || "disciplina");
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    const legendRows = getBiColorLegendData(project, filteredRows, source, groupBy);
    if (legendRows.length === 0) {
      setStatus("No hay categorias para colorear.");
      return;
    }
    const colorMap = normalizeBiColorMap(project.biConfig.colorMap);
    legendRows.forEach((row, index) => {
      const key = buildBiColorKey(groupBy, row.label);
      if (!key) {
        return;
      }
      colorMap[key] = getBiPaletteColor(index);
    });
    project.biConfig.colorMap = colorMap;
    project.biConfig.colorSource = source;
    project.biConfig.colorGroupBy = groupBy;
    renderBiColorPanel(project, filteredRows);
    renderBiWidgets(project, filteredRows);
    saveState();
    setStatus(`Colorimetria automatica aplicada (${legendRows.length} categorias).`);
  });

  els.biClearColorButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const groupBy = normalizeBiGroupBy(els.biColorGroupBySelect?.value || project.biConfig.colorGroupBy || "disciplina");
    const removed = clearBiColorsForGroup(project, groupBy);
    const filteredRows = queryBiProjectRows(project, project.biConfig);
    renderBiColorPanel(project, filteredRows);
    renderBiWidgets(project, filteredRows);
    saveState();
    setStatus(removed > 0
      ? `Colorimetria limpiada para ${getBiGroupLabel(groupBy, project)} (${removed} codigos).`
      : `No habia colores personalizados para ${getBiGroupLabel(groupBy, project)}.`);
  });

  [
    els.biWidgetGroupBySelect,
    els.biWidgetMetricSelect,
    els.biWidgetBreakdownSelect,
    els.biWidgetOptionalMetricSelect,
    els.biWidgetDateDimensionSelect,
    els.biWidgetChartTypeSelect,
    els.biWidgetSortModeSelect
  ].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("input", () => {
      const project = getActiveProject();
      syncBiBuilderSelectionUi(project);
    });
    node.addEventListener("change", () => {
      const project = getActiveProject();
      syncBiBuilderSelectionUi(project);
      applyBiBuilderToSelectedWidget({ announce: false });
    });
  });

  els.biWidgetSourceSelect?.addEventListener("change", () => {
    const widgetSource = normalizeBiSource(els.biWidgetSourceSelect?.value || "all");
    if (els.biCatalogSourceSelect) {
      if (widgetSource === "all" || widgetSource === "deliverable" || widgetSource === "package" || widgetSource === "review-control") {
        els.biCatalogSourceSelect.value = widgetSource;
      }
    }
    renderBiFieldCatalog(getActiveProject());
    syncBiBuilderSelectionUi(getActiveProject());
    applyBiBuilderToSelectedWidget({ announce: false });
  });

  els.biChartTypeButtons?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest("[data-bi-chart-type]");
    if (!(button instanceof HTMLElement)) {
      return;
    }
    const chartType = normalizeBiChartType(button.dataset.biChartType || "bar");
    if (els.biWidgetChartTypeSelect) {
      els.biWidgetChartTypeSelect.value = chartType;
    }
    syncBiBuilderSelectionUi(getActiveProject());
    applyBiBuilderToSelectedWidget({ announce: false });
  });

  els.biCatalogFieldsList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const row = target.closest("[data-bi-groupby], [data-bi-metric]");
    if (!(row instanceof HTMLElement)) {
      return;
    }
    const groupBy = normalizeBiGroupBy(row.dataset.biGroupby || "");
    const metric = normalizeBiMetric(row.dataset.biMetric || "");
    let changed = false;
    if (groupBy && row.dataset.biGroupby && els.biWidgetGroupBySelect) {
      const exists = Array.from(els.biWidgetGroupBySelect.options).some((option) => option.value === groupBy);
      if (exists) {
        els.biWidgetGroupBySelect.value = groupBy;
        changed = true;
      }
    } else if (metric && row.dataset.biMetric && els.biWidgetMetricSelect) {
      els.biWidgetMetricSelect.value = metric;
      changed = true;
    }
    if (changed) {
      syncBiBuilderSelectionUi(getActiveProject());
      renderBiFieldCatalog(getActiveProject());
      applyBiBuilderToSelectedWidget({ announce: false });
    }
  });

  [els.biWidgetTopNInput, els.biWidgetNameInput].forEach((node) => {
    if (!node) {
      return;
    }
    node.addEventListener("change", () => {
      applyBiBuilderToSelectedWidget({ announce: false });
    });
  });

  els.biRefreshButton?.addEventListener("click", () => {
    renderBiPanel(getActiveProject());
    setStatus("Dashboard Dechini BI actualizado.");
  });

  els.biAutoLayoutButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const result = autoArrangeBiWidgets(project, { expandCanvas: true });
    saveState();
    renderBiPanel(project);
    setStatus(`Auto organizar aplicado (${result.arranged} widgets).`);
  });

  els.biGenerateDashboardButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    const result = generateBiStarterDashboard(project);
    saveState();
    renderBiPanel(project);
    if (result.added > 0) {
      setStatus(`Dashboard base generado (${result.added} widgets nuevos).`);
    } else {
      setStatus("No se agregaron widgets (ya existen o se alcanzó el límite).");
    }
  });

  els.biClearCrossFilterButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    project.biConfig.crossFilters = [];
    saveState();
    renderBiPanel(project);
    setStatus("Filtro cruzado BI limpiado.");
  });

  els.biExportRowsButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    exportBiRowsCsv(project);
  });

  els.biExportBoardPngButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    exportBiBoardPng(project);
  });

  els.biAddTextWidgetButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    if (project.biWidgets.length >= BI_MAX_WIDGETS) {
      setStatus(`Maximo de ${BI_MAX_WIDGETS} widgets alcanzado.`);
      return;
    }
    const name = trimOrFallback(els.biWidgetNameInput?.value, "Caja de texto").slice(0, 60);
    const widget = {
      id: uid(),
      kind: "text",
      name,
      source: "all",
      groupBy: "disciplina",
      metric: "count",
      chartType: "table",
      sortMode: "value_desc",
      topN: 10,
      textContent: "Nuevo texto",
      imageSrc: "",
      imageAlt: "",
      showTitle: false,
      showBorder: false,
      showBackground: false,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: "disciplina", metric: "count", chartType: "table" }),
      chartConfig: createDefaultBiChartConfig("table"),
      chartTypeConfig: {},
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      layout: getDefaultBiWidgetLayout(project.biWidgets.length, "text"),
      visualOverride: null
    };
    project.biWidgets.push(widget);
    selectedBiWidgetId = widget.id;
    if (els.biWidgetNameInput) {
      els.biWidgetNameInput.value = "";
    }
    saveState();
    renderBiPanel(project);
    setStatus("Cuadro de texto agregado.");
  });

  els.biAddImageWidgetButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);
    if (project.biWidgets.length >= BI_MAX_WIDGETS) {
      setStatus(`Maximo de ${BI_MAX_WIDGETS} widgets alcanzado.`);
      return;
    }
    const name = trimOrFallback(els.biWidgetNameInput?.value, "Imagen").slice(0, 60);
    const widget = {
      id: uid(),
      kind: "image",
      name,
      source: "all",
      groupBy: "disciplina",
      metric: "count",
      chartType: "table",
      sortMode: "value_desc",
      topN: 10,
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: "disciplina", metric: "count", chartType: "table" }),
      chartConfig: createDefaultBiChartConfig("table"),
      chartTypeConfig: {},
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      layout: getDefaultBiWidgetLayout(project.biWidgets.length, "image"),
      visualOverride: null
    };
    project.biWidgets.push(widget);
    selectedBiWidgetId = widget.id;
    if (els.biWidgetNameInput) {
      els.biWidgetNameInput.value = "";
    }
    saveState();
    renderBiPanel(project);
    if (els.biImageUploadInput) {
      biPendingImageWidgetId = widget.id;
      els.biImageUploadInput.value = "";
      els.biImageUploadInput.click();
    }
    setStatus("Widget de imagen agregado. Selecciona una imagen.");
  });

  els.biImageUploadInput?.addEventListener("change", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    const widgetId = trimOrFallback(biPendingImageWidgetId, selectedBiWidgetId);
    const widget = getBiWidgetById(project, widgetId);
    const file = els.biImageUploadInput?.files?.[0];
    biPendingImageWidgetId = "";
    if (!widget || !file) {
      return;
    }
    if (!file.type.startsWith("image/")) {
      setStatus("Selecciona un archivo de imagen valido.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === "string" ? reader.result : "";
      if (!result) {
        setStatus("No se pudo leer la imagen.");
        return;
      }
      widget.imageSrc = result;
      widget.imageAlt = widget.name || "Imagen";
      saveState();
      renderBiPanel(project);
      setStatus(`Imagen cargada en \"${widget.name}\".`);
    };
    reader.onerror = () => {
      setStatus("Error al cargar la imagen.");
    };
    reader.readAsDataURL(file);
  });

  const applyBiBuilderToSelectedWidget = (options = {}) => {
    const announce = options.announce === true;
    const project = getActiveProject();
    if (!project) {
      return { ok: false, reason: "no_project" };
    }
    ensureBiState(project);
    const selected = getBiSelectedWidget(project);
    if (!selected) {
      if (announce) {
        setStatus("Selecciona un widget para actualizar.");
      }
      return { ok: false, reason: "no_selected" };
    }
    if (normalizeBiWidgetKind(selected.kind || "chart") !== "chart") {
      if (announce) {
        setStatus("Solo se puede actualizar la configuracion de widgets de grafico.");
      }
      return { ok: false, reason: "not_chart" };
    }
    if (selected.locked) {
      if (announce) {
        setStatus("Widget bloqueado. Desbloquea para actualizar propiedades.");
      }
      return { ok: false, reason: "locked" };
    }

    const source = normalizeBiSource(els.biWidgetSourceSelect?.value || selected.source || "all");
    const groupBy = normalizeBiGroupBy(els.biWidgetGroupBySelect?.value || selected.groupBy || "disciplina");
    const metric = normalizeBiMetric(els.biWidgetMetricSelect?.value || selected.metric || "baseunits");
    const chartType = normalizeBiChartType(els.biWidgetChartTypeSelect?.value || selected.chartType || "bar");
    const dataRoles = buildBiDataRolesFromBuilderInput(project, selected);
    const finalGroupBy = normalizeBiGroupBy(dataRoles.dimensions?.[0] || groupBy);
    const finalMetric = normalizeBiMetric(dataRoles.metrics?.[0] || metric);
    const sortMode = normalizeBiSortMode(els.biWidgetSortModeSelect?.value || selected.sortMode || "value_desc");
    const topN = sanitizeBiTopN(els.biWidgetTopNInput?.value ?? selected.topN ?? 10);
    const typedName = trimOrFallback(els.biWidgetNameInput?.value, "");
    const fallbackName = `${getBiMetricLabel(finalMetric)} por ${getBiGroupLabel(finalGroupBy, project)}`;
    const previousSignature = JSON.stringify({
      name: selected.name,
      source: selected.source,
      groupBy: selected.groupBy,
      metric: selected.metric,
      chartType: selected.chartType,
      sortMode: selected.sortMode,
      topN: selected.topN,
      dataRoles: selected.dataRoles
    });

    selected.name = (typedName || selected.name || fallbackName).slice(0, 60);
    selected.source = source;
    selected.groupBy = finalGroupBy;
    selected.metric = finalMetric;
    selected.dataRoles = dataRoles;
    selected.chartType = chartType;
    selected.chartConfig = normalizeBiChartConfig(selected.chartConfig, chartType);
    setBiSpecificChartConfigForWidget(selected, chartType, getBiSpecificChartConfigForWidget(selected, chartType));
    selected.sortMode = sortMode;
    selected.topN = topN;

    const nextSignature = JSON.stringify({
      name: selected.name,
      source: selected.source,
      groupBy: selected.groupBy,
      metric: selected.metric,
      chartType: selected.chartType,
      sortMode: selected.sortMode,
      topN: selected.topN,
      dataRoles: selected.dataRoles
    });
    if (previousSignature === nextSignature) {
      return { ok: true, changed: false };
    }

    saveState();
    renderBiPanel(project);
    if (announce) {
      setStatus(`Widget "${selected.name}" actualizado.`);
    }
    return { ok: true, changed: true, name: selected.name };
  };

  els.biAddWidgetButton?.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    ensureBiState(project);
    if (project.biWidgets.length >= BI_MAX_WIDGETS) {
      setStatus(`Maximo de ${BI_MAX_WIDGETS} widgets alcanzado.`);
      return;
    }

    const topN = sanitizeBiTopN(els.biWidgetTopNInput?.value ?? 10);
    const source = normalizeBiSource(els.biWidgetSourceSelect?.value || "all");
    const groupBy = normalizeBiGroupBy(els.biWidgetGroupBySelect?.value || "disciplina");
    const metric = normalizeBiMetric(els.biWidgetMetricSelect?.value || "baseUnits");
    const chartType = normalizeBiChartType(els.biWidgetChartTypeSelect?.value || "bar");
    const dataRoles = buildBiDataRolesFromBuilderInput(project);
    const finalGroupBy = normalizeBiGroupBy(dataRoles.dimensions?.[0] || groupBy);
    const finalMetric = normalizeBiMetric(dataRoles.metrics?.[0] || metric);
    const sortMode = normalizeBiSortMode(els.biWidgetSortModeSelect?.value || "value_desc");
    const baseName = trimOrFallback(els.biWidgetNameInput?.value, "");
    const defaultName = `${getBiMetricLabel(finalMetric)} por ${getBiGroupLabel(finalGroupBy, project)}`;
    const name = (baseName || defaultName).slice(0, 60);

    const newWidget = {
      id: uid(),
      kind: "chart",
      name,
      source,
      groupBy: finalGroupBy,
      metric: finalMetric,
      chartType,
      sortMode,
      topN,
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles,
      chartConfig: createDefaultBiChartConfig(chartType),
      chartTypeConfig: {
        [chartType]: normalizeBiChartTypeSpecificConfig({}, chartType)
      },
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      layout: getDefaultBiWidgetLayout(project.biWidgets.length),
      visualOverride: null
    };
    project.biWidgets.push(newWidget);
    selectedBiWidgetId = newWidget.id;
    if (els.biWidgetNameInput) {
      els.biWidgetNameInput.value = "";
    }
    saveState();
    renderBiPanel(project);
    setStatus("Widget BI agregado.");
  });

  els.biUpdateWidgetButton?.addEventListener("click", () => {
    applyBiBuilderToSelectedWidget({ announce: true });
  });

  els.biWidgetsGrid?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const duplicateId = trimOrFallback(target.dataset.biDuplicateWidget, "");
    if (duplicateId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      ensureBiState(project);
      const result = duplicateBiWidget(project, duplicateId);
      if (!result.ok) {
        setStatus(result.message || "No se pudo clonar el widget.");
        return;
      }
      selectedBiWidgetId = result.newWidgetId;
      saveState();
      renderBiPanel(project);
      setStatus("Widget clonado.");
      return;
    }

    const lockId = trimOrFallback(target.dataset.biToggleLockWidget, "");
    if (lockId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      ensureBiState(project);
      const result = toggleBiWidgetLock(project, lockId);
      if (!result.ok) {
        setStatus("No se pudo cambiar el bloqueo del widget.");
        return;
      }
      saveState();
      renderBiPanel(project);
      setStatus(result.locked ? "Widget bloqueado." : "Widget desbloqueado.");
      return;
    }

    const removeId = trimOrFallback(target.dataset.biRemoveWidget, "");
    if (removeId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      ensureBiState(project);
      project.biWidgets = project.biWidgets.filter((item) => item.id !== removeId);
      if (selectedBiWidgetId === removeId) {
        selectedBiWidgetId = "";
      }
      saveState();
      renderBiPanel(project);
      setStatus("Widget BI eliminado.");
      return;
    }

    const exportId = trimOrFallback(target.dataset.biExportWidget, "");
    if (exportId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      exportBiWidgetCsv(project, exportId);
      return;
    }

    const exportPngId = trimOrFallback(target.dataset.biExportWidgetPng, "");
    if (exportPngId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      exportBiWidgetPng(project, exportPngId);
      return;
    }

    const selectImageId = trimOrFallback(target.dataset.biSelectImage, "");
    if (selectImageId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      const widget = getBiWidgetById(project, selectImageId);
      if (!widget) {
        return;
      }
      setSelectedBiWidget(project, selectImageId, true);
      if (els.biImageUploadInput) {
        biPendingImageWidgetId = selectImageId;
        els.biImageUploadInput.value = "";
        els.biImageUploadInput.click();
      }
      return;
    }

    const applyImageUrlId = trimOrFallback(target.dataset.biApplyImageUrl, "");
    if (applyImageUrlId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      const widget = getBiWidgetById(project, applyImageUrlId);
      if (!widget) {
        return;
      }
      const input = els.biWidgetsGrid?.querySelector(`[data-bi-image-url=\"${applyImageUrlId}\"]`);
      if (!(input instanceof HTMLInputElement)) {
        return;
      }
      const nextUrl = trimOrFallback(input.value, "").slice(0, 2000);
      widget.imageSrc = nextUrl;
      if (nextUrl && !widget.imageAlt) {
        widget.imageAlt = widget.name || "Imagen";
      }
      saveState();
      renderBiPanel(project);
      setStatus(nextUrl ? "URL de imagen aplicada." : "URL de imagen limpiada.");
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }
    const widgetNode = target.closest("[data-bi-widget-id]");
    if (widgetNode instanceof HTMLElement) {
      const widgetId = trimOrFallback(widgetNode.dataset.biWidgetId, "");
      setSelectedBiWidget(project, widgetId, true);
      return;
    }
    if (target.closest("[data-bi-canvas-surface]")) {
      setSelectedBiWidget(project, "", true);
    }
  });

  els.biWidgetsGrid?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const textWidgetId = trimOrFallback(target.dataset.biTextWidget, "");
    if (textWidgetId) {
      const project = getActiveProject();
      if (!project) {
        return;
      }
      const widget = getBiWidgetById(project, textWidgetId);
      if (!widget) {
        return;
      }
      const nextText = typeof target.textContent === "string"
        ? target.textContent.replace(/\u00a0/g, " ").slice(0, 4000)
        : "";
      widget.textContent = nextText;
      saveState();
      return;
    }
  });

  els.biWidgetsGrid?.addEventListener("pointerdown", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("button")) {
      return;
    }
    if (target.closest("[data-bi-text-widget]")) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureBiState(project);

    const resizeHandle = target.closest("[data-bi-resize-widget]");
    if (resizeHandle instanceof HTMLElement) {
      const widgetId = trimOrFallback(resizeHandle.dataset.biResizeWidget, "");
      const widgetElement = resizeHandle.closest("[data-bi-widget-id]");
      const widget = getBiWidgetById(project, widgetId);
      if (widget && widgetElement instanceof HTMLElement) {
        if (widget.locked) {
          setStatus("Widget bloqueado. Desbloquea para mover o redimensionar.");
          return;
        }
        setSelectedBiWidget(project, widgetId, false);
        startBiWidgetInteraction(project, widget, widgetElement, "resize", event);
      }
      return;
    }

    const dragHandle = target.closest("[data-bi-drag-widget]");
    if (dragHandle instanceof HTMLElement) {
      const widgetId = trimOrFallback(dragHandle.dataset.biDragWidget, "");
      const widgetElement = dragHandle.closest("[data-bi-widget-id]");
      const widget = getBiWidgetById(project, widgetId);
      if (widget && widgetElement instanceof HTMLElement) {
        if (widget.locked) {
          setStatus("Widget bloqueado. Desbloquea para mover o redimensionar.");
          return;
        }
        setSelectedBiWidget(project, widgetId, false);
        startBiWidgetInteraction(project, widget, widgetElement, "drag", event);
      }
    }
  });

  window.addEventListener("pointermove", (event) => {
    updateBiWidgetInteraction(event);
  });
  window.addEventListener("pointerup", (event) => {
    endBiWidgetInteraction(event);
  });
  window.addEventListener("pointercancel", (event) => {
    endBiWidgetInteraction(event);
  });

  els.syncPackagesButton.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const result = syncPackageControlsFromDeliverables(project);
    if (result.missingRequired) {
      setStatus("Faltan campos base: Proyecto, Disciplina, Sistema y Paquete.");
      return;
    }

    saveState();
    renderPackageControlsPanel(project);
    const addedLabel = result.added === 1 ? "1 combinacion agregada" : `${result.added} combinaciones agregadas`;
    const removedLabel = result.removed === 1 ? "1 combinacion eliminada" : `${result.removed} combinaciones eliminadas`;
    setStatus(`Sincronizar paquetes: ${addedLabel}, ${removedLabel}.`);
  });

  els.togglePackageEditButton.addEventListener("click", () => {
    packageEditMode = !packageEditMode;
    updatePackageEditUi();
    renderPackageControlsPanel(getActiveProject());
    setStatus(packageEditMode
      ? "Modo edicion activado en Control de Paquetes."
      : "Modo lectura activado en Control de Paquetes.");
  });

  els.toggleReviewControlsEditButton.addEventListener("click", () => {
    reviewControlsEditMode = !reviewControlsEditMode;
    if (reviewControlsEditMode) {
      closeTrackingPanel();
    }
    updateReviewControlsEditUi();
    renderReviewControlsPanel(getActiveProject());
    setStatus(reviewControlsEditMode
      ? "Modo edicion activado en Control de Flujos de Revision."
      : "Modo lectura activado en Control de Flujos de Revision.");
  });

  els.syncReviewControlsButton.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const result = syncReviewControls(project);
    if (result.missingRequired) {
      setStatus("Faltan campos base: Proyecto, Disciplina y Sistema.");
      return;
    }
    if (result.missingMilestones) {
      setStatus("Primero agrega hitos en la vista Hitos de Flujo de Revision.");
      return;
    }

    saveState();
    renderReviewControlsPanel(project);
    const addedLabel = result.added === 1 ? "1 fila agregada" : `${result.added} filas agregadas`;
    const removedLabel = result.removed === 1 ? "1 fila eliminada" : `${result.removed} filas eliminadas`;
    setStatus(`Sincronizar flujo de revision: ${addedLabel}, ${removedLabel}.`);
  });

  els.packageControlsBody.addEventListener("click", (event) => {
    if (packageEditMode) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.closest("input, select, textarea, button, a")) {
      return;
    }

    const row = target.closest("[data-package-control-row]");
    if (!(row instanceof HTMLElement)) {
      return;
    }

    const packageControlId = trimOrFallback(row.dataset.packageControlRow, "");
    if (!packageControlId) {
      return;
    }

    openPackageTrackingPanel(packageControlId);
  });

  els.packageControlsBody.addEventListener("change", (event) => {
    if (!packageEditMode) {
      setStatus("Activa Editar para modificar control por paquetes.");
      renderPackageControlsPanel(getActiveProject());
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const packageControlId = trimOrFallback(target.dataset.editPackageControl, "");
    const metaKey = trimOrFallback(target.dataset.editPackageMeta, "");
    if (!packageControlId || !metaKey) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    ensurePackageControls(project);
    const packageControl = project.packageControls.find((item) => item.id === packageControlId);
    if (!packageControl) {
      return;
    }

    if (metaKey === "startDate") {
      packageControl.startDate = sanitizeDateInput(target.value || "");
      if (isDateRangeInvalid(packageControl.startDate || "", packageControl.endDate || "")) {
        setStatus(`Paquete ${packageControl.key || ""} con fechas invertidas.`);
      } else {
        setStatus("Control de paquete actualizado.");
      }
    } else if (metaKey === "endDate") {
      packageControl.endDate = sanitizeDateInput(target.value || "");
      if (isDateRangeInvalid(packageControl.startDate || "", packageControl.endDate || "")) {
        setStatus(`Paquete ${packageControl.key || ""} con fechas invertidas.`);
      } else {
        setStatus("Control de paquete actualizado.");
      }
    } else if (metaKey === "realProgress") {
      const parsedValue = sanitizeRealProgress(target.value || "");
      const previousValue = sanitizeRealProgress(packageControl.realProgress);
      packageControl.realProgress = parsedValue === "" ? null : parsedValue;
      if (parsedValue !== "" && (previousValue === "" || previousValue !== parsedValue)) {
        registerPackageRealAdvanceLog(project.id, packageControl.id, previousValue, parsedValue);
      }
      target.value = formatNumberForInput(packageControl.realProgress);
      setStatus("Control de paquete actualizado.");
    } else {
      return;
    }

    saveState();
    renderPackageControlsPanel(project);
  });

  els.reviewControlsBody.addEventListener("click", (event) => {
    if (reviewControlsEditMode) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.closest("input, select, textarea, button, a")) {
      return;
    }

    const row = target.closest("[data-review-control-row]");
    if (!(row instanceof HTMLElement)) {
      return;
    }

    const reviewControlId = trimOrFallback(row.dataset.reviewControlRow, "");
    if (!reviewControlId) {
      return;
    }

    openReviewControlTrackingPanel(reviewControlId);
  });

  els.reviewControlsBody.addEventListener("change", (event) => {
    if (!reviewControlsEditMode) {
      setStatus("Activa Editar para modificar Control de Flujos de Revision.");
      renderReviewControlsPanel(getActiveProject());
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const controlId = trimOrFallback(target.dataset.editReviewControl, "");
    const metaKey = trimOrFallback(target.dataset.editReviewControlMeta, "");
    if (!controlId || !metaKey) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureReviewControls(project);

    const control = project.reviewControls.find((item) => item.id === controlId);
    if (!control) {
      return;
    }

    if (metaKey === "startDate") {
      control.startDate = sanitizeDateInput(target.value || "");
      if (isDateRangeInvalid(control.startDate || "", control.endDate || "")) {
        setStatus("Flujo de revision con fechas invertidas.");
      } else {
        setStatus("Control de flujo de revision actualizado.");
      }
    } else if (metaKey === "endDate") {
      control.endDate = sanitizeDateInput(target.value || "");
      if (isDateRangeInvalid(control.startDate || "", control.endDate || "")) {
        setStatus("Flujo de revision con fechas invertidas.");
      } else {
        setStatus("Control de flujo de revision actualizado.");
      }
    } else if (metaKey === "realProgress") {
      const parsedValue = sanitizeRealProgress(target.value || "");
      const previousValue = sanitizeRealProgress(control.realProgress);
      control.realProgress = parsedValue === "" ? null : parsedValue;
      if (parsedValue !== "" && (previousValue === "" || previousValue !== parsedValue)) {
        registerReviewControlRealAdvanceLog(project.id, control.id, previousValue, parsedValue);
      }
      target.value = formatNumberForInput(control.realProgress);
      setStatus("Control de flujo de revision actualizado.");
    } else {
      return;
    }

    saveState();
    renderReviewControlsPanel(project);
  });

  els.openAddReviewMilestoneButton.addEventListener("click", () => {
    if (!requireReviewFlowEditMode()) {
      return;
    }
    openReviewMilestoneDrawer();
  });

  els.toggleReviewFlowEditButton?.addEventListener("click", () => {
    reviewFlowEditMode = !reviewFlowEditMode;
    if (!reviewFlowEditMode) {
      closeReviewMilestoneDrawer();
    }
    updateReviewFlowEditUi();
    renderReviewFlowPanel(getActiveProject());
    setStatus(reviewFlowEditMode
      ? "Modo edicion activado en Hitos de Flujo de Revisión."
      : "Modo lectura activado en Hitos de Flujo de Revisión.");
  });

  els.addBlankReviewMilestoneButton.addEventListener("click", () => {
    if (!requireReviewFlowEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    insertReviewMilestone(project, "", null);
    saveState();
    renderReviewFlowPanel(project);
    renderReviewControlsPanel(project);
    setStatus("Hito en blanco agregado.");
  });

  els.saveReviewMilestoneButton.addEventListener("click", () => {
    if (!requireReviewFlowEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const name = trimOrFallback(els.reviewMilestoneNameInput?.value, "");
    const weightRaw = trimOrFallback(els.reviewMilestoneWeightInput?.value, "");
    const weight = sanitizeBaseUnits(els.reviewMilestoneWeightInput?.value);

    if (!name) {
      setStatus("Ingresa el nombre del hito.");
      return;
    }
    if (!weightRaw || weight === "" || weight <= 0) {
      setStatus("Unidad productiva debe ser un numero mayor a 0.");
      return;
    }

    insertReviewMilestone(project, name, weight);

    if (els.reviewMilestoneNameInput) {
      els.reviewMilestoneNameInput.value = "";
    }
    if (els.reviewMilestoneWeightInput) {
      els.reviewMilestoneWeightInput.value = "";
    }

    closeReviewMilestoneDrawer();
    saveState();
    renderReviewFlowPanel(project);
    renderReviewControlsPanel(project);
    setStatus("Hito agregado.");
  });

  els.closeReviewMilestoneDrawerButton.addEventListener("click", () => {
    closeReviewMilestoneDrawer();
  });

  els.reviewMilestoneDrawerOverlay.addEventListener("click", (event) => {
    if (event.target === els.reviewMilestoneDrawerOverlay) {
      closeReviewMilestoneDrawer();
    }
  });

  els.reviewMilestonesBody.addEventListener("click", (event) => {
    if (!reviewFlowEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const milestoneId = trimOrFallback(target.dataset.removeReviewMilestone, "");
    if (!milestoneId) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureReviewMilestones(project);

    project.reviewMilestones = project.reviewMilestones.filter((item) => item.id !== milestoneId);
    saveState();
    renderReviewFlowPanel(project);
    renderReviewControlsPanel(project);
    setStatus("Hito eliminado.");
  });

  els.reviewMilestonesBody.addEventListener("change", (event) => {
    if (!reviewFlowEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const milestoneId = trimOrFallback(target.dataset.editReviewMilestone, "");
    const key = trimOrFallback(target.dataset.reviewKey, "");
    if (!milestoneId || !key) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }
    ensureReviewMilestones(project);

    const milestone = project.reviewMilestones.find((item) => item.id === milestoneId);
    if (!milestone) {
      return;
    }

    if (key === "name") {
      const nextName = trimOrFallback(target.value, "");
      milestone.name = nextName;
    } else if (key === "baseUnits") {
      const rawText = trimOrFallback(target.value, "");
      const nextUnits = sanitizeBaseUnits(target.value);
      if (rawText && (nextUnits === "" || nextUnits <= 0)) {
        target.value = formatNumberForInput(milestone.baseUnits);
        setStatus("Unidad productiva debe ser un numero mayor a 0.");
        return;
      }
      milestone.baseUnits = rawText ? nextUnits : null;
      target.value = formatNumberForInput(milestone.baseUnits);
    } else {
      return;
    }

    saveState();
    renderReviewFlowPanel(project);
    renderReviewControlsPanel(project);
    setStatus("Hito actualizado.");
  });

  els.switchProjectButton.addEventListener("click", () => {
    if (state.projects.length === 0) {
      setStatus("No hay proyectos configurados.");
      return;
    }
    openProjectChooser(false);
  });

  els.projectTitleInput.addEventListener("input", () => {
    if (!requireFieldsEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    project.title = trimOrFallback(els.projectTitleInput.value, `MIDP - ${project.name}`);
    scheduleSaveState();
    renderHeader(project);
  });

  els.fieldSeparatorInput.addEventListener("input", () => {
    if (!requireFieldsEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    const clean = sanitizeSeparator(els.fieldSeparatorInput.value);
    project.codeSeparator = clean;
    els.fieldSeparatorInput.value = clean;
    recomputeAllDeliverableCodes(project, false);
    scheduleSaveState();
    renderDeliverablesPanel(project);
    renderNomenclaturePreview(project);
  });

  els.progressControlModeSelect.addEventListener("change", () => {
    if (!requireFieldsEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    project.progressControlMode = normalizeProgressControlMode(els.progressControlModeSelect.value);
    if (project.progressControlMode !== "package" && activeTab === "packages") {
      activeTab = "deliverables";
    }
    scheduleSaveState();
    renderDeliverablesPanel(project);
    renderPackageControlsPanel(project);
    renderReviewControlsPanel(project);
    renderTabState();
    setStatus(project.progressControlMode === "package"
      ? "Control de avance configurado por paquete."
      : "Control de avance configurado por entregable.");
  });

  els.addFieldButton.addEventListener("click", () => {
    if (!requireFieldsEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    if (project.fields.length >= MAX_FIELDS) {
      setStatus(`Maximo de ${MAX_FIELDS} campos alcanzado.`);
      return;
    }

    const newField = createField(`Campo ${project.fields.length + 1}`, false, false, []);
    project.fields.push(newField);
    ensureNomenclatureConfig(project);
    syncDraftValues(project);
    saveState();
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
    setStatus("Campo agregado.");
  });

  els.addFieldRowButton.addEventListener("click", () => {
    if (!requireFieldsEditMode()) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    project.fields.forEach((field) => {
      field.rows.push(createFieldRow("", ""));
    });
    scheduleSaveState();
    renderFieldsPanel(project);
    setStatus("Fila agregada.");
  });

  els.fieldMatrixHead.addEventListener("input", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const fieldId = target.dataset.fieldLabel;
    if (!fieldId) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const field = project.fields.find((item) => item.id === fieldId);
    if (!field) {
      return;
    }

    field.label = trimOrFallback(target.value, "Campo");
    ensureNomenclatureConfig(project);
    scheduleSaveState();
    renderNomenclaturePreview(project);
  });

  els.fieldMatrixHead.addEventListener("change", () => {
    if (!fieldsEditMode) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
  });

  els.fieldMatrixBody.addEventListener("click", (event) => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const position = getMatrixCellPositionFromElement(project, target);
    if (!position) {
      return;
    }

    if (event.shiftKey && matrixSelectionAnchor) {
      matrixSelectionRange = buildMatrixSelectionRange(matrixSelectionAnchor, position);
    } else {
      matrixSelectionAnchor = position;
      matrixSelectionRange = buildMatrixSelectionRange(position, position);
    }

    paintMatrixSelection(project);
  });

  els.fieldMatrixBody.addEventListener("keydown", (event) => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const position = getMatrixCellPositionFromElement(project, target);
    if (position && !matrixSelectionRange) {
      matrixSelectionAnchor = position;
      matrixSelectionRange = buildMatrixSelectionRange(position, position);
      paintMatrixSelection(project);
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      if (!matrixSelectionRange) {
        return;
      }
      const payload = buildMatrixSelectionClipboardText(project, matrixSelectionRange);
      if (!payload) {
        return;
      }
      event.preventDefault();
      copyTextToClipboard(payload);
      setStatus("Rango copiado.");
    }
  });

  els.fieldMatrixBody.addEventListener("paste", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const start = getMatrixCellPositionFromElement(project, target);
    if (!start) {
      return;
    }

    const rawText = event.clipboardData?.getData("text/plain") || "";
    const rows = parseClipboardGrid(rawText);
    if (rows.length === 0) {
      return;
    }

    event.preventDefault();
    applyClipboardGridToMatrix(project, start, rows);
    const adjusted = reconcileAllDeliverableSystemSelections(project);
    syncDraftValues(project);
    recomputeAllDeliverableCodes(project, false);
    saveState();
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
    clearMatrixSelection();
    if (adjusted > 0) {
      setStatus(`Pegado en matriz. ${adjusted} entregables ajustados por dependencias.`);
    } else {
      setStatus("Pegado en matriz.");
    }
  });

  els.fieldMatrixBody.addEventListener("input", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const fieldId = target.dataset.matrixField;
    const key = target.dataset.matrixKey;
    const rowIndex = parseInt(target.dataset.matrixRow || "-1", 10);
    if (!fieldId || (key !== "name" && key !== "code") || rowIndex < 0) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const field = project.fields.find((item) => item.id === fieldId);
    if (!field) {
      return;
    }

    const row = ensureFieldRow(field, rowIndex);
    row[key] = target.value;
    scheduleSaveState();
  });

  els.fieldMatrixBody.addEventListener("change", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (target instanceof HTMLSelectElement) {
      const fieldId = trimOrFallback(target.dataset.matrixSystemParentField, "");
      const rowIndex = parseInt(target.dataset.matrixSystemParentRow || "-1", 10);
      if (fieldId && rowIndex >= 0) {
        const field = project.fields.find((item) => item.id === fieldId);
        if (field) {
          const row = ensureFieldRow(field, rowIndex);
          row.parentDisciplineRowId = trimOrFallback(target.value, "");
          const adjusted = reconcileAllDeliverableSystemSelections(project);
          if (adjusted > 0) {
            setStatus(`Dependencia Sistema -> Disciplina actualizada. ${adjusted} entregables ajustados.`);
          } else {
            setStatus("Dependencia Sistema -> Disciplina actualizada.");
          }
        }
      }
    }

    syncDraftValues(project);
    recomputeAllDeliverableCodes(project, false);
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
  });

  els.fieldsTableBody.addEventListener("change", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const fieldId = target.dataset.fieldCode;
    if (!fieldId) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const field = project.fields.find((item) => item.id === fieldId);
    if (!field) {
      return;
    }

    field.includeInCode = !!target.checked;
    ensureNomenclatureConfig(project);
    if (field.includeInCode) {
      addToNomenclatureOrder(project, field.id);
    } else {
      removeFromNomenclatureOrder(project, field.id);
    }
    recomputeAllDeliverableCodes(project, false);
    saveState();
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
    renderNomenclaturePreview(project);
    setStatus("Configuracion de nomenclatura actualizada.");
  });

  els.fieldsTableBody.addEventListener("click", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const fieldIdRemove = target.dataset.removeField;

    if (!fieldIdRemove) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const field = project.fields.find((item) => item.id === fieldIdRemove);
    if (!field || field.locked) {
      return;
    }

    project.fields = project.fields.filter((item) => item.id !== fieldIdRemove);
    removeFromNomenclatureOrder(project, fieldIdRemove);
    ensureNomenclatureConfig(project);
    syncDraftValues(project);
    recomputeAllDeliverableCodes(project, false);
    saveState();
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
    renderNomenclaturePreview(project);
    setStatus("Campo eliminado.");
  });

  els.nomenclatureDragArea.addEventListener("dragstart", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const item = target.closest("[data-nomenclature-item]");
    if (!(item instanceof HTMLElement)) {
      return;
    }

    const fieldId = trimOrFallback(item.dataset.nomenclatureItem, "");
    if (!fieldId) {
      return;
    }

    draggedNomenclatureFieldId = fieldId;
    item.classList.add("dragging");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", fieldId);
    }
  });

  els.nomenclatureDragArea.addEventListener("dragover", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    if (!draggedNomenclatureFieldId) {
      return;
    }

    event.preventDefault();
    clearNomenclatureDropMarkers();

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const item = target.closest("[data-nomenclature-item]");
    if (!(item instanceof HTMLElement)) {
      return;
    }

    const rect = item.getBoundingClientRect();
    const placeAfter = event.clientX > rect.left + rect.width / 2;
    item.classList.add(placeAfter ? "drop-after" : "drop-before");
  });

  els.nomenclatureDragArea.addEventListener("drop", (event) => {
    if (!fieldsEditMode) {
      return;
    }
    if (!draggedNomenclatureFieldId) {
      return;
    }

    event.preventDefault();
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      clearNomenclatureDropMarkers();
      return;
    }

    const item = target.closest("[data-nomenclature-item]");
    clearNomenclatureDropMarkers();
    if (!(item instanceof HTMLElement)) {
      return;
    }

    const targetFieldId = trimOrFallback(item.dataset.nomenclatureItem, "");
    if (!targetFieldId || targetFieldId === draggedNomenclatureFieldId) {
      return;
    }

    const rect = item.getBoundingClientRect();
    const placeAfter = event.clientX > rect.left + rect.width / 2;
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const changed = moveNomenclatureFieldToTarget(project, draggedNomenclatureFieldId, targetFieldId, placeAfter);
    if (!changed) {
      return;
    }

    recomputeAllDeliverableCodes(project, false);
    saveState();
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
    renderNomenclaturePreview(project);
    setStatus("Orden de nomenclatura actualizado.");
  });

  els.nomenclatureDragArea.addEventListener("dragend", () => {
    const dragging = els.nomenclatureDragArea.querySelector(".nomenclature-item.dragging");
    if (dragging instanceof HTMLElement) {
      dragging.classList.remove("dragging");
    }
    clearNomenclatureDropMarkers();
    draggedNomenclatureFieldId = "";
  });

  els.toggleFieldsEditButton?.addEventListener("click", () => {
    fieldsEditMode = !fieldsEditMode;
    if (!fieldsEditMode) {
      clearMatrixSelection();
    }
    updateFieldsEditUi();
    renderFieldsPanel(getActiveProject());
    setStatus(fieldsEditMode
      ? "Modo edicion activado en Listas."
      : "Modo lectura activado en Listas.");
  });

  els.openAddDeliverableButton.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    primeDraftSelections(project);
    renderDeliverablesPanel(project);
    openDeliverableDrawer();
  });

  els.toggleMidpEditButton.addEventListener("click", () => {
    midpEditMode = !midpEditMode;
    if (!midpEditMode) {
      closeDeliverableDrawer();
    } else {
      closeTrackingPanel();
    }
    updateMidpEditUi();
    renderDeliverablesPanel(getActiveProject());
    setStatus(midpEditMode
      ? "Modo edicion activado en MIDP."
      : "Modo lectura activado en MIDP.");
  });

  if (els.toggleMidpAdvanceDetailsButton) {
    els.toggleMidpAdvanceDetailsButton.addEventListener("click", () => {
      const project = getActiveProject();
      if (!project) {
        return;
      }

      const showDeliverableProgressColumns = getProjectProgressControlMode(project) !== "package";
      if (!showDeliverableProgressColumns) {
        setStatus("En control por paquetes, MIDP no muestra columnas de avance.");
        return;
      }

      project.showMidpAdvanceDetails = !getProjectMidpAdvanceDetailsVisible(project);
      saveState();
      updateMidpEditUi();
      renderDeliverablesPanel(project);
      setStatus(project.showMidpAdvanceDetails === false
        ? "Columnas de incidencia y avances derivados ocultas."
        : "Columnas de incidencia y avances derivados visibles.");
    });
  }

  els.recomputeCodesButton.addEventListener("click", () => {
    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    recomputeAllDeliverableCodes(project, false);
    saveState();
    renderDeliverablesPanel(project);
    setStatus("Nomenclaturas recalculadas.");
  });

  els.addBlankDeliverableButton.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const rowRefs = createEmptyRowRefs(project);
    insertDeliverable(project, rowRefs, false, createEmptyProgressMeta());
    renderDeliverablesPanel(project);
    setStatus("Fila en blanco agregada.");
  });

  els.closeDeliverableDrawerButton.addEventListener("click", () => {
    closeDeliverableDrawer();
  });

  els.deliverableDrawerOverlay.addEventListener("click", (event) => {
    if (event.target === els.deliverableDrawerOverlay) {
      closeDeliverableDrawer();
    }
  });

  els.closeTrackingPanelButton.addEventListener("click", () => {
    closeTrackingPanel();
  });

  els.trackingPanelOverlay.addEventListener("click", (event) => {
    if (event.target === els.trackingPanelOverlay) {
      closeTrackingPanel();
    }
  });

  els.closeQsDrillThroughButton?.addEventListener("click", () => {
    const project = getActiveProject();
    closeQuickSightDrillThrough(project);
  });

  els.qsDrillThroughOverlay?.addEventListener("click", (event) => {
    if (event.target === els.qsDrillThroughOverlay) {
      const project = getActiveProject();
      closeQuickSightDrillThrough(project);
    }
  });

  els.qsDrillThroughSearchInput?.addEventListener("input", () => {
    quickSightDrillThroughQuery = trimOrFallback(els.qsDrillThroughSearchInput?.value || "", "").slice(0, 80);
    renderQuickSightDrillThroughDrawer(getActiveProject());
  });

  els.qsDrillThroughExportButton?.addEventListener("click", () => {
    exportQuickSightDrillThroughCsv(getActiveProject());
  });

  els.qsDrillThroughBody?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const actionButton = target.closest("[data-qs-drill-row-source][data-qs-drill-row-id]");
    if (!(actionButton instanceof HTMLElement)) {
      return;
    }
    const rowSource = trimOrFallback(actionButton.dataset.qsDrillRowSource, "");
    const rowId = trimOrFallback(actionButton.dataset.qsDrillRowId, "");
    if (!rowSource || !rowId) {
      return;
    }
    if (rowSource === "deliverable") {
      openTrackingPanel(rowId);
    } else if (rowSource === "package") {
      openPackageTrackingPanel(rowId);
    } else if (rowSource === "review-control") {
      openReviewControlTrackingPanel(rowId);
    }
  });

  els.addRealAdvanceRowButton.addEventListener("click", () => {
    const context = getTrackingContext();
    if (!context) {
      return;
    }

    context.entity.realAdvances.push({
      id: uid(),
      createdAt: new Date().toISOString(),
      date: "",
      value: null,
      addedPercent: null,
      author: getCurrentUserLabel()
    });
    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.addConsumedHourRowButton.addEventListener("click", () => {
    const context = getTrackingContext();
    if (!context) {
      return;
    }

    context.entity.consumedHours.push({
      id: uid(),
      date: "",
      hours: null,
      note: ""
    });
    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.realAdvancesBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const rowId = target.dataset.removeRealAdvance;
    if (!rowId) {
      return;
    }

    const context = getTrackingContext();
    if (!context) {
      return;
    }

    context.entity.realAdvances = context.entity.realAdvances.filter((item) => item.id !== rowId);
    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.consumedHoursBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const rowId = target.dataset.removeConsumedHour;
    if (!rowId) {
      return;
    }

    const context = getTrackingContext();
    if (!context) {
      return;
    }

    context.entity.consumedHours = context.entity.consumedHours.filter((item) => item.id !== rowId);
    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.realAdvancesBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const rowId = target.dataset.realAdvanceRow;
    const key = target.dataset.realAdvanceKey;
    if (!rowId || !key) {
      return;
    }

    const context = getTrackingContext();
    if (!context) {
      return;
    }

    const row = context.entity.realAdvances.find((item) => item.id === rowId);
    if (!row) {
      return;
    }

    if (key === "date") {
      row.date = sanitizeDateInput(target.value);
    } else if (key === "value") {
      const rawText = trimOrFallback(target.value, "");
      const parsed = sanitizeRealProgress(target.value);
      if (rawText && parsed === "") {
        target.value = formatNumberForInput(row.value);
        setStatus("Avance real debe ser un numero entre 0 y 100.");
        return;
      }
      const previous = sanitizeRealProgress(row.value);
      row.value = parsed === "" ? null : parsed;
      if (parsed === "") {
        row.addedPercent = null;
      } else if (previous === "") {
        row.addedPercent = parsed;
      } else {
        row.addedPercent = Math.round((parsed - previous) * 100) / 100;
      }
      if (!trimOrFallback(row.author || "", "")) {
        row.author = getCurrentUserLabel();
      }
      if (!trimOrFallback(row.createdAt || "", "")) {
        row.createdAt = new Date().toISOString();
      }
      target.value = formatNumberForInput(row.value);
    } else {
      return;
    }

    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.consumedHoursBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const rowId = target.dataset.consumedHourRow;
    const key = target.dataset.consumedHourKey;
    if (!rowId || !key) {
      return;
    }

    const context = getTrackingContext();
    if (!context) {
      return;
    }

    const row = context.entity.consumedHours.find((item) => item.id === rowId);
    if (!row) {
      return;
    }

    if (key === "date") {
      row.date = sanitizeDateInput(target.value);
    } else if (key === "hours") {
      const rawText = trimOrFallback(target.value, "");
      const parsed = sanitizeBaseUnits(target.value);
      if (rawText && parsed === "") {
        target.value = formatNumberForInput(row.hours);
        setStatus("Horas consumidas debe ser un numero mayor o igual a 0.");
        return;
      }
      row.hours = parsed === "" ? null : parsed;
      target.value = formatNumberForInput(row.hours);
    } else if (key === "note") {
      row.note = trimOrFallback(target.value, "").slice(0, TRACKING_MAX_NOTE);
      target.value = row.note;
    } else {
      return;
    }

    saveState();
    renderTrackingPanel(context.project, context.entity, context.entityType);
    setStatus(`${getTrackingStatusLabel(context)} actualizado.`);
  });

  els.deliverableInputs.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) {
      return;
    }

    const fieldId = target.dataset.deliverableInput;
    if (!fieldId) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const draft = getDraftValues(project.id);
    draft[fieldId] = target.value;
    const relation = getDisciplineSystemFields(project.fields);
    if (fieldId === relation.disciplineFieldId || fieldId === relation.systemFieldId) {
      const changed = reconcileSystemSelectionForRowRefs(project, draft);
      if (changed || fieldId === relation.disciplineFieldId) {
        renderDeliverablesPanel(project);
      }
    }
  });

  const updateDraftProgressMeta = () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const draft = getDraftValues(project.id);
    const startDate = sanitizeDateInput(els.drawerStartDateInput?.value || "");
    const endDate = sanitizeDateInput(els.drawerEndDateInput?.value || "");
    const baseUnits = sanitizeBaseUnits(els.drawerBaseUnitsInput?.value || "");
    const realProgress = sanitizeRealProgress(els.drawerRealProgressInput?.value || "");

    draft[DRAFT_META_START] = startDate;
    draft[DRAFT_META_END] = endDate;
    draft[DRAFT_META_BASE_UNITS] = formatNumberForInput(baseUnits);
    draft[DRAFT_META_REAL_PROGRESS] = formatNumberForInput(realProgress);
  };

  [els.drawerStartDateInput, els.drawerEndDateInput, els.drawerBaseUnitsInput, els.drawerRealProgressInput].forEach((input) => {
    if (!(input instanceof HTMLInputElement)) {
      return;
    }
    input.addEventListener("input", updateDraftProgressMeta);
    input.addEventListener("change", updateDraftProgressMeta);
  });

  els.saveDeliverableButton.addEventListener("click", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    const draft = getDraftValues(project.id);
    const rowRefs = {};
    project.fields.forEach((field) => {
      rowRefs[field.id] = trimOrFallback(draft[field.id], "");
    });
    reconcileSystemSelectionForRowRefs(project, rowRefs);

    const startDate = sanitizeDateInput(draft[DRAFT_META_START]);
    const endDate = sanitizeDateInput(draft[DRAFT_META_END]);
    const parsedBaseUnits = sanitizeBaseUnits(draft[DRAFT_META_BASE_UNITS]);
    const parsedRealProgress = sanitizeRealProgress(draft[DRAFT_META_REAL_PROGRESS]);
    insertDeliverable(project, rowRefs, true, {
      startDate,
      endDate,
      baseUnits: parsedBaseUnits === "" ? null : parsedBaseUnits,
      realProgress: parsedRealProgress === "" ? null : parsedRealProgress
    });
    renderDeliverablesPanel(project);
    closeDeliverableDrawer();
    if (isDateRangeInvalid(startDate, endDate)) {
      setStatus("Entregable guardado con fechas invertidas.");
    } else {
      setStatus("Entregable guardado.");
    }
  });

  els.deliverablesBody.addEventListener("click", (event) => {
    if (!midpEditMode) {
      clearDeliverableSelection();
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const position = getDeliverableCellPositionFromElement(project, target);
    if (!position) {
      return;
    }

    if (event.shiftKey && deliverableSelectionAnchor) {
      deliverableSelectionRange = buildMatrixSelectionRange(deliverableSelectionAnchor, position);
    } else {
      deliverableSelectionAnchor = position;
      deliverableSelectionRange = buildMatrixSelectionRange(position, position);
    }
    paintDeliverableSelection(project);
  });

  els.deliverablesBody.addEventListener("keydown", (event) => {
    if (!midpEditMode) {
      return;
    }
    if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "c") {
      return;
    }

    const project = getActiveProject();
    if (!project || !deliverableSelectionRange) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const position = getDeliverableCellPositionFromElement(project, target);
    if (!position) {
      return;
    }

    const payload = buildDeliverableSelectionClipboardText(project, deliverableSelectionRange);
    if (!payload) {
      return;
    }
    event.preventDefault();
    copyTextToClipboard(payload);
    setStatus("Rango MIDP copiado.");
  });

  els.deliverablesBody.addEventListener("paste", (event) => {
    if (!midpEditMode) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const start = getDeliverableCellPositionFromElement(project, target);
    if (!start) {
      return;
    }

    const rawText = event.clipboardData?.getData("text/plain") || "";
    const rows = parseClipboardGrid(rawText);
    if (rows.length === 0) {
      return;
    }

    event.preventDefault();
    const sourceCols = Math.max(1, ...rows.map((line) => Array.isArray(line) ? line.length : 0));
    const targetRange = (deliverableSelectionRange && isCellPositionInsideRange(start, deliverableSelectionRange))
      ? deliverableSelectionRange
      : buildMatrixSelectionRange(start, {
        rowIndex: start.rowIndex + Math.max(0, rows.length - 1),
        colIndex: start.colIndex + Math.max(0, sourceCols - 1)
      });
    const result = applyClipboardGridToDeliverables(project, targetRange, rows);
    if (!result.applied) {
      setStatus("No se pudo pegar en la seleccion MIDP.");
      return;
    }
    saveState();
    renderDeliverablesPanel(project);
    clearDeliverableSelection();
    if (result.errors > 0) {
      setStatus(`Pegado MIDP aplicado (${result.updated} celdas). ${result.errors} valores no validos omitidos.`);
    } else {
      setStatus(`Pegado MIDP aplicado (${result.updated} celdas).`);
    }
  });

  els.deliverablesBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const deliverableId = target.dataset.removeDeliverable;
    if (!deliverableId) {
      if (!midpEditMode) {
        const project = getActiveProject();
        if (getProjectProgressControlMode(project) === "package") {
          setStatus("Control por paquete activo: usa la vista Control de Avance por Paquetes para ver registros.");
          return;
        }
        const row = target.closest("[data-deliverable-row]");
        if (!(row instanceof HTMLElement)) {
          return;
        }
        const rowDeliverableId = row.dataset.deliverableRow;
        if (!rowDeliverableId) {
          return;
        }
        openTrackingPanel(rowDeliverableId);
      }
      return;
    }

    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    project.deliverables = project.deliverables.filter((item) => item.id !== deliverableId);
    saveState();
    renderDeliverablesPanel(project);
    setStatus("Entregable eliminado.");
  });

  els.deliverablesBody.addEventListener("change", (event) => {
    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      renderDeliverablesPanel(getActiveProject());
      return;
    }

    const target = event.target;
    if (!(target instanceof HTMLSelectElement) && !(target instanceof HTMLInputElement)) {
      return;
    }

    const deliverableId = target.dataset.editDeliverable;
    if (!deliverableId) {
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const deliverable = project.deliverables.find((item) => item.id === deliverableId);
    if (!deliverable) {
      return;
    }

    const metaKey = target.dataset.editMeta;
    if (metaKey === "startDate") {
      const nextValue = sanitizeDateInput(target.value);
      deliverable.startDate = nextValue;
      saveState();
      renderDeliverablesPanel(project);
      if (isDateRangeInvalid(deliverable.startDate || "", deliverable.endDate || "")) {
        setStatus(`Fila ${deliverable.ref} con fechas invertidas.`);
      } else {
        setStatus(`Fila ${deliverable.ref} actualizada.`);
      }
      return;
    }

    if (metaKey === "endDate") {
      const nextValue = sanitizeDateInput(target.value);
      deliverable.endDate = nextValue;
      saveState();
      renderDeliverablesPanel(project);
      if (isDateRangeInvalid(deliverable.startDate || "", deliverable.endDate || "")) {
        setStatus(`Fila ${deliverable.ref} con fechas invertidas.`);
      } else {
        setStatus(`Fila ${deliverable.ref} actualizada.`);
      }
      return;
    }

    if (metaKey === "baseUnits") {
      const rawText = trimOrFallback(target.value, "");
      const parsedValue = sanitizeBaseUnits(target.value);
      if (rawText && parsedValue === "") {
        target.value = formatNumberForInput(deliverable.baseUnits);
        setStatus("Unidades productivas base debe ser un numero mayor o igual a 0.");
        return;
      }

      deliverable.baseUnits = parsedValue === "" ? null : parsedValue;
      target.value = formatNumberForInput(deliverable.baseUnits);
      saveState();
      renderDeliverablesPanel(project);
      setStatus(`Fila ${deliverable.ref} actualizada.`);
      return;
    }

    if (metaKey === "realProgress") {
      const rawText = trimOrFallback(target.value, "");
      const parsedValue = sanitizeRealProgress(target.value);
      if (rawText && parsedValue === "") {
        target.value = formatNumberForInput(deliverable.realProgress);
        setStatus("Avance real debe ser un numero entre 0 y 100.");
        return;
      }

      const previousValue = sanitizeRealProgress(deliverable.realProgress);
      const nextValue = parsedValue === "" ? null : parsedValue;
      deliverable.realProgress = nextValue;
      if (parsedValue !== "" && (previousValue === "" || previousValue !== parsedValue)) {
        registerRealAdvanceLog(project.id, deliverable.id, previousValue, parsedValue);
      }
      target.value = formatNumberForInput(deliverable.realProgress);
      saveState();
      renderDeliverablesPanel(project);
      setStatus(`Fila ${deliverable.ref} actualizada.`);
      return;
    }

    if (!(target instanceof HTMLSelectElement)) {
      return;
    }

    const fieldId = target.dataset.editField;
    if (!fieldId) {
      return;
    }

    if (!deliverable.rowRefs) {
      deliverable.rowRefs = {};
    }
    deliverable.rowRefs[fieldId] = target.value;
    reconcileSystemSelectionForRowRefs(project, deliverable.rowRefs);

    applyDeliverableRows(project, deliverable, false);
    saveState();
    renderDeliverablesPanel(project);
    setStatus(`Fila ${deliverable.ref} actualizada.`);
  });

  els.projectChooserList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const selectButton = target.closest("[data-select-project]");
    if (!(selectButton instanceof HTMLElement)) {
      return;
    }

    const projectId = selectButton.dataset.selectProject;
    if (!projectId) {
      return;
    }

    const exists = state.projects.some((item) => item.id === projectId);
    if (!exists) {
      return;
    }

    state.activeProjectId = projectId;
    saveState();
    chooserLocked = false;
    closeProjectChooser(true);
    ensureActiveProject();
    switchTab("deliverables");
    renderAll();
    closeDeliverableDrawer();
    closeReviewMilestoneDrawer();
    closeTrackingPanel();
    setStatus("Proyecto seleccionado.");
  });

  els.addProjectButton.addEventListener("click", () => {
    const rawName = trimOrFallback(els.newProjectNameInput.value, "");
    if (!rawName) {
      setStatus("Ingresa un nombre de proyecto.");
      return;
    }

    const id = buildUniqueProjectId(rawName);
    const project = createDefaultProject(id, rawName);
    state.projects.push(project);
    state.activeProjectId = project.id;
    els.newProjectNameInput.value = "";
    saveState();
    chooserLocked = false;
    closeProjectChooser(true);
    switchTab("deliverables");
    renderAll();
    closeDeliverableDrawer();
    closeReviewMilestoneDrawer();
    closeTrackingPanel();
    setStatus("Proyecto agregado.");
  });

  els.newProjectNameInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    els.addProjectButton.click();
  });

  els.closeChooserButton.addEventListener("click", () => {
    if (chooserLocked) {
      return;
    }
    closeProjectChooser(false);
  });

  document.addEventListener("keydown", (event) => {
    if (activeTab === "bi" && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      toggleBiCommandPalette();
      return;
    }

    if (activeTab === "bi" && biCommandPaletteOpen) {
      if (handleBiCommandPaletteKeydown(event)) {
        return;
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      flushPendingSave();
      saveState(true);
      renderDeliverablesPanel(getActiveProject());
      renderReviewFlowPanel(getActiveProject());
      renderReviewControlsPanel(getActiveProject());
      renderBiPanel(getActiveProject());
      setStatus("Cambios guardados localmente.");
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "n") {
      if (activeTab !== "deliverables") {
        return;
      }
      if (!midpEditMode) {
        setStatus("Activa Editar para modificar campos en MIDP.");
        return;
      }
      event.preventDefault();
      const project = getActiveProject();
      if (!project) {
        return;
      }
      primeDraftSelections(project);
      renderDeliverablesPanel(project);
      openDeliverableDrawer();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === "r") {
      if (!midpEditMode) {
        setStatus("Activa Editar para modificar campos en MIDP.");
        return;
      }
      event.preventDefault();
      const project = getActiveProject();
      if (!project) {
        return;
      }
      recomputeAllDeliverableCodes(project, false);
      saveState();
      renderDeliverablesPanel(project);
      setStatus("Nomenclaturas recalculadas.");
      return;
    }

    if (activeTab === "quicksight") {
      const target = event.target;
      const isTypingTarget = target instanceof HTMLInputElement
        || target instanceof HTMLTextAreaElement
        || target instanceof HTMLSelectElement
        || (target instanceof HTMLElement && target.isContentEditable);
      const keyText = trimOrFallback(event.key, "").toLowerCase();
      const project = getActiveProject();
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !isTypingTarget && keyText === "z") {
        event.preventDefault();
        stepQuickSightHistory(project, event.shiftKey ? "redo" : "undo");
        return;
      }
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !isTypingTarget && keyText === "y") {
        event.preventDefault();
        stepQuickSightHistory(project, "redo");
        return;
      }
      if (!isTypingTarget && project && event.ctrlKey && event.altKey && !event.metaKey) {
        ensureQuickSightState(project);
        if (keyText === "[") {
          event.preventDefault();
          setQuickSightCanvasZoom(getQuickSightCanvasFitZoom(project, "all"), "Zoom ajustado a vista");
          return;
        }
        if (keyText === "]") {
          event.preventDefault();
          setQuickSightCanvasZoom(getQuickSightCanvasFitZoom(project, "width"), "Zoom ajustado al ancho");
          return;
        }
        if (keyText === "1") {
          event.preventDefault();
          setQuickSightCanvasZoom(100, "Zoom pizarra Quicksigth");
          return;
        }
        if (keyText === "+" || keyText === "=") {
          event.preventDefault();
          setQuickSightCanvasZoom(getSteppedQuickSightCanvasZoom(project.quickSightConfig.canvasZoom, 1), "Zoom pizarra Quicksigth");
          return;
        }
        if (keyText === "-" || keyText === "_") {
          event.preventDefault();
          setQuickSightCanvasZoom(getSteppedQuickSightCanvasZoom(project.quickSightConfig.canvasZoom, -1), "Zoom pizarra Quicksigth");
          return;
        }
      }
    }

    if (activeTab === "bi") {
      const target = event.target;
      const isTypingTarget = target instanceof HTMLInputElement
        || target instanceof HTMLTextAreaElement
        || target instanceof HTMLSelectElement
        || (target instanceof HTMLElement && target.isContentEditable);
      const keyText = trimOrFallback(event.key, "").toLowerCase();
      const project = getActiveProject();
      if (project) {
        ensureBiState(project);
      }
      const selectedWidget = project ? getBiSelectedWidget(project) : null;

      if (!isTypingTarget && project && event.ctrlKey && event.altKey && !event.metaKey) {
        if (keyText === "[") {
          event.preventDefault();
          applyBiCanvasZoom(getBiCanvasFitZoom(project, "all"), "Zoom ajustado a vista");
          return;
        }
        if (keyText === "]") {
          event.preventDefault();
          applyBiCanvasZoom(getBiCanvasFitZoom(project, "width"), "Zoom ajustado al ancho");
          return;
        }
        if (keyText === "1") {
          event.preventDefault();
          applyBiCanvasZoom(100, "Zoom de pizarra");
          return;
        }
        if (keyText === "+" || keyText === "=") {
          event.preventDefault();
          applyBiCanvasZoom(getBiSteppedCanvasZoom(project.biConfig.canvasZoom, 1), "Zoom de pizarra");
          return;
        }
        if (keyText === "-" || keyText === "_") {
          event.preventDefault();
          applyBiCanvasZoom(getBiSteppedCanvasZoom(project.biConfig.canvasZoom, -1), "Zoom de pizarra");
          return;
        }
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
        if (!project || !selectedWidget) {
          return;
        }
        event.preventDefault();
        const result = duplicateBiWidget(project, selectedWidget.id);
        if (!result.ok) {
          setStatus(result.message || "No se pudo clonar el widget.");
          return;
        }
        selectedBiWidgetId = result.newWidgetId;
        saveState();
        renderBiPanel(project);
        setStatus("Widget clonado (Ctrl + D).");
        return;
      }

      if (!isTypingTarget && (event.key === "Delete" || event.key === "Backspace")) {
        if (!project || !selectedWidget) {
          return;
        }
        event.preventDefault();
        project.biWidgets = project.biWidgets.filter((item) => item.id !== selectedWidget.id);
        selectedBiWidgetId = "";
        saveState();
        renderBiPanel(project);
        setStatus("Widget eliminado.");
        return;
      }

      if (!isTypingTarget && event.key.toLowerCase() === "l") {
        if (!project || !selectedWidget) {
          return;
        }
        event.preventDefault();
        const result = toggleBiWidgetLock(project, selectedWidget.id);
        saveState();
        renderBiPanel(project);
        setStatus(result.locked ? "Widget bloqueado (L)." : "Widget desbloqueado (L).");
        return;
      }

      if (!isTypingTarget && selectedWidget) {
        const snapEnabled = normalizeBiToggle(project?.biConfig?.snapToGrid, true);
        const snapSize = sanitizeBiInteger(project?.biConfig?.gridSnapSize, 12, 4, 120);
        const step = snapEnabled
          ? (event.shiftKey ? snapSize * 2 : snapSize)
          : (event.shiftKey ? 10 : 1);
        if (event.altKey) {
          let dw = 0;
          let dh = 0;
          if (event.key === "ArrowLeft") dw = -step;
          if (event.key === "ArrowRight") dw = step;
          if (event.key === "ArrowUp") dh = -step;
          if (event.key === "ArrowDown") dh = step;
          if (dw !== 0 || dh !== 0) {
            event.preventDefault();
            if (selectedWidget.locked) {
              setStatus("Widget bloqueado. Desbloquea para redimensionar.");
              return;
            }
            const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
            const minSize = getBiWidgetMinSize(selectedWidget);
            selectedWidget.layout = clampBiWidgetLayoutToCanvas({
              x: selectedWidget.layout?.x || 0,
              y: selectedWidget.layout?.y || 0,
              w: (selectedWidget.layout?.w || minSize.width) + dw,
              h: (selectedWidget.layout?.h || minSize.height) + dh
            }, canvasSize.width, canvasSize.height, minSize);
            if (snapEnabled) {
              selectedWidget.layout = clampBiWidgetLayoutToCanvas(
                snapBiWidgetLayout(selectedWidget.layout, "resize", snapSize, minSize),
                canvasSize.width,
                canvasSize.height,
                minSize
              );
            }
            const widgetNode = findBiWidgetElement(selectedWidget.id);
            if (widgetNode instanceof HTMLElement) {
              widgetNode.style.width = `${selectedWidget.layout.w}px`;
              widgetNode.style.height = `${selectedWidget.layout.h}px`;
            }
            saveState();
            return;
          }
        }
        let dx = 0;
        let dy = 0;
        if (event.key === "ArrowLeft") dx = -step;
        if (event.key === "ArrowRight") dx = step;
        if (event.key === "ArrowUp") dy = -step;
        if (event.key === "ArrowDown") dy = step;
        if (dx !== 0 || dy !== 0) {
          event.preventDefault();
          if (selectedWidget.locked) {
            setStatus("Widget bloqueado. Desbloquea para mover.");
            return;
          }
          const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
          const minSize = getBiWidgetMinSize(selectedWidget);
          selectedWidget.layout = clampBiWidgetLayoutToCanvas({
            x: (selectedWidget.layout?.x || 0) + dx,
            y: (selectedWidget.layout?.y || 0) + dy,
            w: selectedWidget.layout?.w || minSize.width,
            h: selectedWidget.layout?.h || minSize.height
          }, canvasSize.width, canvasSize.height, minSize);
          if (snapEnabled) {
            selectedWidget.layout = clampBiWidgetLayoutToCanvas(
              snapBiWidgetLayout(selectedWidget.layout, "drag", snapSize, minSize),
              canvasSize.width,
              canvasSize.height,
              minSize
            );
          }
          const widgetNode = findBiWidgetElement(selectedWidget.id);
          if (widgetNode instanceof HTMLElement) {
            widgetNode.style.left = `${selectedWidget.layout.x}px`;
            widgetNode.style.top = `${selectedWidget.layout.y}px`;
          }
          saveState();
          return;
        }
      }
    }

    if (event.key !== "Escape") {
      return;
    }

    if (biCommandPaletteOpen) {
      closeBiCommandPalette();
      return;
    }

    closeDeliverableDrawer();
    closeReviewMilestoneDrawer();
    closeTrackingPanel();
  });
}

function renderAll() {
  const project = getActiveProject();
  renderHeader(project);
  renderFieldsPanel(project);
  renderDeliverablesPanel(project);
  renderReviewFlowPanel(project);
  renderReviewControlsPanel(project);
  renderQuickSightPanel(project);
  renderProjectChooser(project);
  renderTabState();
  if (activeTab === "bi") {
    if (biPendingRenderRafId) {
      window.cancelAnimationFrame(biPendingRenderRafId);
      biPendingRenderRafId = 0;
    }
    biPendingRenderRafId = window.requestAnimationFrame(() => {
      biPendingRenderRafId = 0;
      renderBiPanel(getActiveProject());
    });
  }
}

function renderHeader(project) {
  if (!project) {
    els.projectTitle.textContent = "MIDP";
    els.projectSubtitle.textContent = "No hay proyecto activo.";
    return;
  }

  els.projectTitle.textContent = project.title;
  els.projectSubtitle.textContent = `Proyecto activo: ${project.name}`;
}

function renderFieldsPanel(project) {
  clearMatrixSelection();
  if (!project) {
    els.projectTitleInput.value = "";
    els.fieldSeparatorInput.value = "-";
    if (els.progressControlModeSelect) {
      els.progressControlModeSelect.value = "deliverable";
    }
    els.fieldCounter.textContent = `0 / ${MAX_FIELDS} campos`;
    els.fieldMatrixHead.innerHTML = "";
    els.fieldMatrixBody.innerHTML = "";
    els.fieldsTableBody.innerHTML = "";
    if (els.nomenclatureDragArea) {
      els.nomenclatureDragArea.innerHTML = "";
    }
    if (els.nomenclatureOrderHint) {
      els.nomenclatureOrderHint.textContent = "";
    }
    if (els.nomenclaturePreviewText) {
      els.nomenclaturePreviewText.textContent = "(sin configuracion)";
    }
    return;
  }

  ensureNomenclatureConfig(project);
  const editEnabled = fieldsEditMode;
  els.projectTitleInput.value = project.title;
  els.fieldSeparatorInput.value = project.codeSeparator;
  if (els.progressControlModeSelect) {
    els.progressControlModeSelect.value = getProjectProgressControlMode(project);
  }
  els.fieldCounter.textContent = `${project.fields.length} / ${MAX_FIELDS} campos`;

  const relation = getDisciplineSystemFields(project.fields);
  const disciplineRows = relation.disciplineField
    ? getFieldOptions(relation.disciplineField)
    : [];
  const rowsToRender = computeVisualRowCount(project.fields);
  const filteredFieldIndexes = getFilteredFieldIndexes(project.fields, currentSearchQuery);
  els.fieldMatrixHead.innerHTML = `<tr>
    <th class="row-index-head">#</th>
    ${project.fields.map((field) => {
      const dependencyHeader = field.id === relation.systemFieldId
        ? "<th class=\"matrix-code-head\">disciplina</th>"
        : "";
      return `
      <th class="matrix-field-head">
        <input
          type="text"
          class="matrix-head-input"
          data-field-label="${escapeAttribute(field.id)}"
          value="${escapeAttribute(field.label)}"
          maxlength="40"
          ${(field.locked || !editEnabled) ? "readonly" : ""}>
      </th>
      <th class="matrix-code-head">codigo</th>
      ${dependencyHeader}
    `;
    }).join("")}
  </tr>`;

  const bodyRows = [];
  for (let rowIndex = 0; rowIndex < rowsToRender; rowIndex += 1) {
    const isSearchMode = !!currentSearchQuery;
    const rowAllowed = !isSearchMode || filteredFieldIndexes.has(rowIndex);
    if (!rowAllowed) {
      continue;
    }

    const cells = project.fields.map((field) => {
      const row = field.rows[rowIndex];
      const nameValue = row ? row.name : "";
      const codeValue = row ? row.code : "";
      let dependencyCell = "";
      if (field.id === relation.systemFieldId) {
        const selectedDisciplineId = trimOrFallback(row?.parentDisciplineRowId, "");
        const optionsHtml = disciplineRows
          .map((disciplineRow) => {
            const selectedAttr = disciplineRow.id === selectedDisciplineId ? "selected" : "";
            const label = disciplineRow.code
              ? `${disciplineRow.name || "(sin nombre)"} (${disciplineRow.code})`
              : (disciplineRow.name || "(sin nombre)");
            return `<option value="${escapeAttribute(disciplineRow.id)}" ${selectedAttr}>${escapeHtml(label)}</option>`;
          })
          .join("");
        dependencyCell = `
        <td>
          <select
            class="matrix-cell-input"
            data-matrix-system-parent-field="${escapeAttribute(field.id)}"
            data-matrix-system-parent-row="${rowIndex}"
            ${editEnabled ? "" : "disabled"}>
            <option value="">Disciplina</option>
            ${optionsHtml}
          </select>
        </td>`;
      }
      return `
        <td>
          <input
            type="text"
            class="matrix-cell-input"
            data-matrix-field="${escapeAttribute(field.id)}"
            data-matrix-row="${rowIndex}"
            data-matrix-key="name"
            value="${escapeAttribute(nameValue)}"
            placeholder="Nombre"
            ${editEnabled ? "" : "readonly"}>
        </td>
        <td>
          <input
            type="text"
            class="matrix-cell-input matrix-code-input"
            data-matrix-field="${escapeAttribute(field.id)}"
            data-matrix-row="${rowIndex}"
            data-matrix-key="code"
            value="${escapeAttribute(codeValue)}"
            placeholder="Cod."
            ${editEnabled ? "" : "readonly"}>
        </td>
        ${dependencyCell}
      `;
    }).join("");

    bodyRows.push(`<tr><td class="row-index-cell">${rowIndex + 1}</td>${cells}</tr>`);
  }
  els.fieldMatrixBody.innerHTML = bodyRows.join("");

  const orderedFields = getOrderedNomenclatureFields(project);
  if (els.nomenclatureOrderHint) {
    if (orderedFields.length === 0) {
      els.nomenclatureOrderHint.textContent = "Sin apartados seleccionados";
    } else {
      els.nomenclatureOrderHint.textContent = `Orden: ${orderedFields.map((field) => field.label).join(" > ")}`;
    }
  }
  if (els.nomenclatureDragArea) {
    if (orderedFields.length === 0) {
      els.nomenclatureDragArea.innerHTML = "<span class=\"nomenclature-empty\">Activa \"Incluir en nomenclatura\" y luego arrastra para ordenar.</span>";
    } else {
      els.nomenclatureDragArea.innerHTML = orderedFields
        .map((field, index) => `<div class="nomenclature-item" draggable="${editEnabled ? "true" : "false"}" data-nomenclature-item="${escapeAttribute(field.id)}" title="Arrastra para ordenar">
          <span class="nomenclature-drag-handle">|||</span>
          <span class="order-chip">${index + 1}</span>
          <span>${escapeHtml(field.label)}</span>
        </div>`)
        .join("");
    }
  }

  els.fieldsTableBody.innerHTML = project.fields
    .map((field, index) => {
      const removeButton = field.locked
        ? "<span class=\"muted\">Base</span>"
        : `<button type="button" class="danger" data-remove-field="${escapeAttribute(field.id)}" ${editEnabled ? "" : "disabled"}>Quitar</button>`;

      const orderIndex = project.nomenclatureOrder.indexOf(field.id);
      const isIncluded = !!field.includeInCode;
      const orderChip = isIncluded && orderIndex >= 0
        ? `<span class="order-chip">${orderIndex + 1}</span>`
        : "<span class=\"muted\">-</span>";

      return `<tr>
        <td>${index + 1}. ${escapeHtml(field.label)}</td>
        <td><input type="checkbox" data-field-code="${escapeAttribute(field.id)}" ${field.includeInCode ? "checked" : ""} ${editEnabled ? "" : "disabled"}></td>
        <td>${orderChip}</td>
        <td>${removeButton}</td>
      </tr>`;
    })
    .join("");

  renderNomenclaturePreview(project);
}

function getMatrixColumnDescriptors(project) {
  if (!project || !Array.isArray(project.fields)) {
    return [];
  }

  const relation = getDisciplineSystemFields(project.fields);
  const descriptors = [];
  project.fields.forEach((field) => {
    descriptors.push({ fieldId: field.id, key: "name" });
    descriptors.push({ fieldId: field.id, key: "code" });
    if (field.id === relation.systemFieldId) {
      descriptors.push({ fieldId: field.id, key: "parentDisciplineRowId" });
    }
  });
  return descriptors;
}

function getMatrixCellPositionFromElement(project, element) {
  if (!project || !(element instanceof HTMLElement)) {
    return null;
  }

  const cellElement = element.closest(
    "[data-matrix-field][data-matrix-row][data-matrix-key], [data-matrix-system-parent-field][data-matrix-system-parent-row]"
  );
  if (!(cellElement instanceof HTMLElement)) {
    return null;
  }

  const descriptors = getMatrixColumnDescriptors(project);
  if (cellElement.dataset.matrixField) {
    const fieldId = trimOrFallback(cellElement.dataset.matrixField, "");
    const key = trimOrFallback(cellElement.dataset.matrixKey, "");
    const rowIndex = parseInt(cellElement.dataset.matrixRow || "-1", 10);
    if (!fieldId || !key || rowIndex < 0) {
      return null;
    }
    const colIndex = descriptors.findIndex((item) => item.fieldId === fieldId && item.key === key);
    if (colIndex < 0) {
      return null;
    }
    return { rowIndex, colIndex };
  }

  const fieldId = trimOrFallback(cellElement.dataset.matrixSystemParentField, "");
  const rowIndex = parseInt(cellElement.dataset.matrixSystemParentRow || "-1", 10);
  if (!fieldId || rowIndex < 0) {
    return null;
  }
  const colIndex = descriptors.findIndex((item) => item.fieldId === fieldId && item.key === "parentDisciplineRowId");
  if (colIndex < 0) {
    return null;
  }
  return { rowIndex, colIndex };
}

function buildMatrixSelectionRange(anchor, focus) {
  if (!anchor || !focus) {
    return null;
  }

  return {
    startRow: Math.min(anchor.rowIndex, focus.rowIndex),
    endRow: Math.max(anchor.rowIndex, focus.rowIndex),
    startCol: Math.min(anchor.colIndex, focus.colIndex),
    endCol: Math.max(anchor.colIndex, focus.colIndex)
  };
}

function clearMatrixSelection() {
  if (els.fieldMatrixBody) {
    els.fieldMatrixBody.querySelectorAll(".matrix-selected").forEach((node) => {
      node.classList.remove("matrix-selected");
    });
  }
  matrixSelectionAnchor = null;
  matrixSelectionRange = null;
}

function getMatrixCellElement(project, rowIndex, colIndex) {
  const descriptors = getMatrixColumnDescriptors(project);
  const descriptor = descriptors[colIndex];
  if (!descriptor || !els.fieldMatrixBody) {
    return null;
  }

  if (descriptor.key === "parentDisciplineRowId") {
    return els.fieldMatrixBody.querySelector(
      `[data-matrix-system-parent-field="${descriptor.fieldId}"][data-matrix-system-parent-row="${rowIndex}"]`
    );
  }

  return els.fieldMatrixBody.querySelector(
    `[data-matrix-field="${descriptor.fieldId}"][data-matrix-row="${rowIndex}"][data-matrix-key="${descriptor.key}"]`
  );
}

function paintMatrixSelection(project) {
  if (!els.fieldMatrixBody) {
    return;
  }

  els.fieldMatrixBody.querySelectorAll(".matrix-selected").forEach((node) => {
    node.classList.remove("matrix-selected");
  });
  if (!matrixSelectionRange || !project) {
    return;
  }

  for (let rowIndex = matrixSelectionRange.startRow; rowIndex <= matrixSelectionRange.endRow; rowIndex += 1) {
    for (let colIndex = matrixSelectionRange.startCol; colIndex <= matrixSelectionRange.endCol; colIndex += 1) {
      const element = getMatrixCellElement(project, rowIndex, colIndex);
      if (element instanceof HTMLElement) {
        element.classList.add("matrix-selected");
      }
    }
  }
}

function getMatrixCellValueForClipboard(project, descriptor, rowIndex) {
  if (!project || !descriptor) {
    return "";
  }

  const field = project.fields.find((item) => item.id === descriptor.fieldId);
  if (!field) {
    return "";
  }
  const row = field.rows[rowIndex];
  if (!row) {
    return "";
  }

  if (descriptor.key === "name" || descriptor.key === "code") {
    return trimOrFallback(row[descriptor.key], "");
  }

  if (descriptor.key === "parentDisciplineRowId") {
    const pair = getDisciplineSystemFields(project.fields);
    if (!pair.disciplineField) {
      return "";
    }
    const disciplineRow = getFieldRowById(pair.disciplineField, trimOrFallback(row.parentDisciplineRowId, ""));
    if (!disciplineRow) {
      return "";
    }
    return trimOrFallback(disciplineRow.code, "") || trimOrFallback(disciplineRow.name, "");
  }

  return "";
}

function buildMatrixSelectionClipboardText(project, range) {
  if (!project || !range) {
    return "";
  }

  const descriptors = getMatrixColumnDescriptors(project);
  const lines = [];
  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    const cells = [];
    for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
      const descriptor = descriptors[colIndex];
      const value = descriptor ? getMatrixCellValueForClipboard(project, descriptor, rowIndex) : "";
      cells.push(value);
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

function parseClipboardGrid(rawText) {
  const normalized = (rawText || "").replace(/\r/g, "");
  if (!normalized.trim()) {
    return [];
  }

  const rows = normalized.split("\n");
  if (rows.length > 0 && rows[rows.length - 1] === "") {
    rows.pop();
  }
  return rows.map((line) => line.split("\t"));
}

function resolveDisciplineRowIdFromText(project, rawValue) {
  const value = trimOrFallback(rawValue, "");
  if (!value) {
    return "";
  }

  const relation = getDisciplineSystemFields(project.fields);
  if (!relation.disciplineField) {
    return "";
  }

  const rows = getFieldOptions(relation.disciplineField);
  const byId = rows.find((row) => row.id === value);
  if (byId) {
    return byId.id;
  }

  const lookup = normalizeLookup(value);
  if (!lookup) {
    return "";
  }

  const byCode = rows.find((row) => normalizeLookup(row.code) === lookup);
  if (byCode) {
    return byCode.id;
  }

  const byName = rows.find((row) => normalizeLookup(row.name) === lookup);
  if (byName) {
    return byName.id;
  }

  const byLabel = rows.find((row) => normalizeLookup(formatValueWithCode(row.name, row.code)) === lookup);
  if (byLabel) {
    return byLabel.id;
  }

  return "";
}

function applyMatrixCellPastedValue(project, descriptor, rowIndex, rawValue) {
  if (!project || !descriptor || rowIndex < 0) {
    return;
  }

  const field = project.fields.find((item) => item.id === descriptor.fieldId);
  if (!field) {
    return;
  }

  const row = ensureFieldRow(field, rowIndex);
  if (descriptor.key === "name" || descriptor.key === "code") {
    row[descriptor.key] = rawValue;
    return;
  }

  if (descriptor.key === "parentDisciplineRowId") {
    row.parentDisciplineRowId = resolveDisciplineRowIdFromText(project, rawValue);
  }
}

function applyClipboardGridToMatrix(project, start, rows) {
  if (!project || !start || !Array.isArray(rows) || rows.length === 0) {
    return;
  }

  const descriptors = getMatrixColumnDescriptors(project);
  for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
    const rowValues = rows[rowOffset];
    if (!Array.isArray(rowValues)) {
      continue;
    }
    for (let colOffset = 0; colOffset < rowValues.length; colOffset += 1) {
      const colIndex = start.colIndex + colOffset;
      if (colIndex < 0 || colIndex >= descriptors.length) {
        continue;
      }
      const rowIndex = start.rowIndex + rowOffset;
      const descriptor = descriptors[colIndex];
      applyMatrixCellPastedValue(project, descriptor, rowIndex, rowValues[colOffset]);
    }
  }
}

function getDeliverableEditableColumnDescriptors(project) {
  if (!project) {
    return [];
  }

  const showDeliverableProgressColumns = getProjectProgressControlMode(project) !== "package";
  const descriptors = [];
  if (showDeliverableProgressColumns) {
    descriptors.push({ key: "meta:startDate", type: "meta", metaKey: "startDate" });
    descriptors.push({ key: "meta:endDate", type: "meta", metaKey: "endDate" });
  }
  descriptors.push({ key: "meta:baseUnits", type: "meta", metaKey: "baseUnits" });
  if (showDeliverableProgressColumns) {
    descriptors.push({ key: "meta:realProgress", type: "meta", metaKey: "realProgress" });
  }
  project.fields.forEach((field) => {
    descriptors.push({ key: `field:${field.id}`, type: "field", fieldId: field.id });
  });
  return descriptors;
}

function getDeliverableVisibleRows(project) {
  if (!project) {
    return [];
  }
  return filterDeliverablesBySearch(project.deliverables, project, currentSearchQuery);
}

function getDeliverableCellPositionFromElement(project, element) {
  if (!project || !(element instanceof HTMLElement)) {
    return null;
  }

  const control = element.closest("[data-edit-deliverable]");
  if (!(control instanceof HTMLElement)) {
    return null;
  }

  const rowNode = control.closest("tr[data-deliverable-row][data-deliverable-index]");
  if (!(rowNode instanceof HTMLElement)) {
    return null;
  }

  const rowIndex = parseInt(rowNode.dataset.deliverableIndex || "-1", 10);
  if (!Number.isInteger(rowIndex) || rowIndex < 0) {
    return null;
  }

  const colRaw = control.dataset.deliverableColIndex;
  const colIndexFromData = parseInt(colRaw || "-1", 10);
  if (Number.isInteger(colIndexFromData) && colIndexFromData >= 0) {
    return { rowIndex, colIndex: colIndexFromData };
  }

  const descriptors = getDeliverableEditableColumnDescriptors(project);
  const metaKey = trimOrFallback(control.dataset.editMeta, "");
  const fieldId = trimOrFallback(control.dataset.editField, "");
  const key = metaKey ? `meta:${metaKey}` : (fieldId ? `field:${fieldId}` : "");
  if (!key) {
    return null;
  }
  const colIndex = descriptors.findIndex((descriptor) => descriptor.key === key);
  if (colIndex < 0) {
    return null;
  }
  return { rowIndex, colIndex };
}

function isCellPositionInsideRange(position, range) {
  if (!position || !range) {
    return false;
  }
  return position.rowIndex >= range.startRow
    && position.rowIndex <= range.endRow
    && position.colIndex >= range.startCol
    && position.colIndex <= range.endCol;
}

function clearDeliverableSelection() {
  if (els.deliverablesBody) {
    els.deliverablesBody.querySelectorAll(".deliverable-cell-selected").forEach((node) => {
      node.classList.remove("deliverable-cell-selected");
    });
  }
  deliverableSelectionAnchor = null;
  deliverableSelectionRange = null;
}

function paintDeliverableSelection(project) {
  if (!project || !els.deliverablesBody) {
    return;
  }
  els.deliverablesBody.querySelectorAll(".deliverable-cell-selected").forEach((node) => {
    node.classList.remove("deliverable-cell-selected");
  });
  if (!deliverableSelectionRange) {
    return;
  }

  els.deliverablesBody.querySelectorAll("[data-edit-deliverable][data-deliverable-row-index][data-deliverable-col-index]").forEach((node) => {
    if (!(node instanceof HTMLElement)) {
      return;
    }
    const rowIndex = parseInt(node.dataset.deliverableRowIndex || "-1", 10);
    const colIndex = parseInt(node.dataset.deliverableColIndex || "-1", 10);
    if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex) || rowIndex < 0 || colIndex < 0) {
      return;
    }
    if (rowIndex >= deliverableSelectionRange.startRow
      && rowIndex <= deliverableSelectionRange.endRow
      && colIndex >= deliverableSelectionRange.startCol
      && colIndex <= deliverableSelectionRange.endCol) {
      node.classList.add("deliverable-cell-selected");
    }
  });
}

function getDeliverableCellValueForClipboard(project, deliverable, descriptor) {
  if (!project || !deliverable || !descriptor) {
    return "";
  }

  if (descriptor.type === "meta") {
    if (descriptor.metaKey === "startDate" || descriptor.metaKey === "endDate") {
      return sanitizeDateInput(deliverable[descriptor.metaKey] || "");
    }
    if (descriptor.metaKey === "baseUnits") {
      return formatNumberForInput(sanitizeBaseUnits(deliverable.baseUnits));
    }
    if (descriptor.metaKey === "realProgress") {
      return formatNumberForInput(sanitizeRealProgress(deliverable.realProgress));
    }
    return "";
  }

  if (descriptor.type === "field") {
    const field = getFieldById(project.fields, descriptor.fieldId);
    const selectedId = trimOrFallback(deliverable.rowRefs?.[descriptor.fieldId], "");
    const selectedRow = getFieldRowById(field, selectedId);
    if (!selectedRow) {
      return "";
    }
    return formatValueWithCode(selectedRow.name, selectedRow.code);
  }

  return "";
}

function buildDeliverableSelectionClipboardText(project, range) {
  if (!project || !range) {
    return "";
  }
  const visibleRows = getDeliverableVisibleRows(project);
  const descriptors = getDeliverableEditableColumnDescriptors(project);
  if (!visibleRows.length || !descriptors.length) {
    return "";
  }

  const lines = [];
  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    const rowData = visibleRows[rowIndex];
    if (!rowData) {
      continue;
    }
    const cells = [];
    for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
      const descriptor = descriptors[colIndex];
      cells.push(getDeliverableCellValueForClipboard(project, rowData, descriptor));
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

function sanitizeDateFromClipboard(value) {
  const raw = trimOrFallback(value, "");
  if (!raw) {
    return "";
  }
  const direct = sanitizeDateInput(raw);
  if (direct) {
    return direct;
  }

  const dateParts = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (!dateParts) {
    return "";
  }
  const day = String(parseInt(dateParts[1], 10)).padStart(2, "0");
  const month = String(parseInt(dateParts[2], 10)).padStart(2, "0");
  const year = String(parseInt(dateParts[3], 10));
  return sanitizeDateInput(`${year}-${month}-${day}`);
}

function resolveFieldOptionIdFromText(project, field, rowRefs, rawValue) {
  const value = trimOrFallback(rawValue, "");
  if (!field || !value) {
    return "";
  }

  const options = getFieldOptions(field, { project, rowRefs: rowRefs || {} });
  const byId = options.find((row) => row.id === value);
  if (byId) {
    return byId.id;
  }

  const lookup = normalizeLookup(value);
  if (!lookup) {
    return "";
  }

  const byCode = options.find((row) => normalizeLookup(row.code) === lookup);
  if (byCode) {
    return byCode.id;
  }

  const byName = options.find((row) => normalizeLookup(row.name) === lookup);
  if (byName) {
    return byName.id;
  }

  const byLabel = options.find((row) => normalizeLookup(formatValueWithCode(row.name, row.code)) === lookup);
  if (byLabel) {
    return byLabel.id;
  }

  return "";
}

function applyPastedValueToDeliverableCell(project, deliverable, descriptor, rawValue) {
  if (!project || !deliverable || !descriptor) {
    return { changed: false, error: true };
  }

  const rawText = trimOrFallback(rawValue, "");
  if (descriptor.type === "meta") {
    if (descriptor.metaKey === "startDate" || descriptor.metaKey === "endDate") {
      const parsed = sanitizeDateFromClipboard(rawText);
      if (rawText && !parsed) {
        return { changed: false, error: true };
      }
      const previous = sanitizeDateInput(deliverable[descriptor.metaKey] || "");
      if (previous === parsed) {
        return { changed: false, error: false };
      }
      deliverable[descriptor.metaKey] = parsed;
      return { changed: true, error: false };
    }

    if (descriptor.metaKey === "baseUnits") {
      const parsed = sanitizeBaseUnits(rawText);
      if (rawText && parsed === "") {
        return { changed: false, error: true };
      }
      const nextValue = parsed === "" ? null : parsed;
      const previous = sanitizeBaseUnits(deliverable.baseUnits);
      if ((previous === "" ? null : previous) === nextValue) {
        return { changed: false, error: false };
      }
      deliverable.baseUnits = nextValue;
      return { changed: true, error: false };
    }

    if (descriptor.metaKey === "realProgress") {
      const parsed = sanitizeRealProgress(rawText);
      if (rawText && parsed === "") {
        return { changed: false, error: true };
      }
      const previous = sanitizeRealProgress(deliverable.realProgress);
      const nextValue = parsed === "" ? null : parsed;
      if ((previous === "" ? null : previous) === nextValue) {
        return { changed: false, error: false };
      }
      deliverable.realProgress = nextValue;
      if (parsed !== "" && (previous === "" || previous !== parsed)) {
        registerRealAdvanceLog(project.id, deliverable.id, previous, parsed);
      }
      return { changed: true, error: false };
    }
    return { changed: false, error: true };
  }

  if (descriptor.type === "field") {
    const field = getFieldById(project.fields, descriptor.fieldId);
    if (!field) {
      return { changed: false, error: true };
    }
    if (!deliverable.rowRefs) {
      deliverable.rowRefs = {};
    }
    const resolvedId = resolveFieldOptionIdFromText(project, field, deliverable.rowRefs, rawText);
    if (rawText && !resolvedId) {
      return { changed: false, error: true };
    }
    const previousId = trimOrFallback(deliverable.rowRefs[descriptor.fieldId], "");
    const nextId = resolvedId || "";
    if (previousId === nextId) {
      return { changed: false, error: false };
    }
    deliverable.rowRefs[descriptor.fieldId] = nextId;
    reconcileSystemSelectionForRowRefs(project, deliverable.rowRefs);
    applyDeliverableRows(project, deliverable, false);
    return { changed: true, error: false };
  }

  return { changed: false, error: true };
}

function applyClipboardGridToDeliverables(project, range, rows) {
  if (!project || !range || !Array.isArray(rows) || rows.length === 0) {
    return { applied: false, updated: 0, errors: 0 };
  }
  const descriptors = getDeliverableEditableColumnDescriptors(project);
  const visibleRows = getDeliverableVisibleRows(project);
  if (!descriptors.length || !visibleRows.length) {
    return { applied: false, updated: 0, errors: 0 };
  }

  const rowCount = Math.max(0, range.endRow - range.startRow + 1);
  const colCount = Math.max(0, range.endCol - range.startCol + 1);
  if (rowCount === 0 || colCount === 0) {
    return { applied: false, updated: 0, errors: 0 };
  }

  const normalizedGrid = rows.map((line) => Array.isArray(line) ? line : []);
  const sourceRows = normalizedGrid.length;
  const sourceCols = Math.max(1, ...normalizedGrid.map((line) => line.length));
  if (sourceRows === 0 || sourceCols === 0) {
    return { applied: false, updated: 0, errors: 0 };
  }

  let updated = 0;
  let errors = 0;
  let touched = 0;
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const targetRowIndex = range.startRow + rowOffset;
    const deliverable = visibleRows[targetRowIndex];
    if (!deliverable) {
      continue;
    }
    const sourceRow = normalizedGrid[rowOffset % sourceRows];
    for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
      const targetColIndex = range.startCol + colOffset;
      const descriptor = descriptors[targetColIndex];
      if (!descriptor) {
        continue;
      }
      const rawValue = trimOrFallback(sourceRow[colOffset % sourceCols], "");
      const result = applyPastedValueToDeliverableCell(project, deliverable, descriptor, rawValue);
      touched += 1;
      if (result.error) {
        errors += 1;
      } else if (result.changed) {
        updated += 1;
      }
    }
  }

  return { applied: touched > 0, updated, errors };
}

function copyTextToClipboard(text) {
  const safeText = String(text || "");
  if (typeof navigator !== "undefined" && navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(safeText).catch(() => {
      fallbackCopyTextToClipboard(safeText);
    });
    return;
  }
  fallbackCopyTextToClipboard(safeText);
}

function fallbackCopyTextToClipboard(text) {
  const area = document.createElement("textarea");
  area.value = text;
  area.setAttribute("readonly", "readonly");
  area.style.position = "fixed";
  area.style.left = "-9999px";
  document.body.appendChild(area);
  area.focus();
  area.select();
  try {
    document.execCommand("copy");
  } catch {
    // noop
  }
  document.body.removeChild(area);
}

function renderDeliverablesPanel(project) {
  if (!project) {
    els.deliverableInputs.innerHTML = "";
    if (els.drawerStartDateInput) {
      els.drawerStartDateInput.value = "";
    }
    if (els.drawerEndDateInput) {
      els.drawerEndDateInput.value = "";
    }
    if (els.drawerBaseUnitsInput) {
      els.drawerBaseUnitsInput.value = "";
    }
    if (els.drawerRealProgressInput) {
      els.drawerRealProgressInput.value = "";
    }
    els.deliverablesHeader.innerHTML = "";
    els.deliverablesBody.innerHTML = "";
    if (els.deliverablesMetaText) {
      els.deliverablesMetaText.textContent = "";
    }
    clearDeliverableSelection();
    renderPackageControlsPanel(project);
    closeDeliverableDrawer();
    closeTrackingPanel();
    return;
  }

  syncDraftValues(project);
  const draft = getDraftValues(project.id);
  reconcileSystemSelectionForRowRefs(project, draft);
  const editEnabled = midpEditMode;
  const progressMode = getProjectProgressControlMode(project);
  const showDeliverableProgressColumns = progressMode !== "package";
  const showMidpAdvanceDetails = getProjectMidpAdvanceDetailsVisible(project);

  els.deliverableInputs.innerHTML = project.fields
    .map((field) => {
      const options = getFieldOptions(field, { project, rowRefs: draft });
      const selected = trimOrFallback(draft[field.id], "");
      const optionHtml = options
        .map((row) => {
          const selectedAttr = row.id === selected ? "selected" : "";
          const label = row.code
            ? `${row.name || "(sin nombre)"} (${row.code})`
            : (row.name || "(sin nombre)");
          return `<option value="${escapeAttribute(row.id)}" ${selectedAttr}>${escapeHtml(label)}</option>`;
        })
        .join("");

      return `<label>${escapeHtml(field.label)}
        <select data-deliverable-input="${escapeAttribute(field.id)}">
          <option value=""></option>
          ${optionHtml}
        </select>
      </label>`;
    })
    .join("");

  const draftStartDate = sanitizeDateInput(draft[DRAFT_META_START]);
  const draftEndDate = sanitizeDateInput(draft[DRAFT_META_END]);
  const draftBaseUnitsParsed = sanitizeBaseUnits(draft[DRAFT_META_BASE_UNITS]);
  const draftRealProgressParsed = sanitizeRealProgress(draft[DRAFT_META_REAL_PROGRESS]);
  const draftBaseUnitsText = formatNumberForInput(draftBaseUnitsParsed);
  const draftRealProgressText = formatNumberForInput(draftRealProgressParsed);
  draft[DRAFT_META_START] = draftStartDate;
  draft[DRAFT_META_END] = draftEndDate;
  draft[DRAFT_META_BASE_UNITS] = draftBaseUnitsText;
  draft[DRAFT_META_REAL_PROGRESS] = draftRealProgressText;

  if (els.drawerStartDateInput) {
    els.drawerStartDateInput.value = draftStartDate;
  }
  if (els.drawerEndDateInput) {
    els.drawerEndDateInput.value = draftEndDate;
  }
  if (els.drawerBaseUnitsInput) {
    els.drawerBaseUnitsInput.value = draftBaseUnitsText;
  }
  if (els.drawerRealProgressInput) {
    els.drawerRealProgressInput.value = draftRealProgressText;
  }

  const headerColumns = project.fields
    .map((field) => `<th>${escapeHtml(field.label)}</th>`)
    .join("");
  const dateHeaderColumns = showDeliverableProgressColumns
    ? `<th>Fecha de inicio</th>
      <th>Fecha de fin</th>`
    : "";
  const progressHeaderColumns = showDeliverableProgressColumns
    ? `${showMidpAdvanceDetails
      ? `<th>Incidencia Proyecto</th>
      <th>Incidencia Disciplina</th>
      <th>Incidencia Sistema</th>`
      : ""}
      <th>Programado</th>
      ${showMidpAdvanceDetails
      ? `<th>Avance Programado Proyecto</th>
      <th>Avance Programado Disciplina</th>
      <th>Avance Programado Sistema</th>`
      : ""}
      <th>Avance Real</th>
      ${showMidpAdvanceDetails
      ? `<th>Avance Real Proyecto</th>
      <th>Avance Real Disciplina</th>
      <th>Avance Real Sistema</th>`
      : ""}`
    : "";
  els.deliverablesHeader.innerHTML = `<tr>
      <th>#</th>
      <th>Ref</th>
      <th>Nomenclatura</th>
      ${dateHeaderColumns}
      <th>UP Base</th>
      ${progressHeaderColumns}
      ${headerColumns}
      <th>Creado</th>
      <th>Accion</th>
    </tr>`;

  const projectTotalBaseUnits = computeProjectBaseUnitsTotal(project.deliverables);
  const projectTotalBaseUnitsText = formatNumberForInput(projectTotalBaseUnits) || "0";
  const disciplineFieldId = findFieldIdByAlias(project.fields, ["disciplina"]);
  const systemFieldId = findFieldIdByAlias(project.fields, ["sistema"]);
  const disciplineTotals = computeGroupedBaseUnitsTotals(project.deliverables, disciplineFieldId);
  const systemGroupFieldIds = disciplineFieldId && systemFieldId ? [disciplineFieldId, systemFieldId] : [];
  const systemTotals = computeGroupedBaseUnitsTotals(project.deliverables, systemGroupFieldIds);
  const filteredDeliverables = filterDeliverablesBySearch(project.deliverables, project, currentSearchQuery);
  const editableColumnDescriptors = getDeliverableEditableColumnDescriptors(project);
  const editableColumnIndexByKey = new Map(
    editableColumnDescriptors.map((descriptor, index) => [descriptor.key, index])
  );
  const startDateColIndex = editableColumnIndexByKey.get("meta:startDate") ?? -1;
  const endDateColIndex = editableColumnIndexByKey.get("meta:endDate") ?? -1;
  const baseUnitsColIndex = editableColumnIndexByKey.get("meta:baseUnits") ?? -1;
  const realProgressColIndex = editableColumnIndexByKey.get("meta:realProgress") ?? -1;
  if (els.deliverablesMetaText) {
    const total = project.deliverables.length;
    const visible = filteredDeliverables.length;
    const complete = project.deliverables.filter((item) => !!trimOrFallback(item.code || "", "")).length;
    const withSchedule = project.deliverables.filter((item) =>
      !!sanitizeDateInput(item.startDate || "") && !!sanitizeDateInput(item.endDate || "")).length;
    const searchLabel = currentSearchQuery ? ` | filtro activo` : "";
    const modeLabel = midpEditMode ? " | modo: edicion" : " | modo: lectura";
    if (showDeliverableProgressColumns) {
      els.deliverablesMetaText.textContent = `Entregables: ${visible}/${total} visibles | con codigo: ${complete} | UP proyecto: ${projectTotalBaseUnitsText} | con control de avance: ${withSchedule}${modeLabel}${searchLabel}`;
    } else {
      els.deliverablesMetaText.textContent = `Entregables: ${visible}/${total} visibles | con codigo: ${complete} | UP proyecto: ${projectTotalBaseUnitsText} | control: por paquetes${modeLabel}${searchLabel}`;
    }
  }

  els.deliverablesBody.innerHTML = filteredDeliverables
    .map((deliverable, index) => {
      const startDate = sanitizeDateInput(deliverable.startDate || "");
      const endDate = sanitizeDateInput(deliverable.endDate || "");
      const baseUnits = sanitizeBaseUnits(deliverable.baseUnits);
      const baseUnitsText = formatNumberForInput(baseUnits);
      const realProgress = sanitizeRealProgress(deliverable.realProgress);
      const realProgressText = formatNumberForInput(realProgress);
      const realProgressLabel = formatPercent(realProgress, 2);
      const incidenceRatio = computeProjectIncidenceRatio(baseUnits, projectTotalBaseUnits);
      const incidenceText = incidenceRatio === null ? "" : formatPercent(incidenceRatio * 100, 2);
      const incidenceDetail = incidenceText
        ? `${baseUnitsText || "0"} / ${projectTotalBaseUnitsText}`
        : "";
      const disciplineRowId = disciplineFieldId ? trimOrFallback(deliverable.rowRefs?.[disciplineFieldId], "") : "";
      const systemGroupKey = buildDeliverableGroupKey(deliverable, systemGroupFieldIds);
      const disciplineTotal = disciplineRowId ? (disciplineTotals.get(disciplineRowId) ?? "") : "";
      const systemTotal = systemGroupKey ? (systemTotals.get(systemGroupKey) ?? "") : "";
      const disciplineIncidenceRatio = computeProjectIncidenceRatio(baseUnits, disciplineTotal);
      const systemIncidenceRatio = computeProjectIncidenceRatio(baseUnits, systemTotal);
      const disciplineIncidenceText = disciplineIncidenceRatio === null ? "" : formatPercent(disciplineIncidenceRatio * 100, 2);
      const systemIncidenceText = systemIncidenceRatio === null ? "" : formatPercent(systemIncidenceRatio * 100, 2);
      const disciplineIncidenceDetail = disciplineIncidenceText
        ? `${baseUnitsText || "0"} / ${formatNumberForInput(disciplineTotal) || "0"}`
        : "";
      const systemIncidenceDetail = systemIncidenceText
        ? `${baseUnitsText || "0"} / ${formatNumberForInput(systemTotal) || "0"}`
        : "";
      const progress = buildProgressSnapshot(startDate, endDate);
      const progressClass = progress.tone ? ` ${progress.tone}` : "";
      const programmedProjectPercent = computeWeightedProjectProgress(incidenceRatio, progress.percent);
      const programmedProjectText = formatPercent(programmedProjectPercent, 2);
      const programmedProjectDetail = programmedProjectText
        ? `Incidencia Proyecto ${incidenceText} x Programado ${progress.label}`
        : "";
      const programmedDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, progress.percent);
      const programmedDisciplineText = formatPercent(programmedDisciplinePercent, 2);
      const programmedDisciplineDetail = programmedDisciplineText
        ? `Incidencia Disciplina ${disciplineIncidenceText} x Programado ${progress.label}`
        : "";
      const programmedSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, progress.percent);
      const programmedSystemText = formatPercent(programmedSystemPercent, 2);
      const programmedSystemDetail = programmedSystemText
        ? `Incidencia Sistema ${systemIncidenceText} x Programado ${progress.label}`
        : "";
      const realProjectPercent = computeWeightedProjectProgress(incidenceRatio, realProgress);
      const realProjectText = formatPercent(realProjectPercent, 2);
      const realProjectDetail = realProjectText
        ? `Incidencia Proyecto ${incidenceText} x Avance Real ${realProgressLabel || "0%"}`
        : "";
      const realDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, realProgress);
      const realDisciplineText = formatPercent(realDisciplinePercent, 2);
      const realDisciplineDetail = realDisciplineText
        ? `Incidencia Disciplina ${disciplineIncidenceText} x Avance Real ${realProgressLabel || "0%"}`
        : "";
      const realSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, realProgress);
      const realSystemText = formatPercent(realSystemPercent, 2);
      const realSystemDetail = realSystemText
        ? `Incidencia Sistema ${systemIncidenceText} x Avance Real ${realProgressLabel || "0%"}`
        : "";
      const valueColumns = project.fields
        .map((field) => {
          const selectedId = trimOrFallback(deliverable.rowRefs?.[field.id], "");
          const selectedRow = getFieldRowById(field, selectedId);
          const selectedText = selectedRow
            ? formatValueWithCode(selectedRow.name, selectedRow.code)
            : "";

          if (!editEnabled) {
            return `<td><span class="cell-readonly">${escapeHtml(selectedText)}</span></td>`;
          }

          const options = getFieldOptions(field, {
            project,
            rowRefs: deliverable.rowRefs || {}
          });
          const optionHtml = options
            .map((row) => {
              const selectedAttr = row.id === selectedId ? "selected" : "";
              const label = row.code
                ? `${row.name || "(sin nombre)"} (${row.code})`
                : (row.name || "(sin nombre)");
              return `<option value="${escapeAttribute(row.id)}" ${selectedAttr}>${escapeHtml(label)}</option>`;
            })
            .join("");

          return `<td>
            <select class="cell-select" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-field="${escapeAttribute(field.id)}" data-deliverable-row-index="${index}" data-deliverable-col-index="${editableColumnIndexByKey.get(`field:${field.id}`) ?? -1}">
              <option value=""></option>
              ${optionHtml}
            </select>
          </td>`;
        })
        .join("");

      const startDateCell = editEnabled
        ? `<input type="date" class="cell-date-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="startDate" data-deliverable-row-index="${index}" data-deliverable-col-index="${startDateColIndex}" value="${escapeAttribute(startDate)}">`
        : `<span class="cell-readonly">${escapeHtml(formatDateFromInput(startDate) || "")}</span>`;
      const endDateCell = editEnabled
        ? `<input type="date" class="cell-date-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="endDate" data-deliverable-row-index="${index}" data-deliverable-col-index="${endDateColIndex}" value="${escapeAttribute(endDate)}">`
        : `<span class="cell-readonly">${escapeHtml(formatDateFromInput(endDate) || "")}</span>`;
      const baseUnitsCell = editEnabled
        ? `<input type="number" class="cell-number-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="baseUnits" data-deliverable-row-index="${index}" data-deliverable-col-index="${baseUnitsColIndex}" min="0" step="0.01" value="${escapeAttribute(baseUnitsText)}" placeholder="0">`
        : `<span class="cell-readonly">${escapeHtml(baseUnitsText || "")}</span>`;
      const realProgressCell = editEnabled
        ? `<input type="number" class="cell-number-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="realProgress" data-deliverable-row-index="${index}" data-deliverable-col-index="${realProgressColIndex}" min="0" max="100" step="0.01" value="${escapeAttribute(realProgressText)}" placeholder="0">`
        : `<span class="cell-readonly">${escapeHtml(realProgressLabel || "")}</span>`;
      const actionCell = editEnabled
        ? `<button type="button" class="danger" data-remove-deliverable="${escapeAttribute(deliverable.id)}">Eliminar</button>`
        : `<span class="muted">-</span>`;
      const allowDeliverableTracking = getProjectProgressControlMode(project) !== "package";
      const rowClass = !editEnabled && allowDeliverableTracking ? " class=\"row-openable\"" : "";
      const dateCells = showDeliverableProgressColumns
        ? `<td>${startDateCell}</td>
        <td>${endDateCell}</td>`
        : "";
      const progressCells = showDeliverableProgressColumns
        ? `${showMidpAdvanceDetails
          ? `<td>${incidenceText ? `<span class="incidence-chip" title="${escapeAttribute(incidenceDetail)}">${escapeHtml(incidenceText)}</span>` : ""}</td>
        <td>${disciplineIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(disciplineIncidenceDetail)}">${escapeHtml(disciplineIncidenceText)}</span>` : ""}</td>
        <td>${systemIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(systemIncidenceDetail)}">${escapeHtml(systemIncidenceText)}</span>` : ""}</td>`
          : ""}
        <td><span class="progress-chip${progressClass}">${escapeHtml(progress.label)}</span></td>
        ${showMidpAdvanceDetails
          ? `<td>${programmedProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedProjectDetail)}">${escapeHtml(programmedProjectText)}</span>` : ""}</td>
        <td>${programmedDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedDisciplineDetail)}">${escapeHtml(programmedDisciplineText)}</span>` : ""}</td>
        <td>${programmedSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedSystemDetail)}">${escapeHtml(programmedSystemText)}</span>` : ""}</td>`
          : ""}
        <td>${realProgressCell}</td>
        ${showMidpAdvanceDetails
          ? `<td>${realProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(realProjectDetail)}">${escapeHtml(realProjectText)}</span>` : ""}</td>
        <td>${realDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(realDisciplineDetail)}">${escapeHtml(realDisciplineText)}</span>` : ""}</td>
        <td>${realSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(realSystemDetail)}">${escapeHtml(realSystemText)}</span>` : ""}</td>`
          : ""}`
        : "";

      return `<tr data-deliverable-row="${escapeAttribute(deliverable.id)}" data-deliverable-index="${index}"${rowClass}>
        <td>${visibleIndexLabel(project.deliverables, deliverable, index)}</td>
        <td><span class="ref-chip">${escapeHtml(deliverable.ref || "")}</span></td>
        <td>${escapeHtml(deliverable.code || "")}</td>
        ${dateCells}
        <td>${baseUnitsCell}</td>
        ${progressCells}
        ${valueColumns}
        <td>${formatDate(deliverable.createdAt)}</td>
        <td>${actionCell}</td>
      </tr>`;
    })
    .join("");

  clearDeliverableSelection();

  renderPackageControlsPanel(project);
  renderReviewControlsPanel(project);

  if (trackingPanelTargetType === "deliverable" && trackingPanelTargetId) {
    const selectedDeliverable = project.deliverables.find((item) => item.id === trackingPanelTargetId);
    if (!selectedDeliverable) {
      closeTrackingPanel();
    } else {
      renderTrackingPanel(project, selectedDeliverable, "deliverable");
    }
  }
}

function renderPackageControlsPanel(project) {
  if (!els.packageControlsHeader || !els.packageControlsBody || !els.packageControlsMetaText) {
    return;
  }

  els.packageControlsHeader.innerHTML = `<tr>
      <th>#</th>
      <th>Combinacion</th>
      <th>Fecha de inicio</th>
      <th>Fecha de fin</th>
      <th>UP Base</th>
      <th>Incidencia Proyecto</th>
      <th>Incidencia Disciplina</th>
      <th>Incidencia Sistema</th>
      <th>Programado</th>
      <th>Avance Programado Proyecto</th>
      <th>Avance Programado Disciplina</th>
      <th>Avance Programado Sistema</th>
      <th>Avance Real</th>
      <th>Avance Real Proyecto</th>
      <th>Avance Real Disciplina</th>
      <th>Avance Real Sistema</th>
      <th>Proyecto</th>
      <th>Disciplina</th>
      <th>Sistema</th>
      <th>Paquete</th>
      <th>Entregables MIDP</th>
      <th>Creado</th>
    </tr>`;

  if (!project) {
    els.packageControlsBody.innerHTML = "";
    els.packageControlsMetaText.textContent = "";
    return;
  }

  ensurePackageControls(project);
  const editEnabled = packageEditMode;
  const fieldIds = getPackageControlFieldIds(project.fields);
  const hasBaseFields = !!(fieldIds.projectFieldId && fieldIds.disciplineFieldId && fieldIds.systemFieldId && fieldIds.packageFieldId);
  if (!hasBaseFields) {
    els.packageControlsBody.innerHTML = "<tr><td colspan=\"22\" class=\"muted\">Faltan campos base para sincronizar (Proyecto, Disciplina, Sistema, Paquete).</td></tr>";
    els.packageControlsMetaText.textContent = "Configura primero los campos base en Listas.";
    return;
  }

  const metricsByKey = computePackageControlMetrics(project.deliverables, fieldIds);
  const projectTotalBaseUnits = computeProjectBaseUnitsTotal(project.deliverables);
  const projectTotalBaseUnitsText = formatNumberForInput(projectTotalBaseUnits) || "0";
  const groupedTotals = computePackageControlGroupedTotals(metricsByKey);
  const existingKeys = new Set(project.packageControls.map((item) => item.key));
  let unsyncedCount = 0;
  metricsByKey.forEach((_, key) => {
    if (!existingKeys.has(key)) {
      unsyncedCount += 1;
    }
  });

  const projectField = getFieldById(project.fields, fieldIds.projectFieldId);
  const disciplineField = getFieldById(project.fields, fieldIds.disciplineFieldId);
  const systemField = getFieldById(project.fields, fieldIds.systemFieldId);
  const packageField = getFieldById(project.fields, fieldIds.packageFieldId);

  const rows = project.packageControls.map((item, index) => {
    const projectRow = getFieldRowById(projectField, item.projectRowId);
    const disciplineRow = getFieldRowById(disciplineField, item.disciplineRowId);
    const systemRow = getFieldRowById(systemField, item.systemRowId);
    const packageRow = getFieldRowById(packageField, item.packageRowId);
    const combination = buildPackageControlCode(project, [projectRow, disciplineRow, systemRow, packageRow]);
    const metrics = metricsByKey.get(item.key) || {
      projectRowId: item.projectRowId,
      disciplineRowId: item.disciplineRowId,
      systemRowId: item.systemRowId,
      packageRowId: item.packageRowId,
      deliverablesCount: 0,
      baseUnitsTotal: 0,
      startDate: "",
      endDate: "",
      hasInvalidDates: false,
      realPercent: null
    };

    const startDateRaw = sanitizeDateInput(item.startDate || "");
    const endDateRaw = sanitizeDateInput(item.endDate || "");
    const startDateText = formatDateFromInput(startDateRaw);
    const endDateText = formatDateFromInput(endDateRaw);
    const baseUnits = sanitizeBaseUnits(metrics.baseUnitsTotal);
    const baseUnitsText = formatNumberForInput(baseUnits) || "";

    const incidenceRatio = computeProjectIncidenceRatio(baseUnits, projectTotalBaseUnits);
    const incidenceText = incidenceRatio === null ? "" : formatPercent(incidenceRatio * 100, 2);
    const incidenceDetail = incidenceText
      ? `${baseUnitsText || "0"} / ${projectTotalBaseUnitsText}`
      : "";

    const disciplineTotal = metrics.disciplineRowId
      ? (groupedTotals.disciplineTotals.get(metrics.disciplineRowId) ?? "")
      : "";
    const disciplineIncidenceRatio = computeProjectIncidenceRatio(baseUnits, disciplineTotal);
    const disciplineIncidenceText = disciplineIncidenceRatio === null ? "" : formatPercent(disciplineIncidenceRatio * 100, 2);
    const disciplineIncidenceDetail = disciplineIncidenceText
      ? `${baseUnitsText || "0"} / ${formatNumberForInput(disciplineTotal) || "0"}`
      : "";

    const systemGroupKey = buildPackageControlSystemGroupKey(metrics.disciplineRowId, metrics.systemRowId);
    const systemTotal = systemGroupKey
      ? (groupedTotals.systemTotals.get(systemGroupKey) ?? "")
      : "";
    const systemIncidenceRatio = computeProjectIncidenceRatio(baseUnits, systemTotal);
    const systemIncidenceText = systemIncidenceRatio === null ? "" : formatPercent(systemIncidenceRatio * 100, 2);
    const systemIncidenceDetail = systemIncidenceText
      ? `${baseUnitsText || "0"} / ${formatNumberForInput(systemTotal) || "0"}`
      : "";

    // En control por paquetes, el calculo de programado es independiente:
    // solo usa el rango de fechas del propio paquete (editable en esta vista).
    const hasInvalidDates = isDateRangeInvalid(startDateRaw, endDateRaw);
    const progress = hasInvalidDates
      ? { label: "Fechas invertidas", tone: "overdue", percent: null }
      : buildProgressSnapshot(startDateRaw || "", endDateRaw || "");
    const progressClass = progress.tone ? ` ${progress.tone}` : "";

    const programmedProjectPercent = computeWeightedProjectProgress(incidenceRatio, progress.percent);
    const programmedProjectText = formatPercent(programmedProjectPercent, 2);
    const programmedProjectDetail = programmedProjectText
      ? `Incidencia Proyecto ${incidenceText} x Programado ${progress.label}`
      : "";

    const programmedDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, progress.percent);
    const programmedDisciplineText = formatPercent(programmedDisciplinePercent, 2);
    const programmedDisciplineDetail = programmedDisciplineText
      ? `Incidencia Disciplina ${disciplineIncidenceText} x Programado ${progress.label}`
      : "";

    const programmedSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, progress.percent);
    const programmedSystemText = formatPercent(programmedSystemPercent, 2);
    const programmedSystemDetail = programmedSystemText
      ? `Incidencia Sistema ${systemIncidenceText} x Programado ${progress.label}`
      : "";

    const realPercent = sanitizeRealProgress(item.realProgress);
    const realValueForCalc = realPercent === "" ? null : realPercent;
    const realText = formatPercent(realValueForCalc, 2);

    const realProjectPercent = computeWeightedProjectProgress(incidenceRatio, realValueForCalc);
    const realProjectText = formatPercent(realProjectPercent, 2);
    const realProjectDetail = realProjectText
      ? `Incidencia Proyecto ${incidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const realDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, realValueForCalc);
    const realDisciplineText = formatPercent(realDisciplinePercent, 2);
    const realDisciplineDetail = realDisciplineText
      ? `Incidencia Disciplina ${disciplineIncidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const realSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, realValueForCalc);
    const realSystemText = formatPercent(realSystemPercent, 2);
    const realSystemDetail = realSystemText
      ? `Incidencia Sistema ${systemIncidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const projectLabel = projectRow ? formatValueWithCode(projectRow.name, projectRow.code) : "";
    const disciplineLabel = disciplineRow ? formatValueWithCode(disciplineRow.name, disciplineRow.code) : "";
    const systemLabel = systemRow ? formatValueWithCode(systemRow.name, systemRow.code) : "";
    const packageLabel = packageRow ? formatValueWithCode(packageRow.name, packageRow.code) : "";
    const createdLabel = formatDate(item.createdAt);

    const searchBlob = normalizeLookup(
      `${combination} ${startDateText} ${endDateText} ${projectLabel} ${disciplineLabel} ${systemLabel} ${packageLabel} ${metrics.deliverablesCount} ${baseUnitsText} ${incidenceText} ${disciplineIncidenceText} ${systemIncidenceText} ${progress.label} ${programmedProjectText} ${programmedDisciplineText} ${programmedSystemText} ${realText} ${realProjectText} ${realDisciplineText} ${realSystemText} ${createdLabel}`);

    return {
      index,
      packageControlId: item.id,
      combination,
      startDateRaw,
      endDateRaw,
      startDateText,
      endDateText,
      projectLabel,
      disciplineLabel,
      systemLabel,
      packageLabel,
      deliverablesCount: metrics.deliverablesCount,
      baseUnitsText,
      incidenceText,
      incidenceDetail,
      disciplineIncidenceText,
      disciplineIncidenceDetail,
      systemIncidenceText,
      systemIncidenceDetail,
      progressLabel: progress.label,
      progressClass,
      programmedProjectText,
      programmedProjectDetail,
      programmedDisciplineText,
      programmedDisciplineDetail,
      programmedSystemText,
      programmedSystemDetail,
      realText,
      realProjectText,
      realProjectDetail,
      realDisciplineText,
      realDisciplineDetail,
      realSystemText,
      realSystemDetail,
      realInputText: formatNumberForInput(realPercent),
      createdLabel,
      searchBlob
    };
  });

  const filteredRows = currentSearchQuery
    ? rows.filter((row) => row.searchBlob.includes(currentSearchQuery))
    : rows;
  const searchLabel = currentSearchQuery ? " | filtro activo" : "";
  const withSchedule = rows.filter((row) => !!row.startDateText && !!row.endDateText).length;
  els.packageControlsMetaText.textContent = `Paquetes: ${filteredRows.length}/${project.packageControls.length} visibles | combinaciones MIDP: ${metricsByKey.size} | UP proyecto: ${projectTotalBaseUnitsText} | con control de avance: ${withSchedule} | pendientes por sincronizar: ${unsyncedCount}${searchLabel}`;

  if (filteredRows.length === 0) {
    els.packageControlsBody.innerHTML = "<tr><td colspan=\"22\" class=\"muted\">Sin combinaciones sincronizadas.</td></tr>";
    return;
  }

  els.packageControlsBody.innerHTML = filteredRows
    .map((row, index) => {
      const rowClass = editEnabled ? "" : " class=\"row-openable\"";
      return `<tr data-package-control-row="${escapeAttribute(row.packageControlId)}"${rowClass}>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.combination)}</td>
        <td>${editEnabled
    ? `<input type="date" class="cell-date-input" data-edit-package-control="${escapeAttribute(row.packageControlId)}" data-edit-package-meta="startDate" value="${escapeAttribute(row.startDateRaw)}">`
    : `<span class="cell-readonly">${escapeHtml(row.startDateText)}</span>`}</td>
        <td>${editEnabled
    ? `<input type="date" class="cell-date-input" data-edit-package-control="${escapeAttribute(row.packageControlId)}" data-edit-package-meta="endDate" value="${escapeAttribute(row.endDateRaw)}">`
    : `<span class="cell-readonly">${escapeHtml(row.endDateText)}</span>`}</td>
        <td>${escapeHtml(row.baseUnitsText)}</td>
        <td>${row.incidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.incidenceDetail)}">${escapeHtml(row.incidenceText)}</span>` : ""}</td>
        <td>${row.disciplineIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.disciplineIncidenceDetail)}">${escapeHtml(row.disciplineIncidenceText)}</span>` : ""}</td>
        <td>${row.systemIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.systemIncidenceDetail)}">${escapeHtml(row.systemIncidenceText)}</span>` : ""}</td>
        <td><span class="progress-chip${row.progressClass}">${escapeHtml(row.progressLabel)}</span></td>
        <td>${row.programmedProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedProjectDetail)}">${escapeHtml(row.programmedProjectText)}</span>` : ""}</td>
        <td>${row.programmedDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedDisciplineDetail)}">${escapeHtml(row.programmedDisciplineText)}</span>` : ""}</td>
        <td>${row.programmedSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedSystemDetail)}">${escapeHtml(row.programmedSystemText)}</span>` : ""}</td>
        <td>${editEnabled
    ? `<input type="number" class="cell-number-input" data-edit-package-control="${escapeAttribute(row.packageControlId)}" data-edit-package-meta="realProgress" min="0" max="100" step="0.01" value="${escapeAttribute(row.realInputText)}" placeholder="">`
    : (row.realText ? `<span class="project-progress-chip">${escapeHtml(row.realText)}</span>` : "")}</td>
        <td>${row.realProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realProjectDetail)}">${escapeHtml(row.realProjectText)}</span>` : ""}</td>
        <td>${row.realDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realDisciplineDetail)}">${escapeHtml(row.realDisciplineText)}</span>` : ""}</td>
        <td>${row.realSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realSystemDetail)}">${escapeHtml(row.realSystemText)}</span>` : ""}</td>
        <td><span class="cell-readonly">${escapeHtml(row.projectLabel)}</span></td>
        <td><span class="cell-readonly">${escapeHtml(row.disciplineLabel)}</span></td>
        <td><span class="cell-readonly">${escapeHtml(row.systemLabel)}</span></td>
        <td><span class="cell-readonly">${escapeHtml(row.packageLabel)}</span></td>
        <td>${row.deliverablesCount}</td>
        <td>${escapeHtml(row.createdLabel)}</td>
      </tr>`;
    })
    .join("");

  if (trackingPanelTargetType === "package" && trackingPanelTargetId) {
    const selectedPackageControl = project.packageControls.find((item) => item.id === trackingPanelTargetId);
    if (!selectedPackageControl) {
      closeTrackingPanel();
    } else {
      renderTrackingPanel(project, selectedPackageControl, "package");
    }
  }
}

function renderReviewFlowPanel(project) {
  if (!els.reviewMilestonesBody || !els.reviewMilestonesMetaText) {
    return;
  }
  const editEnabled = reviewFlowEditMode;

  if (!project) {
    if (els.reviewMilestoneNameInput) {
      els.reviewMilestoneNameInput.value = "";
    }
    if (els.reviewMilestoneWeightInput) {
      els.reviewMilestoneWeightInput.value = "";
    }
    els.reviewMilestonesBody.innerHTML = "<tr><td class=\"tracking-empty\" colspan=\"5\">Sin proyecto activo</td></tr>";
    els.reviewMilestonesMetaText.textContent = "";
    return;
  }

  ensureReviewMilestones(project);

  const rows = project.reviewMilestones.map((milestone, index) => {
    const baseUnits = sanitizeBaseUnits(milestone.baseUnits);
    const baseUnitsText = formatNumberForInput(baseUnits);
    const createdLabel = formatDate(milestone.createdAt);
    const name = trimOrFallback(milestone.name, "");
    const searchBlob = normalizeLookup(`${name} ${baseUnitsText} ${createdLabel}`);
    return {
      index,
      id: milestone.id,
      name,
      baseUnits: baseUnits === "" ? null : baseUnits,
      baseUnitsText,
      createdLabel,
      searchBlob
    };
  });

  const filteredRows = currentSearchQuery
    ? rows.filter((row) => row.searchBlob.includes(currentSearchQuery))
    : rows;
  const totalBaseUnits = rows.reduce((sum, row) => sum + (row.baseUnits || 0), 0);
  const totalBaseUnitsText = formatNumberForInput(totalBaseUnits) || "0";
  const searchLabel = currentSearchQuery ? " | filtro activo" : "";
  els.reviewMilestonesMetaText.textContent = `Hitos: ${filteredRows.length}/${rows.length} visibles | UP total: ${totalBaseUnitsText}${searchLabel}`;

  if (filteredRows.length === 0) {
    const label = rows.length === 0 ? "Sin hitos registrados." : "Sin resultados para el filtro actual.";
    els.reviewMilestonesBody.innerHTML = `<tr><td class="tracking-empty" colspan="5">${escapeHtml(label)}</td></tr>`;
    return;
  }

  els.reviewMilestonesBody.innerHTML = filteredRows
    .map((row, index) => `<tr>
      <td>${index + 1}</td>
      <td><input type="text" class="cell-text-input" data-edit-review-milestone="${escapeAttribute(row.id)}" data-review-key="name" value="${escapeAttribute(row.name)}" maxlength="120" ${editEnabled ? "" : "readonly"}></td>
      <td><input type="number" class="cell-number-input" data-edit-review-milestone="${escapeAttribute(row.id)}" data-review-key="baseUnits" value="${escapeAttribute(row.baseUnitsText)}" min="0" step="0.01" ${editEnabled ? "" : "readonly"}></td>
      <td>${escapeHtml(row.createdLabel)}</td>
      <td><button type="button" class="danger" data-remove-review-milestone="${escapeAttribute(row.id)}" ${editEnabled ? "" : "disabled"}>Eliminar</button></td>
    </tr>`)
    .join("");
}

function renderReviewControlsPanel(project) {
  if (!els.reviewControlsHeader || !els.reviewControlsBody || !els.reviewControlsMetaText) {
    return;
  }

  els.reviewControlsHeader.innerHTML = `<tr>
      <th>#</th>
      <th>Combinacion</th>
      <th>Hito</th>
      <th>Fecha de inicio</th>
      <th>Fecha de fin</th>
      <th>UP Base</th>
      <th>Incidencia Proyecto</th>
      <th>Incidencia Disciplina</th>
      <th>Incidencia Sistema</th>
      <th>Programado</th>
      <th>Avance Programado Proyecto</th>
      <th>Avance Programado Disciplina</th>
      <th>Avance Programado Sistema</th>
      <th>Avance Real</th>
      <th>Avance Real Proyecto</th>
      <th>Avance Real Disciplina</th>
      <th>Avance Real Sistema</th>
      <th>Proyecto</th>
      <th>Disciplina</th>
      <th>Sistema</th>
      <th>Creado</th>
    </tr>`;

  if (!project) {
    els.reviewControlsBody.innerHTML = "";
    els.reviewControlsMetaText.textContent = "";
    return;
  }

  ensureReviewMilestones(project);
  ensureReviewControls(project);
  const editEnabled = reviewControlsEditMode;
  const fieldIds = getReviewControlFieldIds(project.fields);
  const hasBaseFields = !!(fieldIds.projectFieldId && fieldIds.disciplineFieldId && fieldIds.systemFieldId);
  if (!hasBaseFields) {
    els.reviewControlsBody.innerHTML = "<tr><td colspan=\"21\" class=\"muted\">Faltan campos base para sincronizar (Proyecto, Disciplina, Sistema).</td></tr>";
    els.reviewControlsMetaText.textContent = "Configura primero los campos base en Listas.";
    return;
  }

  const milestones = project.reviewMilestones;
  if (milestones.length === 0) {
    els.reviewControlsBody.innerHTML = "<tr><td colspan=\"21\" class=\"muted\">No hay hitos registrados. Agrega hitos en Hitos de Flujo de Revision.</td></tr>";
    els.reviewControlsMetaText.textContent = "Sin hitos para sincronizar.";
    return;
  }

  const baseCandidates = collectReviewControlBaseCandidates(project, fieldIds);
  const desiredEntries = buildReviewControlDesiredEntries(baseCandidates, milestones);
  const existingKeys = new Set(project.reviewControls.map((item) => item.key));
  let unsyncedCount = 0;
  desiredEntries.forEach((_, key) => {
    if (!existingKeys.has(key)) {
      unsyncedCount += 1;
    }
  });

  const projectField = getFieldById(project.fields, fieldIds.projectFieldId);
  const disciplineField = getFieldById(project.fields, fieldIds.disciplineFieldId);
  const systemField = getFieldById(project.fields, fieldIds.systemFieldId);
  const milestoneMap = new Map(milestones.map((item) => [item.id, item]));

  const controlsWithBase = project.reviewControls.map((item) => {
    const milestone = milestoneMap.get(item.milestoneId) || null;
    const baseUnitsRaw = sanitizeBaseUnits(milestone?.baseUnits);
    const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
    return { item, baseUnits };
  });
  const projectTotalBaseUnits = Math.round(
    controlsWithBase.reduce((sum, row) => sum + (row.baseUnits || 0), 0) * 100
  ) / 100;
  const projectTotalBaseUnitsText = formatNumberForInput(projectTotalBaseUnits) || "0";

  const disciplineTotals = new Map();
  const systemTotals = new Map();
  controlsWithBase.forEach(({ item, baseUnits }) => {
    if (baseUnits <= 0) {
      return;
    }
    const disciplineRowId = trimOrFallback(item.disciplineRowId, "");
    if (disciplineRowId) {
      const next = (disciplineTotals.get(disciplineRowId) || 0) + baseUnits;
      disciplineTotals.set(disciplineRowId, Math.round(next * 100) / 100);
    }
    const systemGroupKey = buildPackageControlSystemGroupKey(item.disciplineRowId, item.systemRowId);
    if (systemGroupKey) {
      const next = (systemTotals.get(systemGroupKey) || 0) + baseUnits;
      systemTotals.set(systemGroupKey, Math.round(next * 100) / 100);
    }
  });

  const rows = project.reviewControls.map((item, index) => {
    const projectRow = getFieldRowById(projectField, item.projectRowId);
    const disciplineRow = getFieldRowById(disciplineField, item.disciplineRowId);
    const systemRow = getFieldRowById(systemField, item.systemRowId);
    const combination = buildPackageControlCode(project, [projectRow, disciplineRow, systemRow]);
    const milestone = milestoneMap.get(item.milestoneId) || null;
    const milestoneName = trimOrFallback(milestone?.name, "");

    const startDateRaw = sanitizeDateInput(item.startDate || "");
    const endDateRaw = sanitizeDateInput(item.endDate || "");
    const startDateText = formatDateFromInput(startDateRaw);
    const endDateText = formatDateFromInput(endDateRaw);

    const baseUnitsRaw = sanitizeBaseUnits(milestone?.baseUnits);
    const baseUnits = baseUnitsRaw === "" ? null : baseUnitsRaw;
    const baseUnitsForCalc = baseUnits === null ? 0 : baseUnits;
    const baseUnitsText = formatNumberForInput(baseUnits);

    const incidenceRatio = computeProjectIncidenceRatio(baseUnitsForCalc, projectTotalBaseUnits);
    const incidenceText = incidenceRatio === null ? "" : formatPercent(incidenceRatio * 100, 2);
    const incidenceDetail = incidenceText
      ? `${baseUnitsText || "0"} / ${projectTotalBaseUnitsText}`
      : "";

    const disciplineTotal = item.disciplineRowId
      ? (disciplineTotals.get(item.disciplineRowId) ?? "")
      : "";
    const disciplineIncidenceRatio = computeProjectIncidenceRatio(baseUnitsForCalc, disciplineTotal);
    const disciplineIncidenceText = disciplineIncidenceRatio === null ? "" : formatPercent(disciplineIncidenceRatio * 100, 2);
    const disciplineIncidenceDetail = disciplineIncidenceText
      ? `${baseUnitsText || "0"} / ${formatNumberForInput(disciplineTotal) || "0"}`
      : "";

    const systemGroupKey = buildPackageControlSystemGroupKey(item.disciplineRowId, item.systemRowId);
    const systemTotal = systemGroupKey
      ? (systemTotals.get(systemGroupKey) ?? "")
      : "";
    const systemIncidenceRatio = computeProjectIncidenceRatio(baseUnitsForCalc, systemTotal);
    const systemIncidenceText = systemIncidenceRatio === null ? "" : formatPercent(systemIncidenceRatio * 100, 2);
    const systemIncidenceDetail = systemIncidenceText
      ? `${baseUnitsText || "0"} / ${formatNumberForInput(systemTotal) || "0"}`
      : "";

    const hasInvalidDates = isDateRangeInvalid(startDateRaw, endDateRaw);
    const progress = hasInvalidDates
      ? { label: "Fechas invertidas", tone: "overdue", percent: null }
      : buildProgressSnapshot(startDateRaw || "", endDateRaw || "");
    const progressClass = progress.tone ? ` ${progress.tone}` : "";

    const programmedProjectPercent = computeWeightedProjectProgress(incidenceRatio, progress.percent);
    const programmedProjectText = formatPercent(programmedProjectPercent, 2);
    const programmedProjectDetail = programmedProjectText
      ? `Incidencia Proyecto ${incidenceText} x Programado ${progress.label}`
      : "";

    const programmedDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, progress.percent);
    const programmedDisciplineText = formatPercent(programmedDisciplinePercent, 2);
    const programmedDisciplineDetail = programmedDisciplineText
      ? `Incidencia Disciplina ${disciplineIncidenceText} x Programado ${progress.label}`
      : "";

    const programmedSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, progress.percent);
    const programmedSystemText = formatPercent(programmedSystemPercent, 2);
    const programmedSystemDetail = programmedSystemText
      ? `Incidencia Sistema ${systemIncidenceText} x Programado ${progress.label}`
      : "";

    const realPercent = sanitizeRealProgress(item.realProgress);
    const realValueForCalc = realPercent === "" ? null : realPercent;
    const realText = formatPercent(realValueForCalc, 2);

    const realProjectPercent = computeWeightedProjectProgress(incidenceRatio, realValueForCalc);
    const realProjectText = formatPercent(realProjectPercent, 2);
    const realProjectDetail = realProjectText
      ? `Incidencia Proyecto ${incidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const realDisciplinePercent = computeWeightedProjectProgress(disciplineIncidenceRatio, realValueForCalc);
    const realDisciplineText = formatPercent(realDisciplinePercent, 2);
    const realDisciplineDetail = realDisciplineText
      ? `Incidencia Disciplina ${disciplineIncidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const realSystemPercent = computeWeightedProjectProgress(systemIncidenceRatio, realValueForCalc);
    const realSystemText = formatPercent(realSystemPercent, 2);
    const realSystemDetail = realSystemText
      ? `Incidencia Sistema ${systemIncidenceText} x Avance Real ${realText || "0%"}`
      : "";

    const projectLabel = projectRow ? formatValueWithCode(projectRow.name, projectRow.code) : "";
    const disciplineLabel = disciplineRow ? formatValueWithCode(disciplineRow.name, disciplineRow.code) : "";
    const systemLabel = systemRow ? formatValueWithCode(systemRow.name, systemRow.code) : "";
    const createdLabel = formatDate(item.createdAt);

    const searchBlob = normalizeLookup(
      `${combination} ${milestoneName} ${startDateText} ${endDateText} ${projectLabel} ${disciplineLabel} ${systemLabel} ${baseUnitsText} ${incidenceText} ${disciplineIncidenceText} ${systemIncidenceText} ${progress.label} ${programmedProjectText} ${programmedDisciplineText} ${programmedSystemText} ${realText} ${createdLabel}`);

    return {
      index,
      reviewControlId: item.id,
      combination,
      milestoneName,
      startDateRaw,
      endDateRaw,
      startDateText,
      endDateText,
      projectLabel,
      disciplineLabel,
      systemLabel,
      baseUnitsText,
      incidenceText,
      incidenceDetail,
      disciplineIncidenceText,
      disciplineIncidenceDetail,
      systemIncidenceText,
      systemIncidenceDetail,
      progressLabel: progress.label,
      progressClass,
      programmedProjectText,
      programmedProjectDetail,
      programmedDisciplineText,
      programmedDisciplineDetail,
      programmedSystemText,
      programmedSystemDetail,
      realText,
      realProjectText,
      realProjectDetail,
      realDisciplineText,
      realDisciplineDetail,
      realSystemText,
      realSystemDetail,
      realInputText: formatNumberForInput(realPercent),
      createdLabel,
      searchBlob
    };
  });

  const filteredRows = currentSearchQuery
    ? rows.filter((row) => row.searchBlob.includes(currentSearchQuery))
    : rows;
  const searchLabel = currentSearchQuery ? " | filtro activo" : "";
  const withSchedule = rows.filter((row) => !!row.startDateText && !!row.endDateText).length;
  els.reviewControlsMetaText.textContent = `Flujos: ${filteredRows.length}/${project.reviewControls.length} visibles | combinaciones base: ${baseCandidates.size} | hitos: ${milestones.length} | UP total: ${projectTotalBaseUnitsText} | con control de avance: ${withSchedule} | pendientes por sincronizar: ${unsyncedCount}${searchLabel}`;

  if (filteredRows.length === 0) {
    els.reviewControlsBody.innerHTML = "<tr><td colspan=\"21\" class=\"muted\">Sin combinaciones sincronizadas.</td></tr>";
    return;
  }

  els.reviewControlsBody.innerHTML = filteredRows
    .map((row, index) => {
      const rowClass = editEnabled ? "" : " class=\"row-openable\"";
      return `<tr data-review-control-row="${escapeAttribute(row.reviewControlId)}"${rowClass}>
      <td>${index + 1}</td>
      <td>${escapeHtml(row.combination)}</td>
      <td>${escapeHtml(row.milestoneName)}</td>
      <td>${editEnabled
    ? `<input type="date" class="cell-date-input" data-edit-review-control="${escapeAttribute(row.reviewControlId)}" data-edit-review-control-meta="startDate" value="${escapeAttribute(row.startDateRaw)}">`
    : `<span class="cell-readonly">${escapeHtml(row.startDateText)}</span>`}</td>
      <td>${editEnabled
    ? `<input type="date" class="cell-date-input" data-edit-review-control="${escapeAttribute(row.reviewControlId)}" data-edit-review-control-meta="endDate" value="${escapeAttribute(row.endDateRaw)}">`
    : `<span class="cell-readonly">${escapeHtml(row.endDateText)}</span>`}</td>
      <td><span class="cell-readonly">${escapeHtml(row.baseUnitsText)}</span></td>
      <td>${row.incidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.incidenceDetail)}">${escapeHtml(row.incidenceText)}</span>` : ""}</td>
      <td>${row.disciplineIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.disciplineIncidenceDetail)}">${escapeHtml(row.disciplineIncidenceText)}</span>` : ""}</td>
      <td>${row.systemIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(row.systemIncidenceDetail)}">${escapeHtml(row.systemIncidenceText)}</span>` : ""}</td>
      <td><span class="progress-chip${row.progressClass}">${escapeHtml(row.progressLabel)}</span></td>
      <td>${row.programmedProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedProjectDetail)}">${escapeHtml(row.programmedProjectText)}</span>` : ""}</td>
      <td>${row.programmedDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedDisciplineDetail)}">${escapeHtml(row.programmedDisciplineText)}</span>` : ""}</td>
      <td>${row.programmedSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(row.programmedSystemDetail)}">${escapeHtml(row.programmedSystemText)}</span>` : ""}</td>
      <td>${editEnabled
    ? `<input type="number" class="cell-number-input" data-edit-review-control="${escapeAttribute(row.reviewControlId)}" data-edit-review-control-meta="realProgress" min="0" max="100" step="0.01" value="${escapeAttribute(row.realInputText)}" placeholder="">`
    : (row.realText ? `<span class="project-progress-chip">${escapeHtml(row.realText)}</span>` : "")}</td>
      <td>${row.realProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realProjectDetail)}">${escapeHtml(row.realProjectText)}</span>` : ""}</td>
      <td>${row.realDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realDisciplineDetail)}">${escapeHtml(row.realDisciplineText)}</span>` : ""}</td>
      <td>${row.realSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(row.realSystemDetail)}">${escapeHtml(row.realSystemText)}</span>` : ""}</td>
      <td><span class="cell-readonly">${escapeHtml(row.projectLabel)}</span></td>
      <td><span class="cell-readonly">${escapeHtml(row.disciplineLabel)}</span></td>
      <td><span class="cell-readonly">${escapeHtml(row.systemLabel)}</span></td>
      <td>${escapeHtml(row.createdLabel)}</td>
    </tr>`;
    })
    .join("");

  if (trackingPanelTargetType === "review-control" && trackingPanelTargetId) {
    const selectedReviewControl = project.reviewControls.find((item) => item.id === trackingPanelTargetId);
    if (!selectedReviewControl) {
      closeTrackingPanel();
    } else {
      renderTrackingPanel(project, selectedReviewControl, "review-control");
    }
  }
}

function ensureBiState(project) {
  if (!project) {
    return;
  }
  project.biConfig = normalizeBiConfig(project.biConfig);
  project.biWidgets = normalizeBiWidgets(project.biWidgets);
}

function renderBiPanel(project) {
  if (!els.biKpiGrid || !els.biWidgetsGrid) {
    return;
  }

  const biPanel = document.getElementById("tab-bi");
  if (project && activeTab !== "bi") {
    return;
  }
  if (
    project
    && activeTab === "bi"
    && biPanel instanceof HTMLElement
    && (biPanel.classList.contains("hidden") || biPanel.clientWidth <= 0 || biPanel.clientHeight <= 0)
  ) {
    if (!biPendingRenderRafId) {
      biPendingRenderRafId = window.requestAnimationFrame(() => {
        biPendingRenderRafId = 0;
        if (activeTab === "bi") {
          renderBiPanel(getActiveProject());
        }
      });
    }
    return;
  }

  if (!project) {
    if (els.biWidgetsMetaText) {
      els.biWidgetsMetaText.textContent = "";
    }
    if (els.biRowsMetaText) {
      els.biRowsMetaText.textContent = "";
    }
    if (els.biCatalogSummaryText) {
      els.biCatalogSummaryText.textContent = "";
    }
    els.biKpiGrid.innerHTML = "";
    els.biWidgetsGrid.innerHTML = "<div class=\"bi-empty\">Sin proyecto activo.</div>";
    if (els.biCatalogFieldsList) {
      els.biCatalogFieldsList.innerHTML = "<div class=\"bi-empty\">Sin proyecto activo.</div>";
    }
    if (els.biColorLegendList) {
      els.biColorLegendList.innerHTML = "<div class=\"bi-empty\">Sin proyecto activo.</div>";
    }
    if (els.biColorSummaryText) {
      els.biColorSummaryText.textContent = "";
    }
    if (els.biCrossFilterText) {
      els.biCrossFilterText.textContent = "Sin proyecto activo.";
    }
    if (els.biInsightsGrid) {
      els.biInsightsGrid.innerHTML = "";
    }
    if (els.biClearCrossFilterButton) {
      els.biClearCrossFilterButton.disabled = true;
    }
    if (els.biRowsHeader) {
      els.biRowsHeader.innerHTML = "";
    }
    if (els.biRowsBody) {
      els.biRowsBody.innerHTML = "";
    }
    renderBiPerformancePanel(null, []);
    refreshBiStudioPanelsUi();
    updateBiStudioOverlayOffset();
    return;
  }

  ensureBiState(project);
  if (selectedBiWidgetId && !project.biWidgets.some((widget) => widget.id === selectedBiWidgetId)) {
    selectedBiWidgetId = "";
  }
  biWidgetSnapshotsByProject[project.id] = {};
  refreshBiGroupByOptions(project);
  refreshBiBreakdownOptions(project);
  refreshBiColorGroupByOptions(project);
  refreshBiVisualPresetOptions(project);
  syncBiInputs(project.biConfig);
  syncBiBuilderInputsFromSelectedWidget(project);
  syncBiBuilderSelectionUi(project);
  renderBiFieldCatalog(project);
  renderBiCrossFilterSummary(project);

  const filteredRows = queryBiProjectRows(project, project.biConfig);
  renderBiColorPanel(project, filteredRows);

  renderBiInsights(project, filteredRows);
  renderBiKpis(filteredRows);
  renderBiWidgets(project, filteredRows);
  renderBiPerformancePanel(project, filteredRows);
  refreshBiStudioPanelsUi();
  updateBiStudioOverlayOffset();
}

function getQuickSightShellColumns() {
  if (quickSightPresentationMode) {
    return "minmax(0, 1fr)";
  }
  if (window.innerWidth <= 1180) {
    return "1fr";
  }
  const dataCol = quickSightPanelState.data ? "280px" : "0px";
  const visualsCol = quickSightPanelState.visuals ? "230px" : "0px";
  const canvasCol = "minmax(0, 1fr)";
  const propsCol = quickSightPanelState.props ? "280px" : "0px";
  return `${dataCol} ${visualsCol} ${canvasCol} ${propsCol}`;
}

function syncQuickSightViewportModeUi() {
  const quickSightActive = activeTab === "quicksight";
  document.body.classList.toggle("qs-tab-compact", quickSightActive);
  document.body.classList.toggle("qs-presentation-mode", quickSightActive && quickSightPresentationMode);
  if (els.appLayout instanceof HTMLElement) {
    els.appLayout.classList.toggle("qs-presentation-mode", quickSightActive && quickSightPresentationMode);
  }
  if (els.qsFrame instanceof HTMLElement) {
    els.qsFrame.classList.toggle("qs-presentation-frame", quickSightPresentationMode);
  }
  if (els.qsPresentationModeButton instanceof HTMLButtonElement) {
    els.qsPresentationModeButton.classList.toggle("active", quickSightPresentationMode);
    els.qsPresentationModeButton.setAttribute("aria-pressed", quickSightPresentationMode ? "true" : "false");
    els.qsPresentationModeButton.textContent = quickSightPresentationMode ? "Salir visualizacion" : "Visualizacion";
    els.qsPresentationModeButton.title = quickSightPresentationMode
      ? "Salir del modo de visualizacion"
      : "Entrar en modo de visualizacion";
  }
}

function toggleQuickSightPresentationMode(forceValue = null) {
  const nextValue = typeof forceValue === "boolean"
    ? forceValue
    : !quickSightPresentationMode;
  if (quickSightPresentationMode === nextValue) {
    syncQuickSightViewportModeUi();
    return false;
  }
  quickSightPresentationMode = nextValue;
  if (!quickSightPresentationMode) {
    setQuickSightEditMenuOpen(false);
  }
  setQuickSightZoomMenuOpen(false);
  syncQuickSightViewportModeUi();
  return true;
}

function getQuickSightExpandedCanvasLayout(canvasSize) {
  const width = Math.max(QUICKSIGHT_VISUAL_MIN_WIDTH, canvasSize.width - 24);
  const height = Math.max(QUICKSIGHT_VISUAL_MIN_HEIGHT, canvasSize.height - 24);
  return {
    x: 12,
    y: 12,
    w: width,
    h: height
  };
}

function toggleQuickSightExpandedVisual(project, visualId) {
  if (!project || !visualId || !project.quickSightVisuals.some((visual) => visual.id === visualId)) {
    quickSightExpandedVisualId = "";
    return false;
  }
  if (quickSightExpandedVisualId === visualId) {
    quickSightExpandedVisualId = "";
    return true;
  }
  quickSightExpandedVisualId = visualId;
  selectedQuickSightVisualId = visualId;
  return true;
}

function syncQuickSightEditMenuInputs(projectOrConfig) {
  if (!projectOrConfig) {
    return;
  }
  const config = projectOrConfig.quickSightConfig ? projectOrConfig.quickSightConfig : projectOrConfig;
  const canvasSize = getQuickSightCanvasSizeFromConfig(config);
  const zoom = sanitizeQuickSightCanvasZoom(config?.canvasZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  if (els.qsCanvasWidthInput instanceof HTMLInputElement) {
    els.qsCanvasWidthInput.value = String(canvasSize.width);
  }
  if (els.qsCanvasHeightInput instanceof HTMLInputElement) {
    els.qsCanvasHeightInput.value = String(canvasSize.height);
  }
  if (els.qsCanvasZoomSelect instanceof HTMLSelectElement) {
    syncQuickSightZoomSelectControl(els.qsCanvasZoomSelect, zoom);
  }
}

function setQuickSightEditMenuOpen(nextOpen) {
  quickSightEditMenuOpen = !!nextOpen;
  if (els.qsEditMenuWrap instanceof HTMLElement) {
    els.qsEditMenuWrap.classList.toggle("open", quickSightEditMenuOpen);
  }
  if (els.qsEditMenuButton instanceof HTMLElement) {
    els.qsEditMenuButton.classList.toggle("active", quickSightEditMenuOpen);
    els.qsEditMenuButton.setAttribute("aria-expanded", quickSightEditMenuOpen ? "true" : "false");
  }
  if (els.qsEditMenu instanceof HTMLElement) {
    els.qsEditMenu.setAttribute("aria-hidden", quickSightEditMenuOpen ? "false" : "true");
  }
}

function setQuickSightZoomMenuOpen(nextOpen) {
  quickSightZoomMenuOpen = !!nextOpen;
  if (els.qsZoomMenuWrap instanceof HTMLElement) {
    els.qsZoomMenuWrap.classList.toggle("open", quickSightZoomMenuOpen);
    const toolbarLeft = els.qsZoomMenuWrap.closest(".qs-toolbar-left");
    if (toolbarLeft instanceof HTMLElement) {
      toolbarLeft.classList.toggle("qs-toolbar-left-menu-open", quickSightZoomMenuOpen);
    }
    const toolbarRow = els.qsZoomMenuWrap.closest(".qs-toolbar-row");
    if (toolbarRow instanceof HTMLElement) {
      toolbarRow.classList.toggle("qs-toolbar-row-menu-open", quickSightZoomMenuOpen);
    }
  }
  if (els.qsCanvasZoomMenu instanceof HTMLElement) {
    els.qsCanvasZoomMenu.classList.toggle("hidden", !quickSightZoomMenuOpen);
  }
  if (els.qsCanvasZoomMenuButton instanceof HTMLButtonElement) {
    els.qsCanvasZoomMenuButton.classList.toggle("active", quickSightZoomMenuOpen);
    els.qsCanvasZoomMenuButton.setAttribute("aria-expanded", quickSightZoomMenuOpen ? "true" : "false");
  }
}

function setQuickSightCanvasZoom(rawZoomValue, statusPrefix = "Zoom pizarra Quicksigth") {
  const project = getActiveProject();
  if (!project) {
    return;
  }
  ensureQuickSightState(project);
  project.quickSightConfig.canvasZoom = sanitizeQuickSightCanvasZoom(
    rawZoomValue,
    project.quickSightConfig.canvasZoom
  );
  syncQuickSightEditMenuInputs(project);
  updateQuickSightZoomMenuUi(project);
  saveState();
  renderQuickSightPanel(project);
  setStatus(`${statusPrefix}: ${project.quickSightConfig.canvasZoom}%.`);
}

function updateQuickSightZoomMenuUi(project) {
  const safeZoom = sanitizeQuickSightCanvasZoom(project?.quickSightConfig?.canvasZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  if (els.qsCanvasZoomMenuLabel instanceof HTMLElement) {
    els.qsCanvasZoomMenuLabel.textContent = `Zoom (${safeZoom}%)`;
  }
  if (els.qsZoomFitAllLabel instanceof HTMLElement) {
    const fitAll = getQuickSightCanvasFitZoom(project, "all");
    els.qsZoomFitAllLabel.textContent = `Ajustar todo (${fitAll}%)`;
  }
  if (els.qsZoomFitWidthLabel instanceof HTMLElement) {
    const fitWidth = getQuickSightCanvasFitZoom(project, "width");
    els.qsZoomFitWidthLabel.textContent = `Ajustar ancho (${fitWidth}%)`;
  }
  if (els.qsZoomDefaultButton instanceof HTMLButtonElement) {
    const isDefault = safeZoom === QUICKSIGHT_CANVAS_ZOOM_DEFAULT;
    els.qsZoomDefaultButton.disabled = isDefault;
    els.qsZoomDefaultButton.classList.toggle("active", isDefault);
  }
  if (els.qsCanvasZoomMenu instanceof HTMLElement) {
    els.qsCanvasZoomMenu
      .querySelectorAll("[data-qs-zoom-value]")
      .forEach((node) => {
        if (!(node instanceof HTMLElement)) {
          return;
        }
        node.classList.toggle("active", trimOrFallback(node.dataset.qsZoomValue, "") === String(safeZoom));
      });
  }
}

function applyQuickSightPanelVisibility(announceMessage = "", animatedPanelKey = "") {
  const syncPanelStateClass = (panelNode, isVisible) => {
    if (!(panelNode instanceof HTMLElement)) {
      return;
    }
    panelNode.classList.toggle("qs-collapsed", !isVisible);
    panelNode.setAttribute("aria-hidden", isVisible ? "false" : "true");
  };

  syncPanelStateClass(els.qsDataPanel, quickSightPanelState.data);
  syncPanelStateClass(els.qsVisualsPanel, quickSightPanelState.visuals);
  syncPanelStateClass(els.qsPropsPanel, quickSightPanelState.props);
  if (els.qsShell instanceof HTMLElement) {
    els.qsShell.style.gridTemplateColumns = getQuickSightShellColumns();
  }

  const buttonStates = [
    [els.qsToggleDataButton, quickSightPanelState.data],
    [els.qsToggleVisualsButton, quickSightPanelState.visuals],
    [els.qsTogglePropsButton, quickSightPanelState.props]
  ];
  buttonStates.forEach(([button, isVisible]) => {
    if (!(button instanceof HTMLElement)) {
      return;
    }
    button.classList.toggle("active", !!isVisible);
    button.setAttribute("aria-pressed", isVisible ? "true" : "false");
  });
  if (announceMessage && activeTab === "quicksight") {
    setStatus(announceMessage);
  }
}

function cloneQuickSightSnapshotData(value, fallback) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return fallback;
  }
}

function getQuickSightHistoryTracker(projectId) {
  const key = trimOrFallback(projectId, "");
  if (!key) {
    return null;
  }
  if (!quickSightHistoryByProject[key]) {
    quickSightHistoryByProject[key] = {
      undo: [],
      redo: [],
      current: null,
      currentKey: "",
      suspended: false
    };
  }
  return quickSightHistoryByProject[key];
}

function captureQuickSightHistorySnapshot(project) {
  if (!project) {
    return null;
  }
  ensureQuickSightState(project);
  return {
    quickSightConfig: cloneQuickSightSnapshotData(project.quickSightConfig, createDefaultQuickSightConfig()),
    quickSightVisuals: cloneQuickSightSnapshotData(project.quickSightVisuals, []),
    selectedQuickSightVisualId: trimOrFallback(selectedQuickSightVisualId, ""),
    quickSightPropsTab: quickSightPropsTab === "interactions" ? "interactions" : "visual"
  };
}

function serializeQuickSightHistorySnapshot(snapshot) {
  if (!snapshot || typeof snapshot !== "object") {
    return "";
  }
  return JSON.stringify({
    quickSightConfig: snapshot.quickSightConfig || createDefaultQuickSightConfig(),
    quickSightVisuals: Array.isArray(snapshot.quickSightVisuals) ? snapshot.quickSightVisuals : []
  });
}

function syncQuickSightHistoryState(project) {
  if (!project?.id) {
    return;
  }
  const tracker = getQuickSightHistoryTracker(project.id);
  if (!tracker) {
    return;
  }
  const snapshot = captureQuickSightHistorySnapshot(project);
  const key = serializeQuickSightHistorySnapshot(snapshot);
  if (!key) {
    return;
  }
  if (tracker.suspended) {
    tracker.current = snapshot;
    tracker.currentKey = key;
    return;
  }
  if (!tracker.currentKey) {
    tracker.current = snapshot;
    tracker.currentKey = key;
    return;
  }
  if (tracker.currentKey === key) {
    tracker.current = snapshot;
    return;
  }
  if (tracker.current) {
    tracker.undo.push(tracker.current);
    if (tracker.undo.length > QUICKSIGHT_HISTORY_LIMIT) {
      tracker.undo.shift();
    }
  }
  tracker.redo = [];
  tracker.current = snapshot;
  tracker.currentKey = key;
}

function syncQuickSightHistoryButtons(project) {
  const tracker = project?.id ? getQuickSightHistoryTracker(project.id) : null;
  const undoDisabled = !tracker || tracker.undo.length === 0;
  const redoDisabled = !tracker || tracker.redo.length === 0;
  if (els.qsUndoButton instanceof HTMLButtonElement) {
    els.qsUndoButton.disabled = undoDisabled;
    els.qsUndoButton.setAttribute("aria-disabled", undoDisabled ? "true" : "false");
    els.qsUndoButton.title = undoDisabled ? "No hay cambios QuickSight para deshacer" : "Deshacer";
  }
  if (els.qsRedoButton instanceof HTMLButtonElement) {
    els.qsRedoButton.disabled = redoDisabled;
    els.qsRedoButton.setAttribute("aria-disabled", redoDisabled ? "true" : "false");
    els.qsRedoButton.title = redoDisabled ? "No hay cambios QuickSight para rehacer" : "Rehacer";
  }
}

function applyQuickSightHistorySnapshot(project, snapshot) {
  if (!project || !snapshot || typeof snapshot !== "object") {
    return false;
  }
  project.quickSightConfig = normalizeQuickSightConfig(snapshot.quickSightConfig);
  project.quickSightVisuals = normalizeQuickSightVisuals(snapshot.quickSightVisuals);
  ensureQuickSightState(project);
  quickSightPropsTab = snapshot.quickSightPropsTab === "interactions" ? "interactions" : "visual";
  const preferredId = trimOrFallback(snapshot.selectedQuickSightVisualId, "");
  selectedQuickSightVisualId = project.quickSightVisuals.some((visual) => visual.id === preferredId)
    ? preferredId
    : (project.quickSightVisuals[0]?.id || "");
  return true;
}

function stepQuickSightHistory(project, direction = "undo") {
  if (!project?.id) {
    return;
  }
  ensureQuickSightState(project);
  syncQuickSightHistoryState(project);
  const tracker = getQuickSightHistoryTracker(project.id);
  if (!tracker) {
    return;
  }
  const mode = direction === "redo" ? "redo" : "undo";
  const sourceStack = mode === "redo" ? tracker.redo : tracker.undo;
  if (sourceStack.length === 0) {
    setStatus(mode === "redo" ? "No hay cambios QuickSight para rehacer." : "No hay cambios QuickSight para deshacer.");
    syncQuickSightHistoryButtons(project);
    return;
  }
  const nextSnapshot = sourceStack.pop();
  const currentSnapshot = tracker.current || captureQuickSightHistorySnapshot(project);
  if (currentSnapshot) {
    const targetStack = mode === "redo" ? tracker.undo : tracker.redo;
    targetStack.push(currentSnapshot);
    if (targetStack.length > QUICKSIGHT_HISTORY_LIMIT) {
      targetStack.shift();
    }
  }
  tracker.suspended = true;
  applyQuickSightHistorySnapshot(project, nextSnapshot);
  tracker.current = nextSnapshot;
  tracker.currentKey = serializeQuickSightHistorySnapshot(nextSnapshot);
  renderQuickSightPanel(project, { softSelect: false, allowEmptySelection: true });
  tracker.suspended = false;
  saveState(true);
  syncQuickSightHistoryButtons(project);
  setStatus(mode === "redo" ? "Cambio QuickSight rehecho." : "Cambio QuickSight deshecho.");
}

function getDefaultQuickSightVisualLayout(index = 0) {
  const safeIndex = Math.max(0, Math.round(index || 0));
  const columns = 3;
  const column = safeIndex % columns;
  const row = Math.floor(safeIndex / columns);
  return {
    x: QUICKSIGHT_CANVAS_PADDING + (column * (QUICKSIGHT_VISUAL_DEFAULT_WIDTH + QUICKSIGHT_VISUAL_GAP_X)),
    y: QUICKSIGHT_CANVAS_PADDING + (row * (QUICKSIGHT_VISUAL_DEFAULT_HEIGHT + QUICKSIGHT_VISUAL_GAP_Y)),
    w: QUICKSIGHT_VISUAL_DEFAULT_WIDTH,
    h: QUICKSIGHT_VISUAL_DEFAULT_HEIGHT
  };
}

function sanitizeQuickSightCanvasZoom(value, fallback = QUICKSIGHT_CANVAS_ZOOM_DEFAULT) {
  const numeric = Math.round(Number(value));
  const fallbackNumeric = Math.round(Number(fallback));
  const safeFallback = Number.isFinite(fallbackNumeric)
    ? Math.max(QUICKSIGHT_CANVAS_ZOOM_MIN, Math.min(QUICKSIGHT_CANVAS_ZOOM_MAX, fallbackNumeric))
    : QUICKSIGHT_CANVAS_ZOOM_DEFAULT;
  if (!Number.isFinite(numeric)) {
    return safeFallback;
  }
  return Math.max(QUICKSIGHT_CANVAS_ZOOM_MIN, Math.min(QUICKSIGHT_CANVAS_ZOOM_MAX, numeric));
}

function getSteppedQuickSightCanvasZoom(currentZoom, direction) {
  const safeZoom = sanitizeQuickSightCanvasZoom(currentZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  const safeDirection = direction >= 0 ? 1 : -1;
  return sanitizeQuickSightCanvasZoom(safeZoom + (QUICKSIGHT_CANVAS_ZOOM_STEP * safeDirection), safeZoom);
}

function syncQuickSightZoomSelectControl(selectNode, zoom) {
  if (!(selectNode instanceof HTMLSelectElement)) {
    return;
  }
  const safeZoom = sanitizeQuickSightCanvasZoom(zoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  const safeValue = String(safeZoom);
  Array.from(selectNode.options).forEach((option) => {
    if (option.dataset.dynamicZoom === "1") {
      option.remove();
    }
  });
  const hasOption = Array.from(selectNode.options).some((option) => option.value === safeValue);
  if (!hasOption) {
    const dynamicOption = document.createElement("option");
    dynamicOption.value = safeValue;
    dynamicOption.textContent = `${safeValue}%`;
    dynamicOption.dataset.dynamicZoom = "1";
    selectNode.insertBefore(dynamicOption, selectNode.firstChild);
  }
  selectNode.value = safeValue;
}

function getQuickSightCanvasUsedArea(project) {
  const canvasSize = getQuickSightCanvasSizeFromConfig(project?.quickSightConfig);
  const visuals = Array.isArray(project?.quickSightVisuals) ? project.quickSightVisuals : [];
  if (visuals.length === 0) {
    return { width: canvasSize.width, height: canvasSize.height };
  }
  let maxRight = 0;
  let maxBottom = 0;
  visuals.forEach((visual, index) => {
    const layout = normalizeQuickSightVisualLayout(visual?.layout, index);
    maxRight = Math.max(maxRight, layout.x + layout.w);
    maxBottom = Math.max(maxBottom, layout.y + layout.h);
  });
  const paddedWidth = Math.round(maxRight + QUICKSIGHT_CANVAS_PADDING + 24);
  const paddedHeight = Math.round(maxBottom + QUICKSIGHT_CANVAS_PADDING + 24);
  return {
    width: Math.max(300, Math.min(canvasSize.width, paddedWidth || canvasSize.width)),
    height: Math.max(240, Math.min(canvasSize.height, paddedHeight || canvasSize.height))
  };
}

function getQuickSightCanvasFitZoom(project, mode = "all") {
  if (!project || !(els.qsCanvasBoard instanceof HTMLElement)) {
    return QUICKSIGHT_CANVAS_ZOOM_DEFAULT;
  }
  const usedArea = getQuickSightCanvasUsedArea(project);
  const viewportWidth = Math.max(120, Math.round(els.qsCanvasBoard.clientWidth - 64));
  const viewportHeight = Math.max(120, Math.round(els.qsCanvasBoard.clientHeight - 64));
  const widthRatio = viewportWidth / Math.max(1, usedArea.width);
  const heightRatio = viewportHeight / Math.max(1, usedArea.height);
  let ratio = mode === "width" ? widthRatio : Math.min(widthRatio, heightRatio);
  if (mode !== "width" && ratio < 0.34 && widthRatio > ratio * 1.45) {
    ratio = widthRatio;
  }
  const rawZoom = Math.floor(ratio * 100);
  return sanitizeQuickSightCanvasZoom(rawZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
}

function getQuickSightCanvasSizeFromConfig(config) {
  const width = sanitizeBiCanvasDimension(
    config?.canvasWidth,
    QUICKSIGHT_CANVAS_DEFAULT_WIDTH,
    QUICKSIGHT_CANVAS_MIN_WIDTH,
    BI_CANVAS_SURFACE_MAX_EDIT_WIDTH
  );
  const height = sanitizeBiCanvasDimension(
    config?.canvasHeight,
    QUICKSIGHT_CANVAS_DEFAULT_HEIGHT,
    QUICKSIGHT_CANVAS_MIN_HEIGHT,
    BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
  );
  return { width, height };
}

function clampQuickSightVisualLayoutToCanvas(layoutInput, config) {
  const canvasSize = getQuickSightCanvasSizeFromConfig(config);
  const normalized = normalizeQuickSightVisualLayout(layoutInput, 0);
  const width = Math.max(QUICKSIGHT_VISUAL_MIN_WIDTH, Math.min(canvasSize.width, normalized.w));
  const height = Math.max(QUICKSIGHT_VISUAL_MIN_HEIGHT, Math.min(canvasSize.height, normalized.h));
  const maxX = Math.max(0, canvasSize.width - width);
  const maxY = Math.max(0, canvasSize.height - height);
  const x = Math.max(0, Math.min(maxX, normalized.x));
  const y = Math.max(0, Math.min(maxY, normalized.y));
  return { x, y, w: width, h: height };
}

function normalizeQuickSightVisualLayout(rawLayout, index = 0) {
  const fallback = getDefaultQuickSightVisualLayout(index);
  const source = rawLayout && typeof rawLayout === "object" && !Array.isArray(rawLayout)
    ? rawLayout
    : {};
  const x = Number.isFinite(source.x) ? Math.max(0, Math.round(source.x)) : fallback.x;
  const y = Number.isFinite(source.y) ? Math.max(0, Math.round(source.y)) : fallback.y;
  const w = Number.isFinite(source.w)
    ? Math.max(QUICKSIGHT_VISUAL_MIN_WIDTH, Math.round(source.w))
    : fallback.w;
  const h = Number.isFinite(source.h)
    ? Math.max(QUICKSIGHT_VISUAL_MIN_HEIGHT, Math.round(source.h))
    : fallback.h;
  return { x, y, w, h };
}

function computeQuickSightSurfaceSize(visuals, config) {
  return getQuickSightCanvasSizeFromConfig(config);
}

function applyQuickSightSurfaceSize(surfaceElement, config) {
  if (!(surfaceElement instanceof HTMLElement)) {
    return;
  }
  const size = computeQuickSightSurfaceSize([], config);
  const zoomPercent = sanitizeQuickSightCanvasZoom(config?.canvasZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  const zoomScale = Math.max(0.25, Math.min(3, zoomPercent / 100));
  surfaceElement.style.width = `${size.width}px`;
  surfaceElement.style.height = `${size.height}px`;
  surfaceElement.style.minWidth = `${size.width}px`;
  surfaceElement.style.minHeight = `${size.height}px`;
  surfaceElement.style.transform = `scale(${zoomScale})`;
  surfaceElement.style.transformOrigin = "top left";
  surfaceElement.dataset.qsCanvasWidth = String(size.width);
  surfaceElement.dataset.qsCanvasHeight = String(size.height);
  surfaceElement.dataset.qsCanvasZoom = String(zoomPercent);
  const wrap = surfaceElement.parentElement;
  if (wrap instanceof HTMLElement && wrap.classList.contains("qs-canvas-zoom-wrap")) {
    wrap.style.width = `${Math.max(1, Math.round(size.width * zoomScale))}px`;
    wrap.style.height = `${Math.max(1, Math.round(size.height * zoomScale))}px`;
  }
}

function clampQuickSightInteractionValue(value, min, max) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return min;
  }
  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    return numeric;
  }
  return Math.min(max, Math.max(min, numeric));
}

function computeQuickSightResizedLayout(layoutStart, deltaX, deltaY, resizeHandle, config) {
  const handle = trimOrFallback(resizeHandle, "se").toLowerCase();
  const canvasSize = getQuickSightCanvasSizeFromConfig(config);
  const startX = Number(layoutStart?.x) || 0;
  const startY = Number(layoutStart?.y) || 0;
  const startW = Math.max(QUICKSIGHT_VISUAL_MIN_WIDTH, Number(layoutStart?.w) || QUICKSIGHT_VISUAL_MIN_WIDTH);
  const startH = Math.max(QUICKSIGHT_VISUAL_MIN_HEIGHT, Number(layoutStart?.h) || QUICKSIGHT_VISUAL_MIN_HEIGHT);
  const startRight = startX + startW;
  const startBottom = startY + startH;
  let nextX = startX;
  let nextY = startY;
  let nextW = startW;
  let nextH = startH;

  if (handle.includes("e")) {
    nextW = clampQuickSightInteractionValue(
      startW + deltaX,
      QUICKSIGHT_VISUAL_MIN_WIDTH,
      canvasSize.width - startX
    );
  } else if (handle.includes("w")) {
    const nextLeft = clampQuickSightInteractionValue(
      startX + deltaX,
      0,
      startRight - QUICKSIGHT_VISUAL_MIN_WIDTH
    );
    nextX = nextLeft;
    nextW = startRight - nextLeft;
  }

  if (handle.includes("s")) {
    nextH = clampQuickSightInteractionValue(
      startH + deltaY,
      QUICKSIGHT_VISUAL_MIN_HEIGHT,
      canvasSize.height - startY
    );
  } else if (handle.includes("n")) {
    const nextTop = clampQuickSightInteractionValue(
      startY + deltaY,
      0,
      startBottom - QUICKSIGHT_VISUAL_MIN_HEIGHT
    );
    nextY = nextTop;
    nextH = startBottom - nextTop;
  }

  return clampQuickSightVisualLayoutToCanvas({
    x: nextX,
    y: nextY,
    w: nextW,
    h: nextH
  }, config);
}

function repaintQuickSightVisualCanvasForCard(project, visual, cardElement) {
  if (!project || !visual || !(cardElement instanceof HTMLElement)) {
    return;
  }
  const chartHostNode = cardElement.querySelector("[data-qs-chart-host-id]");
  if (!(chartHostNode instanceof HTMLElement)) {
    return;
  }
  const snapshot = buildQuickSightVisualSnapshot(project, visual);
  paintQuickSightVisualCanvas(chartHostNode, project, visual, snapshot);
}

function scheduleQuickSightResizeCanvasRepaint(interaction, project, visual) {
  if (!interaction || !project || !visual || !(interaction.cardElement instanceof HTMLElement)) {
    return;
  }
  if (interaction.resizeRepaintRafId) {
    return;
  }
  interaction.resizeRepaintRafId = window.requestAnimationFrame(() => {
    interaction.resizeRepaintRafId = 0;
    repaintQuickSightVisualCanvasForCard(project, visual, interaction.cardElement);
  });
}

function startQuickSightVisualDrag(project, visual, cardElement, event, mode = "drag", resizeHandle = "se") {
  if (!project || !visual || !(cardElement instanceof HTMLElement) || !(event instanceof PointerEvent)) {
    return;
  }
  if (event.button !== 0) {
    return;
  }
  const surfaceElement = cardElement.closest(".qs-canvas-surface");
  if (!(surfaceElement instanceof HTMLElement)) {
    return;
  }
  const visualIndex = Math.max(0, project.quickSightVisuals.findIndex((item) => item.id === visual.id));
  const layout = clampQuickSightVisualLayoutToCanvas(
    normalizeQuickSightVisualLayout(visual.layout, visualIndex),
    project.quickSightConfig
  );
  visual.layout = { ...layout };
  const previousSelectedVisualId = selectedQuickSightVisualId;
  selectedQuickSightVisualId = visual.id;
  if (activeTab === "quicksight" && previousSelectedVisualId !== visual.id) {
    renderQuickSightPanel(project, { softSelect: true });
  }
  if (els.qsCanvasBoard instanceof HTMLElement) {
    els.qsCanvasBoard.querySelectorAll(".qs-visual-card.active").forEach((node) => {
      if (node instanceof HTMLElement) {
        node.classList.remove("active");
      }
    });
  }

  quickSightVisualInteraction = {
    pointerId: event.pointerId,
    projectId: project.id,
    visualId: visual.id,
    startX: event.clientX,
    startY: event.clientY,
    layoutStart: { ...layout },
    mode: mode === "resize" ? "resize" : "drag",
    resizeHandle: trimOrFallback(resizeHandle, "se").toLowerCase(),
    zoomScale: Math.max(
      0.25,
      Math.min(
        3,
        sanitizeQuickSightCanvasZoom(surfaceElement.dataset.qsCanvasZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT) / 100
      )
    ),
    moved: false,
    cardElement,
    surfaceElement,
    resizeRepaintRafId: 0,
    lastResizeRepaintAt: 0
  };
  cardElement.classList.add("active");
  if (quickSightVisualInteraction.mode === "resize") {
    cardElement.classList.add("is-resizing");
  } else {
    cardElement.classList.add("is-dragging");
  }
  if (typeof cardElement.setPointerCapture === "function") {
    try {
      cardElement.setPointerCapture(event.pointerId);
    } catch {
      // Ignore unsupported capture scenarios.
    }
  }
  event.preventDefault();
}

function handleQuickSightVisualDragMove(event) {
  if (!quickSightVisualInteraction || !(event instanceof PointerEvent)) {
    return;
  }
  if (event.pointerId !== quickSightVisualInteraction.pointerId) {
    return;
  }
  const project = state.projects.find((item) => item.id === quickSightVisualInteraction.projectId);
  if (!project) {
    finishQuickSightVisualDrag(event, true);
    return;
  }
  ensureQuickSightState(project);
  const visual = project.quickSightVisuals.find((item) => item.id === quickSightVisualInteraction.visualId);
  if (!visual) {
    finishQuickSightVisualDrag(event, true);
    return;
  }
  const zoomScale = Math.max(0.25, Math.min(3, Number(quickSightVisualInteraction.zoomScale) || 1));
  const deltaX = (event.clientX - quickSightVisualInteraction.startX) / zoomScale;
  const deltaY = (event.clientY - quickSightVisualInteraction.startY) / zoomScale;
  const nextLayout = quickSightVisualInteraction.mode === "resize"
    ? computeQuickSightResizedLayout(
      quickSightVisualInteraction.layoutStart,
      deltaX,
      deltaY,
      quickSightVisualInteraction.resizeHandle,
      project.quickSightConfig
    )
    : clampQuickSightVisualLayoutToCanvas({
      x: quickSightVisualInteraction.layoutStart.x + deltaX,
      y: quickSightVisualInteraction.layoutStart.y + deltaY,
      w: visual.layout?.w ?? quickSightVisualInteraction.layoutStart.w,
      h: visual.layout?.h ?? quickSightVisualInteraction.layoutStart.h
    }, project.quickSightConfig);
  visual.layout = nextLayout;
  if (quickSightVisualInteraction.cardElement instanceof HTMLElement) {
    quickSightVisualInteraction.cardElement.style.left = `${nextLayout.x}px`;
    quickSightVisualInteraction.cardElement.style.top = `${nextLayout.y}px`;
    quickSightVisualInteraction.cardElement.style.width = `${nextLayout.w}px`;
    quickSightVisualInteraction.cardElement.style.height = `${nextLayout.h}px`;
    const sizeNode = quickSightVisualInteraction.cardElement.querySelector(".qs-visual-size");
    if (sizeNode instanceof HTMLElement) {
      sizeNode.textContent = `${nextLayout.w} x ${nextLayout.h}`;
    }
  }
  if (quickSightVisualInteraction.mode === "resize") {
    const now = (typeof performance !== "undefined" && typeof performance.now === "function")
      ? performance.now()
      : Date.now();
    if ((now - (quickSightVisualInteraction.lastResizeRepaintAt || 0)) >= 40) {
      quickSightVisualInteraction.lastResizeRepaintAt = now;
      scheduleQuickSightResizeCanvasRepaint(quickSightVisualInteraction, project, visual);
    }
  }
  quickSightVisualInteraction.moved = true;
  event.preventDefault();
}

function finishQuickSightVisualDrag(event, canceled) {
  if (!quickSightVisualInteraction || !(event instanceof PointerEvent)) {
    return;
  }
  if (event.pointerId !== quickSightVisualInteraction.pointerId) {
    return;
  }
  const interaction = quickSightVisualInteraction;
  quickSightVisualInteraction = null;
  if (interaction.resizeRepaintRafId) {
    window.cancelAnimationFrame(interaction.resizeRepaintRafId);
    interaction.resizeRepaintRafId = 0;
  }
  if (interaction.cardElement instanceof HTMLElement) {
    interaction.cardElement.classList.remove("is-dragging");
    interaction.cardElement.classList.remove("is-resizing");
    if (typeof interaction.cardElement.releasePointerCapture === "function") {
      try {
        interaction.cardElement.releasePointerCapture(event.pointerId);
      } catch {
        // Ignore unsupported capture scenarios.
      }
    }
  }
  if (!canceled && interaction.moved && interaction.mode === "resize") {
    const project = state.projects.find((item) => item.id === interaction.projectId);
    if (project) {
      ensureQuickSightState(project);
      const visual = project.quickSightVisuals.find((item) => item.id === interaction.visualId);
      if (visual) {
        repaintQuickSightVisualCanvasForCard(project, visual, interaction.cardElement);
      }
    }
  }
  if (!canceled && interaction.moved) {
    quickSightSuppressClickUntil = Date.now() + 220;
    saveState();
    if (activeTab === "quicksight") {
      setStatus(interaction.mode === "resize" ? "Visual transformado en pizarra libre." : "Visual movido en pizarra libre.");
    }
  }
}

function getQuickSightMetricOptions() {
  return [
    { value: "count", label: "Cantidad de filas" },
    { value: "baseunits", label: "UP Base total" },
    { value: "baseavg", label: "UP Base promedio" },
    { value: "realavg", label: "Avance real promedio (%)" },
    { value: "programmedavg", label: "Avance programado promedio (%)" },
    { value: "weightedreal", label: "Avance real proyecto (%)" },
    { value: "weightedprogrammed", label: "Avance programado proyecto (%)" },
    { value: "weightedgap", label: "Brecha real - programado (%)" },
    { value: "invaliddates", label: "Fechas invertidas" }
  ];
}

function getQuickSightChartTypeOptions() {
  return [
    { value: "bar", label: "Barras" },
    { value: "bar_horizontal", label: "Barras horizontales" },
    { value: "bar_stacked", label: "Barras apiladas" },
    { value: "line", label: "Lineas" },
    { value: "area", label: "Area" },
    { value: "combo", label: "Combinado" },
    { value: "pie", label: "Circular" },
    { value: "donut", label: "Dona" },
    { value: "treemap", label: "Treemap" },
    { value: "funnel", label: "Embudo" },
    { value: "scatter", label: "Dispersion" },
    { value: "bubble", label: "Burbujas" },
    { value: "waterfall", label: "Waterfall" },
    { value: "radar", label: "Radar" },
    { value: "pareto", label: "Pareto" }
  ];
}

function captureQuickSightPropsSectionState() {
  if (!els.qsPropertiesBody) {
    return;
  }
  const nodes = els.qsPropertiesBody.querySelectorAll("details.qs-props-section[data-qs-props-section][data-qs-props-tab-context]");
  nodes.forEach((node) => {
    if (!(node instanceof HTMLDetailsElement)) {
      return;
    }
    const tab = trimOrFallback(node.dataset.qsPropsTabContext, "");
    const sectionKey = trimOrFallback(node.dataset.qsPropsSection, "");
    if ((tab !== "visual" && tab !== "interactions") || !sectionKey) {
      return;
    }
    if (!quickSightPropsSectionState[tab]) {
      quickSightPropsSectionState[tab] = Object.create(null);
    }
    quickSightPropsSectionState[tab][sectionKey] = node.open;
  });
}

function isQuickSightPropsSectionOpen(tab, sectionKey, defaultOpen = false) {
  if ((tab !== "visual" && tab !== "interactions") || !sectionKey) {
    return defaultOpen;
  }
  const byTab = quickSightPropsSectionState[tab];
  if (!byTab || typeof byTab[sectionKey] !== "boolean") {
    return defaultOpen;
  }
  return byTab[sectionKey];
}

function getQuickSightPropsSectionOpenAttr(tab, sectionKey, defaultOpen = false) {
  return isQuickSightPropsSectionOpen(tab, sectionKey, defaultOpen) ? " open" : "";
}

function buildQuickSightSpecificConfigFieldHtml(field, value) {
  const propToken = `type:${field.key}`;
  if (field.type === "checkbox") {
    return `<label class="qs-props-check"><input type="checkbox" data-qs-prop="${escapeAttribute(propToken)}"${value ? " checked" : ""}><span>${escapeHtml(field.label)}</span></label>`;
  }
  if (field.type === "select") {
    const options = (field.options || [])
      .map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === value ? " selected" : ""}>${escapeHtml(option.label)}</option>`)
      .join("");
    return `<label class="qs-props-row">${escapeHtml(field.label)}<select data-qs-prop="${escapeAttribute(propToken)}">${options}</select></label>`;
  }
  if (field.type === "color") {
    return `<label class="qs-props-row">${escapeHtml(field.label)}<input data-qs-prop="${escapeAttribute(propToken)}" type="color" value="${escapeAttribute(String(value || "#1f2f44"))}"></label>`;
  }
  if (field.type === "text") {
    const maxLengthAttr = Number.isInteger(field.maxLength) ? ` maxlength="${field.maxLength}"` : "";
    return `<label class="qs-props-row">${escapeHtml(field.label)}<input data-qs-prop="${escapeAttribute(propToken)}" type="text"${maxLengthAttr} value="${escapeAttribute(String(value || ""))}"></label>`;
  }
  const minAttr = Number.isFinite(field.min) ? ` min="${field.min}"` : "";
  const maxAttr = Number.isFinite(field.max) ? ` max="${field.max}"` : "";
  const stepAttr = Number.isFinite(field.step) ? ` step="${field.step}"` : "";
  return `<label class="qs-props-row">${escapeHtml(field.label)}<input data-qs-prop="${escapeAttribute(propToken)}" type="number"${minAttr}${maxAttr}${stepAttr} value="${escapeAttribute(String(value ?? ""))}"></label>`;
}

function bindQuickSightCanvasHover(canvas) {
  if (!(canvas instanceof HTMLCanvasElement)) {
    return;
  }
  if (canvas.__qsHoverBound === true) {
    return;
  }
  canvas.__qsHoverBound = true;

  const redrawHoverState = (state, highlightIndex) => {
    if (!state) {
      return;
    }
    state.model = drawBiWidgetChart(
      canvas,
      state.chartType,
      state.labels,
      state.values,
      highlightIndex,
      state.colors,
      state.visualSettings,
      state.labelOffsets,
      state.polarLayout,
      state.rows,
      state.chartConfig,
      state.labelLayoutV2,
      state.typeConfigMap
    );
  };
  let dragState = null;

  canvas.addEventListener("mousemove", (event) => {
    const state = canvas.__qsHoverState;
    if (!state || !Array.isArray(state.rows) || state.rows.length === 0) {
      canvas.style.cursor = "default";
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (dragState) {
      canvas.style.cursor = "grabbing";
      return;
    }
    const model = state.model || { type: "none", items: [] };
    let cursor = "default";
    const labelHoverIndex = resolveBiLabelDragIndex(model, x, y);
    if (labelHoverIndex >= 0) {
      cursor = "grab";
    } else {
      const hoverIndexProbe = resolveBiHoverIndex(model, x, y);
      if (hoverIndexProbe >= 0) {
        cursor = "pointer";
      }
    }
    canvas.style.cursor = cursor;
    const hoverIndex = resolveBiHoverIndex(model, x, y);
    if (hoverIndex === state.hoverIndex) {
      return;
    }
    state.hoverIndex = hoverIndex;
    redrawHoverState(state, hoverIndex);
  });

  canvas.addEventListener("pointerdown", (event) => {
    const state = canvas.__qsHoverState;
    if (!state || !state.visual || !Array.isArray(state.rows) || state.rows.length === 0) {
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const model = state.model || { type: "none", items: [] };
    const labelIndex = resolveBiLabelDragIndex(model, x, y);
    if (labelIndex < 0 || labelIndex >= state.rows.length) {
      return;
    }
    const labelItem = Array.isArray(model.labelItems)
      ? (model.labelItems.find((entry) => entry?.index === labelIndex) || null)
      : null;
    const row = state.rows[labelIndex] || null;
    const rowKey = trimOrFallback(row?.rowKey || labelItem?.rowKey, "");
    const rowLabel = trimOrFallback(row?.label || labelItem?.rowLabel || state.labels[labelIndex] || "", "");
    const currentPlacement = getBiLabelPlacement(state.visual, rowKey, rowLabel);
    const safeOffsets = normalizeBiLabelOffsets(state.labelOffsets);
    const currentLegacy = safeOffsets[String(labelIndex)] || { x: 0, y: 0 };
    const plotWidth = Math.max(1, Number.isFinite(labelItem?.plotWidth) ? labelItem.plotWidth : canvas.clientWidth || 1);
    const plotHeight = Math.max(1, Number.isFinite(labelItem?.plotHeight) ? labelItem.plotHeight : canvas.clientHeight || 1);
    dragState = {
      pointerId: event.pointerId,
      index: labelIndex,
      rowKey,
      rowLabel,
      plotWidth,
      plotHeight,
      startX: x,
      startY: y,
      startOffsetX: Number.isFinite(currentLegacy.x) ? currentLegacy.x : 0,
      startOffsetY: Number.isFinite(currentLegacy.y) ? currentLegacy.y : 0,
      startDxPct: Number.isFinite(currentPlacement?.dxPct)
        ? currentPlacement.dxPct
        : ((Number.isFinite(currentLegacy.x) ? currentLegacy.x : 0) / plotWidth),
      startDyPct: Number.isFinite(currentPlacement?.dyPct)
        ? currentPlacement.dyPct
        : ((Number.isFinite(currentLegacy.y) ? currentLegacy.y : 0) / plotHeight),
      moved: false
    };
    state.labelOffsets = safeOffsets;
    state.visual.labelOffsets = safeOffsets;
    state.hoverIndex = labelIndex;
    event.preventDefault();
    event.stopPropagation();
    canvas.setPointerCapture?.(event.pointerId);
    canvas.style.cursor = "grabbing";
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }
    const state = canvas.__qsHoverState;
    if (!state || !state.visual) {
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const nextX = dragState.startOffsetX + (x - dragState.startX);
    const nextY = dragState.startOffsetY + (y - dragState.startY);
    const safeOffsets = normalizeBiLabelOffsets(state.labelOffsets);
    safeOffsets[String(dragState.index)] = {
      x: Math.max(-400, Math.min(400, Math.round(nextX))),
      y: Math.max(-220, Math.min(220, Math.round(nextY)))
    };
    state.labelOffsets = safeOffsets;
    state.visual.labelOffsets = safeOffsets;
    upsertBiLabelPlacement(state.visual, dragState.rowKey, {
      anchorXPct: 0.5,
      anchorYPct: 0.5,
      dxPct: Math.max(-1.5, Math.min(1.5, dragState.startDxPct + ((x - dragState.startX) / Math.max(1, dragState.plotWidth)))),
      dyPct: Math.max(-1.5, Math.min(1.5, dragState.startDyPct + ((y - dragState.startY) / Math.max(1, dragState.plotHeight)))),
      scope: "plot"
    }, dragState.rowLabel);
    state.labelLayoutV2 = normalizeBiLabelLayoutV2(state.visual.labelLayoutV2);
    state.hoverIndex = dragState.index;
    redrawHoverState(state, dragState.index);
    dragState.moved = Math.abs(x - dragState.startX) > 1 || Math.abs(y - dragState.startY) > 1;
    event.preventDefault();
    event.stopPropagation();
    canvas.style.cursor = "grabbing";
  });

  const finishDrag = (event, canceled = false) => {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }
    const moved = dragState.moved;
    dragState = null;
    canvas.releasePointerCapture?.(event.pointerId);
    if (!canceled && moved) {
      quickSightSuppressClickUntil = Date.now() + 260;
      saveState();
      if (activeTab === "quicksight") {
        setStatus("Posicion de labels guardada.");
      }
    }
    canvas.style.cursor = "default";
    event.preventDefault();
    event.stopPropagation();
  };

  canvas.addEventListener("pointerup", (event) => {
    finishDrag(event, false);
  });

  canvas.addEventListener("pointercancel", (event) => {
    finishDrag(event, true);
  });

  canvas.addEventListener("click", (event) => {
    if (Date.now() < quickSightSuppressClickUntil || dragState) {
      return;
    }
    const state = canvas.__qsHoverState;
    if (!state || !state.visual || !Array.isArray(state.rows) || state.rows.length === 0) {
      return;
    }
    const interactions = normalizeQuickSightInteractionSettings(state.visual.interactions, state.visual);
    state.visual.interactions = interactions;
    if (interactions.clickAction === "none") {
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const model = state.model || { type: "none", items: [] };
    const labelIndex = resolveBiLabelDragIndex(model, x, y);
    if (labelIndex >= 0) {
      return;
    }
    const hoverIndex = resolveBiHoverIndex(model, x, y);
    if (hoverIndex < 0 || hoverIndex >= state.rows.length) {
      return;
    }
    const project = getActiveProject();
    if (!project) {
      return;
    }
    applyQuickSightInteractionFromVisual(project, state.visual, {
      rows: state.rows
    }, hoverIndex, event);
    event.preventDefault();
    event.stopPropagation();
  });

  canvas.addEventListener("mouseleave", () => {
    const state = canvas.__qsHoverState;
    canvas.style.cursor = "default";
    if (!state || state.hoverIndex < 0 || dragState) {
      return;
    }
    state.hoverIndex = -1;
    redrawHoverState(state, -1);
  });
}

function renderQuickSightPanel(project, options = {}) {
  if (!els.qsCanvasBoard || !els.qsFieldsList || !els.qsPropertiesBody) {
    return;
  }
  const softSelect = !!options.softSelect;
  const allowEmptySelection = !!options.allowEmptySelection;
  if (project && activeTab !== "quicksight") {
    renderQuickSightDrillThroughDrawer(project);
    return;
  }
  captureQuickSightPropsSectionState();
  applyQuickSightPanelVisibility("");
  if (!project) {
    quickSightRolePickerOpen = "";
    quickSightRolePickerFocusRole = "";
    quickSightExpandedVisualId = "";
    disposeQuickSightChartsInContainer(els.qsCanvasBoard);
    renderQuickSightDrillThroughDrawer(null);
    updateQuickSightZoomMenuUi(null);
    if (els.qsDataSummaryText) {
      els.qsDataSummaryText.textContent = "";
    }
    if (els.qsCanvasMetaText) {
      els.qsCanvasMetaText.textContent = "";
    }
    if (els.qsSheetTitle) {
      els.qsSheetTitle.textContent = "Hoja 1";
    }
    if (els.qsRoleDimensionList) {
      els.qsRoleDimensionList.innerHTML = "<span class=\"qs-role-chip qs-role-chip-empty\">Sin dimension</span>";
    }
    if (els.qsRoleMetricList) {
      els.qsRoleMetricList.innerHTML = "<span class=\"qs-role-chip qs-role-chip-empty\">Sin metrica</span>";
    }
    if (els.qsRoleOptionalMetricBlock) {
      els.qsRoleOptionalMetricBlock.hidden = true;
      els.qsRoleOptionalMetricBlock.classList.add("hidden");
    }
    if (els.qsRoleOptionalMetricList) {
      els.qsRoleOptionalMetricList.innerHTML = "";
    }
    if (els.qsOptionalMetricSelect) {
      els.qsOptionalMetricSelect.innerHTML = "<option value=\"\">(Automatico: UP Base total)</option>";
      els.qsOptionalMetricSelect.value = "";
    }
    if (els.qsRoleBreakdownList) {
      els.qsRoleBreakdownList.innerHTML = "<span class=\"qs-role-chip qs-role-chip-empty\">Sin grupo/color</span>";
    }
    els.qsFieldsList.innerHTML = "<div class=\"qs-empty\">Sin proyecto activo.</div>";
    els.qsCanvasBoard.innerHTML = "<div class=\"qs-empty\">Sin proyecto activo.</div>";
    els.qsPropertiesBody.innerHTML = "<p class=\"muted\">Sin proyecto activo.</p>";
    if (els.qsPresentationModeButton instanceof HTMLButtonElement) {
      els.qsPresentationModeButton.disabled = true;
    }
    syncQuickSightHistoryButtons(null);
    return;
  }

  ensureQuickSightState(project);
  if (els.qsPresentationModeButton instanceof HTMLButtonElement) {
    els.qsPresentationModeButton.disabled = false;
  }
  if (quickSightExpandedVisualId && !project.quickSightVisuals.some((item) => item.id === quickSightExpandedVisualId)) {
    quickSightExpandedVisualId = "";
  }
  if (!project.quickSightVisuals.some((item) => item.id === selectedQuickSightVisualId)) {
    selectedQuickSightVisualId = allowEmptySelection ? "" : (project.quickSightVisuals[0]?.id || "");
  }
  const activeVisual = project.quickSightVisuals.find((item) => item.id === selectedQuickSightVisualId) || null;

  const scopedRows = getQuickSightPanelBaseRows(project);
  const activeInteractionFilters = normalizeQuickSightInteractionFilters(project.quickSightConfig.interactionFilters);
  const selectedSource = normalizeBiSource(activeVisual?.source || project.quickSightConfig.source || "all");
  const sourceRows = selectedSource === "all"
    ? scopedRows
    : scopedRows.filter((row) => row.source === selectedSource);
  const activeDrillMeta = activeVisual ? getQuickSightVisualDrillMeta(activeVisual, project) : null;
  const groupOptions = getBiGroupByOptions(project);
  const metricOptions = getQuickSightMetricOptions();
  const chartOptions = getQuickSightChartTypeOptions();
  const safeGroupBy = normalizeBiGroupBy(activeVisual?.groupBy || project.quickSightConfig.groupBy || "disciplina");
  const safeBreakdown = activeVisual
    ? normalizeBiOptionalGroupBy(
      getOwnObjectValue(activeVisual, "breakdownDimension", activeVisual?.dataRoles?.breakdownDimension ?? "")
    )
    : normalizeBiOptionalGroupBy(project.quickSightConfig.breakdownDimension || "");
  const safeMetric = normalizeBiMetric(activeVisual?.metric || project.quickSightConfig.metric || "count");
  const safeOptionalMetric = activeVisual
    ? getQuickSightVisualOptionalMetric(activeVisual)
    : getQuickSightVisualOptionalMetric(null, project.quickSightConfig);
  const safeChartType = normalizeBiChartType(activeVisual?.chartType || project.quickSightConfig.chartType || "bar");
  const safeTopN = sanitizeBiTopN(activeVisual?.topN ?? project.quickSightConfig.topN ?? 12);
  const catalogFields = getBiCatalogFields(project, selectedSource);
  const breakdownOptions = [{ value: "", label: "(sin grupo/color)" }, ...groupOptions];
  const panelCapabilities = getQuickSightChartCapabilities(activeVisual ? {
    ...activeVisual,
    chartType: safeChartType
  } : { chartType: safeChartType }, project.quickSightConfig);
  if ((quickSightRolePickerOpen === "optionalMetric" && !panelCapabilities.supportsOptionalMetric)
    || (quickSightRolePickerOpen === "breakdown" && !panelCapabilities.supportsBreakdown)) {
    quickSightRolePickerOpen = "";
    quickSightRolePickerFocusRole = "";
  }

  if (els.qsSourceSelect) {
    els.qsSourceSelect.value = selectedSource;
  }
  if (els.qsGroupBySelect) {
    els.qsGroupBySelect.innerHTML = groupOptions
      .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
      .join("");
    const hasGroup = groupOptions.some((option) => option.value === safeGroupBy);
    els.qsGroupBySelect.value = hasGroup ? safeGroupBy : (groupOptions[0]?.value || "disciplina");
  }
  if (els.qsBreakdownSelect) {
    els.qsBreakdownSelect.innerHTML = breakdownOptions
      .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
      .join("");
    const hasBreakdown = breakdownOptions.some((option) => option.value === safeBreakdown);
    els.qsBreakdownSelect.value = hasBreakdown ? safeBreakdown : "";
  }
  if (els.qsMetricSelect) {
    els.qsMetricSelect.innerHTML = metricOptions
      .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
      .join("");
    els.qsMetricSelect.value = metricOptions.some((option) => option.value === safeMetric)
      ? safeMetric
      : "count";
  }
  if (els.qsOptionalMetricSelect) {
    const rows = [{ value: "", label: "Automatico (UP Base total)" }, ...metricOptions];
    els.qsOptionalMetricSelect.innerHTML = rows
      .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
      .join("");
    els.qsOptionalMetricSelect.value = rows.some((option) => option.value === safeOptionalMetric)
      ? safeOptionalMetric
      : "";
  }
  if (els.qsChartTypeSelect) {
    els.qsChartTypeSelect.innerHTML = chartOptions
      .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
      .join("");
    els.qsChartTypeSelect.value = chartOptions.some((option) => option.value === safeChartType)
      ? safeChartType
      : "bar";
  }
  const resolvedGroupBy = normalizeBiGroupBy(els.qsGroupBySelect?.value || safeGroupBy);
  const resolvedBreakdown = normalizeBiOptionalGroupBy(
    els.qsBreakdownSelect instanceof HTMLSelectElement
      ? els.qsBreakdownSelect.value
      : safeBreakdown
  );
  const resolvedMetric = normalizeBiMetric(els.qsMetricSelect?.value || safeMetric);
  const resolvedOptionalMetric = normalizeBiOptionalMetric(
    els.qsOptionalMetricSelect instanceof HTMLSelectElement
      ? els.qsOptionalMetricSelect.value
      : safeOptionalMetric
  );
  const resolvedChartType = normalizeBiChartType(els.qsChartTypeSelect?.value || safeChartType);
  const roleUiLabels = getQuickSightRoleUiLabels(resolvedChartType);
  const dimensionRoleLabelNode = els.qsRoleDimensionList?.previousElementSibling;
  if (dimensionRoleLabelNode instanceof HTMLElement) {
    dimensionRoleLabelNode.textContent = roleUiLabels.dimension;
  }
  const metricRoleLabelNode = els.qsRoleMetricList?.previousElementSibling;
  if (metricRoleLabelNode instanceof HTMLElement) {
    metricRoleLabelNode.textContent = roleUiLabels.metric;
  }
  const optionalMetricRoleLabelNode = els.qsRoleOptionalMetricList?.previousElementSibling;
  if (optionalMetricRoleLabelNode instanceof HTMLElement) {
    optionalMetricRoleLabelNode.textContent = roleUiLabels.optionalMetric;
  }
  const breakdownRoleLabelNode = els.qsRoleBreakdownList?.previousElementSibling;
  if (breakdownRoleLabelNode instanceof HTMLElement) {
    breakdownRoleLabelNode.textContent = roleUiLabels.breakdown;
  }
  const rolePickerOptionHtml = (options, selectedValue) => options
    .map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === selectedValue ? " selected" : ""}>${escapeHtml(option.label)}</option>`)
    .join("");
  const dimensionPickerHtml = `<select class="qs-role-inline-select" data-qs-role-picker-select="dimension">${rolePickerOptionHtml(groupOptions, resolvedGroupBy)}</select>`;
  const metricPickerHtml = `<select class="qs-role-inline-select" data-qs-role-picker-select="metric">${rolePickerOptionHtml(metricOptions, resolvedMetric)}</select>`;
  const optionalMetricPickerOptions = [{ value: "", label: "Automatico: UP Base total" }, ...metricOptions.filter((option) => option.value !== resolvedMetric)];
  const optionalMetricPickerHtml = panelCapabilities.supportsOptionalMetric
    ? `<select class="qs-role-inline-select" data-qs-role-picker-select="optionalMetric">${rolePickerOptionHtml(optionalMetricPickerOptions, resolvedOptionalMetric)}</select>`
    : "";
  const breakdownPickerOptions = breakdownOptions.filter((option) => !option.value || option.value !== resolvedGroupBy);
  const breakdownPickerHtml = panelCapabilities.supportsBreakdown
    ? `<select class="qs-role-inline-select" data-qs-role-picker-select="breakdown">${rolePickerOptionHtml(breakdownPickerOptions, resolvedBreakdown)}</select>`
    : "";
  if (els.qsChartTypeGrid) {
    els.qsChartTypeGrid.querySelectorAll("[data-qs-chart-type]").forEach((node) => {
      if (!(node instanceof HTMLElement)) {
        return;
      }
      const chartToken = normalizeBiChartType(node.dataset.qsChartType || "");
      node.classList.toggle("active", chartToken === resolvedChartType);
    });
  }
  if (els.qsRoleDimensionList) {
    els.qsRoleDimensionList.innerHTML = quickSightRolePickerOpen === "dimension"
      ? dimensionPickerHtml
      : (resolvedGroupBy
        ? `<button type="button" class="qs-role-chip qs-role-chip-dimension" data-qs-groupby="${escapeAttribute(resolvedGroupBy)}">${escapeHtml(getBiGroupLabel(resolvedGroupBy, project))}</button>`
        : "<span class=\"qs-role-chip qs-role-chip-empty\">Sin dimension</span>");
  }
  if (els.qsRoleMetricList) {
    els.qsRoleMetricList.innerHTML = quickSightRolePickerOpen === "metric"
      ? metricPickerHtml
      : (resolvedMetric
        ? `<button type="button" class="qs-role-chip qs-role-chip-metric" data-qs-metric="${escapeAttribute(resolvedMetric)}">${escapeHtml(getBiMetricLabel(resolvedMetric))}</button>`
        : "<span class=\"qs-role-chip qs-role-chip-empty\">Sin metrica</span>");
  }
  if (els.qsRoleOptionalMetricBlock) {
    els.qsRoleOptionalMetricBlock.hidden = !panelCapabilities.supportsOptionalMetric;
    els.qsRoleOptionalMetricBlock.classList.toggle("hidden", !panelCapabilities.supportsOptionalMetric);
  }
  if (els.qsRoleOptionalMetricList) {
    if (panelCapabilities.supportsOptionalMetric) {
      els.qsRoleOptionalMetricList.innerHTML = quickSightRolePickerOpen === "optionalMetric"
        ? optionalMetricPickerHtml
        : (resolvedOptionalMetric
          ? `<button type="button" class="qs-role-chip qs-role-chip-metric" data-qs-optional-metric="${escapeAttribute(resolvedOptionalMetric)}">${escapeHtml(getBiMetricLabel(resolvedOptionalMetric))}</button>`
          : "<span class=\"qs-role-chip qs-role-chip-empty\">Automatico: UP Base total</span>");
    } else {
      els.qsRoleOptionalMetricList.innerHTML = "";
    }
  }
  if (els.qsRoleBreakdownList) {
    els.qsRoleBreakdownList.innerHTML = quickSightRolePickerOpen === "breakdown"
      ? breakdownPickerHtml
      : (resolvedBreakdown
        ? `<button type="button" class="qs-role-chip qs-role-chip-breakdown" data-qs-breakdown="${escapeAttribute(resolvedBreakdown)}">${escapeHtml(getBiGroupLabel(resolvedBreakdown, project))}</button>`
        : "<span class=\"qs-role-chip qs-role-chip-empty\">Sin grupo/color</span>");
  }
  if (els.qsAddMetricGhostButton) {
    if (panelCapabilities.supportsOptionalMetric) {
      els.qsAddMetricGhostButton.disabled = false;
      els.qsAddMetricGhostButton.textContent = resolvedOptionalMetric
        ? "Cambiar medida eje X"
        : "Agregar medida eje X";
      els.qsAddMetricGhostButton.title = "Asigna o rota la metrica usada como eje X.";
    } else {
      els.qsAddMetricGhostButton.disabled = true;
      els.qsAddMetricGhostButton.textContent = "Una sola medida";
      els.qsAddMetricGhostButton.title = "La metrica adicional solo aplica en tipos como Scatter o Bubble.";
    }
  }
  if (els.qsAddBreakdownGhostButton) {
    const canPickBreakdown = panelCapabilities.supportsBreakdown && groupOptions.some((option) => option.value && option.value !== resolvedGroupBy);
    els.qsAddBreakdownGhostButton.disabled = !canPickBreakdown;
    if (!panelCapabilities.supportsBreakdown) {
      els.qsAddBreakdownGhostButton.textContent = "Sin grupo/color";
      els.qsAddBreakdownGhostButton.title = "Este tipo no usa una dimension secundaria.";
    } else {
      els.qsAddBreakdownGhostButton.textContent = resolvedBreakdown
        ? "Cambiar dimensión"
        : "Agregar una dimensión";
      els.qsAddBreakdownGhostButton.title = "Asigna o rota la dimension usada para GRUPO/COLOR.";
    }
  }
  if (els.qsDeleteVisualButton) {
    els.qsDeleteVisualButton.disabled = !activeVisual;
    els.qsDeleteVisualButton.title = activeVisual
      ? "Eliminar el visual seleccionado"
      : "Selecciona un visual para eliminarlo";
  }
  if (quickSightRolePickerFocusRole) {
    const roleToFocus = quickSightRolePickerFocusRole;
    quickSightRolePickerFocusRole = "";
    window.requestAnimationFrame(() => {
      const picker = document.querySelector(`[data-qs-role-picker-select="${roleToFocus}"]`);
      if (picker instanceof HTMLSelectElement) {
        picker.focus();
        if (typeof picker.showPicker === "function") {
          try {
            picker.showPicker();
          } catch (error) {
            // Fall back to regular focus if the browser blocks programmatic opening.
          }
        }
      }
    });
  }
  if (els.qsTopNInput) {
    els.qsTopNInput.value = String(safeTopN);
  }
  if (els.qsFieldsSearchInput) {
    els.qsFieldsSearchInput.value = trimOrFallback(project.quickSightConfig.fieldsSearch, "");
  }
  if (els.qsSheetTitle) {
    els.qsSheetTitle.textContent = `Hoja 1 | ${project.name}`;
  }
  const canvasSize = getQuickSightCanvasSizeFromConfig(project.quickSightConfig);
  const canvasZoom = sanitizeQuickSightCanvasZoom(project.quickSightConfig.canvasZoom, QUICKSIGHT_CANVAS_ZOOM_DEFAULT);
  const canvasZoomScale = Math.max(0.25, Math.min(3, canvasZoom / 100));
  const scaledCanvasWidth = Math.max(1, Math.round(canvasSize.width * canvasZoomScale));
  const scaledCanvasHeight = Math.max(1, Math.round(canvasSize.height * canvasZoomScale));
  syncQuickSightEditMenuInputs(project);
  updateQuickSightZoomMenuUi(project);
  if (els.qsCanvasMetaText) {
    const interactionText = activeInteractionFilters.length > 0
      ? ` | interacciones: ${activeInteractionFilters.length} filtro(s)`
      : "";
    const drillText = activeDrillMeta?.enabled
      ? ` | drill ${activeDrillMeta.currentLevel}/${activeDrillMeta.totalLevels}`
      : "";
    els.qsCanvasMetaText.textContent = `Visuales: ${project.quickSightVisuals.length} | Filas: ${sourceRows.length}${currentSearchQuery ? " | filtro global" : ""}${interactionText}${drillText} | Pizarra: ${canvasSize.width} x ${canvasSize.height}px | Zoom: ${canvasZoom}%`;
  }

  const searchToken = normalizeLookup(project.quickSightConfig.fieldsSearch || "");
  const filteredCatalogFields = searchToken
    ? catalogFields.filter((field) => normalizeLookup(`${field.source} ${field.name}`).includes(searchToken))
    : catalogFields;

  if (els.qsDataSummaryText) {
    els.qsDataSummaryText.textContent = `${getBiSourceLabel(selectedSource)} | Campos ${filteredCatalogFields.length}/${catalogFields.length}${currentSearchQuery ? " | busqueda global activa" : ""}`;
  }

  els.qsFieldsList.innerHTML = filteredCatalogFields
    .map((field) => {
      const groupBy = trimOrFallback(field.groupBy, "");
      const metric = trimOrFallback(field.metric, "");
      const token = getBiCatalogFieldTypeToken(field.type);
      const roleText = groupBy ? "Dimension" : (metric ? "Metrica" : "No vinculable");
      const activeClass = (groupBy && (groupBy === resolvedGroupBy || groupBy === resolvedBreakdown)) || (metric && metric === resolvedMetric)
        ? " active"
        : "";
      const draggableAttr = (groupBy || metric) ? " draggable=\"true\"" : "";
      return `<button type="button" class="qs-field-item${activeClass}" data-qs-groupby="${escapeAttribute(groupBy)}" data-qs-metric="${escapeAttribute(metric)}" data-qs-field-name="${escapeAttribute(field.name)}"${draggableAttr}>
        <span class="qs-field-kind">${escapeHtml(token)}</span>
        <span class="qs-field-name">${escapeHtml(field.name)}</span>
        <span class="qs-field-source">${escapeHtml(`${field.source} | ${roleText}`)}</span>
      </button>`;
    })
    .join("") || "<div class=\"qs-empty\">No hay campos que coincidan.</div>";

  project.quickSightVisuals.forEach((visual, index) => {
    visual.layout = clampQuickSightVisualLayoutToCanvas(
      normalizeQuickSightVisualLayout(visual.layout, index),
      project.quickSightConfig
    );
  });

  if (!softSelect) {
    disposeQuickSightChartsInContainer(els.qsCanvasBoard);
    if (project.quickSightVisuals.length === 0) {
      els.qsCanvasBoard.innerHTML = `<div class="qs-canvas-zoom-wrap" style="width:${scaledCanvasWidth}px;height:${scaledCanvasHeight}px;">
        <div class="qs-canvas-surface qs-canvas-surface-empty" data-qs-canvas-width="${canvasSize.width}" data-qs-canvas-height="${canvasSize.height}" data-qs-canvas-zoom="${canvasZoom}" style="width:${canvasSize.width}px;height:${canvasSize.height}px;transform:scale(${canvasZoomScale});transform-origin:top left;">
          <div class="qs-empty">AutoGraph: agrega un visual para empezar.</div>
        </div>
      </div>`;
    } else {
      const visibleVisuals = quickSightExpandedVisualId
        ? project.quickSightVisuals.filter((visual) => visual.id === quickSightExpandedVisualId)
        : project.quickSightVisuals.slice();
      const visualCardsHtml = visibleVisuals.map((visual, index) => {
        const activeClass = visual.id === selectedQuickSightVisualId ? "active" : "";
        const storedLayout = normalizeQuickSightVisualLayout(visual.layout, index);
        visual.layout = storedLayout;
        ensureQuickSightVisualDrillState(visual);
        const drillMeta = getQuickSightVisualDrillMeta(visual, project);
        const display = normalizeQuickSightDisplaySettings(visual.display);
        const visualSettings = normalizeBiVisualSettings(visual.visualSettings);
        visual.display = display;
        visual.visualSettings = visualSettings;
        visual.chartConfig = normalizeBiChartConfig(visual.chartConfig, visual.chartType);
        const typeMap = normalizeBiChartTypeConfigMap(visual.chartTypeConfig);
        const safeChartType = normalizeBiChartType(visual.chartType || "bar");
        if (!typeMap[safeChartType]) {
          typeMap[safeChartType] = normalizeBiChartTypeSpecificConfig({}, safeChartType);
        }
        visual.chartTypeConfig = typeMap;
        visual.labelLayoutV2 = normalizeBiLabelLayoutV2(visual.labelLayoutV2);
        visual.polarLayout = normalizeBiCircularLayout(visual.polarLayout);
        const currentGroupLabel = getBiGroupLabel(drillMeta.currentGroupBy, project);
        const drillTag = drillMeta.enabled ? ` | Drill ${drillMeta.currentLevel}/${drillMeta.totalLevels}` : "";
        const subtitle = trimOrFallback(display.subtitle, "")
          || `${getBiSourceLabel(visual.source)} | ${currentGroupLabel} | ${getBiMetricLabel(visual.metric)}${drillTag}`;
        const showHead = display.showTitle || display.showSubtitle;
        const isExpanded = visual.id === quickSightExpandedVisualId;
        const layout = isExpanded
          ? getQuickSightExpandedCanvasLayout(canvasSize)
          : storedLayout;
        const backgroundColor = display.showBackground
          ? hexToRgba(display.backgroundColor, Math.max(0, Math.min(1, display.backgroundOpacity / 100)), "rgba(255,255,255,1)")
          : "transparent";
        const borderColor = display.showBorder
          ? hexToRgba(display.borderColor, Math.max(0, Math.min(1, display.borderOpacity / 100)), "rgba(183,201,223,1)")
          : "transparent";
        const borderWidth = display.showBorder ? display.borderWidth : 0;
        const selectionColor = hexToRgba(
          display.selectionColor,
          Math.max(0.04, Math.min(1, display.selectionOpacity / 100)),
          "rgba(20,149,191,0.18)"
        );
        const cardClasses = [
          "qs-visual-card",
          activeClass ? "active" : "",
          isExpanded ? "is-expanded" : "",
          showHead ? "" : "no-head",
          display.loadAnimation ? "qs-load-anim" : ""
        ].filter(Boolean).join(" ");
        const sizeLabel = drillMeta.enabled
          ? `${layout.w} x ${layout.h} | Drill ${drillMeta.currentLevel}/${drillMeta.totalLevels}`
          : `${layout.w} x ${layout.h}`;
        const expandLabel = isExpanded ? "Minimizar visual" : "Maximizar visual";
        const expandIcon = isExpanded
          ? `<svg viewBox="0 0 16 16" class="qs-visual-icon-svg" aria-hidden="true"><path d="M2 6V2h4"></path><path d="M10 2h4v4"></path><path d="M14 10v4h-4"></path><path d="M6 14H2v-4"></path></svg>`
          : `<svg viewBox="0 0 16 16" class="qs-visual-icon-svg" aria-hidden="true"><path d="M6 2H2v4"></path><path d="M10 2h4v4"></path><path d="M14 10v4h-4"></path><path d="M6 14H2v-4"></path></svg>`;
        return `<article class="${cardClasses}" data-qs-visual-id="${escapeAttribute(visual.id)}" style="left:${layout.x}px;top:${layout.y}px;width:${layout.w}px;height:${layout.h}px;--qs-card-padding:${display.padding}px;--qs-card-bg:${backgroundColor};--qs-card-border-color:${borderColor};--qs-card-border-width:${borderWidth}px;--qs-card-select-color:${selectionColor};--qs-head-height:${showHead ? 39 : 0}px;--qs-title-font-family:'${escapeAttribute(display.titleFontFamily)}';--qs-title-font-size:${display.titleFontSize}px;--qs-title-color:${escapeAttribute(display.titleColor)};--qs-subtitle-font-family:'${escapeAttribute(display.subtitleFontFamily)}';--qs-subtitle-font-size:${display.subtitleFontSize}px;--qs-subtitle-color:${escapeAttribute(display.subtitleColor)};">
          <div class="qs-visual-head${showHead ? "" : " hidden"}" data-qs-drag-visual="${escapeAttribute(visual.id)}">
            <div>
              ${display.showTitle ? `<h4 class="qs-visual-title">${escapeHtml(visual.name)}</h4>` : ""}
              ${display.showSubtitle ? `<p class="qs-visual-meta">${escapeHtml(subtitle)}</p>` : ""}
            </div>
            <div class="qs-visual-head-right">
              <button type="button" class="qs-visual-icon-btn" data-qs-card-action="toggle-expand" data-qs-visual-id="${escapeAttribute(visual.id)}" title="${expandLabel}" aria-label="${expandLabel}">
                ${expandIcon}
              </button>
              <span class="qs-visual-size">${escapeHtml(sizeLabel)}</span>
            </div>
          </div>
          <div class="qs-visual-canvas-wrap">
            <div class="qs-visual-chart-host" data-qs-chart-host-id="${escapeAttribute(visual.id)}">
              <div class="qs-visual-echart-surface hidden" data-qs-echart-surface="${escapeAttribute(visual.id)}"></div>
              <canvas class="qs-visual-canvas" data-qs-canvas-id="${escapeAttribute(visual.id)}"></canvas>
            </div>
          </div>
          <div class="qs-visual-transform-box" aria-hidden="true">
            <span class="qs-transform-handle" data-qs-resize-handle="nw"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="n"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="ne"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="e"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="se"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="s"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="sw"></span>
            <span class="qs-transform-handle" data-qs-resize-handle="w"></span>
          </div>
        </article>`;
      }).join("");
      els.qsCanvasBoard.innerHTML = `<div class="qs-canvas-zoom-wrap" style="width:${scaledCanvasWidth}px;height:${scaledCanvasHeight}px;">
        <div class="qs-canvas-surface" data-qs-canvas-width="${canvasSize.width}" data-qs-canvas-height="${canvasSize.height}" data-qs-canvas-zoom="${canvasZoom}" style="width:${canvasSize.width}px;height:${canvasSize.height}px;transform:scale(${canvasZoomScale});transform-origin:top left;">${visualCardsHtml}</div>
      </div>`;
    }

    if (project.quickSightVisuals.length > 0) {
      const chartHosts = els.qsCanvasBoard.querySelectorAll("[data-qs-chart-host-id]");
      chartHosts.forEach((hostNode) => {
        if (!(hostNode instanceof HTMLElement)) {
          return;
        }
        const visualId = trimOrFallback(hostNode.dataset.qsChartHostId, "");
        const visual = project.quickSightVisuals.find((item) => item.id === visualId);
        if (!visual) {
          return;
        }
        const snapshot = buildQuickSightVisualSnapshot(project, visual, scopedRows);
        paintQuickSightVisualCanvas(hostNode, project, visual, snapshot);
      });
    }
  } else if (els.qsCanvasBoard instanceof HTMLElement) {
    const selectedId = trimOrFallback(selectedQuickSightVisualId, "");
    els.qsCanvasBoard.querySelectorAll(".qs-visual-card").forEach((node) => {
      if (!(node instanceof HTMLElement)) {
        return;
      }
      node.classList.toggle("active", trimOrFallback(node.dataset.qsVisualId, "") === selectedId);
    });
  }

  const selectedVisual = activeVisual;
  if (!selectedVisual) {
    els.qsPropertiesBody.innerHTML = `
      <div class="qs-props-tabs">
        <button type="button" class="qs-props-tab-btn${quickSightPropsTab === "visual" ? " active" : ""}" data-qs-props-tab="visual">Elemento visual</button>
        <button type="button" class="qs-props-tab-btn${quickSightPropsTab === "interactions" ? " active" : ""}" data-qs-props-tab="interactions">Interacciones</button>
      </div>
      <div class="qs-props-tab-panel${quickSightPropsTab === "interactions" ? " hidden" : ""}" data-qs-props-tab-panel="visual">
        <p class="muted">Selecciona un visual para editar propiedades.</p>
      </div>
      <div class="qs-props-tab-panel${quickSightPropsTab === "interactions" ? "" : " hidden"}" data-qs-props-tab-panel="interactions">
        <p class="muted">Selecciona un visual para configurar interacciones.</p>
        <p class="muted">${escapeHtml(getQuickSightInteractionSummary(project))}</p>
        <div class="inline-actions">
          <button type="button" class="secondary" data-qs-action="clear-interaction-filters"${activeInteractionFilters.length === 0 ? " disabled" : ""}>Limpiar filtros activos</button>
        </div>
      </div>
    `;
  } else {
    const selectedLayout = clampQuickSightVisualLayoutToCanvas(
      normalizeQuickSightVisualLayout(selectedVisual.layout, 0),
      project.quickSightConfig
    );
    selectedVisual.layout = selectedLayout;
    const selectedDisplay = normalizeQuickSightDisplaySettings({
      ...selectedVisual.display,
      subtitle: selectedVisual.display?.subtitle ?? selectedVisual.subtitle,
      altText: selectedVisual.display?.altText ?? selectedVisual.altText
    });
    selectedVisual.display = selectedDisplay;
    selectedVisual.subtitle = selectedDisplay.subtitle;
    selectedVisual.altText = selectedDisplay.altText;
    const selectedInteractions = normalizeQuickSightInteractionSettings(selectedVisual.interactions, selectedVisual);
    selectedVisual.interactions = selectedInteractions;
    ensureQuickSightVisualDrillState(selectedVisual);
    const selectedDrillMeta = getQuickSightVisualDrillMeta(selectedVisual, project);
    const selectedVisualSettings = normalizeBiVisualSettings(selectedVisual.visualSettings);
    selectedVisual.visualSettings = selectedVisualSettings;
    selectedVisual.chartConfig = normalizeBiChartConfig(selectedVisual.chartConfig, selectedVisual.chartType);
    const selectedChartType = normalizeBiChartType(selectedVisual.chartType || "bar");
    const selectedTypeMap = normalizeBiChartTypeConfigMap(selectedVisual.chartTypeConfig);
    if (!selectedTypeMap[selectedChartType]) {
      selectedTypeMap[selectedChartType] = normalizeBiChartTypeSpecificConfig({}, selectedChartType);
    }
    selectedVisual.chartTypeConfig = selectedTypeMap;
    const selectedTypeConfig = getBiSpecificChartConfigForWidget(selectedVisual, selectedChartType);
    const selectedCapabilities = getQuickSightChartCapabilities(selectedVisual, project.quickSightConfig);
    const selectedTypeSchema = getBiSpecificConfigSchema("chart", selectedChartType);
    const selectedRenderEngine = normalizeQuickSightRenderEngine(selectedVisual.renderEngine);
    const selectedRenderMeta = selectedCapabilities.renderMeta || getQuickSightRenderEngineMeta(selectedVisual);
    const selectedOptionalMetric = getQuickSightVisualOptionalMetric(selectedVisual, project.quickSightConfig);
    const fontFamilyOptions = Array.from(BI_ALLOWED_FONT_FAMILIES);
    const markerStyleOptions = [
      { value: "circle", label: "Circulo" },
      { value: "square", label: "Cuadrado" },
      { value: "diamond", label: "Rombo" },
      { value: "triangle", label: "Triangulo" },
      { value: "none", label: "Sin marcador" }
    ];
    const lineStyleOptions = [
      { value: "solid", label: "Solida" },
      { value: "dashed", label: "Discontinua" },
      { value: "dotted", label: "Punteada" }
    ];
    const isLineLike = new Set(["line", "timeseries", "area", "combo", "radar", "pareto"]).has(selectedChartType);
    const isMarkerLike = new Set(["line", "timeseries", "area", "combo", "scatter", "bubble", "radar", "pareto"]).has(selectedChartType);
    const isAreaLike = new Set(["area", "combo"]).has(selectedChartType);
    const isBarLike = BI_BAR_FAMILY_TYPES.has(selectedChartType) || new Set(["combo", "waterfall", "pareto", "bullet"]).has(selectedChartType);
    const showBreakdownInSeries = !!selectedCapabilities.supportsBreakdown;
    const interactionTargetGroupBy = getQuickSightInteractionTargetGroupBy(selectedVisual);
    const interactionTargetLabel = selectedInteractions.filterBy === "breakdown"
      ? getBiGroupLabel(normalizeBiOptionalGroupBy(selectedVisual.breakdownDimension || ""), project)
      : getBiGroupLabel(interactionTargetGroupBy, project);
    const drillLevelOptions = [{ value: "", label: "(sin siguiente nivel)" }, ...groupOptions];
    const drillLevelInputsHtml = Array.from({ length: QUICKSIGHT_DRILL_PATH_SLOTS }, (_, index) => {
      const value = selectedInteractions.drillPath[index] || "";
      return `<label class="qs-props-row">
        Nivel ${index + 2}
        <select data-qs-prop="interaction:drillPath${index}">
          ${drillLevelOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === value ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
        </select>
      </label>`;
    }).join("");
    const showSortInSeries = !!selectedCapabilities.supportsSortMode;
    const showReferenceLines = !!selectedCapabilities.supportsReferenceLine;
    const showDataLabelsSection = !!selectedCapabilities.supportsDataLabels;
    const specificFieldsHtml = Array.isArray(selectedTypeSchema.fields) && selectedTypeSchema.fields.length > 0
      ? selectedTypeSchema.fields.map((field) => buildQuickSightSpecificConfigFieldHtml(field, selectedTypeConfig[field.key])).join("")
      : "<p class=\"muted\">Sin configuración adicional para este tipo.</p>";

    els.qsPropertiesBody.innerHTML = `
      <div class="qs-props-tabs">
        <button type="button" class="qs-props-tab-btn${quickSightPropsTab === "visual" ? " active" : ""}" data-qs-props-tab="visual">Elemento visual</button>
        <button type="button" class="qs-props-tab-btn${quickSightPropsTab === "interactions" ? " active" : ""}" data-qs-props-tab="interactions">Interacciones</button>
      </div>

      <div class="qs-props-tab-panel${quickSightPropsTab === "interactions" ? " hidden" : ""}" data-qs-props-tab-panel="visual">
      <details class="qs-props-section" data-qs-props-section="display" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "display", true)}>
        <summary>Configuración de pantalla</summary>
        <div class="qs-props-section-body">
          <label class="qs-props-row">
            Editar titulo
            <input data-qs-prop="name" type="text" maxlength="120" value="${escapeAttribute(selectedVisual.name)}">
          </label>
          <label class="qs-props-row">
            Editar subtitulo
            <input data-qs-prop="display:subtitle" type="text" maxlength="140" value="${escapeAttribute(selectedDisplay.subtitle)}" placeholder="Subtitulo del grafico">
          </label>
          <label class="qs-props-row">
            Texto alternativo
            <textarea data-qs-prop="display:altText" maxlength="260" rows="3" placeholder="Descripcion accesible del visual">${escapeHtml(selectedDisplay.altText)}</textarea>
          </label>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Fuente titulo
              <select data-qs-prop="display:titleFontFamily">
                ${fontFamilyOptions.map((font) => `<option value="${escapeAttribute(font)}"${font === selectedDisplay.titleFontFamily ? " selected" : ""}>${escapeHtml(font)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Tamano titulo
              <input data-qs-prop="display:titleFontSize" type="number" min="10" max="24" step="1" value="${escapeAttribute(String(selectedDisplay.titleFontSize))}">
            </label>
            <label class="qs-props-row">
              Color titulo
              <input data-qs-prop="display:titleColor" type="color" value="${escapeAttribute(selectedDisplay.titleColor)}">
            </label>
            <label class="qs-props-row">
              Fuente subtitulo
              <select data-qs-prop="display:subtitleFontFamily">
                ${fontFamilyOptions.map((font) => `<option value="${escapeAttribute(font)}"${font === selectedDisplay.subtitleFontFamily ? " selected" : ""}>${escapeHtml(font)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Tamano subtitulo
              <input data-qs-prop="display:subtitleFontSize" type="number" min="8" max="18" step="1" value="${escapeAttribute(String(selectedDisplay.subtitleFontSize))}">
            </label>
            <label class="qs-props-row">
              Color subtitulo
              <input data-qs-prop="display:subtitleColor" type="color" value="${escapeAttribute(selectedDisplay.subtitleColor)}">
            </label>
          </div>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              X (px)
              <input data-qs-prop="layoutX" type="number" min="0" step="1" value="${escapeAttribute(String(selectedLayout.x))}">
            </label>
            <label class="qs-props-row">
              Y (px)
              <input data-qs-prop="layoutY" type="number" min="0" step="1" value="${escapeAttribute(String(selectedLayout.y))}">
            </label>
            <label class="qs-props-row">
              Ancho (px)
              <input data-qs-prop="layoutW" type="number" min="${escapeAttribute(String(QUICKSIGHT_VISUAL_MIN_WIDTH))}" step="1" value="${escapeAttribute(String(selectedLayout.w))}">
            </label>
            <label class="qs-props-row">
              Alto (px)
              <input data-qs-prop="layoutH" type="number" min="${escapeAttribute(String(QUICKSIGHT_VISUAL_MIN_HEIGHT))}" step="1" value="${escapeAttribute(String(selectedLayout.h))}">
            </label>
          </div>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Relleno (px)
              <input data-qs-prop="display:padding" type="number" min="0" max="24" step="1" value="${escapeAttribute(String(selectedDisplay.padding))}">
            </label>
            <label class="qs-props-row">
              Tipo de grafico
              <select data-qs-prop="chartType">
                ${chartOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === selectedChartType ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Motor render
              <select data-qs-prop="renderEngine">
                <option value="auto"${selectedRenderEngine === "auto" ? " selected" : ""}>Automatico</option>
                <option value="echarts"${selectedRenderEngine === "echarts" ? " selected" : ""}>ECharts</option>
                <option value="canvas"${selectedRenderEngine === "canvas" ? " selected" : ""}>Canvas clasico</option>
              </select>
            </label>
          </div>
          <p class="muted">Motor actual: ${escapeHtml(selectedRenderMeta.resolved === "echarts" ? "ECharts" : "Canvas")} | ${escapeHtml(selectedRenderMeta.note)}</p>
          <div class="qs-props-check-grid">
            <label class="qs-props-check"><input data-qs-prop="display:showTitle" type="checkbox"${selectedDisplay.showTitle ? " checked" : ""}><span>Mostrar titulo</span></label>
            <label class="qs-props-check"><input data-qs-prop="display:showSubtitle" type="checkbox"${selectedDisplay.showSubtitle ? " checked" : ""}><span>Mostrar subtitulo</span></label>
            <label class="qs-props-check"><input data-qs-prop="display:showBackground" type="checkbox"${selectedDisplay.showBackground ? " checked" : ""}><span>Mostrar fondo</span></label>
            <label class="qs-props-check"><input data-qs-prop="display:showBorder" type="checkbox"${selectedDisplay.showBorder ? " checked" : ""}><span>Mostrar borde</span></label>
            <label class="qs-props-check"><input data-qs-prop="display:loadAnimation" type="checkbox"${selectedDisplay.loadAnimation ? " checked" : ""}><span>Animacion de carga</span></label>
          </div>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Color fondo
              <input data-qs-prop="display:backgroundColor" type="color" value="${escapeAttribute(selectedDisplay.backgroundColor)}">
            </label>
            <label class="qs-props-row">
              Opacidad fondo (%)
              <input data-qs-prop="display:backgroundOpacity" type="number" min="0" max="100" step="1" value="${escapeAttribute(String(selectedDisplay.backgroundOpacity))}">
            </label>
            <label class="qs-props-row">
              Color borde
              <input data-qs-prop="display:borderColor" type="color" value="${escapeAttribute(selectedDisplay.borderColor)}">
            </label>
            <label class="qs-props-row">
              Opacidad borde (%)
              <input data-qs-prop="display:borderOpacity" type="number" min="0" max="100" step="1" value="${escapeAttribute(String(selectedDisplay.borderOpacity))}">
            </label>
            <label class="qs-props-row">
              Grosor borde (px)
              <input data-qs-prop="display:borderWidth" type="number" min="0" max="6" step="1" value="${escapeAttribute(String(selectedDisplay.borderWidth))}">
            </label>
            <label class="qs-props-row">
              Color seleccion
              <input data-qs-prop="display:selectionColor" type="color" value="${escapeAttribute(selectedDisplay.selectionColor)}">
            </label>
            <label class="qs-props-row">
              Opacidad seleccion (%)
              <input data-qs-prop="display:selectionOpacity" type="number" min="0" max="100" step="1" value="${escapeAttribute(String(selectedDisplay.selectionOpacity))}">
            </label>
          </div>
        </div>
      </details>

      <details class="qs-props-section" data-qs-props-section="series" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "series", true)}>
        <summary>Serie de datos</summary>
        <div class="qs-props-section-body">
          <div class="qs-props-layout-grid">
            ${showBreakdownInSeries ? `<label class="qs-props-row">
              Grupo/Color
              <select data-qs-prop="breakdownDimension">
                ${breakdownOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === resolvedBreakdown ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>` : ""}
            ${selectedCapabilities.supportsOptionalMetric ? `<label class="qs-props-row">
              Metrica eje X
              <select data-qs-prop="optionalMetric">
                <option value="">Automatico (UP Base total)</option>
                ${metricOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === selectedOptionalMetric ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>` : ""}
            <label class="qs-props-row">
              Top N
              <input data-qs-prop="topN" type="number" min="1" max="60" step="1" value="${escapeAttribute(String(selectedVisual.topN))}">
            </label>
            ${showSortInSeries ? `<label class="qs-props-row">
              Orden
              <select data-qs-prop="sortMode">
                <option value="value_desc"${selectedVisual.sortMode === "value_desc" ? " selected" : ""}>Valor: mayor a menor</option>
                <option value="value_asc"${selectedVisual.sortMode === "value_asc" ? " selected" : ""}>Valor: menor a mayor</option>
                <option value="label_asc"${selectedVisual.sortMode === "label_asc" ? " selected" : ""}>Etiqueta: A-Z</option>
                <option value="label_desc"${selectedVisual.sortMode === "label_desc" ? " selected" : ""}>Etiqueta: Z-A</option>
              </select>
            </label>` : ""}
          </div>
          ${isLineLike ? `<div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Grosor de linea
              <input data-qs-prop="visual:lineWidth" type="number" min="1" max="6" step="1" value="${escapeAttribute(String(selectedVisualSettings.lineWidth))}">
            </label>
            <label class="qs-props-row">
              Estilo de linea
              <select data-qs-prop="visual:seriesLineStyle">
                ${lineStyleOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === selectedVisualSettings.seriesLineStyle ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>
          </div>` : ""}
          ${isMarkerLike ? `<div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Marcador
              <select data-qs-prop="visual:seriesMarkerStyle">
                ${markerStyleOptions.map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === selectedVisualSettings.seriesMarkerStyle ? " selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Tamano marcador
              <input data-qs-prop="visual:seriesMarkerSize" type="number" min="0" max="16" step="1" value="${escapeAttribute(String(selectedVisualSettings.seriesMarkerSize))}">
            </label>
          </div>` : ""}
          ${isAreaLike ? `<label class="qs-props-row">Opacidad de area<input data-qs-prop="visual:areaOpacity" type="number" min="0.05" max="0.8" step="0.01" value="${escapeAttribute(String(selectedVisualSettings.areaOpacity))}"></label>` : ""}
          ${isBarLike ? `<label class="qs-props-row">Ancho relativo de barra<input data-qs-prop="visual:barWidthRatio" type="number" min="0.2" max="0.9" step="0.01" value="${escapeAttribute(String(selectedVisualSettings.barWidthRatio))}"></label>` : ""}
          ${selectedCapabilities.supportsStack ? `<label class="qs-props-row">Stacking<select data-qs-prop="visual:stackMode"><option value="none"${selectedVisualSettings.stackMode === "none" ? " selected" : ""}>Ninguno</option><option value="normal"${selectedVisualSettings.stackMode === "normal" ? " selected" : ""}>Normal</option><option value="percent"${selectedVisualSettings.stackMode === "percent" ? " selected" : ""}>100%</option></select></label>` : ""}
          ${isLineLike ? `<label class="qs-props-check"><input data-qs-prop="visual:smoothLines" type="checkbox"${selectedVisualSettings.smoothLines ? " checked" : ""}><span>Suavizar lineas</span></label>` : ""}
        </div>
        ${selectedCapabilities.supportsOptionalMetric ? `<p class="muted">Si queda vacio, QuickSight usa UP Base total como eje X para este tipo.</p>` : ""}
      </details>

      ${selectedCapabilities.supportsLegend ? `<details class="qs-props-section" data-qs-props-section="legend" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "legend", false)}>
        <summary>
          <span class="qs-props-summary-title">Leyenda</span>
          <label class="qs-props-summary-switch" data-qs-stop-summary>
            <input data-qs-prop="visual:showLegend" type="checkbox"${selectedVisualSettings.showLegend ? " checked" : ""}>
            <span class="qs-props-switch-track"><span class="qs-props-switch-thumb"></span></span>
          </label>
        </summary>
        <div class="qs-props-section-body">
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Posicion
              <select data-qs-prop="visual:legendPosition">
                <option value="right"${selectedVisualSettings.legendPosition === "right" ? " selected" : ""}>Derecha</option>
                <option value="left"${selectedVisualSettings.legendPosition === "left" ? " selected" : ""}>Izquierda</option>
                <option value="top"${selectedVisualSettings.legendPosition === "top" ? " selected" : ""}>Arriba</option>
                <option value="bottom"${selectedVisualSettings.legendPosition === "bottom" ? " selected" : ""}>Abajo</option>
              </select>
            </label>
            <label class="qs-props-row">
              Maximo items
              <input data-qs-prop="visual:legendMaxItems" type="number" min="3" max="20" step="1" value="${escapeAttribute(String(selectedVisualSettings.legendMaxItems))}">
            </label>
          </div>
        </div>
      </details>` : ""}

      ${selectedCapabilities.usesAxes ? `<details class="qs-props-section" data-qs-props-section="axisY" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "axisY", false)}>
        <summary>Eje Y</summary>
        <div class="qs-props-section-body">
          <label class="qs-props-check"><input data-qs-prop="visual:showAxisLabels" type="checkbox"${selectedVisualSettings.showAxisLabels ? " checked" : ""}><span>Mostrar ejes y etiquetas</span></label>
          <label class="qs-props-row">
            Titulo eje Y
            <input data-qs-prop="visual:axisYLabel" type="text" maxlength="40" value="${escapeAttribute(selectedVisualSettings.axisYLabel)}">
          </label>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Escala
              <select data-qs-prop="visual:axisScale">
                <option value="linear"${selectedVisualSettings.axisScale === "linear" ? " selected" : ""}>Lineal</option>
                <option value="log"${selectedVisualSettings.axisScale === "log" ? " selected" : ""}>Logaritmica</option>
              </select>
            </label>
            <label class="qs-props-row">
              Min eje numerico
              <input data-qs-prop="visual:axisMin" type="number" step="0.1" placeholder="Automatico" value="${escapeAttribute(selectedVisualSettings.axisMin ?? "")}">
            </label>
            <label class="qs-props-row">
              Max eje numerico
              <input data-qs-prop="visual:axisMax" type="number" step="0.1" placeholder="Automatico" value="${escapeAttribute(selectedVisualSettings.axisMax ?? "")}">
            </label>
            <label class="qs-props-row">
              Fuente eje
              <select data-qs-prop="visual:fontFamilyAxis">
                ${fontFamilyOptions.map((font) => `<option value="${escapeAttribute(font)}"${font === selectedVisualSettings.fontFamilyAxis ? " selected" : ""}>${escapeHtml(font)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Tamano texto eje
              <input data-qs-prop="visual:fontSizeAxis" type="number" min="8" max="20" step="1" value="${escapeAttribute(String(selectedVisualSettings.fontSizeAxis))}">
            </label>
          </div>
        </div>
      </details>
      <details class="qs-props-section" data-qs-props-section="axisX" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "axisX", false)}>
        <summary>Eje X</summary>
        <div class="qs-props-section-body">
          <label class="qs-props-row">
            Titulo eje X
            <input data-qs-prop="visual:axisXLabel" type="text" maxlength="40" value="${escapeAttribute(selectedVisualSettings.axisXLabel)}">
          </label>
          ${selectedCapabilities.supportsGrid ? `<label class="qs-props-check"><input data-qs-prop="visual:showGrid" type="checkbox"${selectedVisualSettings.showGrid ? " checked" : ""}><span>Lineas de cuadricula</span></label>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Opacidad cuadricula
              <input data-qs-prop="visual:gridOpacity" type="number" min="0.1" max="1" step="0.05" value="${escapeAttribute(String(selectedVisualSettings.gridOpacity))}">
            </label>
            <label class="qs-props-row">
              Patron cuadricula
              <input data-qs-prop="visual:gridDash" type="number" min="0" max="12" step="1" value="${escapeAttribute(String(selectedVisualSettings.gridDash))}">
            </label>
          </div>` : ""}
        </div>
      </details>` : ""}

      ${showReferenceLines ? `<details class="qs-props-section" data-qs-props-section="reference" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "reference", false)}>
        <summary>Lineas de referencia</summary>
        <div class="qs-props-section-body">
          <label class="qs-props-check"><input data-qs-prop="visual:showTargetLine" type="checkbox"${selectedVisualSettings.showTargetLine ? " checked" : ""}><span>Mostrar linea de referencia</span></label>
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Valor
              <input data-qs-prop="visual:targetLineValue" type="number" step="0.1" value="${escapeAttribute(selectedVisualSettings.targetLineValue ?? "")}">
            </label>
            <label class="qs-props-row">
              Etiqueta
              <input data-qs-prop="visual:targetLineLabel" type="text" maxlength="24" value="${escapeAttribute(selectedVisualSettings.targetLineLabel)}">
            </label>
            <label class="qs-props-row">
              Color
              <input data-qs-prop="visual:targetLineColor" type="color" value="${escapeAttribute(selectedVisualSettings.targetLineColor)}">
            </label>
          </div>
        </div>
      </details>` : ""}

      ${showDataLabelsSection ? `<details class="qs-props-section" data-qs-props-section="labels" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "labels", false)}>
        <summary>
          <span class="qs-props-summary-title">Etiquetas de datos</span>
          <label class="qs-props-summary-switch" data-qs-stop-summary>
            <input data-qs-prop="visual:showDataLabels" type="checkbox"${selectedVisualSettings.showDataLabels ? " checked" : ""}>
            <span class="qs-props-switch-track"><span class="qs-props-switch-thumb"></span></span>
          </label>
        </summary>
        <div class="qs-props-section-body">
          <div class="qs-props-layout-grid">
            <label class="qs-props-row">
              Maximo caracteres
              <input data-qs-prop="visual:labelMaxChars" type="number" min="4" max="40" step="1" value="${escapeAttribute(String(selectedVisualSettings.labelMaxChars))}">
            </label>
            <label class="qs-props-row">
              Decimales
              <input data-qs-prop="visual:valueDecimals" type="number" min="0" max="4" step="1" value="${escapeAttribute(String(selectedVisualSettings.valueDecimals))}">
            </label>
            <label class="qs-props-row">
              Formato numero
              <select data-qs-prop="visual:numberFormat">
                <option value="auto"${selectedVisualSettings.numberFormat === "auto" ? " selected" : ""}>Automatico</option>
                <option value="number"${selectedVisualSettings.numberFormat === "number" ? " selected" : ""}>Numero</option>
                <option value="percent"${selectedVisualSettings.numberFormat === "percent" ? " selected" : ""}>Porcentaje</option>
                <option value="currency_pen"${selectedVisualSettings.numberFormat === "currency_pen" ? " selected" : ""}>Moneda PEN</option>
                <option value="hours"${selectedVisualSettings.numberFormat === "hours" ? " selected" : ""}>Horas</option>
              </select>
            </label>
            <label class="qs-props-row">
              Prefijo
              <input data-qs-prop="visual:valuePrefix" type="text" maxlength="8" value="${escapeAttribute(selectedVisualSettings.valuePrefix)}">
            </label>
            <label class="qs-props-row">
              Sufijo
              <input data-qs-prop="visual:valueSuffix" type="text" maxlength="8" value="${escapeAttribute(selectedVisualSettings.valueSuffix)}">
            </label>
            <label class="qs-props-row">
              Fuente labels
              <select data-qs-prop="visual:fontFamilyLabel">
                ${fontFamilyOptions.map((font) => `<option value="${escapeAttribute(font)}"${font === selectedVisualSettings.fontFamilyLabel ? " selected" : ""}>${escapeHtml(font)}</option>`).join("")}
              </select>
            </label>
            <label class="qs-props-row">
              Tamano labels
              <input data-qs-prop="visual:fontSizeLabels" type="number" min="8" max="22" step="1" value="${escapeAttribute(String(selectedVisualSettings.fontSizeLabels))}">
            </label>
            <label class="qs-props-row">
              Color labels
              <select data-qs-prop="visual:labelColorMode">
                <option value="auto"${selectedVisualSettings.labelColorMode === "auto" ? " selected" : ""}>Automatico</option>
                <option value="manual"${selectedVisualSettings.labelColorMode === "manual" ? " selected" : ""}>Manual</option>
              </select>
            </label>
            <label class="qs-props-row">
              Color manual
              <input data-qs-prop="visual:labelColor" type="color" value="${escapeAttribute(selectedVisualSettings.labelColor)}"${selectedVisualSettings.labelColorMode === "manual" ? "" : " disabled"}>
            </label>
            <label class="qs-props-row">
              Tooltip
              <select data-qs-prop="visual:tooltipMode">
                <option value="full"${selectedVisualSettings.tooltipMode === "full" ? " selected" : ""}>Completo</option>
                <option value="compact"${selectedVisualSettings.tooltipMode === "compact" ? " selected" : ""}>Compacto</option>
              </select>
            </label>
          </div>
        </div>
      </details>` : ""}

      <details class="qs-props-section" data-qs-props-section="specific" data-qs-props-tab-context="visual"${getQuickSightPropsSectionOpenAttr("visual", "specific", false)}>
        <summary>Configuración específica (${escapeHtml(selectedTypeSchema.title)})</summary>
        <div class="qs-props-section-body">
          ${specificFieldsHtml}
        </div>
      </details>

      </div>

      <div class="qs-props-tab-panel${quickSightPropsTab === "interactions" ? "" : " hidden"}" data-qs-props-tab-panel="interactions">
        <details class="qs-props-section" data-qs-props-section="interactions" data-qs-props-tab-context="interactions"${getQuickSightPropsSectionOpenAttr("interactions", "interactions", true)}>
          <summary>Interacciones</summary>
          <div class="qs-props-section-body">
            <p class="muted">${selectedInteractions.clickAction === "drill_down"
              ? "Haz click sobre un dato del visual para bajar al siguiente nivel de la jerarquia."
              : (selectedInteractions.clickAction === "drill_through"
                ? "Haz click sobre un dato del visual para abrir el detalle real de las filas subyacentes."
              : (selectedInteractions.clickAction === "cross_filter"
                ? "Haz click sobre un dato del visual para filtrar el resto de la hoja."
                : "El click del visual no ejecuta ninguna accion."))}</p>
            <div class="qs-props-layout-grid">
              <label class="qs-props-row">
                Accion al hacer click
                <select data-qs-prop="interaction:clickAction">
                  <option value="drill_down"${selectedInteractions.clickAction === "drill_down" ? " selected" : ""}>Drill-down</option>
                  <option value="drill_through"${selectedInteractions.clickAction === "drill_through" ? " selected" : ""}>Drill-through</option>
                  <option value="cross_filter"${selectedInteractions.clickAction === "cross_filter" ? " selected" : ""}>Aplicar filtro cruzado</option>
                  <option value="none"${selectedInteractions.clickAction === "none" ? " selected" : ""}>Sin accion</option>
                </select>
              </label>
              ${new Set(["cross_filter", "drill_through"]).has(selectedInteractions.clickAction) ? `<label class="qs-props-row">
                ${selectedInteractions.clickAction === "drill_through" ? "Campo a detallar" : "Campo a filtrar"}
                <select data-qs-prop="interaction:filterBy">
                  <option value="group"${selectedInteractions.filterBy === "group" ? " selected" : ""}>Dimension principal</option>
                  <option value="breakdown"${selectedInteractions.filterBy === "breakdown" ? " selected" : ""}${normalizeBiOptionalGroupBy(selectedVisual.breakdownDimension || "") ? "" : " disabled"}>Grupo / color</option>
                </select>
              </label>
              ${selectedInteractions.clickAction === "cross_filter" ? `<label class="qs-props-row">
                Alcance
                <select data-qs-prop="interaction:scope">
                  <option value="all"${selectedInteractions.scope === "all" ? " selected" : ""}>Toda la hoja</option>
                  <option value="same_source"${selectedInteractions.scope === "same_source" ? " selected" : ""}>Solo misma fuente</option>
                </select>
              </label>
              <label class="qs-props-check">
                <input data-qs-prop="interaction:appendWithModifier" type="checkbox"${selectedInteractions.appendWithModifier ? " checked" : ""}>
                <span>Permitir multi-filtro con Ctrl/Cmd</span>
              </label>` : ""}`
                : (selectedInteractions.clickAction === "drill_down" ? drillLevelInputsHtml : "")}
            </div>
            ${new Set(["cross_filter", "drill_through"]).has(selectedInteractions.clickAction)
              ? `<p class="muted">Campo objetivo actual: ${escapeHtml(interactionTargetLabel)}</p>`
              : ""}
            ${selectedInteractions.clickAction === "drill_down"
              ? `<p class="muted">${selectedDrillMeta.sequence.length > 1
                ? `Jerarquia configurada: ${escapeHtml(selectedDrillMeta.sequenceText)}`
                : "Configura al menos un siguiente nivel para habilitar el drill-down."}</p>
                 <p class="muted">Nivel actual: ${escapeHtml(getBiGroupLabel(selectedDrillMeta.currentGroupBy, project))}</p>
                 <p class="muted">${escapeHtml(selectedDrillMeta.breadcrumbText)}</p>
                 <div class="inline-actions">
                   <button type="button" class="secondary" data-qs-action="drill-up"${selectedDrillMeta.canDrillUp ? "" : " disabled"}>Subir nivel</button>
                   <button type="button" class="secondary" data-qs-action="drill-reset"${selectedDrillMeta.canDrillUp ? "" : " disabled"}>Reiniciar drill</button>
                 </div>`
              : ""}
            <p class="muted">${escapeHtml(getQuickSightInteractionSummary(project))}</p>
            <div class="inline-actions">
              <button type="button" class="secondary" data-qs-action="clear-interaction-filters"${activeInteractionFilters.length === 0 ? " disabled" : ""}>Limpiar filtros activos</button>
            </div>
          </div>
        </details>
      </div>
    `;
  }
  syncQuickSightHistoryState(project);
  syncQuickSightHistoryButtons(project);
  renderQuickSightDrillThroughDrawer(project);
}

function paintQuickSightVisualCanvas(chartHost, project, visual, snapshot) {
  if (!(chartHost instanceof HTMLElement)) {
    return;
  }
  const canvas = chartHost.querySelector("canvas[data-qs-canvas-id]");
  const echartSurface = chartHost.querySelector("[data-qs-echart-surface]");
  if (!(canvas instanceof HTMLCanvasElement)) {
    return;
  }
  const chartType = normalizeBiChartType(visual.chartType || "bar");
  const visualSettings = normalizeBiVisualSettings(visual.visualSettings);
  visual.visualSettings = visualSettings;
  const chartConfig = normalizeBiChartConfig(visual.chartConfig, chartType);
  visual.chartConfig = chartConfig;
  const typeConfigMap = normalizeBiChartTypeConfigMap(visual.chartTypeConfig);
  if (!typeConfigMap[chartType]) {
    typeConfigMap[chartType] = normalizeBiChartTypeSpecificConfig({}, chartType);
  }
  visual.chartTypeConfig = typeConfigMap;
  const safeLabelLayout = normalizeBiLabelLayoutV2(visual.labelLayoutV2);
  const safeLabelOffsets = normalizeBiLabelOffsets(visual.labelOffsets);
  const safePolarLayout = normalizeBiCircularLayout(visual.polarLayout);
  visual.labelLayoutV2 = safeLabelLayout;
  visual.labelOffsets = safeLabelOffsets;
  visual.polarLayout = safePolarLayout;
  const width = Math.max(220, Math.round(canvas.clientWidth || chartHost.clientWidth || 420));
  const height = Math.max(180, Math.round(canvas.clientHeight || chartHost.clientHeight || 260));
  const dpr = Math.max(1, Math.min(window.devicePixelRatio || 1, 2));
  canvas.width = Math.round(width * dpr);
  canvas.height = Math.round(height * dpr);
  canvas.dataset.biRenderWidth = String(width);
  canvas.dataset.biRenderHeight = String(height);
  canvas.dataset.biDpr = String(dpr);
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  if (!snapshot || !Array.isArray(snapshot.rows) || snapshot.rows.length === 0) {
    disposeQuickSightChartSurface(echartSurface);
    if (echartSurface instanceof HTMLElement) {
      echartSurface.classList.add("hidden");
    }
    chartHost.dataset.qsRenderEngine = "canvas";
    canvas.classList.remove("hidden");
    canvas.style.display = "";
    canvas.__qsHoverState = null;
    canvas.style.cursor = "default";
    drawBiWidgetFallbackPreview(ctx, snapshot || null, width, height, true);
    return;
  }

  const seriesColors = getBiSnapshotSeriesColors(project, snapshot);
  const labels = snapshot.rows.map((row) => row.label);
  const values = snapshot.rows.map((row) => row.value);
  if (echartSurface instanceof HTMLElement && shouldUseQuickSightECharts(visual, snapshot)) {
    const adapter = getQuickSightEChartsAdapter();
    const chartTypeConfig = normalizeBiChartTypeSpecificConfig(
      typeConfigMap[chartType],
      chartType
    );
    echartSurface.classList.remove("hidden");
    const renderResult = adapter?.renderQuickSightChart?.(echartSurface, {
      chartType,
      rows: snapshot.rows,
      colors: seriesColors,
      visualSettings,
      chartConfig,
      typeConfig: chartTypeConfig,
      metric: visual.metric || visual.dataRoles?.metrics?.[0] || "count",
      metricLabel: getBiMetricLabel(visual.metric || visual.dataRoles?.metrics?.[0] || "count")
    });
    if (renderResult?.rendered && renderResult.chart) {
      chartHost.dataset.qsRenderEngine = "echarts";
      canvas.classList.add("hidden");
      canvas.style.display = "none";
      canvas.__qsHoverState = {
        visual,
        rows: snapshot.rows,
        chartType,
        model: { type: "none", items: [] },
        hoverIndex: -1,
        echartsManaged: true
      };
      bindQuickSightEChartsInteractions(echartSurface, renderResult.chart, project, visual, snapshot);
      return;
    }
  }
  disposeQuickSightChartSurface(echartSurface);
  if (echartSurface instanceof HTMLElement) {
    echartSurface.classList.add("hidden");
  }
  chartHost.dataset.qsRenderEngine = "canvas";
  canvas.classList.remove("hidden");
  canvas.style.display = "";
  const initialModel = drawBiWidgetChart(
    canvas,
    chartType,
    labels,
    values,
    -1,
    seriesColors,
    visualSettings,
    safeLabelOffsets,
    safePolarLayout,
    snapshot.rows,
    chartConfig,
    safeLabelLayout,
    typeConfigMap
  );
  canvas.__qsHoverState = {
    visual,
    chartType,
    labels,
    values,
    rows: snapshot.rows,
    colors: seriesColors,
    visualSettings,
    labelOffsets: safeLabelOffsets,
    polarLayout: safePolarLayout,
    chartConfig,
    labelLayoutV2: safeLabelLayout,
    typeConfigMap,
    model: initialModel,
    hoverIndex: -1
  };
  bindQuickSightCanvasHover(canvas);
}

function syncBiInputs(config) {
  if (!config) {
    return;
  }
  if (els.biStartDateInput) {
    els.biStartDateInput.value = sanitizeDateInput(config.startDate || "");
  }
  if (els.biEndDateInput) {
    els.biEndDateInput.value = sanitizeDateInput(config.endDate || "");
  }
  if (els.biTextFilterInput) {
    els.biTextFilterInput.value = trimOrFallback(config.textFilter, "");
  }
  if (els.biGlobalSourceSelect) {
    els.biGlobalSourceSelect.value = normalizeBiSource(config.source || "all");
  }
  if (els.biCrossFilterScopeSelect) {
    els.biCrossFilterScopeSelect.value = normalizeBiCrossFilterScope(config.crossFilterScope || "all");
  }
  if (els.biInvalidOnlyCheckbox) {
    els.biInvalidOnlyCheckbox.checked = !!config.invalidOnly;
  }
  if (els.biUiModeSelect) {
    els.biUiModeSelect.value = normalizeBiUiMode(config.uiMode || "basic");
  }
  const canvasSize = getBiCanvasSizeFromConfig(config);
  if (els.biCanvasWidthInput) {
    els.biCanvasWidthInput.value = String(canvasSize.width);
  }
  if (els.biCanvasHeightInput) {
    els.biCanvasHeightInput.value = String(canvasSize.height);
  }
  if (els.biCanvasPresetSelect) {
    const presetValue = `${canvasSize.width}x${canvasSize.height}`;
    const hasPreset = Array.from(els.biCanvasPresetSelect.options).some((option) => option.value === presetValue);
    els.biCanvasPresetSelect.value = hasPreset ? presetValue : "";
  }
  if (els.biCanvasZoomSelect) {
    const zoom = sanitizeBiCanvasZoom(config.canvasZoom, BI_CANVAS_ZOOM_DEFAULT);
    syncBiZoomSelectControl(els.biCanvasZoomSelect, zoom);
    updateBiZoomMenuUi(getActiveProject());
  }
  if (els.biShowCanvasGridCheckbox) {
    els.biShowCanvasGridCheckbox.checked = normalizeBiToggle(config.showCanvasGrid, false);
  }
  if (els.biSnapToGridCheckbox) {
    els.biSnapToGridCheckbox.checked = normalizeBiToggle(config.snapToGrid, true);
  }
  if (els.biGridSnapSizeInput) {
    els.biGridSnapSizeInput.value = String(sanitizeBiInteger(config.gridSnapSize, 12, 4, 120));
  }
  if (els.biPerformanceModeSelect) {
    els.biPerformanceModeSelect.value = normalizeBiPerformanceMode(config.performanceMode || "balanced");
  }
  if (els.biColorSourceSelect) {
    els.biColorSourceSelect.value = normalizeBiSource(config.colorSource || "all");
  }
  if (els.biColorGroupBySelect) {
    const groupBy = normalizeBiGroupBy(config.colorGroupBy || "disciplina");
    const hasOption = Array.from(els.biColorGroupBySelect.options).some((option) => option.value === groupBy);
    els.biColorGroupBySelect.value = hasOption ? groupBy : (els.biColorGroupBySelect.options[0]?.value || "disciplina");
  }
  const project = getActiveProject();
  const target = project ? resolveBiVisualTarget(project) : { mode: "global", widget: null };
  const selectedWidget = project ? getBiSelectedWidget(project) : null;
  const visual = target.mode === "widget" && target.widget
    ? getBiEffectiveVisualSettings(project, target.widget)
    : normalizeBiVisualSettings(config.visual);
  const widgetScopeWithoutTarget = target.mode === "widget" && !target.widget;
  if (els.biVisualScopeSelect) {
    els.biVisualScopeSelect.value = target.mode;
  }
  if (els.biVisualScopeMeta) {
    if (target.mode === "widget" && target.widget) {
      els.biVisualScopeMeta.textContent = `Widget: ${target.widget.name}`;
    } else if (target.mode === "widget") {
      els.biVisualScopeMeta.textContent = "Selecciona un widget para editar.";
    } else {
      els.biVisualScopeMeta.textContent = "Configuracion global del proyecto.";
    }
  }
  if (els.biWidgetShowTitleCheckbox) {
    els.biWidgetShowTitleCheckbox.checked = selectedWidget ? selectedWidget.showTitle !== false : true;
    els.biWidgetShowTitleCheckbox.toggleAttribute("disabled", !selectedWidget);
  }
  if (els.biWidgetShowBorderCheckbox) {
    els.biWidgetShowBorderCheckbox.checked = selectedWidget ? selectedWidget.showBorder !== false : true;
    els.biWidgetShowBorderCheckbox.toggleAttribute("disabled", !selectedWidget);
  }
  if (els.biWidgetShowBackgroundCheckbox) {
    els.biWidgetShowBackgroundCheckbox.checked = selectedWidget ? selectedWidget.showBackground !== false : true;
    els.biWidgetShowBackgroundCheckbox.toggleAttribute("disabled", !selectedWidget);
  }
  if (els.biWidgetChromeMeta) {
    els.biWidgetChromeMeta.textContent = selectedWidget
      ? `Widget activo: ${selectedWidget.name}`
      : "Selecciona un widget para configurar fondo, borde y titulo.";
  }
  if (els.biApplyVisualScopeButton) {
    els.biApplyVisualScopeButton.disabled = widgetScopeWithoutTarget;
  }
  if (els.biResetVisualScopeButton) {
    els.biResetVisualScopeButton.disabled = widgetScopeWithoutTarget;
  }
  const widgetScopedControls = [
    els.biShowLegendCheckbox,
    els.biShowGridCheckbox,
    els.biShowAxisLabelsCheckbox,
    els.biShowDataLabelsCheckbox,
    els.biAxisXLabelInput,
    els.biAxisYLabelInput,
    els.biLabelMaxCharsInput,
    els.biValueDecimalsInput,
    els.biNumberFormatSelect,
    els.biValuePrefixInput,
    els.biValueSuffixInput,
    els.biFontSizeInput,
    els.biFontFamilyTitleSelect,
    els.biFontFamilyAxisSelect,
    els.biFontFamilyLabelSelect,
    els.biFontFamilyTooltipSelect,
    els.biFontSizeTitleInput,
    els.biFontSizeAxisInput,
    els.biFontSizeLabelsInput,
    els.biFontSizeTooltipInput,
    els.biLabelColorModeSelect,
    els.biLabelColorInput,
    els.biLineWidthInput,
    els.biSeriesLineStyleSelect,
    els.biSeriesMarkerStyleSelect,
    els.biSeriesMarkerSizeInput,
    els.biAreaOpacityInput,
    els.biLegendPositionSelect,
    els.biLegendMaxItemsInput,
    els.biAxisScaleSelect,
    els.biAxisMinInput,
    els.biAxisMaxInput,
    els.biShowTargetLineCheckbox,
    els.biTargetLineValueInput,
    els.biTargetLineLabelInput,
    els.biTargetLineColorInput,
    els.biBarWidthRatioInput,
    els.biGridOpacityInput,
    els.biGridDashInput,
    els.biStackModeSelect,
    els.biTooltipModeSelect,
    els.biSmoothLinesCheckbox,
    els.biApplyVisualPresetButton,
    els.biResetVisualSettingsButton
  ];
  widgetScopedControls.forEach((control) => {
    if (control instanceof HTMLElement) {
      control.toggleAttribute("disabled", widgetScopeWithoutTarget);
    }
  });
  if (els.biShowLegendCheckbox) {
    els.biShowLegendCheckbox.checked = !!visual.showLegend;
  }
  if (els.biShowGridCheckbox) {
    els.biShowGridCheckbox.checked = !!visual.showGrid;
  }
  if (els.biShowAxisLabelsCheckbox) {
    els.biShowAxisLabelsCheckbox.checked = !!visual.showAxisLabels;
  }
  if (els.biShowDataLabelsCheckbox) {
    els.biShowDataLabelsCheckbox.checked = !!visual.showDataLabels;
  }
  if (els.biAxisXLabelInput) {
    els.biAxisXLabelInput.value = trimOrFallback(visual.axisXLabel, "");
  }
  if (els.biAxisYLabelInput) {
    els.biAxisYLabelInput.value = trimOrFallback(visual.axisYLabel, "");
  }
  if (els.biLabelMaxCharsInput) {
    els.biLabelMaxCharsInput.value = String(visual.labelMaxChars);
  }
  if (els.biValueDecimalsInput) {
    els.biValueDecimalsInput.value = String(visual.valueDecimals);
  }
  if (els.biNumberFormatSelect) {
    els.biNumberFormatSelect.value = visual.numberFormat;
  }
  if (els.biValuePrefixInput) {
    els.biValuePrefixInput.value = trimOrFallback(visual.valuePrefix, "");
  }
  if (els.biValueSuffixInput) {
    els.biValueSuffixInput.value = trimOrFallback(visual.valueSuffix, "");
  }
  if (els.biFontSizeInput) {
    els.biFontSizeInput.value = String(visual.fontSize);
  }
  if (els.biFontFamilyTitleSelect) {
    els.biFontFamilyTitleSelect.value = sanitizeBiFontFamily(visual.fontFamilyTitle, "Segoe UI");
  }
  if (els.biFontFamilyAxisSelect) {
    els.biFontFamilyAxisSelect.value = sanitizeBiFontFamily(visual.fontFamilyAxis, "Segoe UI");
  }
  if (els.biFontFamilyLabelSelect) {
    els.biFontFamilyLabelSelect.value = sanitizeBiFontFamily(visual.fontFamilyLabel, "Segoe UI");
  }
  if (els.biFontFamilyTooltipSelect) {
    els.biFontFamilyTooltipSelect.value = sanitizeBiFontFamily(visual.fontFamilyTooltip, "Segoe UI");
  }
  if (els.biFontSizeTitleInput) {
    els.biFontSizeTitleInput.value = String(visual.fontSizeTitle);
  }
  if (els.biFontSizeAxisInput) {
    els.biFontSizeAxisInput.value = String(visual.fontSizeAxis);
  }
  if (els.biFontSizeLabelsInput) {
    els.biFontSizeLabelsInput.value = String(visual.fontSizeLabels);
  }
  if (els.biFontSizeTooltipInput) {
    els.biFontSizeTooltipInput.value = String(visual.fontSizeTooltip);
  }
  if (els.biLabelColorModeSelect) {
    els.biLabelColorModeSelect.value = visual.labelColorMode;
  }
  if (els.biLabelColorInput) {
    els.biLabelColorInput.value = normalizeBiColorHex(visual.labelColor, "#1f2f44");
    els.biLabelColorInput.disabled = visual.labelColorMode !== "manual";
  }
  if (els.biLineWidthInput) {
    els.biLineWidthInput.value = String(visual.lineWidth);
  }
  if (els.biSeriesLineStyleSelect) {
    els.biSeriesLineStyleSelect.value = visual.seriesLineStyle;
  }
  if (els.biSeriesMarkerStyleSelect) {
    els.biSeriesMarkerStyleSelect.value = visual.seriesMarkerStyle;
  }
  if (els.biSeriesMarkerSizeInput) {
    els.biSeriesMarkerSizeInput.value = String(visual.seriesMarkerSize);
  }
  if (els.biAreaOpacityInput) {
    els.biAreaOpacityInput.value = String(visual.areaOpacity);
  }
  if (els.biLegendPositionSelect) {
    els.biLegendPositionSelect.value = visual.legendPosition;
  }
  if (els.biLegendMaxItemsInput) {
    els.biLegendMaxItemsInput.value = String(visual.legendMaxItems);
  }
  if (els.biAxisScaleSelect) {
    els.biAxisScaleSelect.value = visual.axisScale;
  }
  if (els.biAxisMinInput) {
    els.biAxisMinInput.value = visual.axisMin === null ? "" : String(visual.axisMin);
  }
  if (els.biAxisMaxInput) {
    els.biAxisMaxInput.value = visual.axisMax === null ? "" : String(visual.axisMax);
  }
  if (els.biShowTargetLineCheckbox) {
    els.biShowTargetLineCheckbox.checked = !!visual.showTargetLine;
  }
  if (els.biTargetLineValueInput) {
    els.biTargetLineValueInput.value = visual.targetLineValue === null ? "" : String(visual.targetLineValue);
    els.biTargetLineValueInput.disabled = !visual.showTargetLine;
  }
  if (els.biTargetLineLabelInput) {
    els.biTargetLineLabelInput.value = trimOrFallback(visual.targetLineLabel, "Meta");
    els.biTargetLineLabelInput.disabled = !visual.showTargetLine;
  }
  if (els.biTargetLineColorInput) {
    els.biTargetLineColorInput.value = normalizeBiColorHex(visual.targetLineColor, "#b03831");
    els.biTargetLineColorInput.disabled = !visual.showTargetLine;
  }
  if (els.biBarWidthRatioInput) {
    els.biBarWidthRatioInput.value = String(visual.barWidthRatio);
  }
  if (els.biGridOpacityInput) {
    els.biGridOpacityInput.value = String(visual.gridOpacity);
  }
  if (els.biGridDashInput) {
    els.biGridDashInput.value = String(visual.gridDash);
  }
  if (els.biStackModeSelect) {
    els.biStackModeSelect.value = visual.stackMode;
  }
  if (els.biTooltipModeSelect) {
    els.biTooltipModeSelect.value = visual.tooltipMode;
  }
  if (els.biSmoothLinesCheckbox) {
    els.biSmoothLinesCheckbox.checked = !!visual.smoothLines;
  }
  if (els.biVisualPresetSelect) {
    const activePreset = trimOrFallback(config.activeVisualPreset, "corporativo");
    const hasPreset = Array.from(els.biVisualPresetSelect.options).some((option) => option.value === activePreset);
    els.biVisualPresetSelect.value = hasPreset ? activePreset : "corporativo";
  }
  applyBiUiModeVisibility(config.uiMode || "basic");
  syncBiInspectorByWidgetType(project, selectedWidget);
}

function getBiWidgetKindLabel(kind) {
  const safeKind = normalizeBiWidgetKind(kind || "chart");
  if (safeKind === "text") {
    return "texto";
  }
  if (safeKind === "image") {
    return "imagen";
  }
  return "grafico";
}

const BI_ADVANCED_CONTROL_IDS = [
  "biFontFamilyTitleSelect",
  "biFontFamilyAxisSelect",
  "biFontFamilyLabelSelect",
  "biFontFamilyTooltipSelect",
  "biFontSizeTitleInput",
  "biFontSizeAxisInput",
  "biFontSizeLabelsInput",
  "biFontSizeTooltipInput",
  "biSeriesLineStyleSelect",
  "biSeriesMarkerStyleSelect",
  "biSeriesMarkerSizeInput",
  "biAreaOpacityInput",
  "biStackModeSelect",
  "biAxisMinInput",
  "biAxisMaxInput",
  "biGridDashInput",
  "biTooltipModeSelect",
  "biLegendMaxItemsInput",
  "biGridOpacityInput",
  "biBarWidthRatioInput",
  "biValuePrefixInput",
  "biValueSuffixInput"
];

function applyBiUiModeVisibility(mode) {
  const uiMode = normalizeBiUiMode(mode || "basic");
  const advanced = uiMode === "advanced";
  BI_ADVANCED_CONTROL_IDS.forEach((id) => {
    const input = document.getElementById(id);
    if (!(input instanceof HTMLElement)) {
      return;
    }
    const label = input.closest("label");
    if (label instanceof HTMLElement) {
      label.classList.toggle("hidden", !advanced);
    }
  });
}

function syncBiInspectorByWidgetType(project, selectedWidget) {
  const widgetKind = normalizeBiWidgetKind(selectedWidget?.kind || "chart");
  const hasSelection = !!selectedWidget;
  const isChartWidget = !hasSelection || widgetKind === "chart";
  const inspectorMode = normalizeBiRailMode(biRailMode);
  const isPropertiesMode = inspectorMode === "properties";
  const isSettingsMode = inspectorMode === "settings";
  const isBoardMode = inspectorMode === "board";
  const isColorMode = inspectorMode === "colorimetry";
  const isFilterMode = inspectorMode === "filter";
  const isPerformanceMode = inspectorMode === "performance";
  const shouldShowChartBuilder = !hasSelection || widgetKind === "chart";
  const shouldShowChartConfig = isSettingsMode && isChartWidget;
  const chartTypeHint = normalizeBiChartType(selectedWidget?.chartType || els.biWidgetChartTypeSelect?.value || "bar");

  const toggleSection = (section, visible) => {
    if (!(section instanceof HTMLElement)) {
      return;
    }
    section.classList.toggle("hidden", !visible);
  };

  toggleSection(els.biSourceSection, isPropertiesMode && shouldShowChartBuilder);
  toggleSection(els.biDimensionSection, isPropertiesMode && shouldShowChartBuilder);
  toggleSection(els.biMetricSection, isPropertiesMode && shouldShowChartBuilder);
  toggleSection(els.biDataRolesSection, isPropertiesMode && shouldShowChartBuilder);
  toggleSection(els.biChartSection, isPropertiesMode && shouldShowChartBuilder);
  toggleSection(els.biBuilderSection, isPropertiesMode);
  toggleSection(els.biSettingsSection, isSettingsMode);
  toggleSection(els.biElementConfigSection, isSettingsMode);
  toggleSection(els.biBoardSection, isBoardMode);
  toggleSection(els.biColorSection, isColorMode);
  toggleSection(els.biFilterSection, isFilterMode);
  toggleSection(els.biPerformanceSection, isPerformanceMode);
  toggleSection(els.biChartVisualSettingsSection, shouldShowChartConfig);
  if (isSettingsMode) {
    renderBiElementConfigPanel(project, selectedWidget || null, chartTypeHint);
  }

  if (els.biUpdateWidgetButton instanceof HTMLButtonElement) {
    const canUpdateSelected = hasSelection && widgetKind === "chart" && !selectedWidget?.locked;
    els.biUpdateWidgetButton.disabled = !canUpdateSelected;
    if (canUpdateSelected) {
      els.biUpdateWidgetButton.textContent = "Actualizar seleccionado";
    } else if (hasSelection && widgetKind === "chart" && selectedWidget?.locked) {
      els.biUpdateWidgetButton.textContent = "Widget bloqueado";
    } else {
      els.biUpdateWidgetButton.textContent = "Actualizar seleccionado";
    }
  }

  if (els.biStudioTitle instanceof HTMLElement) {
    if (isSettingsMode) {
      els.biStudioTitle.textContent = hasSelection
        ? `Configuracion del ${getBiWidgetKindLabel(widgetKind)}`
        : "Configuracion del grafico";
      return;
    }
    if (isBoardMode) {
      els.biStudioTitle.textContent = "Pizarra";
      return;
    }
    if (isColorMode) {
      els.biStudioTitle.textContent = "Colorimetria";
      return;
    }
    if (isFilterMode) {
      els.biStudioTitle.textContent = "Filtros";
      return;
    }
    if (isPerformanceMode) {
      els.biStudioTitle.textContent = "Rendimiento";
      return;
    }
    if (!hasSelection) {
      els.biStudioTitle.textContent = "Propiedades del grafico";
    } else {
      els.biStudioTitle.textContent = `Propiedades del ${getBiWidgetKindLabel(widgetKind)}`;
    }
  }
}

function normalizeBiVisualScopeMode(value) {
  return trimOrFallback(value, "").toLowerCase() === "widget" ? "widget" : "global";
}

function getBiSelectedWidget(project) {
  if (!project || !selectedBiWidgetId) {
    return null;
  }
  return getBiWidgetById(project, selectedBiWidgetId);
}

function resolveBiVisualTarget(project) {
  const mode = normalizeBiVisualScopeMode(biVisualScopeMode);
  if (mode === "widget") {
    return {
      mode,
      widget: getBiSelectedWidget(project)
    };
  }
  return {
    mode: "global",
    widget: null
  };
}

function applyBiPerformanceVisualTuning(rawVisual, mode) {
  const visual = normalizeBiVisualSettings(rawVisual);
  const safeMode = normalizeBiPerformanceMode(mode);
  if (safeMode === "quality") {
    return visual;
  }
  if (safeMode === "turbo") {
    return normalizeBiVisualSettings({
      ...visual,
      showGrid: false,
      showDataLabels: false,
      smoothLines: false,
      legendMaxItems: Math.min(6, sanitizeBiInteger(visual.legendMaxItems, 8, 1, 20)),
      seriesMarkerSize: Math.min(6, sanitizeBiInteger(visual.seriesMarkerSize, 8, 0, 16)),
      fontSizeLabels: Math.min(10, sanitizeBiInteger(visual.fontSizeLabels, 10, 8, 22)),
      tooltipMode: "compact"
    });
  }
  return normalizeBiVisualSettings({
    ...visual,
    legendMaxItems: Math.min(10, sanitizeBiInteger(visual.legendMaxItems, 10, 1, 20)),
    tooltipMode: visual.tooltipMode === "full" ? "compact" : visual.tooltipMode
  });
}

function getBiEffectiveVisualSettings(project, widget) {
  const globalVisual = normalizeBiVisualSettings(project?.biConfig?.visual);
  const mode = normalizeBiPerformanceMode(project?.biConfig?.performanceMode || "balanced");
  if (!widget || typeof widget.visualOverride !== "object" || widget.visualOverride === null) {
    return applyBiPerformanceVisualTuning(globalVisual, mode);
  }
  const override = normalizeBiVisualSettings(widget.visualOverride);
  return applyBiPerformanceVisualTuning(normalizeBiVisualSettings({
    ...globalVisual,
    ...override
  }), mode);
}

function collectBiVisualSettingsFromInputs() {
  return normalizeBiVisualSettings({
    showLegend: !!els.biShowLegendCheckbox?.checked,
    showGrid: !!els.biShowGridCheckbox?.checked,
    showAxisLabels: !!els.biShowAxisLabelsCheckbox?.checked,
    showDataLabels: !!els.biShowDataLabelsCheckbox?.checked,
    axisXLabel: trimOrFallback(els.biAxisXLabelInput?.value, "").slice(0, 40),
    axisYLabel: trimOrFallback(els.biAxisYLabelInput?.value, "").slice(0, 40),
    labelMaxChars: els.biLabelMaxCharsInput?.value,
    valueDecimals: els.biValueDecimalsInput?.value,
    numberFormat: els.biNumberFormatSelect?.value,
    valuePrefix: trimOrFallback(els.biValuePrefixInput?.value, "").slice(0, 8),
    valueSuffix: trimOrFallback(els.biValueSuffixInput?.value, "").slice(0, 8),
    fontSize: els.biFontSizeInput?.value,
    fontFamilyTitle: els.biFontFamilyTitleSelect?.value,
    fontFamilyAxis: els.biFontFamilyAxisSelect?.value,
    fontFamilyLabel: els.biFontFamilyLabelSelect?.value,
    fontFamilyTooltip: els.biFontFamilyTooltipSelect?.value,
    fontSizeTitle: els.biFontSizeTitleInput?.value,
    fontSizeAxis: els.biFontSizeAxisInput?.value,
    fontSizeLabels: els.biFontSizeLabelsInput?.value,
    fontSizeTooltip: els.biFontSizeTooltipInput?.value,
    labelColorMode: els.biLabelColorModeSelect?.value,
    labelColor: els.biLabelColorInput?.value,
    lineWidth: els.biLineWidthInput?.value,
    seriesLineStyle: els.biSeriesLineStyleSelect?.value,
    seriesMarkerStyle: els.biSeriesMarkerStyleSelect?.value,
    seriesMarkerSize: els.biSeriesMarkerSizeInput?.value,
    areaOpacity: els.biAreaOpacityInput?.value,
    smoothLines: !!els.biSmoothLinesCheckbox?.checked,
    legendPosition: els.biLegendPositionSelect?.value,
    legendMaxItems: els.biLegendMaxItemsInput?.value,
    axisScale: els.biAxisScaleSelect?.value,
    axisMin: els.biAxisMinInput?.value,
    axisMax: els.biAxisMaxInput?.value,
    showTargetLine: !!els.biShowTargetLineCheckbox?.checked,
    targetLineValue: els.biTargetLineValueInput?.value,
    targetLineLabel: trimOrFallback(els.biTargetLineLabelInput?.value, "").slice(0, 24),
    targetLineColor: els.biTargetLineColorInput?.value,
    barWidthRatio: els.biBarWidthRatioInput?.value,
    gridOpacity: els.biGridOpacityInput?.value,
    gridDash: els.biGridDashInput?.value,
    stackMode: els.biStackModeSelect?.value,
    tooltipMode: els.biTooltipModeSelect?.value
  });
}

function getBiBuiltInVisualPresets() {
  return [
    {
      id: "corporativo",
      name: "Corporativo",
      settings: {
        showLegend: true,
        showGrid: true,
        showAxisLabels: true,
        showDataLabels: false,
        numberFormat: "auto",
        fontSize: 10,
        lineWidth: 2,
        areaOpacity: 0.2,
        smoothLines: false,
        gridOpacity: 0.35,
        gridDash: 4,
        legendPosition: "right",
        legendMaxItems: 8
      }
    },
    {
      id: "minimal",
      name: "Minimal",
      settings: {
        showLegend: false,
        showGrid: false,
        showAxisLabels: true,
        showDataLabels: true,
        numberFormat: "number",
        fontSize: 10,
        lineWidth: 2,
        areaOpacity: 0.18,
        smoothLines: true,
        gridOpacity: 0.2,
        gridDash: 0,
        legendPosition: "bottom",
        legendMaxItems: 6
      }
    },
    {
      id: "contraste",
      name: "Contraste",
      settings: {
        showLegend: true,
        showGrid: true,
        showAxisLabels: true,
        showDataLabels: true,
        numberFormat: "number",
        fontSize: 11,
        lineWidth: 3,
        areaOpacity: 0.28,
        smoothLines: false,
        gridOpacity: 0.6,
        gridDash: 2,
        legendPosition: "right",
        legendMaxItems: 10
      }
    },
    {
      id: "presentacion",
      name: "Presentación",
      settings: {
        showLegend: true,
        showGrid: true,
        showAxisLabels: true,
        showDataLabels: true,
        numberFormat: "auto",
        fontSize: 12,
        lineWidth: 3,
        areaOpacity: 0.24,
        smoothLines: true,
        gridOpacity: 0.35,
        gridDash: 5,
        legendPosition: "bottom",
        legendMaxItems: 8
      }
    }
  ];
}

function refreshBiVisualPresetOptions(project) {
  if (!els.biVisualPresetSelect) {
    return;
  }
  const builtIns = getBiBuiltInVisualPresets();
  const customEntries = Object.entries(project?.biConfig?.visualPresets || {})
    .map(([id, item]) => ({
      id,
      name: trimOrFallback(item?.name, id)
    }))
    .sort((a, b) => a.name.localeCompare(b.name, "es"));
  const builtInHtml = builtIns
    .map((item) => `<option value="${escapeAttribute(item.id)}">${escapeHtml(item.name)}</option>`)
    .join("");
  const customHtml = customEntries.length > 0
    ? customEntries.map((item) => `<option value="${escapeAttribute(item.id)}">${escapeHtml(item.name)}</option>`).join("")
    : "";
  els.biVisualPresetSelect.innerHTML = `${builtInHtml}${customHtml}`;
  const activePreset = trimOrFallback(project?.biConfig?.activeVisualPreset, "corporativo");
  const hasActive = Array.from(els.biVisualPresetSelect.options).some((option) => option.value === activePreset);
  els.biVisualPresetSelect.value = hasActive ? activePreset : "corporativo";
}

function resolveBiVisualPreset(project, presetId) {
  const builtIn = getBiBuiltInVisualPresets().find((item) => item.id === presetId);
  if (builtIn) {
    return normalizeBiVisualSettings(builtIn.settings);
  }
  const custom = project?.biConfig?.visualPresets?.[presetId];
  if (!custom || typeof custom !== "object") {
    return null;
  }
  return normalizeBiVisualSettings(custom.settings || custom.visual || custom);
}

function saveCustomBiVisualPreset(project, name, settings) {
  if (!project) {
    return;
  }
  const safeName = trimOrFallback(name, "").slice(0, 30);
  if (!safeName) {
    return;
  }
  const slug = slugify(safeName) || "preset";
  const presetId = `custom_${slug}`;
  if (!project.biConfig.visualPresets || typeof project.biConfig.visualPresets !== "object") {
    project.biConfig.visualPresets = {};
  }
  project.biConfig.visualPresets[presetId] = {
    name: safeName,
    settings: normalizeBiVisualSettings(settings)
  };
  project.biConfig.activeVisualPreset = presetId;
}

function setSelectedBiWidget(project, widgetId, syncInspector) {
  const resolvedId = trimOrFallback(widgetId, "");
  const exists = project && resolvedId
    ? project.biWidgets.some((widget) => widget.id === resolvedId)
    : false;
  selectedBiWidgetId = exists ? resolvedId : "";

  if (els.biWidgetsGrid) {
    els.biWidgetsGrid.querySelectorAll("[data-bi-widget-id]").forEach((node) => {
      if (!(node instanceof HTMLElement)) {
        return;
      }
      const active = trimOrFallback(node.dataset.biWidgetId, "") === selectedBiWidgetId;
      node.classList.toggle("selected", active);
    });
  }

  if (syncInspector && project) {
    syncBiInputs(project.biConfig);
    syncBiBuilderInputsFromSelectedWidget(project);
    syncBiBuilderSelectionUi(project);
  }
}

function sanitizeBiCanvasDimension(value, fallback, min, max) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return Math.max(min, Math.min(max, Math.round(fallback)));
  }
  return Math.max(min, Math.min(max, Math.round(numeric)));
}

function sanitizeBiCanvasZoom(value, fallback = BI_CANVAS_ZOOM_DEFAULT) {
  const numeric = Math.round(Number(value));
  const fallbackNumeric = Math.round(Number(fallback));
  const safeFallback = Number.isFinite(fallbackNumeric)
    ? Math.max(BI_CANVAS_ZOOM_MIN, Math.min(BI_CANVAS_ZOOM_MAX, fallbackNumeric))
    : BI_CANVAS_ZOOM_DEFAULT;
  if (!Number.isFinite(numeric)) {
    return safeFallback;
  }
  return Math.max(BI_CANVAS_ZOOM_MIN, Math.min(BI_CANVAS_ZOOM_MAX, numeric));
}

function getBiCanvasZoomScale(config) {
  const zoom = sanitizeBiCanvasZoom(
    config?.canvasZoom ?? config?.zoomPizarra ?? config?.boardZoom,
    BI_CANVAS_ZOOM_DEFAULT
  );
  return zoom / 100;
}

function getBiSteppedCanvasZoom(currentZoom, direction) {
  const safeZoom = sanitizeBiCanvasZoom(currentZoom, BI_CANVAS_ZOOM_DEFAULT);
  const safeDirection = direction >= 0 ? 1 : -1;
  return sanitizeBiCanvasZoom(safeZoom + (BI_CANVAS_ZOOM_STEP * safeDirection), safeZoom);
}

function getBiCanvasUsedArea(project) {
  const canvasSize = getBiCanvasSizeFromConfig(project?.biConfig);
  const widgets = Array.isArray(project?.biWidgets) ? project.biWidgets : [];
  if (widgets.length === 0) {
    return { width: canvasSize.width, height: canvasSize.height };
  }

  let maxRight = 0;
  let maxBottom = 0;
  widgets.forEach((widget, index) => {
    const normalized = normalizeBiWidget(widget, index);
    const layout = normalizeBiWidgetLayout(normalized.layout, index, normalized);
    maxRight = Math.max(maxRight, layout.x + layout.w);
    maxBottom = Math.max(maxBottom, layout.y + layout.h);
  });

  const paddedWidth = Math.round(maxRight + 48);
  const paddedHeight = Math.round(maxBottom + 48);
  const width = Math.max(220, Math.min(canvasSize.width, paddedWidth || canvasSize.width));
  const height = Math.max(180, Math.min(canvasSize.height, paddedHeight || canvasSize.height));
  return { width, height };
}

function getBiCanvasFitZoom(project, mode = "all") {
  if (!project || !(els.biWidgetsGrid instanceof HTMLElement)) {
    return BI_CANVAS_ZOOM_DEFAULT;
  }
  const usedArea = getBiCanvasUsedArea(project);
  const viewportWidth = Math.max(120, Math.round(els.biWidgetsGrid.clientWidth - 64));
  const viewportHeight = Math.max(120, Math.round(els.biWidgetsGrid.clientHeight - 64));
  const widthRatio = viewportWidth / Math.max(1, usedArea.width);
  const heightRatio = viewportHeight / Math.max(1, usedArea.height);
  let ratio = mode === "width" ? widthRatio : Math.min(widthRatio, heightRatio);
  if (mode !== "width" && ratio < 0.34 && widthRatio > ratio * 1.45) {
    ratio = widthRatio;
  }
  const rawZoom = Math.floor(ratio * 100);
  return sanitizeBiCanvasZoom(rawZoom, BI_CANVAS_ZOOM_DEFAULT);
}

function syncBiZoomSelectControl(selectNode, zoom) {
  if (!(selectNode instanceof HTMLSelectElement)) {
    return;
  }
  const safeZoom = sanitizeBiCanvasZoom(zoom, BI_CANVAS_ZOOM_DEFAULT);
  const safeValue = String(safeZoom);
  Array.from(selectNode.options).forEach((option) => {
    if (option.dataset.dynamicZoom === "1") {
      option.remove();
    }
  });
  const hasOption = Array.from(selectNode.options).some((option) => option.value === safeValue);
  if (!hasOption) {
    const dynamicOption = document.createElement("option");
    dynamicOption.value = safeValue;
    dynamicOption.textContent = `${safeValue}%`;
    dynamicOption.dataset.dynamicZoom = "1";
    selectNode.insertBefore(dynamicOption, selectNode.firstChild);
  }
  selectNode.value = safeValue;
}

function updateBiZoomMenuUi(project) {
  const safeZoom = sanitizeBiCanvasZoom(project?.biConfig?.canvasZoom, BI_CANVAS_ZOOM_DEFAULT);
  if (els.biCanvasZoomMenuButton instanceof HTMLButtonElement) {
    els.biCanvasZoomMenuButton.textContent = `Zoom (${safeZoom}%) v`;
  }
  if (els.biZoomFitAllLabel instanceof HTMLElement) {
    const fitAll = getBiCanvasFitZoom(project, "all");
    els.biZoomFitAllLabel.textContent = `Ajustar todo (${fitAll}%)`;
  }
  if (els.biZoomFitWidthLabel instanceof HTMLElement) {
    const fitWidth = getBiCanvasFitZoom(project, "width");
    els.biZoomFitWidthLabel.textContent = `Ajustar ancho (${fitWidth}%)`;
  }
  if (els.biZoomDefaultButton instanceof HTMLButtonElement) {
    const isDefault = safeZoom === 100;
    els.biZoomDefaultButton.disabled = isDefault;
    els.biZoomDefaultButton.classList.toggle("active", isDefault);
  }
  if (els.biCanvasZoomMenu instanceof HTMLElement) {
    els.biCanvasZoomMenu
      .querySelectorAll("[data-bi-zoom-value]")
      .forEach((node) => {
        if (!(node instanceof HTMLElement)) {
          return;
        }
        node.classList.toggle("active", trimOrFallback(node.dataset.biZoomValue, "") === String(safeZoom));
      });
  }
}

function getBiCanvasSizeFromConfig(config) {
  const width = sanitizeBiCanvasDimension(
    config?.canvasWidth ?? config?.pizarraAncho ?? config?.boardWidth,
    BI_CANVAS_SURFACE_MIN_WIDTH,
    BI_CANVAS_SURFACE_MIN_EDIT_WIDTH,
    BI_CANVAS_SURFACE_MAX_EDIT_WIDTH
  );
  const height = sanitizeBiCanvasDimension(
    config?.canvasHeight ?? config?.pizarraAlto ?? config?.boardHeight,
    BI_CANVAS_SURFACE_HEIGHT,
    BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT,
    BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
  );
  return { width, height };
}

function getBiSourceLabel(source) {
  if (source === "deliverable") {
    return "MIDP";
  }
  if (source === "package") {
    return "Paquetes";
  }
  if (source === "review-control") {
    return "Flujo revision";
  }
  if (source === "all") {
    return "Todas";
  }
  return "General";
}

function getBiFieldRowData(fields, fieldId, rowId) {
  if (!fieldId || !rowId) {
    return { name: "", code: "", label: "" };
  }
  const field = getFieldById(fields, fieldId);
  const row = getFieldRowById(field, rowId);
  if (!row) {
    return { name: "", code: "", label: "" };
  }
  return {
    name: trimOrFallback(row.name, ""),
    code: trimOrFallback(row.code, ""),
    label: formatValueWithCode(row.name, row.code)
  };
}

function collectDeliverableMetricsByPackageKey(project, fieldIds) {
  if (!project) {
    return new Map();
  }
  if (!fieldIds?.projectFieldId || !fieldIds?.disciplineFieldId || !fieldIds?.systemFieldId || !fieldIds?.packageFieldId) {
    return new Map();
  }
  return computePackageControlMetrics(project.deliverables, fieldIds);
}

function isBiFixedField(field) {
  const token = normalizeLookup(field?.label || "").replace(/\s+/g, "");
  if (!token) {
    return false;
  }
  return Array.from(BI_FIXED_FIELD_ALIASES).some((alias) => token.includes(alias));
}

function getBiCustomFieldDefs(project) {
  if (!project || !Array.isArray(project.fields)) {
    return [];
  }
  return project.fields
    .filter((field) => !isBiFixedField(field))
    .map((field) => ({ id: field.id, label: trimOrFallback(field.label, "Campo") }))
    .filter((field) => !!field.id);
}

function getBiCustomValuesFromRefs(project, rowRefs) {
  const values = {};
  const customFields = getBiCustomFieldDefs(project);
  customFields.forEach((field) => {
    const fieldEntity = getFieldById(project.fields, field.id);
    const rowId = trimOrFallback(rowRefs?.[field.id], "");
    const row = getFieldRowById(fieldEntity, rowId);
    values[field.id] = row ? formatValueWithCode(row.name, row.code) : "";
  });
  return values;
}

function getBiGroupByOptions(project) {
  const fixed = [
    { value: "disciplina", label: "Disciplina" },
    { value: "sistema", label: "Sistema" },
    { value: "paquete", label: "Paquete" },
    { value: "creador", label: "Creador" },
    { value: "fase", label: "Fase" },
    { value: "sector", label: "Sector" },
    { value: "nivel", label: "Nivel" },
    { value: "tipo", label: "Tipo" },
    { value: "hito", label: "Hito" },
    { value: "proyecto", label: "Proyecto" },
    { value: "fuente", label: "Fuente" },
    { value: "mes_inicio", label: "Mes de inicio" },
    { value: "mes_fin", label: "Mes de fin" }
  ];
  const custom = getBiCustomFieldDefs(project).map((field) => ({
    value: `field:${field.id}`,
    label: field.label
  }));
  return [...fixed, ...custom];
}

function normalizeBiOptionalMetric(value) {
  const metric = trimOrFallback(value, "");
  return metric ? normalizeBiMetric(metric) : "";
}

function normalizeBiDateRole(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (!token) {
    return "";
  }
  if (token === "startdate" || token === "inicio" || token === "fechainicio") {
    return "startDate";
  }
  if (token === "enddate" || token === "fin" || token === "fechafin") {
    return "endDate";
  }
  if (token === "createdat" || token === "creado" || token === "created") {
    return "createdAt";
  }
  return "";
}

function normalizeBiOptionalGroupBy(value) {
  const token = trimOrFallback(value, "");
  if (!token) {
    return "";
  }
  return normalizeBiGroupBy(token);
}

function hasOwnObjectKey(value, key) {
  return !!value
    && typeof value === "object"
    && !Array.isArray(value)
    && Object.prototype.hasOwnProperty.call(value, key);
}

function getOwnObjectValue(value, key, fallback = undefined) {
  return hasOwnObjectKey(value, key) ? value[key] : fallback;
}

function getBiChartCapabilities(chartType) {
  const type = normalizeBiChartType(chartType || "bar");
  const noAxisTypes = new Set(["pie", "donut", "treemap", "funnel", "gauge", "scorecard", "table", "pivot", "sankey", "timeline"]);
  const legendTypes = new Set(["pie", "donut"]);
  const stackTypes = new Set(["bar", "bar_horizontal", "bar_stacked", "area", "combo"]);
  const breakdownTypes = new Set([
    "bar", "bar_horizontal", "bar_stacked", "line", "timeseries", "area", "combo",
    "pie", "donut", "scatter", "bubble",
    "treemap", "funnel", "waterfall", "radar", "pareto",
    "pivot", "sankey"
  ]);
  const optionalMetricTypes = new Set(["scatter", "bubble", "bullet", "candlestick", "boxplot"]);
  const dateRequiredTypes = new Set(["timeseries", "timeline", "candlestick"]);
  const referenceLineTypes = new Set(["bar", "bar_horizontal", "bar_stacked", "line", "timeseries", "area", "combo", "pareto", "bullet"]);
  const dataLabelTypes = new Set(["bar", "bar_horizontal", "bar_stacked", "line", "timeseries", "area", "combo", "pie", "donut", "scatter", "bubble", "treemap", "funnel", "waterfall", "radar", "pareto"]);
  const sortModeTypes = new Set(["bar", "bar_horizontal", "bar_stacked", "line", "timeseries", "area", "combo", "pie", "donut", "scatter", "bubble", "treemap", "funnel", "waterfall", "radar", "pareto", "gauge", "bullet", "boxplot", "candlestick", "scorecard", "table"]);
  return {
    type,
    usesAxes: !noAxisTypes.has(type),
    supportsLegend: legendTypes.has(type),
    supportsGrid: !noAxisTypes.has(type),
    supportsStack: stackTypes.has(type),
    supportsBreakdown: breakdownTypes.has(type),
    supportsOptionalMetric: optionalMetricTypes.has(type),
    supportsReferenceLine: referenceLineTypes.has(type),
    supportsDataLabels: dataLabelTypes.has(type),
    supportsSortMode: sortModeTypes.has(type),
    supportsDateDimension: dateRequiredTypes.has(type) || type === "line" || type === "area" || type === "combo",
    requiresDateDimension: dateRequiredTypes.has(type)
  };
}

function buildBiDataRolesFromBuilderInput(project, sourceWidget = null) {
  const chartType = normalizeBiChartType(els.biWidgetChartTypeSelect?.value || sourceWidget?.chartType || "bar");
  const groupBy = normalizeBiGroupBy(els.biWidgetGroupBySelect?.value || sourceWidget?.groupBy || "disciplina");
  const metric = normalizeBiMetric(els.biWidgetMetricSelect?.value || sourceWidget?.metric || "count");
  const defaults = buildBiDefaultDataRoles({ groupBy, metric, chartType });
  const capabilities = getBiChartCapabilities(chartType);
  const breakdownRaw = normalizeBiGroupBy(els.biWidgetBreakdownSelect?.value || sourceWidget?.dataRoles?.breakdownDimension || "");
  const optionalMetricRaw = normalizeBiOptionalMetric(els.biWidgetOptionalMetricSelect?.value || sourceWidget?.dataRoles?.optionalMetrics?.[0] || "");
  const dateRaw = normalizeBiDateRole(els.biWidgetDateDimensionSelect?.value || sourceWidget?.dataRoles?.dateDimension || "");
  const breakdownDimension = capabilities.supportsBreakdown
    ? (breakdownRaw || defaults.breakdownDimension || "")
    : null;
  const optionalMetrics = capabilities.supportsOptionalMetric && optionalMetricRaw
    ? [optionalMetricRaw]
    : [];
  const dateDimension = capabilities.supportsDateDimension
    ? (dateRaw || defaults.dateDimension || (capabilities.requiresDateDimension ? "startDate" : ""))
    : null;
  return normalizeBiDataRoles({
    dimensions: [groupBy],
    metrics: [metric],
    optionalMetrics,
    breakdownDimension,
    dateDimension
  }, { groupBy, metric, chartType });
}

function refreshBiGroupByOptions(project) {
  if (!els.biWidgetGroupBySelect) {
    return;
  }
  const options = getBiGroupByOptions(project);
  const currentValue = normalizeBiGroupBy(els.biWidgetGroupBySelect.value || "disciplina");
  els.biWidgetGroupBySelect.innerHTML = options
    .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  const valid = options.some((item) => item.value === currentValue);
  els.biWidgetGroupBySelect.value = valid ? currentValue : (options[0]?.value || "disciplina");
}

function refreshBiBreakdownOptions(project) {
  if (!els.biWidgetBreakdownSelect) {
    return;
  }
  const options = getBiGroupByOptions(project);
  const currentValue = normalizeBiGroupBy(els.biWidgetBreakdownSelect.value || "");
  const rows = [{ value: "", label: "(sin breakdown)" }, ...options];
  els.biWidgetBreakdownSelect.innerHTML = rows
    .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  const valid = rows.some((item) => item.value === currentValue);
  els.biWidgetBreakdownSelect.value = valid ? currentValue : "";
}

function refreshBiColorGroupByOptions(project) {
  if (!els.biColorGroupBySelect) {
    return;
  }
  const options = getBiGroupByOptions(project);
  const currentValue = normalizeBiGroupBy(els.biColorGroupBySelect.value || "disciplina");
  els.biColorGroupBySelect.innerHTML = options
    .map((option) => `<option value="${escapeAttribute(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  const valid = options.some((item) => item.value === currentValue);
  els.biColorGroupBySelect.value = valid ? currentValue : (options[0]?.value || "disciplina");
}

const BI_VISUAL_CONTROL_IDS = [
  "biShowLegendCheckbox",
  "biShowGridCheckbox",
  "biShowAxisLabelsCheckbox",
  "biShowDataLabelsCheckbox",
  "biAxisXLabelInput",
  "biAxisYLabelInput",
  "biLabelMaxCharsInput",
  "biValueDecimalsInput",
  "biNumberFormatSelect",
  "biValuePrefixInput",
  "biValueSuffixInput",
  "biFontSizeInput",
  "biFontFamilyTitleSelect",
  "biFontFamilyAxisSelect",
  "biFontFamilyLabelSelect",
  "biFontFamilyTooltipSelect",
  "biFontSizeTitleInput",
  "biFontSizeAxisInput",
  "biFontSizeLabelsInput",
  "biFontSizeTooltipInput",
  "biLabelColorModeSelect",
  "biLabelColorInput",
  "biLineWidthInput",
  "biSeriesLineStyleSelect",
  "biSeriesMarkerStyleSelect",
  "biSeriesMarkerSizeInput",
  "biAreaOpacityInput",
  "biStackModeSelect",
  "biLegendPositionSelect",
  "biLegendMaxItemsInput",
  "biAxisScaleSelect",
  "biAxisMinInput",
  "biAxisMaxInput",
  "biShowTargetLineCheckbox",
  "biTargetLineValueInput",
  "biTargetLineLabelInput",
  "biTargetLineColorInput",
  "biBarWidthRatioInput",
  "biGridOpacityInput",
  "biGridDashInput",
  "biTooltipModeSelect",
  "biSmoothLinesCheckbox"
];

const BI_CONFIG_SECTION_LABELS = {
  general: "General",
  axes: "Ejes",
  series: "Series",
  labels: "Labels",
  numbers: "Formato",
  typography: "Tipografia",
  legend: "Leyenda",
  target: "Meta",
  interaction: "Interaccion"
};

function createBiChartProfile(chartType) {
  const type = normalizeBiChartType(chartType || "bar");
  const common = [
    "biLabelMaxCharsInput",
    "biValueDecimalsInput",
    "biNumberFormatSelect",
    "biValuePrefixInput",
    "biValueSuffixInput",
    "biFontSizeInput",
    "biFontFamilyTitleSelect",
    "biFontFamilyAxisSelect",
    "biFontFamilyLabelSelect",
    "biFontFamilyTooltipSelect",
    "biFontSizeTitleInput",
    "biFontSizeAxisInput",
    "biFontSizeLabelsInput",
    "biFontSizeTooltipInput",
    "biLabelColorModeSelect",
    "biLabelColorInput",
    "biTooltipModeSelect"
  ];
  const legend = ["biShowLegendCheckbox", "biLegendPositionSelect", "biLegendMaxItemsInput"];
  const axes = [
    "biShowGridCheckbox",
    "biShowAxisLabelsCheckbox",
    "biAxisXLabelInput",
    "biAxisYLabelInput",
    "biAxisScaleSelect",
    "biAxisMinInput",
    "biAxisMaxInput",
    "biGridOpacityInput",
    "biGridDashInput"
  ];
  const lineSeries = [
    "biLineWidthInput",
    "biSeriesLineStyleSelect",
    "biSeriesMarkerStyleSelect",
    "biSeriesMarkerSizeInput",
    "biSmoothLinesCheckbox"
  ];
  const markersOnly = ["biSeriesMarkerStyleSelect", "biSeriesMarkerSizeInput"];
  const target = ["biShowTargetLineCheckbox", "biTargetLineValueInput", "biTargetLineLabelInput", "biTargetLineColorInput"];
  const all = new Set();
  const sections = [];
  const addControls = (controls) => controls.forEach((id) => all.add(id));
  const addSections = (names) => names.forEach((name) => {
    if (!sections.includes(name)) {
      sections.push(name);
    }
  });

  addControls(common);
  addSections(["general", "labels", "numbers", "typography", "interaction"]);

  if (type === "line" || type === "timeseries") {
    addControls(legend);
    addControls(axes);
    addControls(lineSeries);
    addControls(target);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["axes", "series", "legend", "target"]);
    return { controls: all, sections, description: "Configuracion para grafico de lineas/serie temporal." };
  }
  if (type === "area") {
    addControls(legend);
    addControls(axes);
    addControls(lineSeries);
    addControls(["biAreaOpacityInput", "biStackModeSelect", "biShowDataLabelsCheckbox"]);
    addControls(target);
    addSections(["axes", "series", "legend", "target"]);
    return { controls: all, sections, description: "Configuracion para grafico de area." };
  }
  if (type === "combo") {
    addControls(legend);
    addControls(axes);
    addControls(lineSeries);
    addControls(["biAreaOpacityInput", "biStackModeSelect", "biBarWidthRatioInput", "biShowDataLabelsCheckbox"]);
    addControls(target);
    addSections(["axes", "series", "legend", "target"]);
    return { controls: all, sections, description: "Configuracion para grafico combinado (lineas + barras)." };
  }
  if (BI_BAR_FAMILY_TYPES.has(type) || type === "pareto" || type === "waterfall") {
    addControls(legend);
    addControls(axes);
    addControls(["biBarWidthRatioInput", "biStackModeSelect", "biShowDataLabelsCheckbox"]);
    addControls(target);
    addSections(["axes", "series", "legend", "target"]);
    return {
      controls: all,
      sections,
      description: BI_HORIZONTAL_BAR_TYPES.has(type)
        ? "Configuracion para grafico de barras horizontales."
        : "Configuracion para grafico de barras/cartesiano."
    };
  }
  if (type === "scatter" || type === "bubble") {
    addControls(axes);
    addControls(markersOnly);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["axes", "series"]);
    return { controls: all, sections, description: "Configuracion para dispersion/burbujas (sin leyenda ni stacking)." };
  }
  if (type === "pie" || type === "donut") {
    addControls(legend);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["legend"]);
    return { controls: all, sections, description: "Configuracion para grafico circular/dona (sin ejes ni rejilla)." };
  }
  if (type === "treemap" || type === "funnel") {
    addControls(legend);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["legend", "series"]);
    return { controls: all, sections, description: "Configuracion para jerarquia/proceso (sin ejes)." };
  }
  if (type === "radar") {
    addControls(lineSeries);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["series"]);
    return { controls: all, sections, description: "Configuracion para radar." };
  }
  if (type === "gauge" || type === "bullet") {
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["series"]);
    return { controls: all, sections, description: "Configuración para medidor/viñeta (sin ejes/rejilla)." };
  }
  if (type === "boxplot" || type === "candlestick") {
    addControls(axes);
    addControls(["biShowDataLabelsCheckbox"]);
    addSections(["axes", "series"]);
    return { controls: all, sections, description: "Configuracion para distribucion/vela." };
  }
  if (type === "table" || type === "pivot" || type === "scorecard") {
    addSections(["numbers", "typography"]);
    return { controls: all, sections, description: "Configuracion para tabla/KPI (sin ejes, rejilla ni series)." };
  }
  if (type === "sankey" || type === "timeline") {
    addSections(["numbers", "typography", "interaction"]);
    return { controls: all, sections, description: "Configuracion para flujo/cronograma (enfocado en formato)." };
  }
  addControls(legend);
  addControls(axes);
  addControls(["biShowDataLabelsCheckbox"]);
  addSections(["axes", "series", "legend"]);
  return { controls: all, sections, description: "Configuracion de grafico general." };
}

function renderBiChartConfigProfile(profile) {
  if (els.biChartConfigIntro instanceof HTMLElement) {
    els.biChartConfigIntro.textContent = trimOrFallback(profile?.description, "");
  }
  if (els.biChartConfigGroups instanceof HTMLElement) {
    const sections = Array.isArray(profile?.sections) ? profile.sections : [];
    els.biChartConfigGroups.innerHTML = sections
      .map((section) => {
        const label = BI_CONFIG_SECTION_LABELS[section] || section;
        return `<span class="bi-config-chip">${escapeHtml(label)}</span>`;
      })
      .join("");
    els.biChartConfigGroups.classList.toggle("hidden", sections.length === 0);
  }
}

function getBiSpecificConfigSchema(kind, chartType) {
  const safeKind = normalizeBiWidgetKind(kind || "chart");
  if (safeKind === "text") {
    return {
      title: "Configuración de texto",
      hint: "Controla tipografía, alineación y espaciado del cuadro de texto.",
      fields: [
        { key: "fontFamily", label: "Fuente", type: "select", options: Array.from(BI_ALLOWED_FONT_FAMILIES).map((item) => ({ value: item, label: item })) },
        { key: "fontSize", label: "Tamaño (px)", type: "number", min: 12, max: 180, step: 1 },
        { key: "fontWeight", label: "Peso", type: "number", min: 300, max: 900, step: 100 },
        { key: "fontStyle", label: "Estilo", type: "select", options: [{ value: "normal", label: "Normal" }, { value: "italic", label: "Cursiva" }] },
        { key: "textAlign", label: "Alineación", type: "select", options: [{ value: "left", label: "Izquierda" }, { value: "center", label: "Centro" }, { value: "right", label: "Derecha" }] },
        { key: "lineHeight", label: "Interlineado", type: "number", min: 0.9, max: 2.4, step: 0.01 },
        { key: "color", label: "Color texto", type: "color" },
        { key: "padding", label: "Padding (px)", type: "number", min: 0, max: 48, step: 1 }
      ]
    };
  }
  if (safeKind === "image") {
    return {
      title: "Configuración de imagen",
      hint: "Controla ajuste, opacidad y marco del widget de imagen.",
      fields: [
        { key: "fit", label: "Ajuste", type: "select", options: [{ value: "contain", label: "Contain" }, { value: "cover", label: "Cover" }, { value: "fill", label: "Fill" }] },
        { key: "opacity", label: "Opacidad", type: "number", min: 0.05, max: 1, step: 0.01 },
        { key: "borderRadius", label: "Radio borde (px)", type: "number", min: 0, max: 40, step: 1 },
        { key: "frameColor", label: "Color marco", type: "color" },
        { key: "background", label: "Fondo imagen", type: "color" }
      ]
    };
  }

  const safeType = normalizeBiChartType(chartType || "bar");
  const base = {
    title: `Configuración de ${safeType}`,
    hint: "Opciones específicas del tipo de gráfico seleccionado.",
    fields: []
  };
  if (safeType === "pie") {
    base.fields = [
      { key: "labelMinPercent", label: "Label mínimo (%)", type: "number", min: 0, max: 100, step: 0.1 },
      { key: "startAngle", label: "?ngulo inicial", type: "number", min: -360, max: 360, step: 1 }
    ];
    return base;
  }
  if (safeType === "donut") {
    base.fields = [
      { key: "labelMinPercent", label: "Label mínimo (%)", type: "number", min: 0, max: 100, step: 0.1 },
      { key: "innerRadiusRatio", label: "Radio interno", type: "number", min: 0.35, max: 0.8, step: 0.01 },
      { key: "centerTitle", label: "Título central", type: "text", maxLength: 20 }
    ];
    return base;
  }
  if (safeType === "treemap") {
    base.fields = [
      { key: "minTilePercent", label: "Mín. área (%)", type: "number", min: 0, max: 25, step: 0.1 },
      { key: "showValueText", label: "Mostrar valores", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "funnel") {
    base.fields = [
      { key: "sortMode", label: "Orden", type: "select", options: [{ value: "desc", label: "Mayor a menor" }, { value: "asc", label: "Menor a mayor" }] },
      { key: "minSegmentPercent", label: "Mín. segmento (%)", type: "number", min: 0, max: 60, step: 0.1 }
    ];
    return base;
  }
  if (safeType === "gauge") {
    base.fields = [
      { key: "minValue", label: "Mínimo", type: "number", min: -1000000, max: 1000000, step: 0.1 },
      { key: "maxValue", label: "Máximo", type: "number", min: -1000000, max: 1000000, step: 0.1 },
      { key: "targetValue", label: "Meta", type: "number", min: -1000000, max: 1000000, step: 0.1 }
    ];
    return base;
  }
  if (safeType === "table") {
    base.fields = [
      { key: "rowLimit", label: "Máx. filas", type: "number", min: 1, max: 200, step: 1 },
      { key: "showIndex", label: "Mostrar ?ndice", type: "checkbox" },
      { key: "compact", label: "Modo compacto", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "pivot") {
    base.fields = [
      { key: "rowLimit", label: "Máx. filas", type: "number", min: 1, max: 200, step: 1 },
      { key: "showTotals", label: "Mostrar totales", type: "checkbox" },
      { key: "compact", label: "Modo compacto", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "scorecard") {
    base.fields = [
      { key: "showSource", label: "Mostrar fuente", type: "checkbox" },
      { key: "showGroups", label: "Mostrar grupos", type: "checkbox" },
      { key: "prefix", label: "Prefijo", type: "text", maxLength: 10 },
      { key: "suffix", label: "Sufijo", type: "text", maxLength: 10 }
    ];
    return base;
  }
  if (safeType === "waterfall") {
    base.fields = [
      { key: "showTotalBar", label: "Mostrar barra total", type: "checkbox" },
      { key: "totalLabel", label: "Etiqueta total", type: "text", maxLength: 18 }
    ];
    return base;
  }
  if (safeType === "scatter") {
    base.fields = [
      { key: "pointAlpha", label: "Opacidad punto", type: "number", min: 0.15, max: 1, step: 0.01 },
      { key: "minPointRadius", label: "Radio mínimo", type: "number", min: 3, max: 40, step: 1 }
    ];
    return base;
  }
  if (safeType === "bubble") {
    base.fields = [
      { key: "pointAlpha", label: "Opacidad burbuja", type: "number", min: 0.15, max: 1, step: 0.01 },
      { key: "minBubbleRadius", label: "Radio mínimo", type: "number", min: 3, max: 60, step: 1 },
      { key: "maxBubbleRadius", label: "Radio máximo", type: "number", min: 3, max: 80, step: 1 }
    ];
    return base;
  }
  if (safeType === "bullet") {
    base.fields = [
      { key: "targetValue", label: "Meta (%)", type: "number", min: 0, max: 100, step: 0.1 },
      { key: "lowBand", label: "Banda baja (%)", type: "number", min: 0, max: 100, step: 0.1 },
      { key: "midBand", label: "Banda media (%)", type: "number", min: 0, max: 100, step: 0.1 }
    ];
    return base;
  }
  if (safeType === "boxplot") {
    base.fields = [
      { key: "showMedianValue", label: "Mostrar mediana", type: "checkbox" },
      { key: "whiskerMultiplier", label: "Multiplicador bigotes", type: "number", min: 0.5, max: 3.5, step: 0.1 }
    ];
    return base;
  }
  if (safeType === "candlestick") {
    base.fields = [
      { key: "upColor", label: "Color alza", type: "color" },
      { key: "downColor", label: "Color baja", type: "color" },
      { key: "wickWidth", label: "Grosor mecha", type: "number", min: 1, max: 4, step: 1 }
    ];
    return base;
  }
  if (safeType === "sankey") {
    base.fields = [
      { key: "maxLinks", label: "Máx. enlaces", type: "number", min: 1, max: 200, step: 1 },
      { key: "minLinkPercent", label: "Mín. ancho (%)", type: "number", min: 0, max: 50, step: 0.1 },
      { key: "showValues", label: "Mostrar valores", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "timeline") {
    base.fields = [
      { key: "rowLimit", label: "Máx. filas", type: "number", min: 1, max: 200, step: 1 },
      { key: "showDates", label: "Mostrar fechas", type: "checkbox" },
      { key: "minBarPercent", label: "Mín. barra (%)", type: "number", min: 0.2, max: 50, step: 0.1 }
    ];
    return base;
  }
  if (safeType === "radar") {
    base.fields = [
      { key: "maxScale", label: "Escala máxima", type: "number", min: 1, max: 10000, step: 1 },
      { key: "showPointValues", label: "Mostrar valores", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "pareto") {
    base.fields = [
      { key: "targetPercent", label: "Meta acumulada (%)", type: "number", min: 1, max: 100, step: 0.1 },
      { key: "showCumulativeLine", label: "Mostrar línea acumulada", type: "checkbox" }
    ];
    return base;
  }
  if (safeType === "line" || safeType === "timeseries" || safeType === "area" || safeType === "combo" || BI_BAR_FAMILY_TYPES.has(safeType)) {
    base.fields = [{ key: "labelMinValue", label: "Mín. valor para label", type: "number", min: -1000000, max: 1000000, step: 0.1 }];
    return base;
  }
  return base;
}

function buildBiSpecificConfigControlHtml(field, value, disabled, kind, chartType) {
  const sharedAttrs = `data-bi-specific-key="${escapeAttribute(field.key)}" data-bi-specific-kind="${escapeAttribute(kind)}" data-bi-specific-chart-type="${escapeAttribute(chartType || "")}"`;
  const disabledAttr = disabled ? " disabled" : "";
  if (field.type === "checkbox") {
    return `<label class="bi-check-line"><input type="checkbox" ${sharedAttrs}${value ? " checked" : ""}${disabledAttr}><span>${escapeHtml(field.label)}</span></label>`;
  }
  if (field.type === "select") {
    const options = (field.options || [])
      .map((option) => `<option value="${escapeAttribute(option.value)}"${option.value === value ? " selected" : ""}>${escapeHtml(option.label)}</option>`)
      .join("");
    return `<label>${escapeHtml(field.label)}<select ${sharedAttrs}${disabledAttr}>${options}</select></label>`;
  }
  if (field.type === "color") {
    return `<label>${escapeHtml(field.label)}<input type="color" ${sharedAttrs} value="${escapeAttribute(String(value || "#1f2f44"))}"${disabledAttr}></label>`;
  }
  if (field.type === "text") {
    const maxLengthAttr = Number.isInteger(field.maxLength) ? ` maxlength="${field.maxLength}"` : "";
    return `<label>${escapeHtml(field.label)}<input type="text" ${sharedAttrs}${maxLengthAttr} value="${escapeAttribute(String(value || ""))}"${disabledAttr}></label>`;
  }
  const minAttr = Number.isFinite(field.min) ? ` min="${field.min}"` : "";
  const maxAttr = Number.isFinite(field.max) ? ` max="${field.max}"` : "";
  const stepAttr = Number.isFinite(field.step) ? ` step="${field.step}"` : "";
  return `<label>${escapeHtml(field.label)}<input type="number" ${sharedAttrs}${minAttr}${maxAttr}${stepAttr} value="${escapeAttribute(String(value ?? ""))}"${disabledAttr}></label>`;
}

function renderBiElementConfigPanel(project, selectedWidget, chartTypeHint) {
  if (!(els.biElementConfigSection instanceof HTMLElement) || !(els.biElementConfigBody instanceof HTMLElement)) {
    return;
  }
  const hasSelection = !!selectedWidget;
  const kind = normalizeBiWidgetKind(selectedWidget?.kind || "chart");
  const chartType = normalizeBiChartType(chartTypeHint || selectedWidget?.chartType || "bar");
  const schema = getBiSpecificConfigSchema(kind, chartType);
  if (els.biElementConfigTitle instanceof HTMLElement) {
    els.biElementConfigTitle.textContent = schema.title;
  }
  if (els.biElementConfigHint instanceof HTMLElement) {
    els.biElementConfigHint.textContent = hasSelection
      ? schema.hint
      : "Selecciona un widget para aplicar cambios persistentes en su configuración específica.";
  }

  let configValues = {};
  if (kind === "text") {
    configValues = normalizeBiTextConfig(selectedWidget?.textConfig);
  } else if (kind === "image") {
    configValues = normalizeBiImageConfig(selectedWidget?.imageConfig);
  } else {
    configValues = getBiSpecificChartConfigForWidget(selectedWidget, chartType);
  }

  if (!Array.isArray(schema.fields) || schema.fields.length === 0) {
    els.biElementConfigBody.innerHTML = "<div class=\"muted\">Este tipo no tiene controles adicionales en esta versión.</div>";
    return;
  }
  const disabled = !hasSelection;
  els.biElementConfigBody.innerHTML = schema.fields
    .map((field) => buildBiSpecificConfigControlHtml(field, configValues[field.key], disabled, kind, chartType))
    .join("");
}

function applyBiSpecificConfigChange(input, options = {}) {
  if (!(input instanceof HTMLElement)) {
    return;
  }
  const key = trimOrFallback(input.dataset.biSpecificKey, "");
  if (!key) {
    return;
  }
  const project = getActiveProject();
  if (!project) {
    return;
  }
  ensureBiState(project);
  const widget = getBiSelectedWidget(project);
  if (!widget) {
    return;
  }
  const kind = normalizeBiWidgetKind(input.dataset.biSpecificKind || widget.kind || "chart");
  const chartType = normalizeBiChartType(input.dataset.biSpecificChartType || widget.chartType || "bar");
  let rawValue = "";
  if (input instanceof HTMLInputElement && input.type === "checkbox") {
    rawValue = input.checked;
  } else if (input instanceof HTMLInputElement || input instanceof HTMLSelectElement) {
    rawValue = input.value;
  } else {
    return;
  }

  if (kind === "text") {
    const next = normalizeBiTextConfig({ ...normalizeBiTextConfig(widget.textConfig), [key]: rawValue });
    widget.textConfig = next;
  } else if (kind === "image") {
    const next = normalizeBiImageConfig({ ...normalizeBiImageConfig(widget.imageConfig), [key]: rawValue });
    widget.imageConfig = next;
  } else {
    const current = getBiSpecificChartConfigForWidget(widget, chartType);
    const next = normalizeBiChartTypeSpecificConfig({ ...current, [key]: rawValue }, chartType);
    setBiSpecificChartConfigForWidget(widget, chartType, next);
  }
  saveState();
  const filteredRows = queryBiProjectRows(project, project.biConfig);
  renderBiWidgets(project, filteredRows);
  if (options.refreshInspector) {
    syncBiInputs(project.biConfig);
  }
}

function toggleBiControlLabel(input, visible) {
  if (!(input instanceof HTMLElement)) {
    return;
  }
  const label = input.closest("label");
  if (label instanceof HTMLElement) {
    label.classList.toggle("bi-hidden-by-chart", !visible);
  }
  if ("disabled" in input) {
    input.disabled = !visible;
  }
}

function syncBiChartContextControls(project, chartType, widget) {
  const capabilities = getBiChartCapabilities(chartType);
  const defaults = buildBiDefaultDataRoles({
    groupBy: normalizeBiGroupBy(els.biWidgetGroupBySelect?.value || widget?.groupBy || "disciplina"),
    metric: normalizeBiMetric(els.biWidgetMetricSelect?.value || widget?.metric || "count"),
    chartType
  });
  const dataRoles = normalizeBiDataRoles(widget?.dataRoles, {
    groupBy: defaults.dimensions?.[0] || "disciplina",
    metric: defaults.metrics?.[0] || "count",
    chartType
  });
  if (els.biDataRolesSection instanceof HTMLElement) {
    const hasDataRoleControls = capabilities.supportsBreakdown || capabilities.supportsOptionalMetric || capabilities.supportsDateDimension;
    els.biDataRolesSection.classList.toggle("hidden", !hasDataRoleControls);
  }

  toggleBiControlLabel(els.biWidgetBreakdownSelect, capabilities.supportsBreakdown);
  toggleBiControlLabel(els.biWidgetOptionalMetricSelect, capabilities.supportsOptionalMetric);
  toggleBiControlLabel(els.biWidgetDateDimensionSelect, capabilities.supportsDateDimension);

  if (els.biWidgetBreakdownSelect) {
    const value = capabilities.supportsBreakdown
      ? normalizeBiGroupBy(els.biWidgetBreakdownSelect.value || dataRoles.breakdownDimension || defaults.breakdownDimension || "")
      : "";
    els.biWidgetBreakdownSelect.value = value || "";
  }
  if (els.biWidgetOptionalMetricSelect) {
    const value = capabilities.supportsOptionalMetric
      ? normalizeBiOptionalMetric(els.biWidgetOptionalMetricSelect.value || dataRoles.optionalMetrics?.[0] || "")
      : "";
    const hasOption = Array.from(els.biWidgetOptionalMetricSelect.options).some((option) => option.value === value);
    els.biWidgetOptionalMetricSelect.value = hasOption ? value : "";
  }
  if (els.biWidgetDateDimensionSelect) {
    const value = capabilities.supportsDateDimension
      ? (normalizeBiDateRole(els.biWidgetDateDimensionSelect.value || dataRoles.dateDimension || defaults.dateDimension || "") || (capabilities.requiresDateDimension ? "startDate" : ""))
      : "";
    const hasOption = Array.from(els.biWidgetDateDimensionSelect.options).some((option) => option.value === value);
    els.biWidgetDateDimensionSelect.value = hasOption ? value : "";
  }

  const profile = createBiChartProfile(chartType);
  BI_VISUAL_CONTROL_IDS.forEach((controlId) => {
    const node = document.getElementById(controlId);
    toggleBiControlLabel(node, profile.controls.has(controlId));
  });
  renderBiChartConfigProfile(profile);

  const legendVisible = profile.controls.has("biShowLegendCheckbox");
  const axisVisible = profile.controls.has("biAxisXLabelInput") || profile.controls.has("biShowAxisLabelsCheckbox");
  const stackVisible = profile.controls.has("biStackModeSelect");

  if (els.biDimensionSection instanceof HTMLElement) {
    const title = els.biDimensionSection.querySelector("h4");
    if (title instanceof HTMLElement) {
      if (!axisVisible) {
        title.textContent = "Dimension";
      } else if (chartType === "bar" || chartType === "bullet") {
        title.textContent = "Dimension - Y axis";
      } else {
        title.textContent = "Dimension - X axis";
      }
    }
  }
  if (els.biMetricSection instanceof HTMLElement) {
    const title = els.biMetricSection.querySelector("h4");
    if (title instanceof HTMLElement) {
      if (!axisVisible) {
        title.textContent = "Metrica";
      } else if (chartType === "bar" || chartType === "bullet") {
        title.textContent = "Metric - X axis";
      } else {
        title.textContent = "Metric - Y axis";
      }
    }
  }

  if (!legendVisible && els.biShowLegendCheckbox instanceof HTMLInputElement) {
    els.biShowLegendCheckbox.checked = false;
  }
  if (!axisVisible) {
    if (els.biShowAxisLabelsCheckbox instanceof HTMLInputElement) {
      els.biShowAxisLabelsCheckbox.checked = false;
    }
    if (els.biShowGridCheckbox instanceof HTMLInputElement) {
      els.biShowGridCheckbox.checked = false;
    }
  }
  if (!stackVisible && els.biStackModeSelect instanceof HTMLSelectElement) {
    els.biStackModeSelect.value = "none";
  }
}

function syncBiBuilderSelectionUi(project) {
  const chartType = normalizeBiChartType(els.biWidgetChartTypeSelect?.value || "bar");
  const selectedWidget = project ? getBiSelectedWidget(project) : null;

  if (els.biWidgetGroupBySelect) {
    const groupBy = normalizeBiGroupBy(els.biWidgetGroupBySelect.value || "disciplina");
    if (els.biSelectedDimensionChip) {
      els.biSelectedDimensionChip.textContent = getBiGroupLabel(groupBy, project);
    }
  }

  if (els.biWidgetMetricSelect) {
    const metric = normalizeBiMetric(els.biWidgetMetricSelect.value || "baseunits");
    if (els.biSelectedMetricChip) {
      els.biSelectedMetricChip.textContent = getBiMetricLabel(metric);
    }
  }

  if (els.biWidgetChartTypeSelect && els.biChartTypeButtons) {
    els.biChartTypeButtons
      .querySelectorAll("[data-bi-chart-type]")
      .forEach((node) => {
        if (!(node instanceof HTMLElement)) {
          return;
        }
        node.classList.toggle("active", normalizeBiChartType(node.dataset.biChartType || "") === chartType);
      });
  }

  if (els.biWidgetSortModeSelect) {
    const sortMode = normalizeBiSortMode(els.biWidgetSortModeSelect.value || "value_desc");
    els.biWidgetSortModeSelect.value = sortMode;
  }

  syncBiChartContextControls(project, chartType, selectedWidget);
  if (normalizeBiRailMode(biRailMode) === "settings") {
    renderBiElementConfigPanel(project, selectedWidget || null, chartType);
  }
}

function syncBiBuilderInputsFromSelectedWidget(project) {
  if (!project) {
    return;
  }
  const widget = getBiSelectedWidget(project);
  if (!widget || normalizeBiWidgetKind(widget.kind || "chart") !== "chart") {
    return;
  }
  const chartType = normalizeBiChartType(widget.chartType || "bar");
  const dataRoles = normalizeBiDataRoles(widget.dataRoles, {
    groupBy: normalizeBiGroupBy(widget.groupBy || "disciplina"),
    metric: normalizeBiMetric(widget.metric || "baseunits"),
    chartType
  });
  if (els.biWidgetNameInput) {
    els.biWidgetNameInput.value = trimOrFallback(widget.name, "").slice(0, 60);
  }
  if (els.biWidgetSourceSelect) {
    els.biWidgetSourceSelect.value = normalizeBiSource(widget.source || "all");
  }
  if (els.biWidgetGroupBySelect) {
    const nextGroupBy = normalizeBiGroupBy(dataRoles.dimensions?.[0] || widget.groupBy || "disciplina");
    const hasGroup = Array.from(els.biWidgetGroupBySelect.options).some((option) => option.value === nextGroupBy);
    els.biWidgetGroupBySelect.value = hasGroup ? nextGroupBy : (els.biWidgetGroupBySelect.options[0]?.value || "disciplina");
  }
  if (els.biWidgetMetricSelect) {
    els.biWidgetMetricSelect.value = normalizeBiMetric(dataRoles.metrics?.[0] || widget.metric || "baseunits");
  }
  if (els.biWidgetBreakdownSelect) {
    const nextBreakdown = normalizeBiGroupBy(dataRoles.breakdownDimension || "");
    const hasBreakdown = Array.from(els.biWidgetBreakdownSelect.options).some((option) => option.value === nextBreakdown);
    els.biWidgetBreakdownSelect.value = hasBreakdown ? nextBreakdown : "";
  }
  if (els.biWidgetOptionalMetricSelect) {
    const nextOptionalMetric = normalizeBiOptionalMetric(dataRoles.optionalMetrics?.[0] || "");
    const hasOptionalMetric = Array.from(els.biWidgetOptionalMetricSelect.options).some((option) => option.value === nextOptionalMetric);
    els.biWidgetOptionalMetricSelect.value = hasOptionalMetric ? nextOptionalMetric : "";
  }
  if (els.biWidgetDateDimensionSelect) {
    const nextDateDimension = normalizeBiDateRole(dataRoles.dateDimension || "");
    const hasDateDimension = Array.from(els.biWidgetDateDimensionSelect.options).some((option) => option.value === nextDateDimension);
    els.biWidgetDateDimensionSelect.value = hasDateDimension ? nextDateDimension : "";
  }
  if (els.biWidgetChartTypeSelect) {
    els.biWidgetChartTypeSelect.value = chartType;
  }
  if (els.biWidgetTopNInput) {
    els.biWidgetTopNInput.value = String(sanitizeBiTopN(widget.topN ?? 10));
  }
  if (els.biWidgetSortModeSelect) {
    els.biWidgetSortModeSelect.value = normalizeBiSortMode(widget.sortMode || "value_desc");
  }
}

function getBiCatalogFieldTypeToken(type) {
  if (type === "date") return "CAL";
  if (type === "number") return "123";
  if (type === "percent") return "%";
  if (type === "calc") return "fx";
  return "ABC";
}

function getBiProjectQueryApi() {
  if (typeof window === "undefined" || !window.MIDPBIProjectQuery || typeof window.MIDPBIProjectQuery !== "object") {
    return null;
  }
  return window.MIDPBIProjectQuery;
}

function createBiProjectQueryService(project) {
  const api = getBiProjectQueryApi();
  if (!project || !api || typeof api.createProjectQueryService !== "function") {
    return null;
  }
  return api.createProjectQueryService({
    project,
    getSourceLabel: getBiSourceLabel,
    getCustomFieldDefs: getBiCustomFieldDefs,
    normalizeSource: normalizeBiSource,
    collectRows: collectBiRows,
    filterRows: filterBiRows,
    trimOrFallback
  });
}

function getBiCatalogFields(project, source) {
  const service = createBiProjectQueryService(project);
  if (!service || typeof service.getCatalogFields !== "function") {
    return [];
  }
  return service.getCatalogFields(source || "all");
}

function buildBiQuerySchemaFields(project, source = "all") {
  const service = createBiProjectQueryService(project);
  if (!service || typeof service.getSchemaFields !== "function") {
    return [];
  }
  return service.getSchemaFields(source);
}

function createBiProjectQueryAdapter(project) {
  const service = createBiProjectQueryService(project);
  if (!service || typeof service.createAdapter !== "function") {
    return null;
  }
  return service.createAdapter();
}

function queryBiProjectRows(project, rawConfig = null) {
  const service = createBiProjectQueryService(project);
  if (!service || typeof service.queryRows !== "function") {
    return [];
  }
  return service.queryRows(rawConfig);
}

function queryQuickSightProjectRows(project) {
  if (!project) {
    return [];
  }
  const baseRows = queryBiProjectRows(project, {
    source: "all",
    startDate: "",
    endDate: "",
    textFilter: "",
    invalidOnly: false,
    crossFilters: [],
    crossFilterScope: "all"
  });
  return filterQuickSightRowsByInteractions(baseRows, project.quickSightConfig);
}

function renderBiFieldCatalog(project) {
  if (!els.biCatalogFieldsList) {
    return;
  }

  if (!project) {
    if (els.biCatalogSummaryText) {
      els.biCatalogSummaryText.textContent = "";
    }
    els.biCatalogFieldsList.innerHTML = "<div class=\"bi-empty\">Sin proyecto activo.</div>";
    return;
  }

  const selectedSource = normalizeBiSource(els.biCatalogSourceSelect?.value || "deliverable");
  if (els.biCatalogSourceSelect && els.biCatalogSourceSelect.value !== selectedSource) {
    els.biCatalogSourceSelect.value = selectedSource;
  }
  const query = normalizeLookup(els.biCatalogSearchInput?.value || "");
  const fields = getBiCatalogFields(project, selectedSource);
  const currentGroupBy = normalizeBiGroupBy(els.biWidgetGroupBySelect?.value || "disciplina");
  const currentMetric = normalizeBiMetric(els.biWidgetMetricSelect?.value || "baseunits");
  const filtered = query
    ? fields.filter((field) => normalizeLookup(`${field.source} ${field.name}`).includes(query))
    : fields;

  if (els.biCatalogSummaryText) {
    els.biCatalogSummaryText.textContent = `Fuente: ${getBiSourceLabel(selectedSource)} | Campos: ${filtered.length}/${fields.length}`;
  }

  if (filtered.length === 0) {
    els.biCatalogFieldsList.innerHTML = "<div class=\"bi-empty\">No hay campos para ese filtro.</div>";
    return;
  }

  els.biCatalogFieldsList.innerHTML = filtered
    .map((field) => {
      const groupBy = trimOrFallback(field.groupBy, "");
      const metric = trimOrFallback(field.metric, "");
      const pickable = !!groupBy || !!metric;
      const isActive = (groupBy && groupBy === currentGroupBy) || (metric && metric === currentMetric);
      const pickableClass = pickable ? " pickable" : "";
      const activeClass = isActive ? " active" : "";
      return `<div class="bi-field-item${pickableClass}${activeClass}"${groupBy ? ` data-bi-groupby="${escapeAttribute(groupBy)}"` : ""}${metric ? ` data-bi-metric="${escapeAttribute(metric)}"` : ""}>
      <span class="bi-field-type">${escapeHtml(getBiCatalogFieldTypeToken(field.type))}</span>
      <span class="bi-field-name">${escapeHtml(field.name)}</span>
      <span class="bi-field-source">${escapeHtml(field.source)}</span>
    </div>`;
    })
    .join("");
}

function getBiColorLegendData(project, rows, source, groupBy) {
  if (!project) {
    return [];
  }
  const safeRows = Array.isArray(rows) ? rows : [];
  const selectedSource = normalizeBiSource(source || "all");
  const selectedGroupBy = normalizeBiGroupBy(groupBy || "disciplina");
  const scopedRows = selectedSource === "all"
    ? safeRows
    : safeRows.filter((row) => row.source === selectedSource);
  return buildBiGroupedRows(scopedRows, selectedGroupBy)
    .sort((a, b) => {
      if (b.count !== a.count) {
        return b.count - a.count;
      }
      return a.label.localeCompare(b.label, "es");
    })
    .slice(0, BI_COLOR_ROWS_LIMIT)
    .map((item, index) => ({
      label: item.label,
      color: getBiColorFromConfig(project, selectedGroupBy, item.label, index),
      index
    }));
}

function clearBiColorsForGroup(project, groupBy) {
  if (!project?.biConfig) {
    return 0;
  }
  const safeGroupBy = normalizeBiGroupBy(groupBy || "disciplina");
  const colorMap = normalizeBiColorMap(project.biConfig.colorMap);
  const prefix = `${safeGroupBy}::`;
  let removed = 0;
  Object.keys(colorMap).forEach((key) => {
    if (!key.startsWith(prefix)) {
      return;
    }
    delete colorMap[key];
    removed += 1;
  });
  project.biConfig.colorMap = colorMap;
  return removed;
}

function renderBiColorPanel(project, rows) {
  if (!els.biColorLegendList || !els.biColorSummaryText) {
    return;
  }
  if (!project) {
    els.biColorLegendList.innerHTML = "<div class=\"bi-empty\">Sin proyecto activo.</div>";
    els.biColorSummaryText.textContent = "";
    return;
  }

  ensureBiState(project);
  const selectedSource = normalizeBiSource(els.biColorSourceSelect?.value || project.biConfig.colorSource || "all");
  const selectedGroupBy = normalizeBiGroupBy(els.biColorGroupBySelect?.value || project.biConfig.colorGroupBy || "disciplina");
  project.biConfig.colorSource = selectedSource;
  project.biConfig.colorGroupBy = selectedGroupBy;
  if (els.biColorSourceSelect && els.biColorSourceSelect.value !== selectedSource) {
    els.biColorSourceSelect.value = selectedSource;
  }
  if (els.biColorGroupBySelect && els.biColorGroupBySelect.value !== selectedGroupBy) {
    els.biColorGroupBySelect.value = selectedGroupBy;
  }

  const legendRows = getBiColorLegendData(project, rows, selectedSource, selectedGroupBy);
  const groupLabel = getBiGroupLabel(selectedGroupBy, project);
  const sourceLabel = getBiSourceLabel(selectedSource);
  els.biColorSummaryText.textContent = `Fuente: ${sourceLabel} | Dimension: ${groupLabel} | Codigos: ${legendRows.length}`;
  if (legendRows.length === 0) {
    els.biColorLegendList.innerHTML = "<div class=\"bi-empty\">No hay categorias disponibles para colorear.</div>";
    return;
  }

  els.biColorLegendList.innerHTML = legendRows
    .map((row) => `<div class="bi-color-row">
      <span class="bi-color-label" title="${escapeAttribute(row.label)}">
        <span class="bi-color-bullet" style="background:${escapeAttribute(row.color)};"></span>
        <span class="bi-color-name">${escapeHtml(row.label)}</span>
      </span>
      <input class="bi-color-picker" type="color" value="${escapeAttribute(row.color)}" data-bi-color-label="${escapeAttribute(row.label)}" data-bi-color-groupby="${escapeAttribute(selectedGroupBy)}">
    </div>`)
    .join("");
}

function getBiRuntimeStats(projectId) {
  if (!projectId) {
    return {
      lastRenderMs: 0,
      avgRenderMs: 0,
      renderCount: 0,
      widgets: 0,
      rows: 0,
      points: 0
    };
  }
  return biRuntimeStatsByProject[projectId] || {
    lastRenderMs: 0,
    avgRenderMs: 0,
    renderCount: 0,
    widgets: 0,
    rows: 0,
    points: 0
  };
}

function updateBiRuntimeStats(projectId, nextStats) {
  if (!projectId || !nextStats || typeof nextStats !== "object") {
    return;
  }
  const previous = getBiRuntimeStats(projectId);
  const renderCount = (previous.renderCount || 0) + 1;
  const lastRenderMs = sanitizeBiDecimal(nextStats.lastRenderMs, 0, 0, 200000);
  const avgRenderMs = ((previous.avgRenderMs || 0) * (renderCount - 1) + lastRenderMs) / renderCount;
  biRuntimeStatsByProject[projectId] = {
    lastRenderMs,
    avgRenderMs: Math.round(avgRenderMs * 100) / 100,
    renderCount,
    widgets: sanitizeBiInteger(nextStats.widgets, 0, 0, 9999),
    rows: sanitizeBiInteger(nextStats.rows, 0, 0, 999999),
    points: sanitizeBiInteger(nextStats.points, 0, 0, 9999999)
  };
}

function runBiAutoOptimize(project, mode) {
  if (!project || !Array.isArray(project.biWidgets)) {
    return { updatedWidgets: 0, topNChanges: 0, visualChanges: 0, tableChanges: 0 };
  }
  const profile = normalizeBiPerformanceMode(mode || project?.biConfig?.performanceMode || "balanced");
  let updatedWidgets = 0;
  let topNChanges = 0;
  let visualChanges = 0;
  let tableChanges = 0;
  const topNCap = profile === "turbo" ? 18 : (profile === "balanced" ? 32 : 50);
  const rowLimitCap = profile === "turbo" ? 40 : (profile === "balanced" ? 80 : 160);
  project.biWidgets.forEach((widget, index) => {
    const normalizedWidget = normalizeBiWidget(widget, index);
    if (normalizeBiWidgetKind(normalizedWidget.kind) !== "chart") {
      return;
    }
    let changed = false;
    const currentTopN = sanitizeBiTopN(normalizedWidget.topN ?? 10);
    if (currentTopN > topNCap) {
      normalizedWidget.topN = topNCap;
      topNChanges += 1;
      changed = true;
    }
    const chartType = normalizeBiChartType(normalizedWidget.chartType || "bar");
    if (chartType === "table" || chartType === "pivot" || chartType === "timeline" || chartType === "sankey") {
      const current = getBiSpecificChartConfigForWidget(normalizedWidget, chartType);
      const safeRowLimit = sanitizeBiInteger(current.rowLimit, 60, 1, 200);
      if (safeRowLimit > rowLimitCap) {
        setBiSpecificChartConfigForWidget(normalizedWidget, chartType, {
          ...current,
          rowLimit: rowLimitCap
        });
        tableChanges += 1;
        changed = true;
      }
    }
    if (!changed) {
      return;
    }
    widget.topN = normalizedWidget.topN;
    widget.chartTypeConfig = normalizeBiChartTypeConfigMap(normalizedWidget.chartTypeConfig);
    updatedWidgets += 1;
  });
  return { updatedWidgets, topNChanges, visualChanges, tableChanges };
}

function renderBiPerformancePanel(project, rows) {
  if (
    !(els.biPerformanceModeSelect instanceof HTMLSelectElement)
    || !(els.biPerfRenderMsText instanceof HTMLElement)
    || !(els.biPerfWidgetsText instanceof HTMLElement)
    || !(els.biPerfPointsText instanceof HTMLElement)
    || !(els.biPerfRowsText instanceof HTMLElement)
    || !(els.biPerfSummaryText instanceof HTMLElement)
  ) {
    return;
  }

  if (!project) {
    els.biPerformanceModeSelect.value = "balanced";
    els.biPerfRenderMsText.textContent = "-";
    els.biPerfWidgetsText.textContent = "-";
    els.biPerfPointsText.textContent = "-";
    els.biPerfRowsText.textContent = "-";
    els.biPerfSummaryText.textContent = "Sin proyecto activo.";
    return;
  }

  ensureBiState(project);
  const runtime = getBiRuntimeStats(project.id);
  const safeMode = normalizeBiPerformanceMode(project.biConfig.performanceMode || "balanced");
  if (els.biPerformanceModeSelect.value !== safeMode) {
    els.biPerformanceModeSelect.value = safeMode;
  }
  const widgets = Array.isArray(project.biWidgets) ? project.biWidgets : [];
  const chartWidgets = widgets.filter((item) => normalizeBiWidgetKind(item?.kind || "chart") === "chart");
  const textWidgets = widgets.filter((item) => normalizeBiWidgetKind(item?.kind || "chart") === "text");
  const imageWidgets = widgets.filter((item) => normalizeBiWidgetKind(item?.kind || "chart") === "image");
  const safeRows = Array.isArray(rows) ? rows : [];
  const avgByWidget = runtime.widgets > 0
    ? Math.round((runtime.points / runtime.widgets) * 10) / 10
    : 0;
  els.biPerfRenderMsText.textContent = `${Math.max(0, Math.round(runtime.lastRenderMs))} ms`;
  els.biPerfWidgetsText.textContent = `${widgets.length} (${chartWidgets.length} gráficos)`;
  els.biPerfPointsText.textContent = `${runtime.points.toLocaleString("es-PE")} | prom: ${avgByWidget.toLocaleString("es-PE")}`;
  els.biPerfRowsText.textContent = safeRows.length.toLocaleString("es-PE");

  const modeLabel = safeMode === "turbo"
    ? "Turbo"
    : (safeMode === "quality" ? "Calidad máxima" : "Balanceado");
  const summaryParts = [
    `Perfil activo: ${modeLabel}`,
    `Render promedio: ${Math.round(runtime.avgRenderMs)} ms`,
    `Texto: ${textWidgets.length}`,
    `Imagen: ${imageWidgets.length}`
  ];
  if (safeMode === "turbo") {
    summaryParts.push("Se prioriza velocidad: menos labels/rejilla.");
  } else if (safeMode === "quality") {
    summaryParts.push("Se prioriza fidelidad visual.");
  } else {
    summaryParts.push("Equilibrio entre claridad y performance.");
  }
  els.biPerfSummaryText.textContent = summaryParts.join(" | ");
}

function getBiLocalDataApi() {
  if (typeof window === "undefined" || !window.MIDPBILocalData || typeof window.MIDPBILocalData !== "object") {
    return null;
  }
  return window.MIDPBILocalData;
}

function createBiLocalDataService(project = null) {
  const api = getBiLocalDataApi();
  if (!api || typeof api.createLocalProjectDataService !== "function") {
    return null;
  }
  return api.createLocalProjectDataService({
    project,
    ensurePackageControls,
    ensureReviewControls,
    ensureReviewMilestones,
    getPackageControlFieldIds,
    collectDeliverableMetricsByPackageKey,
    getFieldRowData: getBiFieldRowData,
    findFieldIdByAlias,
    getCustomValuesFromRefs: getBiCustomValuesFromRefs,
    sanitizeDateInput,
    isDateRangeInvalid,
    buildProgressSnapshot,
    sanitizeBaseUnits,
    sanitizeRealProgress,
    getSourceLabel: getBiSourceLabel,
    trimOrFallback,
    buildPackageControlCode,
    computeProjectIncidenceRatio,
    computeWeightedProjectProgress,
    normalizeLookup,
    formatNumberForInput,
    formatPercent,
    normalizeSource: normalizeBiSource,
    normalizeCrossFilterScope: normalizeBiCrossFilterScope,
    normalizeCrossFilters: normalizeBiCrossFilters,
    rowMatchesCrossFilter: rowMatchesBiCrossFilter,
    getGlobalSearchQuery: () => currentSearchQuery
  });
}

function collectBiRows(project) {
  const service = createBiLocalDataService(project);
  if (!service || typeof service.collectRows !== "function") {
    return [];
  }
  return service.collectRows();
}

function filterBiRows(rows, config) {
  const service = createBiLocalDataService();
  if (!service || typeof service.filterRows !== "function") {
    return [];
  }
  return service.filterRows(rows, config);
}

function formatBiSignedPercent(value, decimals = 2) {
  if (!Number.isFinite(value)) {
    return "0%";
  }
  const absText = formatPercent(Math.abs(value), decimals) || "0%";
  if (value > 0) {
    return `+${absText}`;
  }
  if (value < 0) {
    return `-${absText}`;
  }
  return "0%";
}

function pickTopBiGroup(groups, selector, descending = true) {
  if (!Array.isArray(groups) || groups.length === 0 || typeof selector !== "function") {
    return null;
  }
  const safe = groups.filter((item) => item && item.label && item.label !== "(sin dato)");
  if (safe.length === 0) {
    return null;
  }
  safe.sort((left, right) => {
    const a = selector(left);
    const b = selector(right);
    const leftValue = Number.isFinite(a) ? a : 0;
    const rightValue = Number.isFinite(b) ? b : 0;
    return descending ? rightValue - leftValue : leftValue - rightValue;
  });
  return safe[0] || null;
}

function renderBiInsights(project, rows) {
  if (!(els.biInsightsGrid instanceof HTMLElement)) {
    return;
  }

  if (!project || !Array.isArray(rows) || rows.length === 0) {
    els.biInsightsGrid.innerHTML = "<article class=\"bi-insight-card\"><h4>Insights</h4><p class=\"muted\">Agrega datos o quita filtros para ver recomendaciones inteligentes.</p></article>";
    return;
  }

  const groupedByDiscipline = buildBiGroupedRows(rows, "disciplina");
  const groupedBySystem = buildBiGroupedRows(rows, "sistema");
  const groupedBySource = buildBiGroupedRows(rows, "fuente");
  const invalidDatesCount = rows.filter((item) => item.invalidDates).length;
  const weightedProgrammed = rows.reduce((sum, item) => sum + (item.weightedProgrammed || 0), 0);
  const weightedReal = rows.reduce((sum, item) => sum + (item.weightedReal || 0), 0);
  const weightedGap = weightedReal - weightedProgrammed;
  const weightedGapClass = weightedGap >= 0 ? "positive" : "negative";

  const topDisciplineByUnits = pickTopBiGroup(groupedByDiscipline, (item) => item.baseUnits, true);
  const topSystemByUnits = pickTopBiGroup(groupedBySystem, (item) => item.baseUnits, true);
  const topSourceByRows = pickTopBiGroup(groupedBySource, (item) => item.count, true);
  const lagDiscipline = pickTopBiGroup(
    groupedByDiscipline,
    (item) => (item.realAvg ?? 0) - (item.programmedAvg ?? 0),
    false
  );

  const cards = [];
  cards.push({
    tone: weightedGapClass,
    title: "Salud global del avance",
    value: formatBiSignedPercent(weightedGap, 2),
    detail: `Real proyecto ${formatPercent(weightedReal, 2) || "0%"} vs Programado ${formatPercent(weightedProgrammed, 2) || "0%"}`,
    action: null
  });

  if (topDisciplineByUnits) {
    cards.push({
      tone: "neutral",
      title: "Disciplina con mayor carga",
      value: topDisciplineByUnits.label,
      detail: `${formatNumberForInput(topDisciplineByUnits.baseUnits) || "0"} UP base`,
      action: {
        label: "Filtrar disciplina",
        action: "filter_group",
        groupBy: "disciplina",
        value: topDisciplineByUnits.label
      }
    });
  }

  if (lagDiscipline) {
    const lagDelta = (lagDiscipline.realAvg ?? 0) - (lagDiscipline.programmedAvg ?? 0);
    cards.push({
      tone: lagDelta < 0 ? "warning" : "positive",
      title: "Disciplina más rezagada",
      value: lagDiscipline.label,
      detail: `Brecha ${formatBiSignedPercent(lagDelta, 2)} (Real ${formatPercent(lagDiscipline.realAvg, 2) || "0%"} / Prog ${formatPercent(lagDiscipline.programmedAvg, 2) || "0%"})`,
      action: {
        label: "Analizar disciplina",
        action: "filter_group",
        groupBy: "disciplina",
        value: lagDiscipline.label
      }
    });
  }

  if (topSystemByUnits) {
    cards.push({
      tone: "neutral",
      title: "Sistema con mayor incidencia",
      value: topSystemByUnits.label,
      detail: `${formatNumberForInput(topSystemByUnits.baseUnits) || "0"} UP base`,
      action: {
        label: "Filtrar sistema",
        action: "filter_group",
        groupBy: "sistema",
        value: topSystemByUnits.label
      }
    });
  }

  cards.push({
    tone: invalidDatesCount > 0 ? "warning" : "positive",
    title: "Control de fechas",
    value: invalidDatesCount > 0 ? `${invalidDatesCount} invertidas` : "Sin errores",
    detail: invalidDatesCount > 0
      ? "Hay filas con fecha fin menor a fecha inicio."
      : "No se detectaron rangos de fecha invertidos.",
    action: invalidDatesCount > 0
      ? { label: "Ver inconsistencias", action: "invalid_dates", groupBy: "", value: "" }
      : null
  });

  if (topSourceByRows) {
    cards.push({
      tone: "neutral",
      title: "Fuente dominante",
      value: topSourceByRows.label,
      detail: `${topSourceByRows.count} filas activas`,
      action: {
        label: "Filtrar fuente",
        action: "filter_group",
        groupBy: "fuente",
        value: topSourceByRows.label
      }
    });
  }

  els.biInsightsGrid.innerHTML = cards
    .map((card) => {
      const action = card.action;
      const actionHtml = action
        ? `<button type="button" class="secondary mini-button" data-bi-insight-action="${escapeAttribute(action.action)}" data-bi-insight-groupby="${escapeAttribute(action.groupBy || "")}" data-bi-insight-value="${escapeAttribute(action.value || "")}">${escapeHtml(action.label)}</button>`
        : "";
      return `<article class="bi-insight-card ${escapeAttribute(card.tone || "neutral")}">
        <h4>${escapeHtml(card.title)}</h4>
        <strong>${escapeHtml(card.value)}</strong>
        <p>${escapeHtml(card.detail)}</p>
        <div class="bi-insight-actions">${actionHtml}</div>
      </article>`;
    })
    .join("");
}

function renderBiKpis(rows) {
  if (!els.biKpiGrid) {
    return;
  }

  const totalRows = rows.length;
  const totalBaseUnits = rows.reduce((sum, row) => sum + (row.baseUnits || 0), 0);
  const realRows = rows.filter((row) => row.realProgress !== null);
  const programmedRows = rows.filter((row) => row.programmedPercent !== null);
  const realAvg = realRows.length === 0
    ? null
    : realRows.reduce((sum, row) => sum + (row.realProgress || 0), 0) / realRows.length;
  const programmedAvg = programmedRows.length === 0
    ? null
    : programmedRows.reduce((sum, row) => sum + (row.programmedPercent || 0), 0) / programmedRows.length;
  const weightedReal = rows.reduce((sum, row) => sum + (row.weightedReal || 0), 0);
  const weightedProgrammed = rows.reduce((sum, row) => sum + (row.weightedProgrammed || 0), 0);
  const invalidDates = rows.filter((row) => row.invalidDates).length;
  const sourceCount = new Set(rows.map((row) => row.source)).size;

  const cards = [
    { label: "Filas activas", value: String(totalRows) },
    { label: "Fuentes activas", value: String(sourceCount) },
    { label: "UP Base total", value: formatNumberForInput(totalBaseUnits) || "0" },
    { label: "Avance real promedio", value: formatPercent(realAvg, 2) || "0%" },
    { label: "Avance programado promedio", value: formatPercent(programmedAvg, 2) || "0%" },
    { label: "Avance real proyecto", value: formatPercent(weightedReal, 2) || "0%" },
    { label: "Avance programado proyecto", value: formatPercent(weightedProgrammed, 2) || "0%" },
    { label: "Fechas invertidas", value: String(invalidDates) }
  ];

  els.biKpiGrid.innerHTML = cards
    .map((card) => `<article class="bi-kpi-card">
      <span class="bi-kpi-label">${escapeHtml(card.label)}</span>
      <strong class="bi-kpi-value">${escapeHtml(card.value)}</strong>
    </article>`)
    .join("");
}

function getBiGroupLabel(groupBy, project) {
  if (typeof groupBy === "string" && groupBy.startsWith("field:")) {
    const fieldId = trimOrFallback(groupBy.slice(6), "");
    if (!fieldId) {
      return "Campo";
    }
    const field = getFieldById(project?.fields, fieldId);
    return trimOrFallback(field?.label, "Campo");
  }
  if (groupBy === "proyecto") return "Proyecto";
  if (groupBy === "disciplina") return "Disciplina";
  if (groupBy === "sistema") return "Sistema";
  if (groupBy === "paquete") return "Paquete";
  if (groupBy === "creador") return "Creador";
  if (groupBy === "fase") return "Fase";
  if (groupBy === "sector") return "Sector";
  if (groupBy === "nivel") return "Nivel";
  if (groupBy === "tipo") return "Tipo";
  if (groupBy === "hito") return "Hito";
  if (groupBy === "fuente") return "Fuente";
  if (groupBy === "mes_inicio") return "Mes inicio";
  if (groupBy === "mes_fin") return "Mes fin";
  return "Grupo";
}

function getBiMetricLabel(metric) {
  if (metric === "count") return "Cantidad de filas";
  if (metric === "baseunits") return "UP Base total";
  if (metric === "baseavg") return "UP Base promedio";
  if (metric === "basemax") return "UP Base max";
  if (metric === "basemin") return "UP Base min";
  if (metric === "realavg") return "Avance real promedio (%)";
  if (metric === "realmax") return "Avance real max (%)";
  if (metric === "realmin") return "Avance real min (%)";
  if (metric === "programmedavg") return "Avance programado promedio (%)";
  if (metric === "programmedmax") return "Avance programado max (%)";
  if (metric === "programmedmin") return "Avance programado min (%)";
  if (metric === "weightedreal") return "Avance real proyecto (%)";
  if (metric === "weightedprogrammed") return "Avance programado proyecto (%)";
  if (metric === "weightedgap") return "Brecha real - programado (%)";
  if (metric === "invaliddates") return "Fechas invertidas";
  return "Metrica";
}

function formatBiMonthToken(token) {
  if (!token || !/^\d{4}-\d{2}$/.test(token)) {
    return "";
  }
  return `${token.slice(5, 7)}/${token.slice(0, 4)}`;
}

function getBiRowGroupValue(row, groupBy) {
  if (typeof groupBy === "string" && groupBy.startsWith("field:")) {
    const fieldId = trimOrFallback(groupBy.slice(6), "");
    return trimOrFallback(row?.customValues?.[fieldId], "(sin dato)") || "(sin dato)";
  }
  if (groupBy === "proyecto") return row.projectLabel || "(sin dato)";
  if (groupBy === "disciplina") return row.disciplineLabel || "(sin dato)";
  if (groupBy === "sistema") return row.systemLabel || "(sin dato)";
  if (groupBy === "paquete") return row.packageLabel || "(sin dato)";
  if (groupBy === "creador") return row.creatorLabel || "(sin dato)";
  if (groupBy === "fase") return row.phaseLabel || "(sin dato)";
  if (groupBy === "sector") return row.sectorLabel || "(sin dato)";
  if (groupBy === "nivel") return row.levelLabel || "(sin dato)";
  if (groupBy === "tipo") return row.typeLabel || "(sin dato)";
  if (groupBy === "hito") return row.milestoneLabel || "(sin dato)";
  if (groupBy === "fuente") return row.sourceLabel || "(sin dato)";
  if (groupBy === "mes_inicio") return formatBiMonthToken(row.monthStart) || "(sin dato)";
  if (groupBy === "mes_fin") return formatBiMonthToken(row.monthEnd) || "(sin dato)";
  return "(sin dato)";
}

function rowMatchesBiCrossFilter(row, filter, scope) {
  if (!row || !filter) {
    return true;
  }
  const normalizedScope = normalizeBiCrossFilterScope(scope || "all");
  if (normalizedScope === "same_source") {
    const filterSource = normalizeBiSource(filter.source || "all");
    if (filterSource !== "all" && row.source !== filterSource) {
      return true;
    }
  }
  const rowValue = getBiRowGroupValue(row, filter.groupBy);
  const left = normalizeLookup(rowValue || "");
  const right = normalizeLookup(filter.label || "");
  return !!left && !!right && left === right;
}

function rowMatchesQuickSightInteractionFilter(row, filter) {
  if (!row || !filter) {
    return true;
  }
  return rowMatchesBiCrossFilter(row, filter, filter.scope || "all");
}

function filterQuickSightRowsByInteractions(rows, quickSightConfig) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const filters = normalizeQuickSightInteractionFilters(quickSightConfig?.interactionFilters);
  if (filters.length === 0) {
    return safeRows;
  }
  return safeRows.filter((row) => filters.every((filter) => rowMatchesQuickSightInteractionFilter(row, filter)));
}

function getQuickSightInteractionTargetGroupBy(visual) {
  const safeVisual = visual && typeof visual === "object" ? visual : {};
  const interactions = normalizeQuickSightInteractionSettings(safeVisual.interactions, safeVisual);
  const breakdownGroupBy = normalizeBiOptionalGroupBy(
    getOwnObjectValue(
      safeVisual,
      "breakdownDimension",
      safeVisual.dataRoles?.breakdownDimension ?? ""
    ) || ""
  );
  if (interactions.filterBy === "breakdown" && breakdownGroupBy) {
    return breakdownGroupBy;
  }
  return normalizeBiGroupBy(
    safeVisual.groupBy
    || safeVisual.dataRoles?.dimensions?.[0]
    || "disciplina"
  );
}

function getQuickSightInteractionTargetLabel(snapshotRow, visual) {
  const row = snapshotRow && typeof snapshotRow === "object" ? snapshotRow : {};
  const safeVisual = visual && typeof visual === "object" ? visual : {};
  const interactions = normalizeQuickSightInteractionSettings(safeVisual.interactions, safeVisual);
  if (interactions.filterBy === "breakdown") {
    return trimOrFallback(row.breakdownLabel, "");
  }
  return trimOrFallback(row.groupLabel || row.label, "");
}

function toggleQuickSightInteractionFilter(project, visual, snapshotRow, appendMode = false) {
  if (!project || !visual || !snapshotRow) {
    return false;
  }
  ensureQuickSightState(project);
  const interactions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
  visual.interactions = interactions;
  if (interactions.clickAction !== "cross_filter") {
    return false;
  }
  const groupBy = getQuickSightInteractionTargetGroupBy(visual);
  const label = getQuickSightInteractionTargetLabel(snapshotRow, visual);
  if (!groupBy || !label) {
    return false;
  }
  const nextEntry = normalizeQuickSightInteractionFilterEntry({
    groupBy,
    label,
    source: visual.source || "all",
    scope: interactions.scope,
    visualId: visual.id
  });
  if (!nextEntry) {
    return false;
  }

  const current = normalizeQuickSightInteractionFilters(project.quickSightConfig.interactionFilters);
  const isSameEntry = (entry) => entry
    && entry.groupBy === nextEntry.groupBy
    && entry.source === nextEntry.source
    && entry.scope === nextEntry.scope
    && normalizeLookup(entry.label) === normalizeLookup(nextEntry.label);

  if (appendMode) {
    const exists = current.some((entry) => isSameEntry(entry));
    project.quickSightConfig.interactionFilters = exists
      ? current.filter((entry) => !isSameEntry(entry))
      : [...current, nextEntry];
    return true;
  }

  const currentEntry = current[0] || null;
  if (isSameEntry(currentEntry) && current.length === 1) {
    project.quickSightConfig.interactionFilters = [];
    return true;
  }
  project.quickSightConfig.interactionFilters = [nextEntry];
  return true;
}

function clearQuickSightInteractionFilters(project) {
  if (!project) {
    return false;
  }
  ensureQuickSightState(project);
  const current = normalizeQuickSightInteractionFilters(project.quickSightConfig.interactionFilters);
  if (current.length === 0) {
    return false;
  }
  project.quickSightConfig.interactionFilters = [];
  return true;
}

function getQuickSightInteractionSummary(project) {
  const filters = normalizeQuickSightInteractionFilters(project?.quickSightConfig?.interactionFilters);
  if (filters.length === 0) {
    return "Sin filtros interactivos activos.";
  }
  return filters
    .map((item) => {
      const scopeText = item.scope === "same_source" ? "misma fuente" : "global";
      return `${getBiGroupLabel(item.groupBy, project)} = ${item.label} (${scopeText})`;
    })
    .join(" | ");
}

function openQuickSightDrillThroughDrawerOverlay() {
  if (!(els.qsDrillThroughOverlay instanceof HTMLElement)) {
    return;
  }
  if (quickSightDrillThroughCloseTimer !== null) {
    clearTimeout(quickSightDrillThroughCloseTimer);
    quickSightDrillThroughCloseTimer = null;
  }
  els.qsDrillThroughOverlay.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.qsDrillThroughOverlay?.classList.add("open");
  });
}

function hideQuickSightDrillThroughDrawerOverlay() {
  if (!(els.qsDrillThroughOverlay instanceof HTMLElement) || els.qsDrillThroughOverlay.classList.contains("hidden")) {
    return;
  }
  els.qsDrillThroughOverlay.classList.remove("open");
  if (quickSightDrillThroughCloseTimer !== null) {
    clearTimeout(quickSightDrillThroughCloseTimer);
  }
  quickSightDrillThroughCloseTimer = window.setTimeout(() => {
    els.qsDrillThroughOverlay?.classList.add("hidden");
    quickSightDrillThroughCloseTimer = null;
  }, 220);
}

function getQuickSightDrillThroughSelection(project) {
  if (!project) {
    return null;
  }
  ensureQuickSightState(project);
  return normalizeQuickSightDrillThroughSelection(project.quickSightConfig.drillThroughSelection);
}

function clearQuickSightDrillThroughSelection(project) {
  if (!project) {
    return false;
  }
  ensureQuickSightState(project);
  if (!project.quickSightConfig.drillThroughSelection) {
    return false;
  }
  project.quickSightConfig.drillThroughSelection = null;
  return true;
}

function createQuickSightDrillThroughSelection(project, visual, snapshotRow) {
  if (!project || !visual || !snapshotRow) {
    return null;
  }
  const interactions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
  const breakdownGroupBy = normalizeBiOptionalGroupBy(
    getOwnObjectValue(
      visual,
      "breakdownDimension",
      visual.dataRoles?.breakdownDimension ?? ""
    ) || ""
  );
  const drillMeta = getQuickSightVisualDrillMeta(visual, project);
  const groupBy = interactions.filterBy === "breakdown" && breakdownGroupBy
    ? breakdownGroupBy
    : drillMeta.currentGroupBy;
  const label = interactions.filterBy === "breakdown" && breakdownGroupBy
    ? trimOrFallback(snapshotRow.breakdownLabel, "")
    : trimOrFallback(snapshotRow.groupLabel || snapshotRow.label, "");
  return normalizeQuickSightDrillThroughSelection({
    visualId: visual.id,
    groupBy,
    label,
    source: visual.source || "all",
    openedAt: new Date().toISOString()
  });
}

function getQuickSightDrillThroughRows(project, selection = null) {
  const safeSelection = selection || getQuickSightDrillThroughSelection(project);
  if (!project || !safeSelection) {
    return [];
  }
  const baseRows = getQuickSightPanelBaseRows(project);
  return baseRows.filter((row) => {
    if (safeSelection.source !== "all" && row.source !== safeSelection.source) {
      return false;
    }
    return rowMatchesQuickSightDrillFilter(row, safeSelection);
  });
}

function buildQuickSightDrillThroughRowActionHtml(row) {
  const rowSource = trimOrFallback(row?.source, "");
  const itemId = trimOrFallback(row?.itemId, "");
  if (!itemId || !new Set(["deliverable", "package", "review-control"]).has(rowSource)) {
    return "";
  }
  return `<button type="button" class="secondary mini-button" data-qs-drill-row-source="${escapeAttribute(rowSource)}" data-qs-drill-row-id="${escapeAttribute(itemId)}">Seguimiento</button>`;
}

function renderQuickSightDrillThroughDrawer(project) {
  if (!els.qsDrillThroughOverlay || !els.qsDrillThroughHeader || !els.qsDrillThroughBody) {
    return;
  }
  if (!project || activeTab !== "quicksight") {
    hideQuickSightDrillThroughDrawerOverlay();
    return;
  }
  const selection = getQuickSightDrillThroughSelection(project);
  if (!selection) {
    hideQuickSightDrillThroughDrawerOverlay();
    return;
  }
  const visual = project.quickSightVisuals.find((item) => item.id === selection.visualId) || null;
  if (!visual) {
    project.quickSightConfig.drillThroughSelection = null;
    hideQuickSightDrillThroughDrawerOverlay();
    return;
  }

  const rows = getQuickSightDrillThroughRows(project, selection);
  const searchToken = normalizeLookup(quickSightDrillThroughQuery || "");
  const filteredRows = searchToken
    ? rows.filter((row) => row.searchBlob.includes(searchToken))
    : rows;
  const visibleRows = filteredRows.slice(0, BI_DETAIL_ROW_LIMIT);
  const baseUnitsTotal = rows.reduce((sum, row) => sum + (Number(row.baseUnits) || 0), 0);
  const realValues = rows.map((row) => Number(row.realProgress)).filter((value) => Number.isFinite(value));
  const programmedValues = rows.map((row) => Number(row.programmedPercent)).filter((value) => Number.isFinite(value));
  const avgReal = realValues.length > 0
    ? realValues.reduce((sum, value) => sum + value, 0) / realValues.length
    : null;
  const avgProgrammed = programmedValues.length > 0
    ? programmedValues.reduce((sum, value) => sum + value, 0) / programmedValues.length
    : null;
  const columns = [
    "#",
    "Accion",
    "Fuente",
    "Item",
    "Proyecto",
    "Disciplina",
    "Sistema",
    "Paquete",
    "Hito",
    "Inicio",
    "Fin",
    "UP Base",
    "Prog.",
    "Real",
    "Estado"
  ];

  if (els.qsDrillThroughTitle) {
    els.qsDrillThroughTitle.textContent = "Drill-through QuickSight";
  }
  if (els.qsDrillThroughSubtitle) {
    els.qsDrillThroughSubtitle.textContent = `${visual.name} | ${getBiGroupLabel(selection.groupBy, project)} = ${selection.label}`;
  }
  if (els.qsDrillThroughMeta) {
    const capLabel = filteredRows.length > BI_DETAIL_ROW_LIMIT
      ? `Mostrando ${visibleRows.length} de ${filteredRows.length} filas`
      : `${filteredRows.length} filas`;
    const searchLabel = searchToken ? " | filtro local activo" : "";
    els.qsDrillThroughMeta.textContent = `${capLabel}${searchLabel}`;
  }
  if (els.qsDrillThroughKpis) {
    const kpis = [
      { label: "Filas base", value: String(rows.length) },
      { label: "UP base", value: formatNumberForInput(baseUnitsTotal) || "0" },
      { label: "Real promedio", value: avgReal === null ? "-" : (formatPercent(avgReal, 2) || "0%") },
      { label: "Programado promedio", value: avgProgrammed === null ? "-" : (formatPercent(avgProgrammed, 2) || "0%") }
    ];
    els.qsDrillThroughKpis.innerHTML = kpis
      .map((item) => `<div class="qs-drillthrough-kpi"><span>${escapeHtml(item.label)}</span><strong>${escapeHtml(item.value)}</strong></div>`)
      .join("");
  }
  if (els.qsDrillThroughSearchInput && els.qsDrillThroughSearchInput.value !== quickSightDrillThroughQuery) {
    els.qsDrillThroughSearchInput.value = quickSightDrillThroughQuery;
  }
  els.qsDrillThroughHeader.innerHTML = `<tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;
  if (visibleRows.length === 0) {
    els.qsDrillThroughBody.innerHTML = `<tr><td colspan="${columns.length}" class="muted">Sin filas subyacentes para esta seleccion.</td></tr>`;
  } else {
    els.qsDrillThroughBody.innerHTML = visibleRows
      .map((row, index) => {
        const programmedText = row.programmedPercent === null ? "" : formatPercent(row.programmedPercent, 2);
        const realText = row.realProgress === null ? "" : formatPercent(row.realProgress, 2);
        return `<tr>
          <td>${index + 1}</td>
          <td>${buildQuickSightDrillThroughRowActionHtml(row)}</td>
          <td>${escapeHtml(row.sourceLabel || getBiSourceLabel(row.source))}</td>
          <td>${escapeHtml(row.itemLabel || "")}</td>
          <td>${escapeHtml(row.projectLabel || "")}</td>
          <td>${escapeHtml(row.disciplineLabel || "")}</td>
          <td>${escapeHtml(row.systemLabel || "")}</td>
          <td>${escapeHtml(row.packageLabel || "")}</td>
          <td>${escapeHtml(row.milestoneLabel || "")}</td>
          <td>${escapeHtml(formatDateFromInput(row.startDate) || "")}</td>
          <td>${escapeHtml(formatDateFromInput(row.endDate) || "")}</td>
          <td>${escapeHtml(formatNumberForInput(row.baseUnits) || "0")}</td>
          <td>${programmedText ? `<span class="project-progress-chip">${escapeHtml(programmedText)}</span>` : ""}</td>
          <td>${realText ? `<span class="project-progress-chip">${escapeHtml(realText)}</span>` : ""}</td>
          <td><span class="progress-chip${row.invalidDates ? " overdue" : ""}">${escapeHtml(row.statusLabel || "")}</span></td>
        </tr>`;
      })
      .join("");
  }
  openQuickSightDrillThroughDrawerOverlay();
}

function closeQuickSightDrillThrough(project) {
  quickSightDrillThroughQuery = "";
  if (els.qsDrillThroughSearchInput) {
    els.qsDrillThroughSearchInput.value = "";
  }
  hideQuickSightDrillThroughDrawerOverlay();
  if (!project) {
    return;
  }
  const changed = clearQuickSightDrillThroughSelection(project);
  if (!changed) {
    return;
  }
  saveState();
  renderQuickSightPanel(project, { softSelect: true, allowEmptySelection: true });
  setStatus("Detalle QuickSight cerrado.");
}

function exportQuickSightDrillThroughCsv(project) {
  if (!project) {
    return;
  }
  const selection = getQuickSightDrillThroughSelection(project);
  if (!selection) {
    setStatus("No hay detalle QuickSight activo para exportar.");
    return;
  }
  const rows = getQuickSightDrillThroughRows(project, selection);
  const searchToken = normalizeLookup(quickSightDrillThroughQuery || "");
  const filteredRows = searchToken
    ? rows.filter((row) => row.searchBlob.includes(searchToken))
    : rows;
  if (filteredRows.length === 0) {
    setStatus("No hay filas en el detalle QuickSight para exportar.");
    return;
  }
  const headers = ["Fuente", "Item", "Proyecto", "Disciplina", "Sistema", "Paquete", "Hito", "Inicio", "Fin", "UP Base", "Programado", "Real", "Estado"];
  const csvRows = filteredRows.map((row) => ([
    row.sourceLabel || getBiSourceLabel(row.source),
    row.itemLabel || "",
    row.projectLabel || "",
    row.disciplineLabel || "",
    row.systemLabel || "",
    row.packageLabel || "",
    row.milestoneLabel || "",
    formatDateFromInput(row.startDate) || "",
    formatDateFromInput(row.endDate) || "",
    formatNumberForInput(row.baseUnits) || "0",
    row.programmedPercent === null ? "" : (formatPercent(row.programmedPercent, 2) || ""),
    row.realProgress === null ? "" : (formatPercent(row.realProgress, 2) || ""),
    row.statusLabel || ""
  ]));
  downloadCsv(buildBiExportFileName("dechini_quicksight_drillthrough", project), headers, csvRows);
  setStatus(`Detalle QuickSight exportado (${filteredRows.length} filas).`);
}

function getQuickSightEChartsApi() {
  if (typeof window === "undefined" || !window.echarts || typeof window.echarts.getInstanceByDom !== "function") {
    return null;
  }
  return window.echarts;
}

async function getQuickSightVisualExportDataUrl(visualId, exportScale = 3) {
  if (!(els.qsCanvasBoard instanceof HTMLElement) || !visualId) {
    return "";
  }
  let chartHost = null;
  els.qsCanvasBoard.querySelectorAll("[data-qs-chart-host-id]").forEach((node) => {
    if (chartHost || !(node instanceof HTMLElement)) {
      return;
    }
    if (trimOrFallback(node.dataset.qsChartHostId, "") === visualId) {
      chartHost = node;
    }
  });
  if (!(chartHost instanceof HTMLElement)) {
    return "";
  }
  if (trimOrFallback(chartHost.dataset.qsRenderEngine, "") === "echarts") {
    const surface = chartHost.querySelector("[data-qs-echart-surface]");
    const api = getQuickSightEChartsApi();
    if (surface instanceof HTMLElement && api) {
      const instance = api.getInstanceByDom(surface);
      if (instance && typeof instance.getDataURL === "function") {
        return instance.getDataURL({
          type: "png",
          pixelRatio: exportScale,
          backgroundColor: "#ffffff"
        });
      }
    }
  }
  const canvas = chartHost.querySelector("canvas[data-qs-canvas-id]");
  if (canvas instanceof HTMLCanvasElement && canvas.width > 0 && canvas.height > 0) {
    return canvas.toDataURL("image/png");
  }
  return "";
}

async function exportQuickSightBoardPng(project) {
  if (!project) {
    return;
  }
  ensureQuickSightState(project);
  if (!Array.isArray(project.quickSightVisuals) || project.quickSightVisuals.length === 0) {
    setStatus("No hay visuales QuickSight para publicar/exportar.");
    return;
  }
  const canvasSize = getQuickSightCanvasSizeFromConfig(project.quickSightConfig);
  const boardCanvas = document.createElement("canvas");
  const exportScale = 3;
  boardCanvas.width = canvasSize.width * exportScale;
  boardCanvas.height = canvasSize.height * exportScale;
  const ctx = boardCanvas.getContext("2d");
  if (!ctx) {
    setStatus("No se pudo generar PNG de QuickSight.");
    return;
  }
  ctx.scale(exportScale, exportScale);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvasSize.width, canvasSize.height);

  for (let index = 0; index < project.quickSightVisuals.length; index += 1) {
    const visual = normalizeQuickSightVisual(project.quickSightVisuals[index], index);
    const layout = clampQuickSightVisualLayoutToCanvas(
      normalizeQuickSightVisualLayout(visual.layout, index),
      project.quickSightConfig
    );
    const display = normalizeQuickSightDisplaySettings(visual.display);
    const x = Math.max(0, Math.round(layout.x));
    const y = Math.max(0, Math.round(layout.y));
    const w = Math.max(220, Math.round(layout.w));
    const h = Math.max(160, Math.round(layout.h));
    const showHead = display.showTitle || display.showSubtitle;
    const headHeight = showHead ? 39 : 0;
    const innerPadding = Math.max(4, sanitizeBiInteger(display.padding, 8, 0, 24));
    if (display.showBackground) {
      ctx.fillStyle = hexToRgba(
        display.backgroundColor,
        Math.max(0, Math.min(1, display.backgroundOpacity / 100)),
        "rgba(255,255,255,1)"
      );
      ctx.fillRect(x, y, w, h);
    }
    if (display.showBorder && display.borderWidth > 0) {
      ctx.strokeStyle = hexToRgba(
        display.borderColor,
        Math.max(0, Math.min(1, display.borderOpacity / 100)),
        "rgba(183,201,223,1)"
      );
      ctx.lineWidth = Math.max(1, display.borderWidth);
      ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    }
    if (showHead) {
      const textX = x + innerPadding;
      ctx.textAlign = "left";
      ctx.textBaseline = "alphabetic";
      if (display.showTitle) {
        ctx.fillStyle = display.titleColor;
        ctx.font = `600 ${Math.max(10, display.titleFontSize)}px ${display.titleFontFamily}`;
        ctx.fillText(trimOrFallback(visual.name, "Visual"), textX, y + 18);
      }
      if (display.showSubtitle) {
        const drillMeta = getQuickSightVisualDrillMeta(visual, project);
        const subtitle = trimOrFallback(display.subtitle, "")
          || `${getBiSourceLabel(visual.source)} | ${getBiGroupLabel(drillMeta.currentGroupBy, project)} | ${getBiMetricLabel(visual.metric)}`;
        ctx.fillStyle = display.subtitleColor;
        ctx.font = `${Math.max(8, display.subtitleFontSize)}px ${display.subtitleFontFamily}`;
        ctx.fillText(subtitle, textX, y + 32);
      }
    }
    const chartX = x + innerPadding;
    const chartY = y + headHeight + innerPadding;
    const chartW = Math.max(40, w - (innerPadding * 2));
    const chartH = Math.max(40, h - headHeight - (innerPadding * 2));
    const dataUrl = await getQuickSightVisualExportDataUrl(visual.id, exportScale);
    if (!dataUrl) {
      continue;
    }
    const image = await loadBiImage(dataUrl);
    if (!image) {
      continue;
    }
    ctx.drawImage(image, chartX, chartY, chartW, chartH);
  }

  downloadDataUrl(buildBiExportFileName("dechini_quicksight_hoja", project, "png"), boardCanvas.toDataURL("image/png"));
  setStatus("QuickSight publicado localmente como PNG.");
}

function advanceQuickSightVisualDrill(project, visual, snapshotRow) {
  if (!project || !visual || !snapshotRow) {
    return { changed: false, reason: "invalid" };
  }
  ensureQuickSightState(project);
  ensureQuickSightVisualDrillState(visual);
  const drillMeta = getQuickSightVisualDrillMeta(visual, project);
  if (!drillMeta.enabled) {
    return { changed: false, reason: "disabled" };
  }
  if (drillMeta.sequence.length < 2) {
    return { changed: false, reason: "no_path" };
  }
  if (!drillMeta.canDrillDown) {
    return {
      changed: false,
      reason: "max_level",
      currentGroupBy: drillMeta.currentGroupBy
    };
  }
  const currentGroupBy = drillMeta.currentGroupBy;
  const targetLabel = trimOrFallback(snapshotRow.groupLabel || snapshotRow.label, "");
  if (!targetLabel) {
    return { changed: false, reason: "missing_label" };
  }
  visual.drillState = normalizeQuickSightDrillState({
    filters: [
      ...drillMeta.filters,
      { groupBy: currentGroupBy, label: targetLabel }
    ]
  }, visual);
  const nextMeta = getQuickSightVisualDrillMeta(visual, project);
  return {
    changed: true,
    fromGroupBy: currentGroupBy,
    toGroupBy: nextMeta.currentGroupBy,
    label: targetLabel,
    currentLevel: nextMeta.currentLevel,
    totalLevels: nextMeta.totalLevels
  };
}

function rewindQuickSightVisualDrill(project, visual) {
  if (!project || !visual) {
    return { changed: false, reason: "invalid" };
  }
  ensureQuickSightState(project);
  ensureQuickSightVisualDrillState(visual);
  const drillMeta = getQuickSightVisualDrillMeta(visual, project);
  if (!drillMeta.canDrillUp) {
    return { changed: false, reason: "at_root" };
  }
  visual.drillState = normalizeQuickSightDrillState({
    filters: drillMeta.filters.slice(0, -1)
  }, visual);
  const nextMeta = getQuickSightVisualDrillMeta(visual, project);
  return {
    changed: true,
    currentGroupBy: nextMeta.currentGroupBy,
    currentLevel: nextMeta.currentLevel,
    totalLevels: nextMeta.totalLevels
  };
}

function resetQuickSightVisualDrill(project, visual) {
  if (!project || !visual) {
    return false;
  }
  ensureQuickSightState(project);
  ensureQuickSightVisualDrillState(visual);
  const drillMeta = getQuickSightVisualDrillMeta(visual, project);
  if (drillMeta.filters.length === 0) {
    return false;
  }
  visual.drillState = createDefaultQuickSightDrillState();
  return true;
}

function applyQuickSightInteractionFromVisual(project, visual, snapshot, rowIndex, event) {
  if (!project || !visual || !snapshot || !Array.isArray(snapshot.rows)) {
    return;
  }
  const selectedRow = snapshot.rows[rowIndex];
  if (!selectedRow) {
    return;
  }
  const interactions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
  visual.interactions = interactions;
  if (interactions.clickAction === "none") {
    return;
  }
  if (interactions.clickAction === "drill_through") {
    const nextSelection = createQuickSightDrillThroughSelection(project, visual, selectedRow);
    if (!nextSelection) {
      setStatus("No se pudo abrir el detalle QuickSight.");
      return;
    }
    const currentSelection = getQuickSightDrillThroughSelection(project);
    const isSameSelection = currentSelection
      && currentSelection.visualId === nextSelection.visualId
      && currentSelection.groupBy === nextSelection.groupBy
      && currentSelection.source === nextSelection.source
      && normalizeLookup(currentSelection.label) === normalizeLookup(nextSelection.label);
    if (isSameSelection) {
      closeQuickSightDrillThrough(project);
      return;
    }
    quickSightDrillThroughQuery = "";
    project.quickSightConfig.drillThroughSelection = nextSelection;
    selectedQuickSightVisualId = visual.id;
    saveState();
    renderQuickSightPanel(project, { softSelect: false });
    renderQuickSightDrillThroughDrawer(project);
    setStatus(`Drill-through QuickSight: ${getBiGroupLabel(nextSelection.groupBy, project)} = ${nextSelection.label}.`);
    return;
  }
  if (interactions.clickAction === "drill_down") {
    const result = advanceQuickSightVisualDrill(project, visual, selectedRow);
    if (!result.changed) {
      if (result.reason === "no_path") {
        setStatus("Configura al menos un siguiente nivel para usar drill-down.");
      } else if (result.reason === "max_level") {
        setStatus(`Ya estas en el ultimo nivel de drill: ${getBiGroupLabel(result.currentGroupBy, project)}.`);
      } else {
        setStatus("No se pudo aplicar el drill-down.");
      }
      return;
    }
    selectedQuickSightVisualId = visual.id;
    saveState();
    renderQuickSightPanel(project, { softSelect: false });
    setStatus(`Drill QuickSight: ${getBiGroupLabel(result.fromGroupBy, project)} = ${result.label} -> ${getBiGroupLabel(result.toGroupBy, project)}.`);
    return;
  }
  const allowAppend = interactions.appendWithModifier !== false;
  const appendMode = allowAppend && !!(event && (event.ctrlKey || event.metaKey));
  const changed = toggleQuickSightInteractionFilter(project, visual, selectedRow, appendMode);
  if (!changed) {
    return;
  }
  selectedQuickSightVisualId = visual.id;
  saveState();
  renderQuickSightPanel(project, { softSelect: false });
  const hasFilter = normalizeQuickSightInteractionFilters(project.quickSightConfig.interactionFilters).length > 0;
  if (hasFilter) {
    const targetGroupBy = getQuickSightInteractionTargetGroupBy(visual);
    const targetLabel = getQuickSightInteractionTargetLabel(selectedRow, visual);
    const modeLabel = appendMode ? " (multi)" : "";
    setStatus(`Filtro QuickSight aplicado${modeLabel}: ${getBiGroupLabel(targetGroupBy, project)} = ${targetLabel}`);
  } else {
    setStatus("Filtro QuickSight removido.");
  }
}

function renderBiCrossFilterSummary(project) {
  if (!els.biCrossFilterText || !els.biClearCrossFilterButton) {
    return;
  }
  const filters = normalizeBiCrossFilters(project?.biConfig?.crossFilters);
  if (filters.length === 0) {
    els.biCrossFilterText.textContent = "Sin filtro cruzado activo.";
    els.biClearCrossFilterButton.disabled = true;
    return;
  }
  const scope = normalizeBiCrossFilterScope(project?.biConfig?.crossFilterScope || "all");
  const summary = filters
    .map((item) => {
      const sourceTag = scope === "same_source"
        ? ` [${getBiSourceLabel(item.source || "all")}]`
        : "";
      return `${getBiGroupLabel(item.groupBy, project)} = ${item.label}${sourceTag}`;
    })
    .join(" | ");
  const scopeText = scope === "same_source" ? "alcance misma fuente" : "alcance global";
  els.biCrossFilterText.textContent = `Filtro cruzado (${scopeText}): ${summary}`;
  els.biClearCrossFilterButton.disabled = false;
}

function buildBiGroupedRows(rows, groupBy) {
  const groups = new Map();
  rows.forEach((row) => {
    const label = getBiRowGroupValue(row, groupBy);
    if (!groups.has(label)) {
      groups.set(label, {
        label,
        count: 0,
        baseUnits: 0,
        baseMax: null,
        baseMin: null,
        realSum: 0,
        realCount: 0,
        realMax: null,
        realMin: null,
        programmedSum: 0,
        programmedCount: 0,
        programmedMax: null,
        programmedMin: null,
        weightedReal: 0,
        weightedProgrammed: 0,
        invalidDates: 0
      });
    }

    const item = groups.get(label);
    const baseValue = Number.isFinite(row.baseUnits) ? row.baseUnits : 0;
    item.count += 1;
    item.baseUnits += baseValue;
    item.baseMax = item.baseMax === null ? baseValue : Math.max(item.baseMax, baseValue);
    item.baseMin = item.baseMin === null ? baseValue : Math.min(item.baseMin, baseValue);
    if (row.realProgress !== null) {
      const realValue = row.realProgress || 0;
      item.realSum += realValue;
      item.realCount += 1;
      item.realMax = item.realMax === null ? realValue : Math.max(item.realMax, realValue);
      item.realMin = item.realMin === null ? realValue : Math.min(item.realMin, realValue);
    }
    if (row.programmedPercent !== null) {
      const programmedValue = row.programmedPercent || 0;
      item.programmedSum += programmedValue;
      item.programmedCount += 1;
      item.programmedMax = item.programmedMax === null ? programmedValue : Math.max(item.programmedMax, programmedValue);
      item.programmedMin = item.programmedMin === null ? programmedValue : Math.min(item.programmedMin, programmedValue);
    }
    item.weightedReal += row.weightedReal || 0;
    item.weightedProgrammed += row.weightedProgrammed || 0;
    if (row.invalidDates) {
      item.invalidDates += 1;
    }
  });

  return Array.from(groups.values()).map((group) => ({
    ...group,
    realAvg: group.realCount ? (group.realSum / group.realCount) : null,
    programmedAvg: group.programmedCount ? (group.programmedSum / group.programmedCount) : null
  }));
}

function getBiMetricValue(group, metric) {
  if (metric === "count") return group.count;
  if (metric === "baseunits") return group.baseUnits;
  if (metric === "baseavg") return group.count > 0 ? (group.baseUnits / group.count) : 0;
  if (metric === "basemax") return group.baseMax ?? 0;
  if (metric === "basemin") return group.baseMin ?? 0;
  if (metric === "realavg") return group.realAvg ?? 0;
  if (metric === "realmax") return group.realMax ?? 0;
  if (metric === "realmin") return group.realMin ?? 0;
  if (metric === "programmedavg") return group.programmedAvg ?? 0;
  if (metric === "programmedmax") return group.programmedMax ?? 0;
  if (metric === "programmedmin") return group.programmedMin ?? 0;
  if (metric === "weightedreal") return group.weightedReal ?? 0;
  if (metric === "weightedprogrammed") return group.weightedProgrammed ?? 0;
  if (metric === "weightedgap") return (group.weightedReal ?? 0) - (group.weightedProgrammed ?? 0);
  if (metric === "invaliddates") return group.invalidDates;
  return 0;
}

function isBiPercentMetric(metric) {
  return new Set([
    "realavg",
    "realmax",
    "realmin",
    "programmedavg",
    "programmedmax",
    "programmedmin",
    "weightedreal",
    "weightedprogrammed",
    "weightedgap"
  ]).has(normalizeBiMetric(metric || ""));
}

function formatBiMetricValue(metric, value) {
  if (metric === "count" || metric === "invaliddates") {
    return String(Math.round(value || 0));
  }
  if (metric === "baseunits" || metric === "baseavg" || metric === "basemax" || metric === "basemin") {
    return formatNumberForInput(value) || "0";
  }
  return formatPercent(value, 2) || "0%";
}

function getBiPaletteColor(index) {
  const palette = ["#2f7ed8", "#24a36c", "#f39c12", "#b551e0", "#e74c3c", "#16a085", "#34495e", "#f06292", "#7cb342", "#5c6bc0"];
  return palette[index % palette.length];
}

function normalizeBiColorHex(color, fallback = "#2f7ed8") {
  const raw = trimOrFallback(color, "");
  if (!raw) {
    return fallback;
  }
  const shortMatch = raw.match(/^#([0-9a-fA-F]{3})$/);
  if (shortMatch) {
    const [r, g, b] = shortMatch[1].split("");
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  const fullMatch = raw.match(/^#([0-9a-fA-F]{6})$/);
  if (fullMatch) {
    return `#${fullMatch[1]}`.toLowerCase();
  }
  return fallback;
}

function buildBiColorKey(groupBy, label) {
  const safeGroup = normalizeBiGroupBy(groupBy || "disciplina");
  const safeLabel = normalizeLookup(label || "");
  if (!safeLabel) {
    return "";
  }
  return `${safeGroup}::${safeLabel}`;
}

function normalizeBiColorMap(rawMap) {
  if (!rawMap || typeof rawMap !== "object") {
    return {};
  }
  const entries = Object.entries(rawMap);
  const normalized = {};
  entries.forEach(([key, value]) => {
    const safeKey = trimOrFallback(key, "");
    if (!safeKey) {
      return;
    }
    normalized[safeKey] = normalizeBiColorHex(String(value || ""), "#2f7ed8");
  });
  return normalized;
}

function getBiColorFromConfig(project, groupBy, label, index) {
  const key = buildBiColorKey(groupBy, label);
  const fallback = getBiPaletteColor(index);
  if (!key || !project?.biConfig?.colorMap || typeof project.biConfig.colorMap !== "object") {
    return fallback;
  }
  const configured = project.biConfig.colorMap[key];
  return normalizeBiColorHex(configured, fallback);
}

function getBiSeriesColors(project, groupBy, labels) {
  if (!Array.isArray(labels)) {
    return [];
  }
  return labels.map((label, index) => getBiColorFromConfig(project, groupBy, label, index));
}

function getBiSnapshotSeriesColors(project, snapshot) {
  const rows = Array.isArray(snapshot?.rows) ? snapshot.rows : [];
  const colorDimension = normalizeBiGroupBy(snapshot?.colorByDimension || snapshot?.groupBy || "disciplina");
  const colorLabels = rows.map((row) => trimOrFallback(row?.colorKey, row?.label || ""));
  return getBiSeriesColors(project, colorDimension, colorLabels);
}

function hexToRgba(color, alpha, fallback = "rgba(47, 126, 216, 0.2)") {
  const safeHex = normalizeBiColorHex(color, "");
  const safeAlpha = Number.isFinite(alpha) ? Math.max(0, Math.min(1, alpha)) : 1;
  if (!safeHex) {
    return fallback;
  }
  const raw = safeHex.replace("#", "");
  const red = parseInt(raw.slice(0, 2), 16);
  const green = parseInt(raw.slice(2, 4), 16);
  const blue = parseInt(raw.slice(4, 6), 16);
  return `rgba(${red}, ${green}, ${blue}, ${safeAlpha})`;
}

function biHexToRgb(color, fallback = "#ffffff") {
  const safeHex = normalizeBiColorHex(color, fallback);
  const raw = safeHex.replace("#", "");
  return {
    r: parseInt(raw.slice(0, 2), 16),
    g: parseInt(raw.slice(2, 4), 16),
    b: parseInt(raw.slice(4, 6), 16)
  };
}

function biRelativeLuminance(color) {
  const rgb = biHexToRgb(color, "#ffffff");
  const channel = (value) => {
    const normalized = value / 255;
    if (normalized <= 0.03928) {
      return normalized / 12.92;
    }
    return ((normalized + 0.055) / 1.055) ** 2.4;
  };
  const r = channel(rgb.r);
  const g = channel(rgb.g);
  const b = channel(rgb.b);
  return (0.2126 * r) + (0.7152 * g) + (0.0722 * b);
}

function biContrastRatio(colorA, colorB) {
  const luminanceA = biRelativeLuminance(colorA);
  const luminanceB = biRelativeLuminance(colorB);
  const lighter = Math.max(luminanceA, luminanceB);
  const darker = Math.min(luminanceA, luminanceB);
  return (lighter + 0.05) / (darker + 0.05);
}

function getBiAutoContrastTextColor(backgroundColor) {
  const safeBackground = normalizeBiColorHex(backgroundColor || "#ffffff", "#ffffff");
  const dark = "#1f2f44";
  const light = "#ffffff";
  const darkRatio = biContrastRatio(safeBackground, dark);
  const lightRatio = biContrastRatio(safeBackground, light);
  return darkRatio >= lightRatio ? dark : light;
}

function resolveBiLabelColor(visualSettings, backgroundColor = "#ffffff", fallback = "#1f2f44") {
  const visual = normalizeBiVisualSettings(visualSettings || {});
  if (visual.labelColorMode === "manual") {
    return normalizeBiColorHex(visual.labelColor, fallback);
  }
  return getBiAutoContrastTextColor(backgroundColor);
}

function getBiFontFamilyByLayer(visualSettings, layer) {
  const visual = normalizeBiVisualSettings(visualSettings || {});
  if (layer === "title") {
    return sanitizeBiFontFamily(visual.fontFamilyTitle, "Segoe UI");
  }
  if (layer === "axis") {
    return sanitizeBiFontFamily(visual.fontFamilyAxis, "Segoe UI");
  }
  if (layer === "tooltip") {
    return sanitizeBiFontFamily(visual.fontFamilyTooltip, "Segoe UI");
  }
  return sanitizeBiFontFamily(visual.fontFamilyLabel, "Segoe UI");
}

function getBiFontSizeByLayer(visualSettings, layer, fallback) {
  const visual = normalizeBiVisualSettings(visualSettings || {});
  if (layer === "title") {
    return sanitizeBiInteger(visual.fontSizeTitle, fallback ?? 16, 10, 32);
  }
  if (layer === "axis") {
    return sanitizeBiInteger(visual.fontSizeAxis, fallback ?? 10, 8, 20);
  }
  if (layer === "tooltip") {
    return sanitizeBiInteger(visual.fontSizeTooltip, fallback ?? 11, 8, 20);
  }
  return sanitizeBiInteger(visual.fontSizeLabels, fallback ?? visual.fontSize, 8, 22);
}

function applyBiCanvasFont(ctx, visualSettings, layer, options = {}) {
  if (!ctx) {
    return "";
  }
  const weight = options.bold ? 700 : (options.weight || 400);
  const italic = options.italic ? "italic " : "";
  const defaultSize = Number.isFinite(options.fallbackSize) ? options.fallbackSize : 10;
  const size = getBiFontSizeByLayer(visualSettings, layer, defaultSize);
  const family = getBiFontFamilyByLayer(visualSettings, layer);
  const font = `${italic}${weight} ${size}px ${family}`;
  ctx.font = font;
  return font;
}

function resolveBiLineDashPattern(visualSettings) {
  const visual = normalizeBiVisualSettings(visualSettings || {});
  if (visual.seriesLineStyle === "dashed") {
    return [10, 6];
  }
  if (visual.seriesLineStyle === "dotted") {
    return [2, 5];
  }
  return [];
}

function drawBiMarker(ctx, x, y, markerStyle, size, color) {
  if (!ctx) {
    return;
  }
  const style = trimOrFallback(markerStyle, "circle").toLowerCase();
  const radius = Math.max(0, Math.round(size || 0));
  if (!Number.isFinite(x) || !Number.isFinite(y) || radius <= 0 || style === "none") {
    return;
  }
  const fillColor = normalizeBiColorHex(color, "#2f7ed8");
  ctx.save();
  ctx.fillStyle = fillColor;
  ctx.beginPath();
  if (style === "square") {
    ctx.rect(x - radius, y - radius, radius * 2, radius * 2);
  } else if (style === "diamond") {
    ctx.moveTo(x, y - radius);
    ctx.lineTo(x + radius, y);
    ctx.lineTo(x, y + radius);
    ctx.lineTo(x - radius, y);
    ctx.closePath();
  } else if (style === "triangle") {
    ctx.moveTo(x, y - radius);
    ctx.lineTo(x + radius, y + radius);
    ctx.lineTo(x - radius, y + radius);
    ctx.closePath();
  } else {
    ctx.arc(x, y, radius, 0, Math.PI * 2);
  }
  ctx.fill();
  ctx.restore();
}

function getBiRenderedValuesForVisual(values, visualSettings, chartType) {
  const safeValues = Array.isArray(values)
    ? values.map((value) => (Number.isFinite(value) ? value : 0))
    : [];
  const safeType = normalizeBiChartType(chartType || "bar");
  const stackable = safeType === "bar" || safeType === "line" || safeType === "area" || safeType === "combo";
  const visual = normalizeBiVisualSettings(visualSettings || {});
  if (!stackable || visual.stackMode === "none" || safeValues.length === 0) {
    return safeValues;
  }
  if (visual.stackMode === "percent") {
    const positiveValues = safeValues.map((value) => Math.max(0, value));
    const total = positiveValues.reduce((sum, value) => sum + value, 0);
    if (total <= 0) {
      return safeValues.map(() => 0);
    }
    let cumulative = 0;
    return positiveValues.map((value) => {
      cumulative += value;
      return (cumulative / total) * 100;
    });
  }
  let cumulative = 0;
  return safeValues.map((value) => {
    cumulative += value;
    return cumulative;
  });
}

function resizeCanvasForDisplay(canvas, fallbackHeight) {
  const rawDprOverride = Number(canvas?.dataset?.biDpr || "");
  const hasDprOverride = Number.isFinite(rawDprOverride) && rawDprOverride > 0;
  const dpr = hasDprOverride ? rawDprOverride : (window.devicePixelRatio || 1);
  const widthOverride = Number(canvas?.dataset?.biRenderWidth || "");
  const heightOverride = Number(canvas?.dataset?.biRenderHeight || "");
  const widthBase = Number.isFinite(widthOverride) && widthOverride > 0
    ? widthOverride
    : (canvas.clientWidth || canvas.width || 340);
  const heightBase = Number.isFinite(heightOverride) && heightOverride > 0
    ? heightOverride
    : (canvas.clientHeight || canvas.height || fallbackHeight || 160);
  const width = Math.max(200, Math.floor(widthBase));
  const height = Math.max(120, Math.floor(heightBase));
  canvas.width = Math.floor(width * dpr);
  canvas.height = Math.floor(height * dpr);
  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  return { ctx, width, height };
}

function formatBiVisualNumber(value, settings, options = {}) {
  const visual = normalizeBiVisualSettings(settings);
  const decimals = sanitizeBiInteger(visual.valueDecimals, 1, 0, 4);
  const numeric = Number(value || 0);
  const usePercent = visual.numberFormat === "percent"
    || (visual.numberFormat === "auto" && options.percentHint === true);
  const useCurrencyPen = visual.numberFormat === "currency_pen";
  const useHours = visual.numberFormat === "hours";
  const baseText = numeric.toLocaleString("es-PE", {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals
  });
  let formatted = baseText;
  if (useCurrencyPen) {
    formatted = `S/ ${baseText}`;
  } else if (usePercent) {
    formatted = `${baseText}%`;
  } else if (useHours) {
    formatted = `${baseText} h`;
  }
  const prefix = trimOrFallback(visual.valuePrefix, "");
  const suffix = trimOrFallback(visual.valueSuffix, "");
  return `${prefix}${formatted}${suffix}`;
}

function truncateBiLabel(label, settings, fallbackMax) {
  const visual = normalizeBiVisualSettings(settings);
  const maxChars = sanitizeBiInteger(visual.labelMaxChars, fallbackMax, 4, 40);
  return trimOrFallback(label, "").slice(0, maxChars);
}

function drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, steps, color, settings) {
  const visual = normalizeBiVisualSettings(settings);
  const safeSteps = Math.max(2, steps || 5);
  ctx.save();
  const safeOpacity = sanitizeBiDecimal(visual.gridOpacity, 0.35, 0.1, 1);
  ctx.globalAlpha = safeOpacity;
  ctx.strokeStyle = color || "#e4ebf3";
  ctx.lineWidth = 1;
  const dash = sanitizeBiInteger(visual.gridDash, 4, 0, 12);
  if (dash > 0) {
    ctx.setLineDash([dash, dash]);
  } else {
    ctx.setLineDash([]);
  }
  for (let index = 0; index < safeSteps; index += 1) {
    const y = top + ((plotHeight / (safeSteps - 1)) * index);
    ctx.beginPath();
    ctx.moveTo(left, y);
    ctx.lineTo(left + plotWidth, y);
    ctx.stroke();
  }
  ctx.restore();
}

function resolveBiAxisRange(values, settings, options = {}) {
  const visual = normalizeBiVisualSettings(settings);
  const safeValues = Array.isArray(values)
    ? values.filter((value) => Number.isFinite(value))
    : [];
  if (safeValues.length === 0) {
    return { min: 0, max: 1, scale: "linear" };
  }

  let dataMin = Math.min(...safeValues);
  let dataMax = Math.max(...safeValues);
  if (options.forceZeroBaseline !== false) {
    dataMin = Math.min(0, dataMin);
  }
  let min = visual.axisMin === null ? dataMin : visual.axisMin;
  let max = visual.axisMax === null ? dataMax : visual.axisMax;
  if (!Number.isFinite(min)) {
    min = dataMin;
  }
  if (!Number.isFinite(max)) {
    max = dataMax;
  }
  if (max <= min) {
    max = min + 1;
  }
  const scale = visual.axisScale === "log" ? "log" : "linear";
  if (scale === "log") {
    const positiveValues = safeValues.filter((value) => value > 0);
    const minPositive = positiveValues.length ? Math.min(...positiveValues) : 1;
    if (min <= 0) {
      min = minPositive > 0 ? minPositive : 1;
    }
    if (max <= min) {
      max = min * 10;
    }
  }
  return { min, max, scale };
}

function normalizeBiAxisValue(value, axisRange) {
  const axis = axisRange || { min: 0, max: 1, scale: "linear" };
  const safeValue = Number.isFinite(value) ? value : 0;
  if (axis.scale === "log") {
    const min = axis.min > 0 ? axis.min : 1;
    const max = axis.max > min ? axis.max : (min * 10);
    const safe = safeValue > 0 ? safeValue : min;
    const minLog = Math.log10(min);
    const maxLog = Math.log10(max);
    const safeLog = Math.log10(Math.max(min, Math.min(max, safe)));
    return maxLog === minLog ? 0 : ((safeLog - minLog) / (maxLog - minLog));
  }
  const denominator = axis.max - axis.min;
  if (!Number.isFinite(denominator) || denominator <= 0) {
    return 0;
  }
  return (safeValue - axis.min) / denominator;
}

function drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, settings, options = {}) {
  if (!ctx) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  if (!visual.showTargetLine || visual.targetLineValue === null) {
    return;
  }
  const targetValue = Number(visual.targetLineValue);
  if (!Number.isFinite(targetValue)) {
    return;
  }
  if ((axisRange?.scale || "linear") === "log" && targetValue <= 0) {
    return;
  }
  const normalizedRaw = normalizeBiAxisValue(targetValue, axisRange);
  if (!Number.isFinite(normalizedRaw)) {
    return;
  }
  const normalized = Math.max(0, Math.min(1, normalizedRaw));
  const y = top + plotHeight - (normalized * plotHeight);
  const lineColor = normalizeBiColorHex(visual.targetLineColor, "#b03831");
  ctx.save();
  ctx.strokeStyle = lineColor;
  ctx.lineWidth = Math.max(1, visual.lineWidth - 1);
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  ctx.moveTo(left, y);
  ctx.lineTo(left + plotWidth, y);
  ctx.stroke();
  ctx.setLineDash([]);
  const label = trimOrFallback(visual.targetLineLabel, "Meta");
  const valueText = formatBiVisualNumber(targetValue, visual, { percentHint: options.percentHint === true });
  ctx.textAlign = "right";
  ctx.fillStyle = lineColor;
  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1), bold: true });
  const labelY = Math.max(12, Math.min(top + plotHeight - 2, y - 4));
  ctx.fillText(`${label}: ${valueText}`, left + plotWidth - 3, labelY);
  ctx.restore();
}

function drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, settings) {
  const visual = normalizeBiVisualSettings(settings);
  const axisTitleColor = resolveBiLabelColor(visual, "#ffffff", "#425264");
  const axisXTitle = trimOrFallback(visual.axisXLabel, "");
  const axisYTitle = trimOrFallback(visual.axisYLabel, "");
  if (axisXTitle) {
    ctx.save();
    ctx.textAlign = "center";
    ctx.fillStyle = axisTitleColor;
    applyBiCanvasFont(ctx, visual, "axis", { bold: true, fallbackSize: Math.max(10, visual.fontSizeAxis) });
    ctx.fillText(axisXTitle, left + (plotWidth / 2), top + plotHeight + 26);
    ctx.restore();
  }
  if (axisYTitle) {
    ctx.save();
    ctx.translate(Math.max(10, left - 34), top + (plotHeight / 2));
    ctx.rotate(-Math.PI / 2);
    ctx.textAlign = "center";
    ctx.fillStyle = axisTitleColor;
    applyBiCanvasFont(ctx, visual, "axis", { bold: true, fallbackSize: Math.max(10, visual.fontSizeAxis) });
    ctx.fillText(axisYTitle, 0, 0);
    ctx.restore();
  }
}

function resolveBiLegendAnchor(position, width, height, centerX, centerY, radius, limit) {
  const safeLimit = Math.max(1, limit || 1);
  const safePosition = trimOrFallback(position, "right").toLowerCase();
  if (safePosition === "left") {
    return { x: 8, y: Math.max(16, centerY - ((safeLimit - 1) * 10)), align: "left" };
  }
  if (safePosition === "top") {
    return { x: 8, y: 16, align: "left" };
  }
  if (safePosition === "bottom") {
    return { x: 8, y: Math.max(16, height - (safeLimit * 20) - 8), align: "left" };
  }
  return {
    x: Math.max(centerX + radius + 14, Math.floor(width * 0.58)),
    y: Math.max(16, centerY - ((safeLimit - 1) * 10)),
    align: "left"
  };
}

function clampBiValue(value, min, max) {
  const safeValue = Number(value);
  if (!Number.isFinite(safeValue)) {
    return min;
  }
  if (safeValue < min) {
    return min;
  }
  if (safeValue > max) {
    return max;
  }
  return safeValue;
}

function resolveBiCircularLayoutGeometry(ctx, width, height, legendTexts, visual, circularLayout) {
  const safeWidth = Math.max(20, Math.floor(width));
  const safeHeight = Math.max(20, Math.floor(height));
  const safeVisual = normalizeBiVisualSettings(visual);
  const safeLayout = normalizeBiCircularLayout(circularLayout);
  const padding = 8;
  const gap = 10;
  const legendFontSize = getBiFontSizeByLayer(safeVisual, "label", 10);
  const rowHeight = Math.max(16, Math.round(Math.max(9, legendFontSize) * 1.45));
  const available = {
    x: padding,
    y: padding,
    w: Math.max(20, safeWidth - (padding * 2)),
    h: Math.max(20, safeHeight - (padding * 2))
  };

  const legendItems = Array.isArray(legendTexts) ? legendTexts : [];
  const requestedLegendRows = Math.max(0, Math.min(
    legendItems.length,
    sanitizeBiInteger(safeVisual.legendMaxItems, 8, 3, 20)
  ));
  const legendPosition = trimOrFallback(safeVisual.legendPosition, "right").toLowerCase();
  const legendEnabled = !!safeVisual.showLegend && requestedLegendRows > 0;

  ctx.save();
  applyBiCanvasFont(ctx, safeVisual, "label", { fallbackSize: legendFontSize });
  const legendTextWidths = legendItems.map((text) => ctx.measureText(text).width);
  ctx.restore();
  const maxLegendTextWidth = legendTextWidths.length > 0 ? Math.ceil(Math.max(...legendTextWidths)) : 0;
  const legendWidthPreferred = Math.max(92, maxLegendTextWidth + 22);
  const legendHeightPreferred = Math.max(rowHeight + 6, requestedLegendRows * rowHeight + 8);

  let chartArea = { ...available };
  let legendArea = null;

  if (legendEnabled) {
    if (legendPosition === "left" || legendPosition === "right") {
      const maxSideWidth = Math.max(86, Math.floor(available.w * 0.55));
      const sideWidth = Math.max(86, Math.min(maxSideWidth, legendWidthPreferred));
      if (available.w - sideWidth - gap >= 96) {
        if (legendPosition === "left") {
          legendArea = { x: available.x, y: available.y, w: sideWidth, h: available.h };
          chartArea = {
            x: legendArea.x + legendArea.w + gap,
            y: available.y,
            w: Math.max(20, available.w - legendArea.w - gap),
            h: available.h
          };
        } else {
          legendArea = {
            x: available.x + available.w - sideWidth,
            y: available.y,
            w: sideWidth,
            h: available.h
          };
          chartArea = {
            x: available.x,
            y: available.y,
            w: Math.max(20, available.w - legendArea.w - gap),
            h: available.h
          };
        }
      }
    } else {
      const maxBandHeight = Math.max(rowHeight + 6, Math.floor(available.h * 0.46));
      const bandHeight = Math.max(rowHeight + 6, Math.min(maxBandHeight, legendHeightPreferred));
      if (available.h - bandHeight - gap >= 96) {
        if (legendPosition === "top") {
          legendArea = { x: available.x, y: available.y, w: available.w, h: bandHeight };
          chartArea = {
            x: available.x,
            y: legendArea.y + legendArea.h + gap,
            w: available.w,
            h: Math.max(20, available.h - legendArea.h - gap)
          };
        } else {
          legendArea = {
            x: available.x,
            y: available.y + available.h - bandHeight,
            w: available.w,
            h: bandHeight
          };
          chartArea = {
            x: available.x,
            y: available.y,
            w: available.w,
            h: Math.max(20, available.h - legendArea.h - gap)
          };
        }
      }
    }
  }

  const radius = Math.max(18, Math.floor((Math.min(chartArea.w, chartArea.h) / 2) - 6));
  const centerMinX = chartArea.x + radius + 2;
  const centerMaxX = chartArea.x + chartArea.w - radius - 2;
  const centerMinY = chartArea.y + radius + 2;
  const centerMaxY = chartArea.y + chartArea.h - radius - 2;
  const centerX = Math.round(clampBiValue(
    chartArea.x + (chartArea.w / 2) + safeLayout.chart.x,
    Math.min(centerMinX, centerMaxX),
    Math.max(centerMinX, centerMaxX)
  ));
  const centerY = Math.round(clampBiValue(
    chartArea.y + (chartArea.h / 2) + safeLayout.chart.y,
    Math.min(centerMinY, centerMaxY),
    Math.max(centerMinY, centerMaxY)
  ));

  let legendRows = [];
  let legendBox = null;
  if (legendArea) {
    const maxRowsByHeight = Math.max(0, Math.floor((legendArea.h - 6) / rowHeight));
    const visibleRows = Math.max(0, Math.min(requestedLegendRows, maxRowsByHeight));
    if (visibleRows > 0) {
      legendRows = legendItems.slice(0, visibleRows);
      const rowsWidths = legendRows.map((text) => {
        ctx.save();
        applyBiCanvasFont(ctx, safeVisual, "label", { fallbackSize: legendFontSize });
        const widthText = ctx.measureText(text).width;
        ctx.restore();
        return widthText;
      });
      const legendContentWidth = (rowsWidths.length > 0 ? Math.ceil(Math.max(...rowsWidths)) : 0) + 16;
      const baseX = legendPosition === "top" || legendPosition === "bottom"
        ? legendArea.x + ((legendArea.w - legendContentWidth) / 2)
        : legendArea.x + 4;
      const baseY = legendArea.y + ((legendArea.h - (visibleRows * rowHeight)) / 2) + Math.round(rowHeight * 0.72);
      const minLegendX = legendArea.x + 2;
      const maxLegendX = legendArea.x + legendArea.w - legendContentWidth - 2;
      const minLegendY = legendArea.y + Math.round(rowHeight * 0.72);
      const maxLegendY = legendArea.y + legendArea.h - (visibleRows * rowHeight) + Math.round(rowHeight * 0.72);
      const legendX = Math.round(clampBiValue(
        baseX + safeLayout.legend.x,
        Math.min(minLegendX, maxLegendX),
        Math.max(minLegendX, maxLegendX)
      ));
      const legendStartY = Math.round(clampBiValue(
        baseY + safeLayout.legend.y,
        Math.min(minLegendY, maxLegendY),
        Math.max(minLegendY, maxLegendY)
      ));
      legendBox = {
        x: legendX - 2,
        y: legendStartY - Math.round(rowHeight * 0.72),
        w: legendContentWidth + 4,
        h: visibleRows * rowHeight
      };
      legendRows = legendRows.map((text, index) => ({
        text,
        x: legendX,
        y: legendStartY + (index * rowHeight)
      }));
    }
  }

  return {
    chartArea,
    centerX,
    centerY,
    radius,
    legendRows,
    legendBox
  };
}

function drawBiSmoothedPath(ctx, points) {
  if (!Array.isArray(points) || points.length === 0) {
    return;
  }
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  if (points.length === 1) {
    return;
  }
  for (let index = 0; index < points.length - 1; index += 1) {
    const p0 = points[index - 1] || points[index];
    const p1 = points[index];
    const p2 = points[index + 1];
    const p3 = points[index + 2] || p2;
    const cp1x = p1.x + ((p2.x - p0.x) / 6);
    const cp1y = p1.y + ((p2.y - p0.y) / 6);
    const cp2x = p2.x - ((p3.x - p1.x) / 6);
    const cp2y = p2.y - ((p3.y - p1.y) / 6);
    ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
  }
}

function drawBiLinearPath(ctx, points) {
  if (!Array.isArray(points) || points.length === 0) {
    return;
  }
  ctx.beginPath();
  points.forEach((point, index) => {
    if (index === 0) {
      ctx.moveTo(point.x, point.y);
    } else {
      ctx.lineTo(point.x, point.y);
    }
  });
}

function drawBiSeriesPath(ctx, points, smoothLines) {
  if (smoothLines) {
    drawBiSmoothedPath(ctx, points);
    return;
  }
  drawBiLinearPath(ctx, points);
}

function drawBiBarChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "bar");
  const minLabelValue = sanitizeBiDecimal(specific.labelMinValue, 0, -1000000, 1000000);
  const percentHint = visual.stackMode === "percent";
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const axisRange = resolveBiAxisRange(values, visual, { forceZeroBaseline: true });
  const left = (visual.axisYLabel ? 58 : 44);
  const right = 10;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, values.length);
  const barWidth = Math.max(4, slot * sanitizeBiDecimal(visual.barWidthRatio, 0.64, 0.2, 0.9));

  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.stroke();

  const fontSize = Math.max(9, visual.fontSize);
  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: fontSize });
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(axisRange.max, visual, { percentHint }), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(axisRange.min, visual, { percentHint }), left - 4, top + plotHeight + 2);
  }

  values.forEach((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const barHeight = Math.max(0, normalized * plotHeight);
    const x = left + slot * index + ((slot - barWidth) / 2);
    const y = top + plotHeight - barHeight;
    ctx.fillStyle = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    ctx.fillRect(x, y, barWidth, barHeight);
    if (visual.showDataLabels && value >= minLabelValue) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
      ctx.fillText(formatBiVisualNumber(value, visual, { percentHint }), x + (barWidth / 2), Math.max(10, y - 4));
    }
    if (visual.showAxisLabels && labels.length <= 12) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: fontSize });
      ctx.fillText(truncateBiLabel(labels[index], visual, 8), x + (barWidth / 2), top + plotHeight + 12);
    }
  });
  drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, visual, { percentHint });
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiLineChart(ctx, width, height, labels, values, colors, settings, labelOffsets, capture, rowsData, labelLayoutV2, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }

  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "line");
  const minLabelValue = sanitizeBiDecimal(specific.labelMinValue, 0, -1000000, 1000000);
  const percentHint = visual.stackMode === "percent";
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const offsets = normalizeBiLabelOffsets(labelOffsets);
  const layoutV2 = normalizeBiLabelLayoutV2(labelLayoutV2);
  const safeRowsData = Array.isArray(rowsData) ? rowsData : [];
  const collector = capture && typeof capture === "object" ? capture : null;
  if (collector) {
    collector.labelItems = [];
  }
  const axisRange = resolveBiAxisRange(values, visual, { forceZeroBaseline: false });
  const left = (visual.axisYLabel ? 58 : 44);
  const right = 10;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const stepX = values.length <= 1 ? 0 : (plotWidth / (values.length - 1));

  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.stroke();

  const lineColor = normalizeBiColorHex(colors?.[0], "#2f7ed8");
  ctx.strokeStyle = lineColor;
  ctx.lineWidth = visual.lineWidth;
  ctx.setLineDash(resolveBiLineDashPattern(visual));
  const points = values.map((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const x = left + (stepX * index);
    const y = top + plotHeight - (Math.max(0, normalized) * plotHeight);
    return { x, y, value, label: labels[index] || "" };
  });
  drawBiSeriesPath(ctx, points, visual.smoothLines);
  ctx.stroke();
  ctx.setLineDash([]);

  points.forEach((point, index) => {
    const markerColor = normalizeBiColorHex(colors?.[index], lineColor);
    drawBiMarker(ctx, point.x, point.y, visual.seriesMarkerStyle, visual.seriesMarkerSize, markerColor);
    if (visual.showDataLabels && point.value >= minLabelValue) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
      const text = formatBiVisualNumber(point.value, visual, { percentHint });
      const rowKey = trimOrFallback(safeRowsData[index]?.rowKey, "");
      const fallbackKey = buildBiLabelFallbackKey(point.label);
      const placement = getBiLabelPlacementFromLayout(layoutV2, rowKey, point.label || fallbackKey || "");
      const legacyOffset = offsets[String(index)] || { x: 0, y: 0 };
      const offsetX = placement
        ? (placement.dxPct * plotWidth)
        : (Number.isFinite(legacyOffset.x) ? legacyOffset.x : 0);
      const offsetY = placement
        ? (placement.dyPct * plotHeight)
        : (Number.isFinite(legacyOffset.y) ? legacyOffset.y : 0);
      const labelX = point.x + offsetX;
      const labelY = Math.max(10, point.y - 6) + offsetY;
      ctx.fillText(text, labelX, labelY);
      if (collector) {
        const textWidth = ctx.measureText(text).width;
        const textHeight = Math.max(10, visual.fontSize + 2);
        collector.labelItems.push({
          index,
          rowKey,
          rowLabel: point.label,
          text,
          x: labelX - (textWidth / 2) - 4,
          y: labelY - textHeight - 1,
          w: textWidth + 8,
          h: textHeight + 6,
          anchorX: labelX,
          anchorY: labelY,
          plotWidth,
          plotHeight
        });
      }
    }
  });

  drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, visual, { percentHint });

  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(axisRange.max, visual, { percentHint }), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(axisRange.min, visual, { percentHint }), left - 4, top + plotHeight + 2);
  }

  if (visual.showAxisLabels && labels.length <= 12) {
    ctx.textAlign = "center";
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisTextColor;
    labels.forEach((label, index) => {
      const x = left + (stepX * index);
      ctx.fillText(truncateBiLabel(label, visual, 8), x, top + plotHeight + 12);
    });
  }
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiDonutChart(ctx, width, height, labels, values, colors, settings, circularLayout, capture, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    if (capture && typeof capture === "object") {
      capture.polarModel = { type: "donut", items: [], chartDragRect: null, legendDragRect: null };
    }
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "donut");
  const legendTextColor = resolveBiLabelColor(visual, "#ffffff", "#4a5d73");

  const safeValues = values.map((value) => (Number.isFinite(value) && value > 0 ? value : 0));
  const total = safeValues.reduce((sum, value) => sum + value, 0);
  if (total <= 0) {
    if (capture && typeof capture === "object") {
      capture.polarModel = { type: "donut", items: [], chartDragRect: null, legendDragRect: null };
    }
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 12 });
    ctx.fillStyle = "#748395";
    ctx.textAlign = "center";
    ctx.fillText("Sin datos", Math.floor(width / 2), Math.floor(height / 2));
    return;
  }

  const legendLimit = Math.min(labels.length, sanitizeBiInteger(visual.legendMaxItems, 8, 3, 20));
  const legendTexts = labels
    .slice(0, legendLimit)
    .map((label, index) => `${truncateBiLabel(label, visual, 16)} (${formatBiVisualNumber(safeValues[index], visual)})`);
  const layout = resolveBiCircularLayoutGeometry(ctx, width, height, legendTexts, visual, circularLayout);
  const centerX = layout.centerX;
  const centerY = layout.centerY;
  const radius = layout.radius;
  const innerRadius = Math.floor(radius * sanitizeBiDecimal(specific.innerRadiusRatio, 0.58, 0.35, 0.8));
  const minLabelPercent = sanitizeBiDecimal(specific.labelMinPercent, 4, 0, 100);
  let start = -Math.PI / 2;
  const arcItems = [];

  safeValues.forEach((value, index) => {
    const angle = (value / total) * Math.PI * 2;
    const end = start + angle;
    const sliceColor = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    arcItems.push({
      index,
      kind: "arc",
      cx: centerX,
      cy: centerY,
      innerRadius,
      outerRadius: radius,
      startAngle: start,
      endAngle: end
    });
    ctx.fillStyle = sliceColor;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, start, end);
    ctx.closePath();
    ctx.fill();
    if (visual.showDataLabels) {
      const mid = start + (angle / 2);
      const labelRadius = innerRadius + ((radius - innerRadius) * 0.65);
      const labelX = centerX + (Math.cos(mid) * labelRadius);
      const labelY = centerY + (Math.sin(mid) * labelRadius);
      const pct = (value / total) * 100;
      if (pct >= minLabelPercent) {
        ctx.fillStyle = resolveBiLabelColor(visual, sliceColor, "#1f2f44");
        applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
        ctx.textAlign = "center";
        ctx.fillText(formatPercent(pct, visual.valueDecimals), labelX, labelY);
      }
    }
    start = end;
  });

  ctx.fillStyle = "#fff";
  ctx.beginPath();
  ctx.arc(centerX, centerY, innerRadius, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = resolveBiLabelColor(visual, "#ffffff", "#3b4a5f");
  applyBiCanvasFont(ctx, visual, "title", { bold: true, fallbackSize: 12 });
  ctx.textAlign = "center";
  ctx.fillText(trimOrFallback(specific.centerTitle, "Total"), centerX, centerY - 2);
  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 11 });
  ctx.fillText(formatBiVisualNumber(total, visual), centerX, centerY + 14);

  if (layout.legendRows.length > 0) {
    ctx.textAlign = "left";
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels) });
    for (let index = 0; index < layout.legendRows.length; index += 1) {
      const row = layout.legendRows[index];
      const y = row.y;
      ctx.fillStyle = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
      ctx.fillRect(row.x, y - 8, 10, 10);
      ctx.fillStyle = legendTextColor;
      ctx.fillText(row.text, row.x + 14, y);
    }
  }

  if (capture && typeof capture === "object") {
    capture.polarModel = {
      type: "donut",
      items: arcItems,
      chartDragRect: {
        x: centerX - radius,
        y: centerY - radius,
        w: radius * 2,
        h: radius * 2
      },
      legendDragRect: layout.legendBox
    };
  }
}

function drawBiPieChart(ctx, width, height, labels, values, colors, settings, circularLayout, capture, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    if (capture && typeof capture === "object") {
      capture.polarModel = { type: "pie", items: [], chartDragRect: null, legendDragRect: null };
    }
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "pie");
  const legendTextColor = resolveBiLabelColor(visual, "#ffffff", "#4a5d73");
  const minLabelPercent = sanitizeBiDecimal(specific.labelMinPercent, 4, 0, 100);
  const safeValues = values.map((value) => (Number.isFinite(value) && value > 0 ? value : 0));
  const total = safeValues.reduce((sum, value) => sum + value, 0);
  if (total <= 0) {
    if (capture && typeof capture === "object") {
      capture.polarModel = { type: "pie", items: [], chartDragRect: null, legendDragRect: null };
    }
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 12 });
    ctx.fillStyle = "#748395";
    ctx.textAlign = "center";
    ctx.fillText("Sin datos", Math.floor(width / 2), Math.floor(height / 2));
    return;
  }

  const legendLimit = Math.min(labels.length, sanitizeBiInteger(visual.legendMaxItems, 8, 3, 20));
  const legendTexts = labels
    .slice(0, legendLimit)
    .map((label, index) => {
      const pct = total > 0 ? ((safeValues[index] / total) * 100) : 0;
      return `${truncateBiLabel(label, visual, 12)} ${formatPercent(pct, visual.valueDecimals)}`;
    });
  const layout = resolveBiCircularLayoutGeometry(ctx, width, height, legendTexts, visual, circularLayout);
  const centerX = layout.centerX;
  const centerY = layout.centerY;
  const radius = layout.radius;
  const startOffset = sanitizeBiDecimal(specific.startAngle, -90, -360, 360);
  let start = (startOffset * Math.PI) / 180;
  const arcItems = [];
  safeValues.forEach((value, index) => {
    const angle = (value / total) * Math.PI * 2;
    const end = start + angle;
    const sliceColor = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    arcItems.push({
      index,
      kind: "arc",
      cx: centerX,
      cy: centerY,
      innerRadius: 0,
      outerRadius: radius,
      startAngle: start,
      endAngle: end
    });
    ctx.fillStyle = sliceColor;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, start, end);
    ctx.closePath();
    ctx.fill();
    if (visual.showDataLabels) {
      const pct = total > 0 ? ((value / total) * 100) : 0;
      if (pct >= minLabelPercent) {
        const mid = start + (angle / 2);
        const labelX = centerX + (Math.cos(mid) * (radius * 0.65));
        const labelY = centerY + (Math.sin(mid) * (radius * 0.65));
        ctx.fillStyle = resolveBiLabelColor(visual, sliceColor, "#1f2f44");
        applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
        ctx.textAlign = "center";
        ctx.fillText(formatPercent(pct, visual.valueDecimals), labelX, labelY);
      }
    }
    start = end;
  });

  if (layout.legendRows.length > 0) {
    ctx.textAlign = "left";
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels) });
    for (let index = 0; index < layout.legendRows.length; index += 1) {
      const row = layout.legendRows[index];
      const y = row.y;
      ctx.fillStyle = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
      ctx.fillRect(row.x, y - 8, 10, 10);
      ctx.fillStyle = legendTextColor;
      ctx.fillText(row.text, row.x + 14, y);
    }
  }

  if (capture && typeof capture === "object") {
    capture.polarModel = {
      type: "pie",
      items: arcItems,
      chartDragRect: {
        x: centerX - radius,
        y: centerY - radius,
        w: radius * 2,
        h: radius * 2
      },
      legendDragRect: layout.legendBox
    };
  }
}

function drawBiAreaChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }

  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "area");
  const minLabelValue = sanitizeBiDecimal(specific.labelMinValue, 0, -1000000, 1000000);
  const percentHint = visual.stackMode === "percent";
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const axisRange = resolveBiAxisRange(values, visual, { forceZeroBaseline: false });
  const left = (visual.axisYLabel ? 58 : 44);
  const right = 10;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const stepX = values.length <= 1 ? 0 : (plotWidth / (values.length - 1));

  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.stroke();

  const points = values.map((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const x = left + (stepX * index);
    const y = top + plotHeight - (Math.max(0, normalized) * plotHeight);
    return { x, y, value, label: labels[index] || "" };
  });
  drawBiSeriesPath(ctx, points, visual.smoothLines);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.lineTo(left, top + plotHeight);
  ctx.closePath();
  const areaColor = normalizeBiColorHex(colors?.[0], "#2f7ed8");
  ctx.fillStyle = hexToRgba(areaColor, visual.areaOpacity);
  ctx.fill();

  ctx.strokeStyle = areaColor;
  ctx.lineWidth = visual.lineWidth;
  ctx.setLineDash(resolveBiLineDashPattern(visual));
  drawBiSeriesPath(ctx, points, visual.smoothLines);
  ctx.stroke();
  ctx.setLineDash([]);
  points.forEach((point) => {
    drawBiMarker(ctx, point.x, point.y, visual.seriesMarkerStyle, visual.seriesMarkerSize, areaColor);
  });

  if (visual.showDataLabels) {
    ctx.textAlign = "center";
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
    points.forEach((point) => {
      if (point.value < minLabelValue) {
        return;
      }
      ctx.fillStyle = labelTextColor;
      ctx.fillText(formatBiVisualNumber(point.value, visual, { percentHint }), point.x, Math.max(10, point.y - 6));
    });
  }

  drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, visual, { percentHint });

  if (visual.showAxisLabels && labels.length <= 12) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "center";
    labels.forEach((label, index) => {
      const x = left + (stepX * index);
      ctx.fillText(truncateBiLabel(label, visual, 8), x, top + plotHeight + 12);
    });
  }
  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(axisRange.max, visual, { percentHint }), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(axisRange.min, visual, { percentHint }), left - 4, top + plotHeight + 2);
  }
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiComboChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }

  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "combo");
  const minLabelValue = sanitizeBiDecimal(specific.labelMinValue, 0, -1000000, 1000000);
  const percentHint = visual.stackMode === "percent";
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const axisRange = resolveBiAxisRange(values, visual, { forceZeroBaseline: false });
  const left = (visual.axisYLabel ? 58 : 44);
  const right = 10;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, values.length);
  const barWidth = Math.max(6, slot * sanitizeBiDecimal(visual.barWidthRatio, 0.64, 0.2, 0.9));
  const stepX = values.length <= 1 ? 0 : (plotWidth / (values.length - 1));

  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.stroke();

  values.forEach((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const barHeight = Math.max(0, normalized * plotHeight);
    const x = left + slot * index + ((slot - barWidth) / 2);
    const y = top + plotHeight - barHeight;
    const slotColor = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    ctx.fillStyle = hexToRgba(slotColor, 0.45, "rgba(58, 136, 220, 0.45)");
    ctx.fillRect(x, y, barWidth, barHeight);
    if (visual.showDataLabels && value >= minLabelValue) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
      ctx.fillText(formatBiVisualNumber(value, visual, { percentHint }), x + (barWidth / 2), Math.max(10, y - 4));
    }
  });

  const comboLineColor = normalizeBiColorHex(colors?.[0], "#1f5ea8");
  ctx.strokeStyle = comboLineColor;
  ctx.lineWidth = visual.lineWidth;
  ctx.setLineDash(resolveBiLineDashPattern(visual));
  const points = values.map((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const x = left + (stepX * index);
    const y = top + plotHeight - (Math.max(0, normalized) * plotHeight);
    return { x, y, value, label: labels[index] || "" };
  });
  drawBiSeriesPath(ctx, points, visual.smoothLines);
  ctx.stroke();
  ctx.setLineDash([]);
  points.forEach((point) => {
    drawBiMarker(ctx, point.x, point.y, visual.seriesMarkerStyle, visual.seriesMarkerSize, comboLineColor);
  });
  drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, visual, { percentHint });
  if (visual.showAxisLabels && labels.length <= 12) {
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "center";
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    labels.forEach((label, index) => {
      const x = left + (stepX * index);
      ctx.fillText(truncateBiLabel(label, visual, 8), x, top + plotHeight + 12);
    });
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(axisRange.max, visual, { percentHint }), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(axisRange.min, visual, { percentHint }), left - 4, top + plotHeight + 2);
  }
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiTreemapChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "treemap");
  const defaultTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const minTilePercent = sanitizeBiDecimal(specific.minTilePercent, 1, 0, 25) / 100;
  const showValueText = normalizeBiToggle(specific.showValueText, true);
  const safeValues = values.map((value) => (Number.isFinite(value) && value > 0 ? value : 0));
  const total = safeValues.reduce((sum, value) => sum + value, 0);
  if (total <= 0) {
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 12 });
    ctx.fillStyle = "#748395";
    ctx.textAlign = "center";
    ctx.fillText("Sin datos", Math.floor(width / 2), Math.floor(height / 2));
    return;
  }

  const left = 8;
  const top = 8;
  const plotWidth = Math.max(20, width - 16);
  const plotHeight = Math.max(20, height - 16);
  let offsetX = left;
  safeValues.forEach((value, index) => {
    const ratio = value / total;
    if (ratio < minTilePercent) {
      return;
    }
    const blockWidth = Math.max(18, Math.round(plotWidth * ratio));
    const finalWidth = Math.min(blockWidth, left + plotWidth - offsetX);
    const blockColor = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    ctx.fillStyle = blockColor;
    ctx.fillRect(offsetX, top, finalWidth, plotHeight);
    ctx.fillStyle = visual.labelColorMode === "manual"
      ? defaultTextColor
      : resolveBiLabelColor(visual, blockColor, "#ffffff");
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels) });
    const labelText = truncateBiLabel(labels[index], visual, 12);
    const valueText = (visual.showDataLabels && showValueText) ? ` ${formatBiVisualNumber(safeValues[index], visual)}` : "";
    ctx.fillText(`${labelText}${valueText}`, offsetX + 6, top + 18);
    offsetX += finalWidth;
    if (offsetX >= left + plotWidth) {
      offsetX = left + plotWidth;
    }
  });
}

function drawBiFunnelChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "funnel");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#24384f");
  const percentHint = visual.stackMode === "percent";
  const minSegmentRatio = sanitizeBiDecimal(specific.minSegmentPercent, 2, 0, 60) / 100;
  const sortMode = trimOrFallback(specific.sortMode, "desc").toLowerCase() === "asc" ? "asc" : "desc";
  const pairs = labels.map((label, index) => ({
    label,
    value: Number.isFinite(values[index]) ? values[index] : 0,
    color: normalizeBiColorHex(colors?.[index], getBiPaletteColor(index))
  })).sort((a, b) => sortMode === "asc" ? a.value - b.value : b.value - a.value);
  const sortedLabels = pairs.map((item) => item.label);
  const sortedValues = pairs.map((item) => item.value);
  const sortedColors = pairs.map((item) => item.color);
  const max = Math.max(1, ...sortedValues);
  const topY = 10;
  const rowHeight = Math.max(20, Math.floor((height - 20) / sortedValues.length));
  const centerX = Math.floor(width * 0.45);
  const maxWidth = Math.max(40, Math.floor(width * 0.72));
  sortedValues.forEach((value, index) => {
    const ratio = Math.max(0, value / max);
    if (ratio < minSegmentRatio) {
      return;
    }
    const topWidth = Math.max(20, maxWidth * ratio);
    const nextRatio = index === sortedValues.length - 1 ? 0.18 : Math.max(0, sortedValues[index + 1] / max);
    const bottomWidth = Math.max(16, maxWidth * nextRatio);
    const y = topY + (index * rowHeight);

    ctx.fillStyle = sortedColors[index];
    ctx.beginPath();
    ctx.moveTo(centerX - (topWidth / 2), y);
    ctx.lineTo(centerX + (topWidth / 2), y);
    ctx.lineTo(centerX + (bottomWidth / 2), y + rowHeight - 2);
    ctx.lineTo(centerX - (bottomWidth / 2), y + rowHeight - 2);
    ctx.closePath();
    ctx.fill();

    ctx.fillStyle = labelTextColor;
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels) });
    ctx.fillText(truncateBiLabel(sortedLabels[index], visual, 16), 8, y + (rowHeight / 2));
    if (visual.showDataLabels) {
      ctx.textAlign = "right";
      ctx.fillText(formatBiVisualNumber(value, visual, { percentHint }), width - 8, y + (rowHeight / 2));
    }
    ctx.textAlign = "left";
  });
}

function drawBiGaugeChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "gauge");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#2c3e50");
  const firstValue = Number.isFinite(values[0]) ? values[0] : 0;
  const minValue = sanitizeBiDecimal(specific.minValue, 0, -1000000, 1000000);
  const maxValue = sanitizeBiDecimal(specific.maxValue, 100, minValue, 1000000);
  const targetValue = sanitizeBiDecimal(specific.targetValue, 80, minValue, maxValue);
  const ratio = Math.max(0, Math.min(1, ((firstValue - minValue) / Math.max(1e-6, maxValue - minValue))));
  const targetRatio = Math.max(0, Math.min(1, ((targetValue - minValue) / Math.max(1e-6, maxValue - minValue))));
  const centerX = Math.floor(width / 2);
  const centerY = Math.floor(height * 0.78);
  const radius = Math.max(40, Math.floor(Math.min(width * 0.42, height * 0.70)));
  const start = Math.PI;
  const end = 0;

  ctx.lineWidth = Math.max(10, visual.lineWidth * 5);
  ctx.strokeStyle = "#dde5ef";
  ctx.beginPath();
  ctx.arc(centerX, centerY, radius, start, end, false);
  ctx.stroke();

  const gaugeColor = normalizeBiColorHex(colors?.[0], "");
  ctx.strokeStyle = gaugeColor || (ratio < 0.4 ? "#e67e22" : (ratio < 0.75 ? "#f1c40f" : "#24a36c"));
  ctx.beginPath();
  ctx.arc(centerX, centerY, radius, start, start - (Math.PI * ratio), true);
  ctx.stroke();
  const targetAngle = start - (Math.PI * targetRatio);
  const targetInner = radius - Math.max(8, visual.lineWidth * 3.6);
  const targetOuter = radius + Math.max(2, visual.lineWidth * 1.4);
  ctx.strokeStyle = normalizeBiColorHex(visual.targetLineColor, "#b03831");
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(centerX + (Math.cos(targetAngle) * targetInner), centerY + (Math.sin(targetAngle) * targetInner));
  ctx.lineTo(centerX + (Math.cos(targetAngle) * targetOuter), centerY + (Math.sin(targetAngle) * targetOuter));
  ctx.stroke();

  ctx.fillStyle = labelTextColor;
  applyBiCanvasFont(ctx, visual, "title", { bold: true, fallbackSize: 22 });
  ctx.textAlign = "center";
  ctx.fillText(formatBiVisualNumber(firstValue, visual), centerX, centerY - 10);
  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 11 });
  ctx.fillStyle = labelTextColor;
  ctx.fillText(truncateBiLabel(labels[0], visual, 22), centerX, centerY + 14);
}

function drawBiWaterfallChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "waterfall");
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const cumulatives = [];
  let sum = 0;
  values.forEach((value) => {
    sum += value;
    cumulatives.push(sum);
  });
  const max = Math.max(1, ...cumulatives.map((value) => Math.abs(value)));
  const left = (visual.axisYLabel ? 58 : 44);
  const right = 10;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, values.length);
  const barWidth = Math.max(6, slot * 0.58);
  const zeroY = top + (plotHeight / 2);

  ctx.strokeStyle = "#d6dde6";
  ctx.beginPath();
  ctx.moveTo(left, zeroY);
  ctx.lineTo(left + plotWidth, zeroY);
  ctx.stroke();

  let previous = 0;
  values.forEach((value, index) => {
    const current = previous + value;
    const yFrom = zeroY - ((previous / max) * (plotHeight / 2));
    const yTo = zeroY - ((current / max) * (plotHeight / 2));
    const y = Math.min(yFrom, yTo);
    const h = Math.max(2, Math.abs(yTo - yFrom));
    const x = left + slot * index + ((slot - barWidth) / 2);
    const defaultColor = value >= 0 ? "#2f7ed8" : "#e06b6b";
    ctx.fillStyle = normalizeBiColorHex(colors?.[index], defaultColor);
    ctx.fillRect(x, y, barWidth, h);
    if (visual.showDataLabels) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels - 1) });
      ctx.fillText(formatBiVisualNumber(value, visual), x + (barWidth / 2), Math.max(10, y - 4));
    }
    previous = current;
  });
  if (normalizeBiToggle(specific.showTotalBar, true)) {
    const total = cumulatives[cumulatives.length - 1] || 0;
    ctx.fillStyle = "#1f4d78";
    applyBiCanvasFont(ctx, visual, "label", { bold: true, fallbackSize: Math.max(9, visual.fontSizeLabels) });
    ctx.textAlign = "right";
    ctx.fillText(`${trimOrFallback(specific.totalLabel, "Total")}: ${formatBiVisualNumber(total, visual)}`, width - 6, 14);
  }
  if (visual.showAxisLabels && labels.length <= 12) {
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "center";
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    labels.forEach((label, index) => {
      const x = left + slot * index + ((slot - barWidth) / 2) + (barWidth / 2);
      ctx.fillText(truncateBiLabel(label, visual, 8), x, top + plotHeight + 12);
    });
  }
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiRadarChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "radar");
  const axisColor = resolveBiLabelColor(visual, "#ffffff", "#5a6d84");
  const labelColor = resolveBiLabelColor(visual, "#ffffff", "#2e4258");
  const safeValues = values.map((value) => (Number.isFinite(value) ? Math.max(0, value) : 0));
  const maxRaw = Math.max(1, ...safeValues);
  const axisRange = resolveBiAxisRange(safeValues, { ...visual, axisMin: 0 }, { forceZeroBaseline: true });
  const maxScale = sanitizeBiDecimal(specific.maxScale, maxRaw, 1, 100000);
  const maxValue = Math.max(1, axisRange.max, maxScale);
  const showPointValues = normalizeBiToggle(specific.showPointValues, true);
  const count = Math.max(3, safeValues.length);
  const centerX = Math.floor(width / 2);
  const centerY = Math.floor(height / 2) + 6;
  const radius = Math.max(24, Math.floor(Math.min(width, height) * 0.33));
  const ringCount = 5;

  if (visual.showGrid) {
    ctx.strokeStyle = hexToRgba("#d9e5f2", visual.gridOpacity, "rgba(217,229,242,0.6)");
    ctx.lineWidth = 1;
    const dash = sanitizeBiInteger(visual.gridDash, 4, 0, 12);
    ctx.setLineDash(dash > 0 ? [dash, dash] : []);
    for (let ring = 1; ring <= ringCount; ring += 1) {
      const ringRadius = (radius * ring) / ringCount;
      ctx.beginPath();
      for (let index = 0; index < count; index += 1) {
        const angle = (-Math.PI / 2) + ((Math.PI * 2 * index) / count);
        const x = centerX + Math.cos(angle) * ringRadius;
        const y = centerY + Math.sin(angle) * ringRadius;
        if (index === 0) {
          ctx.moveTo(x, y);
        } else {
          ctx.lineTo(x, y);
        }
      }
      ctx.closePath();
      ctx.stroke();
    }
    ctx.setLineDash([]);
  }

  ctx.strokeStyle = "#cfd9e7";
  ctx.lineWidth = 1;
  for (let index = 0; index < count; index += 1) {
    const angle = (-Math.PI / 2) + ((Math.PI * 2 * index) / count);
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.lineTo(centerX + Math.cos(angle) * radius, centerY + Math.sin(angle) * radius);
    ctx.stroke();
  }

  const lineColor = normalizeBiColorHex(colors?.[0], "#2f7ed8");
  const points = safeValues.map((value, index) => {
    const normalized = Math.max(0, Math.min(1, value / maxValue));
    const angle = (-Math.PI / 2) + ((Math.PI * 2 * index) / count);
    const px = centerX + Math.cos(angle) * (radius * normalized);
    const py = centerY + Math.sin(angle) * (radius * normalized);
    return { index, x: px, y: py, value, angle };
  });

  ctx.fillStyle = hexToRgba(lineColor, Math.max(0.08, Math.min(0.5, visual.areaOpacity + 0.1)), "rgba(47,126,216,0.25)");
  ctx.strokeStyle = lineColor;
  ctx.lineWidth = visual.lineWidth;
  ctx.setLineDash(resolveBiLineDashPattern(visual));
  ctx.beginPath();
  points.forEach((point, index) => {
    if (index === 0) {
      ctx.moveTo(point.x, point.y);
    } else {
      ctx.lineTo(point.x, point.y);
    }
  });
  ctx.closePath();
  ctx.fill();
  ctx.stroke();
  ctx.setLineDash([]);

  points.forEach((point, index) => {
    const pointColor = normalizeBiColorHex(colors?.[index], lineColor);
    drawBiMarker(ctx, point.x, point.y, visual.seriesMarkerStyle, visual.seriesMarkerSize, pointColor);
  });

  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisColor;
    ctx.textAlign = "left";
    for (let ring = 1; ring <= ringCount; ring += 1) {
      const value = (maxValue * ring) / ringCount;
      ctx.fillText(formatBiVisualNumber(value, visual), centerX + 6, centerY - ((radius * ring) / ringCount) + 2);
    }
  }

  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(9, visual.fontSizeLabels) });
  labels.forEach((label, index) => {
    const angle = (-Math.PI / 2) + ((Math.PI * 2 * index) / count);
    const labelRadius = radius + 14;
    const lx = centerX + Math.cos(angle) * labelRadius;
    const ly = centerY + Math.sin(angle) * labelRadius;
    ctx.fillStyle = labelColor;
    ctx.textAlign = Math.cos(angle) > 0.3 ? "left" : (Math.cos(angle) < -0.3 ? "right" : "center");
    ctx.fillText(truncateBiLabel(label, visual, 10), lx, ly);
  });

  if (visual.showDataLabels && showPointValues) {
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
    ctx.textAlign = "center";
    points.forEach((point) => {
      ctx.fillStyle = labelColor;
      ctx.fillText(formatBiVisualNumber(point.value, visual), point.x, point.y - 8);
    });
  }
}

function drawBiParetoChart(ctx, width, height, labels, values, colors, settings, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "pareto");
  const showCumulativeLine = normalizeBiToggle(specific.showCumulativeLine, true);
  const targetPercent = sanitizeBiDecimal(specific.targetPercent, 80, 1, 100);
  const axisTextColor = resolveBiLabelColor(visual, "#ffffff", "#6a7686");
  const labelTextColor = resolveBiLabelColor(visual, "#ffffff", "#425064");
  const safeValues = values.map((value) => (Number.isFinite(value) ? Math.max(0, value) : 0));
  const total = safeValues.reduce((sum, value) => sum + value, 0);
  const cumulativePct = [];
  let accumulated = 0;
  safeValues.forEach((value) => {
    accumulated += value;
    cumulativePct.push(total <= 0 ? 0 : ((accumulated / total) * 100));
  });

  const axisRange = resolveBiAxisRange(safeValues, { ...visual, axisMin: 0 }, { forceZeroBaseline: true });
  const left = 44;
  const right = 44;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, safeValues.length);
  const barWidth = Math.max(5, slot * sanitizeBiDecimal(visual.barWidthRatio, 0.62, 0.2, 0.9));
  const lineColor = normalizeBiColorHex(colors?.[0], "#235fa3");

  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.lineTo(left + plotWidth, top);
  ctx.stroke();

  safeValues.forEach((value, index) => {
    const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
    const barHeight = Math.max(0, normalized * plotHeight);
    const x = left + slot * index + ((slot - barWidth) / 2);
    const y = top + plotHeight - barHeight;
    ctx.fillStyle = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    ctx.fillRect(x, y, barWidth, barHeight);
    if (visual.showDataLabels) {
      ctx.fillStyle = labelTextColor;
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
      ctx.fillText(formatBiVisualNumber(value, visual), x + (barWidth / 2), Math.max(10, y - 4));
    }
  });

  const points = cumulativePct.map((pct, index) => {
    const x = left + (slot * index) + (slot / 2);
    const y = top + plotHeight - ((pct / 100) * plotHeight);
    return { x, y, value: pct };
  });
  if (showCumulativeLine) {
    ctx.strokeStyle = lineColor;
    ctx.lineWidth = visual.lineWidth;
    ctx.setLineDash(resolveBiLineDashPattern(visual));
    drawBiSeriesPath(ctx, points, visual.smoothLines);
    ctx.stroke();
    ctx.setLineDash([]);
    points.forEach((point) => {
      drawBiMarker(ctx, point.x, point.y, visual.seriesMarkerStyle, visual.seriesMarkerSize, lineColor);
    });
  }

  if (visual.showDataLabels && showCumulativeLine) {
    ctx.textAlign = "center";
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
    points.forEach((point) => {
      ctx.fillStyle = lineColor;
      ctx.fillText(formatPercent(point.value, visual.valueDecimals), point.x, Math.max(10, point.y - 8));
    });
  }

  drawBiTargetLine(ctx, left, top, plotWidth, plotHeight, axisRange, {
    ...visual,
    showTargetLine: true,
    targetLineValue: targetPercent
  }, { percentHint: false });

  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = axisTextColor;
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(axisRange.max, visual), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(0, visual), left - 4, top + plotHeight + 2);
    ctx.fillText("100%", width - 4, top + 8);
    ctx.fillText("0%", width - 4, top + plotHeight + 2);
  }

  if (visual.showAxisLabels && labels.length <= 12) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = labelTextColor;
    ctx.textAlign = "center";
    labels.forEach((label, index) => {
      const x = left + (slot * index) + (slot / 2);
      ctx.fillText(truncateBiLabel(label, visual, 8), x, top + plotHeight + 12);
    });
  }

  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiScatterLikeChart(ctx, width, height, labels, values, colors, settings, rows, bubbleMode, labelOffsets, capture, labelLayoutV2, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, bubbleMode ? "bubble" : "scatter");
  const safeRows = Array.isArray(rows) ? rows : [];
  const offsets = normalizeBiLabelOffsets(labelOffsets);
  const layoutV2 = normalizeBiLabelLayoutV2(labelLayoutV2);
  const collector = capture && typeof capture === "object" ? capture : null;
  if (collector) {
    collector.labelItems = [];
  }
  const minRadius = bubbleMode
    ? sanitizeBiInteger(specific.minBubbleRadius, 6, 3, 60)
    : sanitizeBiInteger(specific.minPointRadius, 7, 3, 40);
  const maxRadius = bubbleMode
    ? sanitizeBiInteger(specific.maxBubbleRadius, 40, minRadius, 80)
    : minRadius;
  const pointsData = values.map((value, index) => {
    const row = safeRows[index] || {};
    const xValue = Number.isFinite(row.xValue) ? row.xValue : (index + 1);
    const yValue = Number.isFinite(value) ? value : 0;
    const bubbleSize = bubbleMode
      ? (Number.isFinite(row.bubbleSize) ? row.bubbleSize : Math.max(minRadius, Math.round(Math.abs(yValue) * 0.16)))
      : minRadius;
    return {
      xValue,
      yValue,
      bubbleSize: Math.max(minRadius, Math.min(maxRadius, bubbleSize))
    };
  });
  const xValues = pointsData.map((item) => item.xValue);
  const yValues = pointsData.map((item) => item.yValue);
  const xMin = Math.min(...xValues, 0);
  const xMax = Math.max(...xValues, 1);
  const yMin = Math.min(...yValues, 0);
  const yMax = Math.max(...yValues, 1);
  const left = 52;
  const right = 14;
  const top = 12;
  const bottom = (visual.showAxisLabels ? 32 : 18) + (visual.axisXLabel ? 14 : 0);
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const xRange = xMax - xMin <= 0 ? 1 : (xMax - xMin);
  const yRange = yMax - yMin <= 0 ? 1 : (yMax - yMin);
  if (visual.showGrid) {
    drawBiHorizontalGrid(ctx, left, top, plotWidth, plotHeight, 5, "#e4ecf5", visual);
  }
  ctx.strokeStyle = "#d6dde6";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotHeight);
  ctx.lineTo(left + plotWidth, top + plotHeight);
  ctx.stroke();
  pointsData.forEach((item, index) => {
    const x = left + (((item.xValue - xMin) / xRange) * plotWidth);
    const y = top + plotHeight - (((item.yValue - yMin) / yRange) * plotHeight);
    const color = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    const pointAlpha = sanitizeBiDecimal(specific.pointAlpha, bubbleMode ? 0.95 : 1, 0.15, 1);
    ctx.fillStyle = bubbleMode ? hexToRgba(color, pointAlpha, color) : hexToRgba(color, pointAlpha, color);
    ctx.strokeStyle = color;
    ctx.beginPath();
    ctx.arc(x, y, item.bubbleSize, 0, Math.PI * 2);
    ctx.fill();
    if (bubbleMode) {
      ctx.lineWidth = 1.4;
      ctx.stroke();
    }
    if (visual.showDataLabels) {
      ctx.fillStyle = resolveBiLabelColor(visual, "#ffffff", "#2f425a");
      ctx.textAlign = "center";
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
      const text = formatBiVisualNumber(item.yValue, visual);
      const rowKey = trimOrFallback(safeRows[index]?.rowKey, "");
      const fallbackKey = buildBiLabelFallbackKey(labels[index] || "");
      const placement = getBiLabelPlacementFromLayout(layoutV2, rowKey, labels[index] || fallbackKey || "");
      const legacyOffset = offsets[String(index)] || { x: 0, y: 0 };
      const offsetX = placement
        ? (placement.dxPct * plotWidth)
        : (Number.isFinite(legacyOffset.x) ? legacyOffset.x : 0);
      const offsetY = placement
        ? (placement.dyPct * plotHeight)
        : (Number.isFinite(legacyOffset.y) ? legacyOffset.y : 0);
      const labelX = x + offsetX;
      const labelY = Math.max(10, y - item.bubbleSize - 2) + offsetY;
      ctx.fillText(text, labelX, labelY);
      if (collector) {
        const textWidth = ctx.measureText(text).width;
        const textHeight = Math.max(10, visual.fontSizeLabels + 2);
        collector.labelItems.push({
          index,
          rowKey,
          rowLabel: labels[index] || "",
          text,
          x: labelX - (textWidth / 2) - 4,
          y: labelY - textHeight - 1,
          w: textWidth + 8,
          h: textHeight + 6,
          anchorX: labelX,
          anchorY: labelY,
          plotWidth,
          plotHeight
        });
      }
    }
  });
  if (visual.showAxisLabels) {
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = "#6a7686";
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(yMax, visual), left - 4, top + 8);
    ctx.fillText(formatBiVisualNumber(yMin, visual), left - 4, top + plotHeight + 2);
    ctx.textAlign = "left";
    ctx.fillText(formatBiVisualNumber(xMin, visual), left, top + plotHeight + 14);
    ctx.fillText(formatBiVisualNumber(xMax, visual), left + plotWidth - 16, top + plotHeight + 14);
  }
  drawBiAxisTitles(ctx, left, top, plotWidth, plotHeight, visual);
}

function drawBiBulletChart(ctx, width, height, labels, values, colors, settings, rows, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  if (!Array.isArray(values) || values.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "bullet");
  const safeRows = Array.isArray(rows) ? rows : [];
  const maxValue = Math.max(1, ...values.map((item) => Math.abs(item)));
  const targetPercent = sanitizeBiDecimal(specific.targetValue, Number.isFinite(visual.targetLineValue) ? visual.targetLineValue : 80, 0, 100);
  const lowBand = sanitizeBiDecimal(specific.lowBand, 40, 0, 100);
  const midBand = sanitizeBiDecimal(specific.midBand, 70, lowBand, 100);
  const left = 12;
  const right = 12;
  const top = 10;
  const rowHeight = Math.max(18, Math.floor((height - top - 8) / Math.max(1, values.length)));
  const trackWidth = Math.max(40, width - left - right - 86);
  values.forEach((value, index) => {
    const y = top + (index * rowHeight);
    const label = truncateBiLabel(labels[index] || "", visual, 16);
    const ratio = Math.max(0, Math.min(1, (Math.abs(value) / maxValue)));
    const barW = Math.max(1, Math.round(trackWidth * ratio));
    const targetX = left + 76 + Math.round(trackWidth * Math.max(0, Math.min(1, targetPercent / 100)));
    ctx.fillStyle = "#edf3fb";
    ctx.fillRect(left + 76, y + 4, trackWidth, rowHeight - 8);
    ctx.fillStyle = "rgba(77, 132, 199, 0.14)";
    ctx.fillRect(left + 76, y + 4, Math.round(trackWidth * (lowBand / 100)), rowHeight - 8);
    ctx.fillStyle = "rgba(44, 153, 110, 0.14)";
    ctx.fillRect(left + 76 + Math.round(trackWidth * (lowBand / 100)), y + 4, Math.round(trackWidth * ((midBand - lowBand) / 100)), rowHeight - 8);
    ctx.fillStyle = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    ctx.fillRect(left + 76, y + 4, barW, rowHeight - 8);
    ctx.strokeStyle = normalizeBiColorHex(visual.targetLineColor || "#b03831", "#b03831");
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(targetX, y + 2);
    ctx.lineTo(targetX, y + rowHeight - 2);
    ctx.stroke();
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
    ctx.fillStyle = "#2c405a";
    ctx.textAlign = "left";
    ctx.fillText(label, left, y + Math.round(rowHeight * 0.64));
    ctx.textAlign = "right";
    ctx.fillText(formatBiVisualNumber(value, visual), width - 4, y + Math.round(rowHeight * 0.64));
  });
}

function drawBiBoxplotChart(ctx, width, height, labels, values, colors, settings, rows, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  const safeRows = Array.isArray(rows) ? rows : [];
  if (safeRows.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "boxplot");
  const showMedianValue = normalizeBiToggle(specific.showMedianValue, true);
  const max = Math.max(1, ...safeRows.map((row) => Number(row?.box?.max) || 0));
  const min = Math.min(0, ...safeRows.map((row) => Number(row?.box?.min) || 0));
  const range = max - min <= 0 ? 1 : (max - min);
  const left = 44;
  const right = 10;
  const top = 12;
  const bottom = 30;
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, safeRows.length);
  safeRows.forEach((row, index) => {
    const box = row?.box || {};
    const color = normalizeBiColorHex(colors?.[index], getBiPaletteColor(index));
    const toY = (v) => top + plotHeight - (((Number(v) - min) / range) * plotHeight);
    const xCenter = left + (slot * index) + (slot / 2);
    const boxW = Math.max(8, slot * 0.42);
    const yMin = toY(box.min || 0);
    const yQ1 = toY(box.q1 || 0);
    const yMedian = toY(box.median || 0);
    const yQ3 = toY(box.q3 || 0);
    const yMax = toY(box.max || 0);
    ctx.strokeStyle = color;
    ctx.lineWidth = 1.4;
    ctx.beginPath();
    ctx.moveTo(xCenter, yMin);
    ctx.lineTo(xCenter, yQ1);
    ctx.moveTo(xCenter, yQ3);
    ctx.lineTo(xCenter, yMax);
    ctx.stroke();
    ctx.fillStyle = hexToRgba(color, 0.26, "rgba(47,126,216,0.26)");
    ctx.fillRect(xCenter - (boxW / 2), yQ3, boxW, Math.max(2, yQ1 - yQ3));
    ctx.strokeRect(xCenter - (boxW / 2), yQ3, boxW, Math.max(2, yQ1 - yQ3));
    ctx.strokeStyle = color;
    ctx.beginPath();
    ctx.moveTo(xCenter - (boxW / 2), yMedian);
    ctx.lineTo(xCenter + (boxW / 2), yMedian);
    ctx.stroke();
    if (showMedianValue) {
      applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(8, visual.fontSizeLabels - 1) });
      ctx.fillStyle = "#2a3f56";
      ctx.textAlign = "center";
      ctx.fillText(formatBiVisualNumber(Number(box.median) || 0, visual), xCenter, Math.max(10, yMedian - 4));
    }
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = "#5f7186";
    ctx.textAlign = "center";
    ctx.fillText(truncateBiLabel(labels[index] || "", visual, 10), xCenter, top + plotHeight + 12);
  });
}

function drawBiCandlestickChart(ctx, width, height, labels, values, colors, settings, rows, typeConfig) {
  ctx.clearRect(0, 0, width, height);
  const safeRows = Array.isArray(rows) ? rows : [];
  if (safeRows.length === 0) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "candlestick");
  const upColor = normalizeBiColorHex(specific.upColor, "#25a368");
  const downColor = normalizeBiColorHex(specific.downColor, "#d7565b");
  const wickWidth = sanitizeBiInteger(specific.wickWidth, 1, 1, 4);
  const max = Math.max(1, ...safeRows.map((row) => Number(row?.candle?.high) || 0));
  const min = Math.min(0, ...safeRows.map((row) => Number(row?.candle?.low) || 0));
  const range = max - min <= 0 ? 1 : (max - min);
  const left = 44;
  const right = 10;
  const top = 12;
  const bottom = 30;
  const plotWidth = Math.max(10, width - left - right);
  const plotHeight = Math.max(10, height - top - bottom);
  const slot = plotWidth / Math.max(1, safeRows.length);
  safeRows.forEach((row, index) => {
    const candle = row?.candle || {};
    const open = Number(candle.open) || 0;
    const close = Number(candle.close) || 0;
    const high = Number(candle.high) || Math.max(open, close);
    const low = Number(candle.low) || Math.min(open, close);
    const bullish = close >= open;
    const color = bullish ? upColor : downColor;
    const toY = (v) => top + plotHeight - (((v - min) / range) * plotHeight);
    const xCenter = left + (slot * index) + (slot / 2);
    const bodyW = Math.max(6, slot * 0.44);
    const yOpen = toY(open);
    const yClose = toY(close);
    const yHigh = toY(high);
    const yLow = toY(low);
    ctx.strokeStyle = color;
    ctx.lineWidth = wickWidth;
    ctx.beginPath();
    ctx.moveTo(xCenter, yHigh);
    ctx.lineTo(xCenter, yLow);
    ctx.stroke();
    ctx.fillStyle = hexToRgba(color, bullish ? 0.38 : 0.26, bullish ? "rgba(47,157,95,0.38)" : "rgba(212,78,71,0.26)");
    const bodyTop = Math.min(yOpen, yClose);
    const bodyHeight = Math.max(2, Math.abs(yClose - yOpen));
    ctx.fillRect(xCenter - (bodyW / 2), bodyTop, bodyW, bodyHeight);
    ctx.strokeRect(xCenter - (bodyW / 2), bodyTop, bodyW, bodyHeight);
    applyBiCanvasFont(ctx, visual, "axis", { fallbackSize: Math.max(9, visual.fontSizeAxis) });
    ctx.fillStyle = "#5f7186";
    ctx.textAlign = "center";
    ctx.fillText(truncateBiLabel(labels[index] || "", visual, 10), xCenter, top + plotHeight + 12);
  });
}

function drawBiHoverHighlight(ctx, model, highlightIndex) {
  if (!ctx || !model || !Array.isArray(model.items) || highlightIndex < 0) {
    return;
  }

  const item = model.items.find((entry) => entry.index === highlightIndex);
  if (!item) {
    return;
  }

  ctx.save();
  ctx.shadowColor = "rgba(41, 128, 185, 0.45)";
  ctx.shadowBlur = 16;
  ctx.strokeStyle = "#1d6fb7";
  ctx.fillStyle = "rgba(29, 111, 183, 0.12)";
  ctx.lineWidth = 2;

  if (item.kind === "rect") {
    ctx.fillRect(item.x, item.y, item.w, item.h);
    ctx.strokeRect(item.x, item.y, item.w, item.h);
    ctx.restore();
    return;
  }

  if (item.kind === "point") {
    ctx.beginPath();
    ctx.arc(item.x, item.y, Math.max(7, item.r * 0.8), 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
    ctx.restore();
    return;
  }

  if (item.kind === "arc") {
    ctx.beginPath();
    ctx.arc(item.cx, item.cy, item.outerRadius + 4, item.startAngle, item.endAngle);
    ctx.arc(item.cx, item.cy, Math.max(0, item.innerRadius - 3), item.endAngle, item.startAngle, true);
    ctx.closePath();
    ctx.fill();
    ctx.stroke();
    ctx.restore();
    return;
  }

  ctx.restore();
}

function buildSeriesModel(type, snapshotRows, labels, values, chartConfig, typeConfig) {
  const safeType = normalizeBiChartType(type || "bar");
  return {
    type: safeType,
    rows: Array.isArray(snapshotRows) ? snapshotRows : [],
    labels: Array.isArray(labels) ? labels : [],
    values: Array.isArray(values) ? values : [],
    chartConfig: normalizeBiChartConfig(chartConfig, safeType),
    typeConfig: normalizeBiChartTypeSpecificConfig(typeConfig, safeType)
  };
}

function drawSeriesModel(ctx, width, height, model, colors, visual, labelOffsets, circularLayout, geometry, labelLayoutV2) {
  if (!ctx || !model) {
    return;
  }
  const safeType = normalizeBiChartType(model.type || "bar");
  const safeLabels = model.labels || [];
  const safeValues = model.values || [];
  const safeRows = model.rows || [];
  const safeTypeConfig = normalizeBiChartTypeSpecificConfig(model.typeConfig, safeType);
  if (safeType === "line" || safeType === "timeseries") {
    drawBiLineChart(ctx, width, height, safeLabels, safeValues, colors, visual, labelOffsets, geometry, safeRows, labelLayoutV2, safeTypeConfig);
    return;
  }
  if (safeType === "area") {
    drawBiAreaChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "combo") {
    drawBiComboChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "donut") {
    drawBiDonutChart(ctx, width, height, safeLabels, safeValues, colors, visual, circularLayout, geometry, safeTypeConfig);
    return;
  }
  if (safeType === "pie") {
    drawBiPieChart(ctx, width, height, safeLabels, safeValues, colors, visual, circularLayout, geometry, safeTypeConfig);
    return;
  }
  if (safeType === "scatter") {
    drawBiScatterLikeChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeRows, false, labelOffsets, geometry, labelLayoutV2, safeTypeConfig);
    return;
  }
  if (safeType === "bubble") {
    drawBiScatterLikeChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeRows, true, labelOffsets, geometry, labelLayoutV2, safeTypeConfig);
    return;
  }
  if (safeType === "bullet") {
    drawBiBulletChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeRows, safeTypeConfig);
    return;
  }
  if (safeType === "boxplot") {
    drawBiBoxplotChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeRows, safeTypeConfig);
    return;
  }
  if (safeType === "candlestick") {
    drawBiCandlestickChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeRows, safeTypeConfig);
    return;
  }
  if (safeType === "treemap") {
    drawBiTreemapChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "funnel") {
    drawBiFunnelChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "gauge") {
    drawBiGaugeChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "waterfall") {
    drawBiWaterfallChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "radar") {
    drawBiRadarChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  if (safeType === "pareto") {
    drawBiParetoChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
    return;
  }
  drawBiBarChart(ctx, width, height, safeLabels, safeValues, colors, visual, safeTypeConfig);
}

function drawBiWidgetChart(canvas, chartType, labels, values, highlightIndex, colors, settings, labelOffsets, circularLayout, rowsData, chartConfig, labelLayoutV2, typeConfigMap) {
  if (!(canvas instanceof HTMLCanvasElement)) {
    return { type: "none", items: [] };
  }

  const { ctx, width, height } = resizeCanvasForDisplay(canvas, 160);
  const visual = normalizeBiVisualSettings(settings);
  const safeLabels = Array.isArray(labels) ? labels : [];
  const safeValues = Array.isArray(values)
    ? values.map((value) => (Number.isFinite(value) ? value : 0))
    : [];
  const safeColors = Array.isArray(colors) ? colors.map((color, index) => normalizeBiColorHex(color, getBiPaletteColor(index))) : [];

  if (safeLabels.length === 0 || safeValues.length === 0) {
    ctx.clearRect(0, 0, width, height);
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 12 });
    ctx.fillStyle = "#748395";
    ctx.textAlign = "center";
    ctx.fillText("Sin datos para este widget", Math.floor(width / 2), Math.floor(height / 2));
    return { type: "none", items: [] };
  }

  const safeType = normalizeBiChartType(chartType || "bar");
  const renderedValues = getBiRenderedValuesForVisual(safeValues, visual, safeType);
  const typeConfig = typeConfigMap && typeof typeConfigMap === "object"
    ? typeConfigMap[safeType]
    : null;
  const seriesModel = buildSeriesModel(safeType, rowsData, safeLabels, renderedValues, chartConfig, typeConfig);
  const geometry = { labelItems: [], polarModel: null };
  drawSeriesModel(
    ctx,
    width,
    height,
    seriesModel,
    safeColors,
    visual,
    labelOffsets,
    circularLayout,
    geometry,
    labelLayoutV2
  );

  const model = buildBiHoverModel(safeType, width, height, safeLabels, renderedValues, {
    visualSettings: visual,
    labelItems: geometry.labelItems,
    polarModel: geometry.polarModel,
    rowsData: rowsData
  });
  const activeIndex = Number.isInteger(highlightIndex) ? highlightIndex : -1;
  if (activeIndex >= 0) {
    drawBiHoverHighlight(ctx, model, activeIndex);
  }
  return model;
}

function buildBiHoverModel(chartType, width, height, labels, values, options) {
  const safeLabels = Array.isArray(labels) ? labels : [];
  const safeValues = Array.isArray(values) ? values : [];
  if (safeLabels.length === 0 || safeValues.length === 0) {
    return { type: "none", items: [] };
  }

  const type = normalizeBiChartType(chartType || "bar");
  const visual = normalizeBiVisualSettings(options?.visualSettings || {});
  const labelItems = Array.isArray(options?.labelItems)
    ? options.labelItems.filter((item) => item && Number.isInteger(item.index))
    : [];
  const items = [];
  if (type === "line" || type === "timeseries" || type === "area" || type === "combo") {
    const left = (visual.axisYLabel ? 58 : 44);
    const right = 10;
    const top = 12;
    const bottom = (visual.showAxisLabels ? 30 : 16) + (visual.axisXLabel ? 14 : 0);
    const plotWidth = Math.max(10, width - left - right);
    const plotHeight = Math.max(10, height - top - bottom);
    const stepX = safeValues.length <= 1 ? 0 : (plotWidth / (safeValues.length - 1));
    const axisRange = resolveBiAxisRange(safeValues, visual, { forceZeroBaseline: false });
    safeValues.forEach((value, index) => {
      const normalized = Math.max(0, Math.min(1, normalizeBiAxisValue(value, axisRange)));
      const x = left + (stepX * index);
      const y = top + plotHeight - (Math.max(0, normalized) * plotHeight);
      items.push({ index, kind: "point", x, y, r: 12 });
    });
    return { type, items, labelItems };
  }

  if (type === "scatter" || type === "bubble") {
    const safeRows = Array.isArray(options?.rowsData) ? options.rowsData : [];
    const pointsData = safeValues.map((value, index) => {
      const row = safeRows[index] || {};
      const xValue = Number.isFinite(row.xValue) ? row.xValue : (index + 1);
      const yValue = Number.isFinite(value) ? value : 0;
      const bubbleSize = type === "bubble"
        ? (Number.isFinite(row.bubbleSize) ? row.bubbleSize : Math.max(8, Math.round(Math.abs(yValue) * 0.16)))
        : 8;
      return {
        index,
        xValue,
        yValue,
        radius: Math.max(6, Math.min(40, bubbleSize))
      };
    });
    const xValues = pointsData.map((item) => item.xValue);
    const yValues = pointsData.map((item) => item.yValue);
    const xMin = Math.min(...xValues, 0);
    const xMax = Math.max(...xValues, 1);
    const yMin = Math.min(...yValues, 0);
    const yMax = Math.max(...yValues, 1);
    const xRange = xMax - xMin <= 0 ? 1 : (xMax - xMin);
    const yRange = yMax - yMin <= 0 ? 1 : (yMax - yMin);
    const left = 52;
    const right = 14;
    const top = 12;
    const bottom = (visual.showAxisLabels ? 32 : 18) + (visual.axisXLabel ? 14 : 0);
    const plotWidth = Math.max(10, width - left - right);
    const plotHeight = Math.max(10, height - top - bottom);
    pointsData.forEach((item) => {
      const x = left + (((item.xValue - xMin) / xRange) * plotWidth);
      const y = top + plotHeight - (((item.yValue - yMin) / yRange) * plotHeight);
      items.push({ index: item.index, kind: "point", x, y, r: item.radius + 2 });
    });
    return { type, items, labelItems };
  }

  if (type === "radar") {
    const safePositive = safeValues.map((value) => (Number.isFinite(value) ? Math.max(0, value) : 0));
    const max = Math.max(1, ...safePositive);
    const count = Math.max(3, safePositive.length);
    const centerX = Math.floor(width / 2);
    const centerY = Math.floor(height / 2) + 6;
    const radius = Math.max(24, Math.floor(Math.min(width, height) * 0.33));
    safePositive.forEach((value, index) => {
      const angle = (-Math.PI / 2) + ((Math.PI * 2 * index) / count);
      const normalized = Math.max(0, Math.min(1, value / max));
      const x = centerX + Math.cos(angle) * (radius * normalized);
      const y = centerY + Math.sin(angle) * (radius * normalized);
      items.push({ index, kind: "point", x, y, r: 14 });
    });
    return { type, items };
  }

  if (type === "bar" || type === "waterfall" || type === "pareto") {
    const left = 38;
    const right = 8;
    const top = 10;
    const bottom = type === "bar" ? 26 : (type === "pareto" ? 30 : 24);
    const plotWidth = Math.max(10, width - left - right);
    const plotHeight = Math.max(10, height - top - bottom);
    const slot = plotWidth / Math.max(1, safeValues.length);
    const barWidth = Math.max(6, slot * (type === "bar" ? 0.64 : (type === "pareto" ? 0.62 : 0.58)));
    const max = Math.max(1, ...safeValues.map((value) => Math.abs(value)));
    if (type === "bar" || type === "pareto") {
      safeValues.forEach((value, index) => {
        const normalized = max <= 0 ? 0 : (Math.abs(value) / max);
        const barHeight = Math.max(0, normalized * plotHeight);
        const x = left + slot * index + ((slot - barWidth) / 2);
        const y = top + plotHeight - barHeight;
        items.push({ index, kind: "rect", x, y, w: barWidth, h: Math.max(2, barHeight) });
      });
      return { type, items };
    }

    const zeroY = top + (plotHeight / 2);
    let previous = 0;
    safeValues.forEach((value, index) => {
      const current = previous + value;
      const yFrom = zeroY - ((previous / max) * (plotHeight / 2));
      const yTo = zeroY - ((current / max) * (plotHeight / 2));
      const y = Math.min(yFrom, yTo);
      const h = Math.max(2, Math.abs(yTo - yFrom));
      const x = left + slot * index + ((slot - barWidth) / 2);
      items.push({ index, kind: "rect", x, y, w: barWidth, h });
      previous = current;
    });
    return { type, items };
  }

  if (type === "bullet") {
    const left = 12;
    const right = 12;
    const top = 10;
    const rowHeight = Math.max(18, Math.floor((height - top - 8) / Math.max(1, safeValues.length)));
    const trackWidth = Math.max(40, width - left - right - 86);
    safeValues.forEach((value, index) => {
      const ratio = Math.max(0, Math.min(1, Math.abs(value) / Math.max(1, ...safeValues.map((item) => Math.abs(item)))));
      const y = top + (index * rowHeight);
      const w = Math.max(1, Math.round(trackWidth * ratio));
      items.push({ index, kind: "rect", x: left + 76, y: y + 4, w, h: Math.max(4, rowHeight - 8) });
    });
    return { type, items };
  }

  if (type === "boxplot" || type === "candlestick") {
    const safeRows = Array.isArray(options?.rowsData) ? options.rowsData : [];
    const left = 44;
    const right = 10;
    const top = 12;
    const bottom = 30;
    const plotWidth = Math.max(10, width - left - right);
    const plotHeight = Math.max(10, height - top - bottom);
    const slot = plotWidth / Math.max(1, safeRows.length || safeValues.length);
    safeRows.forEach((row, index) => {
      const xCenter = left + (slot * index) + (slot / 2);
      if (type === "boxplot") {
        const boxW = Math.max(8, slot * 0.42);
        const box = row?.box || {};
        const allValues = safeRows.map((item) => [
          Number(item?.box?.min) || 0,
          Number(item?.box?.q1) || 0,
          Number(item?.box?.median) || 0,
          Number(item?.box?.q3) || 0,
          Number(item?.box?.max) || 0
        ]).flat();
        const min = Math.min(0, ...allValues);
        const max = Math.max(1, ...allValues);
        const range = max - min <= 0 ? 1 : (max - min);
        const toY = (value) => top + plotHeight - (((value - min) / range) * plotHeight);
        const yQ1 = toY(Number(box.q1) || 0);
        const yQ3 = toY(Number(box.q3) || 0);
        items.push({
          index,
          kind: "rect",
          x: xCenter - (boxW / 2),
          y: Math.min(yQ1, yQ3),
          w: boxW,
          h: Math.max(2, Math.abs(yQ1 - yQ3))
        });
      } else {
        const bodyW = Math.max(6, slot * 0.44);
        const candle = row?.candle || {};
        const allValues = safeRows.map((item) => [
          Number(item?.candle?.low) || 0,
          Number(item?.candle?.high) || 0,
          Number(item?.candle?.open) || 0,
          Number(item?.candle?.close) || 0
        ]).flat();
        const min = Math.min(0, ...allValues);
        const max = Math.max(1, ...allValues);
        const range = max - min <= 0 ? 1 : (max - min);
        const toY = (value) => top + plotHeight - (((value - min) / range) * plotHeight);
        const yOpen = toY(Number(candle.open) || 0);
        const yClose = toY(Number(candle.close) || 0);
        items.push({
          index,
          kind: "rect",
          x: xCenter - (bodyW / 2),
          y: Math.min(yOpen, yClose),
          w: bodyW,
          h: Math.max(2, Math.abs(yClose - yOpen))
        });
      }
    });
    return { type, items };
  }

  if (type === "donut" || type === "pie") {
    const polarModel = options?.polarModel;
    if (polarModel && Array.isArray(polarModel.items)) {
      return {
        type,
        items: polarModel.items,
        chartDragRect: polarModel.chartDragRect || null,
        legendDragRect: polarModel.legendDragRect || null
      };
    }
    const safePositive = safeValues.map((value) => (Number.isFinite(value) && value > 0 ? value : 0));
    const total = safePositive.reduce((sum, value) => sum + value, 0);
    if (total <= 0) {
      return { type, items: [] };
    }
    const centerX = Math.floor(width * (type === "donut" ? 0.33 : 0.40));
    const centerY = Math.floor(height / 2);
    const outerRadius = Math.max(28, Math.floor(Math.min(width * (type === "donut" ? 0.56 : 0.52), height) * (type === "donut" ? 0.34 : 0.40)));
    const innerRadius = type === "donut" ? Math.floor(outerRadius * 0.58) : 0;
    let start = -Math.PI / 2;
    safePositive.forEach((value, index) => {
      const angle = (value / total) * Math.PI * 2;
      const end = start + angle;
      items.push({
        index,
        kind: "arc",
        cx: centerX,
        cy: centerY,
        innerRadius,
        outerRadius,
        startAngle: start,
        endAngle: end
      });
      start = end;
    });
    return { type, items };
  }

  if (type === "treemap") {
    const safePositive = safeValues.map((value) => (Number.isFinite(value) && value > 0 ? value : 0));
    const total = safePositive.reduce((sum, value) => sum + value, 0);
    if (total <= 0) {
      return { type, items: [] };
    }
    const left = 8;
    const top = 8;
    const plotWidth = Math.max(20, width - 16);
    const plotHeight = Math.max(20, height - 16);
    let offsetX = left;
    safePositive.forEach((value, index) => {
      const ratio = value / total;
      const blockWidth = Math.max(18, Math.round(plotWidth * ratio));
      const w = Math.min(blockWidth, left + plotWidth - offsetX);
      items.push({ index, kind: "rect", x: offsetX, y: top, w, h: plotHeight });
      offsetX += w;
    });
    return { type, items };
  }

  if (type === "funnel") {
    const topY = 10;
    const rowHeight = Math.max(20, Math.floor((height - 20) / safeValues.length));
    safeValues.forEach((value, index) => {
      items.push({
        index,
        kind: "rect",
        x: 0,
        y: topY + (index * rowHeight),
        w: width,
        h: rowHeight
      });
    });
    return { type, items };
  }

  if (type === "gauge" || type === "scorecard") {
    items.push({ index: 0, kind: "rect", x: 0, y: 0, w: width, h: height });
    return { type, items };
  }

  return { type: "none", items: [] };
}

function biAngleBetween(start, end, angle) {
  const twoPi = Math.PI * 2;
  let normalizedStart = start;
  let normalizedEnd = end;
  let normalizedAngle = angle;
  while (normalizedEnd < normalizedStart) {
    normalizedEnd += twoPi;
  }
  while (normalizedAngle < normalizedStart) {
    normalizedAngle += twoPi;
  }
  while (normalizedAngle > normalizedEnd) {
    normalizedAngle -= twoPi;
  }
  return normalizedAngle >= normalizedStart && normalizedAngle <= normalizedEnd;
}

function resolveBiHoverIndex(model, x, y) {
  if (!model || !Array.isArray(model.items) || model.items.length === 0) {
    return -1;
  }

  if (model.type === "line" || model.type === "timeseries" || model.type === "area" || model.type === "combo" || model.type === "radar") {
    let bestIndex = -1;
    let bestDistance = Number.POSITIVE_INFINITY;
    model.items.forEach((item) => {
      const dx = x - item.x;
      const dy = y - item.y;
      const distance = Math.sqrt((dx * dx) + (dy * dy));
      if (distance < bestDistance) {
        bestDistance = distance;
        bestIndex = item.index;
      }
    });
    if (bestDistance <= 24) {
      return bestIndex;
    }
    if (model.type === "radar") {
      return -1;
    }
    let segmentDistance = Number.POSITIVE_INFINITY;
    let segmentIndex = -1;
    for (let i = 0; i < model.items.length - 1; i += 1) {
      const start = model.items[i];
      const end = model.items[i + 1];
      if (!start || !end) {
        continue;
      }
      const vx = end.x - start.x;
      const vy = end.y - start.y;
      const len2 = (vx * vx) + (vy * vy);
      if (len2 <= 0.0001) {
        continue;
      }
      const tRaw = (((x - start.x) * vx) + ((y - start.y) * vy)) / len2;
      const t = Math.max(0, Math.min(1, tRaw));
      const px = start.x + (vx * t);
      const py = start.y + (vy * t);
      const dx = x - px;
      const dy = y - py;
      const distance = Math.sqrt((dx * dx) + (dy * dy));
      if (distance < segmentDistance) {
        segmentDistance = distance;
        segmentIndex = t < 0.5 ? start.index : end.index;
      }
    }
    return segmentDistance <= 10 ? segmentIndex : -1;
  }

  for (const item of model.items) {
    if (item.kind === "rect") {
      const inside = x >= item.x && x <= (item.x + item.w) && y >= item.y && y <= (item.y + item.h);
      if (inside) {
        return item.index;
      }
      continue;
    }
    if (item.kind === "point") {
      const dx = x - item.x;
      const dy = y - item.y;
      if ((dx * dx) + (dy * dy) <= (item.r * item.r)) {
        return item.index;
      }
      continue;
    }
    if (item.kind === "arc") {
      const dx = x - item.cx;
      const dy = y - item.cy;
      const radius = Math.sqrt((dx * dx) + (dy * dy));
      if (radius < item.innerRadius || radius > item.outerRadius) {
        continue;
      }
      const angle = Math.atan2(dy, dx);
      if (biAngleBetween(item.startAngle, item.endAngle, angle)) {
        return item.index;
      }
    }
  }

  return -1;
}

function resolveBiLabelDragIndex(model, x, y) {
  if (!model || !Array.isArray(model.labelItems) || model.labelItems.length === 0) {
    return -1;
  }
  for (let i = model.labelItems.length - 1; i >= 0; i -= 1) {
    const item = model.labelItems[i];
    if (!item || !Number.isInteger(item.index)) {
      continue;
    }
    const inside = x >= item.x && x <= (item.x + item.w) && y >= item.y && y <= (item.y + item.h);
    if (inside) {
      return item.index;
    }
  }
  return -1;
}

function resolveBiPolarDragTarget(model, x, y) {
  if (!model || typeof model !== "object") {
    return "";
  }
  const isInside = (box) => box
    && Number.isFinite(box.x)
    && Number.isFinite(box.y)
    && Number.isFinite(box.w)
    && Number.isFinite(box.h)
    && x >= box.x
    && x <= (box.x + box.w)
    && y >= box.y
    && y <= (box.y + box.h);
  if (isInside(model.legendDragRect)) {
    return "legend";
  }
  if (isInside(model.chartDragRect)) {
    return "chart";
  }
  return "";
}

function toggleBiCrossFilter(project, groupBy, label, source, appendMode = false) {
  if (!project || !groupBy || !label) {
    return false;
  }
  ensureBiState(project);
  const nextEntry = normalizeBiCrossFilterEntry({ groupBy, label, source });
  if (!nextEntry) {
    return false;
  }

  const current = normalizeBiCrossFilters(project.biConfig.crossFilters);
  const isSameEntry = (entry) => entry
    && entry.groupBy === nextEntry.groupBy
    && entry.source === nextEntry.source
    && normalizeLookup(entry.label) === normalizeLookup(nextEntry.label);

  if (appendMode) {
    const exists = current.some((entry) => isSameEntry(entry));
    if (exists) {
      project.biConfig.crossFilters = current.filter((entry) => !isSameEntry(entry));
    } else {
      project.biConfig.crossFilters = [...current, nextEntry];
    }
    return true;
  }

  const currentEntry = current[0] || null;
  if (isSameEntry(currentEntry) && current.length === 1) {
    project.biConfig.crossFilters = [];
    return true;
  }
  project.biConfig.crossFilters = [nextEntry];
  return true;
}

function applyBiCrossFilterFromWidget(project, snapshot, rowIndex, event) {
  if (!project || !snapshot || !Array.isArray(snapshot.rows)) {
    return;
  }
  const selected = snapshot.rows[rowIndex];
  if (!selected) {
    return;
  }
  const appendMode = !!(event && (event.ctrlKey || event.metaKey));
  const changed = toggleBiCrossFilter(project, snapshot.groupBy, selected.label, snapshot.source, appendMode);
  if (!changed) {
    return;
  }
  saveState();
  renderBiPanel(project);
  const hasFilter = normalizeBiCrossFilters(project.biConfig.crossFilters).length > 0;
  if (hasFilter) {
    const modeLabel = appendMode ? " (multi)" : "";
    setStatus(`Filtro cruzado aplicado${modeLabel}: ${getBiGroupLabel(snapshot.groupBy, project)} = ${selected.label}`);
  } else {
    setStatus("Filtro cruzado removido.");
  }
}

function bindBiCanvasHover(canvas, tooltip, snapshot, chartType, onSelect, initialModel, colors, settings, widget) {
  if (!(canvas instanceof HTMLCanvasElement) || !(tooltip instanceof HTMLElement) || !snapshot) {
    return;
  }

  const labels = snapshot.rows.map((row) => row.label);
  const values = snapshot.rows.map((row) => row.value);
  if (labels.length === 0 || values.length === 0) {
    return;
  }
  const safeType = normalizeBiChartType(chartType || "bar");
  const visual = normalizeBiVisualSettings(settings);
  const renderedValues = getBiRenderedValuesForVisual(values, visual, safeType);
  const canDragLayout = (safeType === "pie" || safeType === "donut") && !!widget;
  let model = initialModel || drawBiWidgetChart(
    canvas,
    chartType,
    labels,
    values,
    -1,
    colors,
    settings,
    widget?.labelOffsets,
    widget?.polarLayout,
    snapshot.rows,
    widget?.chartConfig,
    widget?.labelLayoutV2,
    widget?.chartTypeConfig
  );
  if (!model || !Array.isArray(model.items) || !model.items.length) {
    return;
  }
  const canDragLabels = () => !!widget && Array.isArray(model?.labelItems) && model.labelItems.length > 0;
  tooltip.style.fontFamily = getBiFontFamilyByLayer(visual, "tooltip");
  tooltip.style.fontSize = `${getBiFontSizeByLayer(visual, "tooltip", 11)}px`;
  let highlightedIndex = -1;
  let suppressClickUntil = 0;
  let dragState = null;

  const redraw = (highlight = highlightedIndex) => {
    model = drawBiWidgetChart(
      canvas,
      chartType,
      labels,
      values,
      highlight,
      colors,
      settings,
      widget?.labelOffsets,
      widget?.polarLayout,
      snapshot.rows,
      widget?.chartConfig,
      widget?.labelLayoutV2,
      widget?.chartTypeConfig
    );
  };

  const hideTooltip = () => {
    if (dragState) {
      return;
    }
    canvas.style.cursor = "default";
    tooltip.classList.add("hidden");
    if (highlightedIndex !== -1) {
      highlightedIndex = -1;
      redraw(-1);
    }
  };

  canvas.addEventListener("mouseleave", hideTooltip);
  canvas.addEventListener("mousemove", (event) => {
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (dragState) {
      canvas.style.cursor = "grabbing";
      return;
    }
    let cursor = "default";
    if (canDragLabels()) {
      const labelHover = resolveBiLabelDragIndex(model, x, y);
      if (labelHover >= 0) {
        cursor = "grab";
      }
    }
    if (cursor === "default" && canDragLayout) {
      const layoutTarget = resolveBiPolarDragTarget(model, x, y);
      if (layoutTarget) {
        cursor = "grab";
      }
    }
    canvas.style.cursor = cursor;
    const index = resolveBiHoverIndex(model, x, y);
    if (index < 0 || index >= snapshot.rows.length) {
      hideTooltip();
      return;
    }
    if (index !== highlightedIndex) {
      highlightedIndex = index;
      redraw(highlightedIndex);
    }
    const row = snapshot.rows[index];
    const renderedValue = Number.isFinite(renderedValues[index]) ? renderedValues[index] : row.value;
    const valueDisplay = formatBiVisualNumber(renderedValue, visual, {
      percentHint: isBiPercentMetric(snapshot.metric)
        || visual.stackMode === "percent"
    });
    if (visual.tooltipMode === "compact") {
      tooltip.innerHTML = `<strong>${escapeHtml(row.label)}</strong>${escapeHtml(valueDisplay)}`;
    } else {
      const stackInfo = visual.stackMode !== "none"
        ? `Stack: ${escapeHtml(visual.stackMode === "percent" ? "100%" : "normal")} | Base: ${escapeHtml(formatBiVisualNumber(row.value, visual, { percentHint: isBiPercentMetric(snapshot.metric) }))}`
        : "";
      const lines = [
        `<strong>${escapeHtml(row.label)}</strong>`,
        escapeHtml(valueDisplay),
        `Filas: ${escapeHtml(String(row.count || 0))}`,
        `UP: ${escapeHtml(formatNumberForInput(row.baseUnits) || "0")}`,
        stackInfo
      ];
      tooltip.innerHTML = lines.filter((line) => !!line).join("<br>");
    }
    tooltip.style.left = `${x}px`;
    tooltip.style.top = `${y}px`;
    tooltip.classList.remove("hidden");
  });

  canvas.addEventListener("pointerdown", (event) => {
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;

    if (canDragLabels()) {
      const labelIndex = resolveBiLabelDragIndex(model, x, y);
      if (labelIndex >= 0 && labelIndex < snapshot.rows.length) {
        event.preventDefault();
        const labelItem = Array.isArray(model?.labelItems)
          ? (model.labelItems.find((entry) => entry?.index === labelIndex) || null)
          : null;
        const rowKey = trimOrFallback(snapshot.rows[labelIndex]?.rowKey || labelItem?.rowKey, "");
        const rowLabel = trimOrFallback(snapshot.rows[labelIndex]?.label || labelItem?.rowLabel || labels[labelIndex] || "", "");
        const safeOffsets = normalizeBiLabelOffsets(widget.labelOffsets);
        const currentLegacy = safeOffsets[String(labelIndex)] || { x: 0, y: 0 };
        const currentPlacement = getBiLabelPlacement(widget, rowKey, rowLabel);
        const plotWidth = Math.max(1, Number.isFinite(labelItem?.plotWidth) ? labelItem.plotWidth : 1);
        const plotHeight = Math.max(1, Number.isFinite(labelItem?.plotHeight) ? labelItem.plotHeight : 1);
        dragState = {
          type: "label",
          pointerId: event.pointerId,
          index: labelIndex,
          rowKey,
          rowLabel,
          plotWidth,
          plotHeight,
          startX: x,
          startY: y,
          startOffsetX: Number.isFinite(currentLegacy.x) ? currentLegacy.x : 0,
          startOffsetY: Number.isFinite(currentLegacy.y) ? currentLegacy.y : 0,
          startDxPct: Number.isFinite(currentPlacement?.dxPct)
            ? currentPlacement.dxPct
            : ((Number.isFinite(currentLegacy.x) ? currentLegacy.x : 0) / plotWidth),
          startDyPct: Number.isFinite(currentPlacement?.dyPct)
            ? currentPlacement.dyPct
            : ((Number.isFinite(currentLegacy.y) ? currentLegacy.y : 0) / plotHeight),
          moved: false
        };
        suppressClickUntil = Date.now() + 200;
        canvas.setPointerCapture?.(event.pointerId);
        canvas.style.cursor = "grabbing";
        return;
      }
    }

    if (!canDragLayout) {
      return;
    }
    const layoutTarget = resolveBiPolarDragTarget(model, x, y);
    if (!layoutTarget) {
      return;
    }
    event.preventDefault();
    const safeLayout = normalizeBiCircularLayout(widget.polarLayout);
    const currentOffset = layoutTarget === "chart" ? safeLayout.chart : safeLayout.legend;
    dragState = {
      type: "layout",
      target: layoutTarget,
      pointerId: event.pointerId,
      startX: x,
      startY: y,
      startOffsetX: Number.isFinite(currentOffset.x) ? currentOffset.x : 0,
      startOffsetY: Number.isFinite(currentOffset.y) ? currentOffset.y : 0,
      moved: false
    };
    canvas.setPointerCapture?.(event.pointerId);
    canvas.style.cursor = "grabbing";
    tooltip.classList.add("hidden");
  });

  canvas.addEventListener("pointerup", (event) => {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }
    const moved = dragState.moved;
    const dragType = dragState.type;
    const dragTarget = dragState.target;
    dragState = null;
    canvas.releasePointerCapture?.(event.pointerId);
    if (moved) {
      saveState();
      if (dragType === "label") {
        setStatus("Etiqueta del grafico movida.");
      } else if (dragTarget === "legend") {
        setStatus("Leyenda reposicionada dentro del widget.");
      } else {
        setStatus("Grafico reposicionado dentro del widget.");
      }
      suppressClickUntil = Date.now() + 280;
    }
    canvas.style.cursor = "default";
  });

  canvas.addEventListener("pointercancel", (event) => {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }
    dragState = null;
    canvas.releasePointerCapture?.(event.pointerId);
    canvas.style.cursor = "default";
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const nextX = dragState.startOffsetX + (x - dragState.startX);
    const nextY = dragState.startOffsetY + (y - dragState.startY);
    if (dragState.type === "label") {
      const safeOffsets = normalizeBiLabelOffsets(widget.labelOffsets);
      safeOffsets[String(dragState.index)] = {
        x: Math.max(-400, Math.min(400, Math.round(nextX))),
        y: Math.max(-220, Math.min(220, Math.round(nextY)))
      };
      widget.labelOffsets = safeOffsets;
      if (dragState.rowKey || dragState.rowLabel) {
        const dxPct = Math.max(-1.5, Math.min(1.5, dragState.startDxPct + ((x - dragState.startX) / Math.max(1, dragState.plotWidth))));
        const dyPct = Math.max(-1.5, Math.min(1.5, dragState.startDyPct + ((y - dragState.startY) / Math.max(1, dragState.plotHeight))));
        upsertBiLabelPlacement(widget, dragState.rowKey, {
          anchorXPct: 0.5,
          anchorYPct: 0.5,
          dxPct,
          dyPct,
          scope: "plot"
        }, dragState.rowLabel || labels[dragState.index] || "");
      }
      redraw(dragState.index);
      highlightedIndex = dragState.index;
    } else {
      const safeLayout = normalizeBiCircularLayout(widget.polarLayout);
      if (dragState.target === "chart") {
        safeLayout.chart = {
          x: Math.max(-520, Math.min(520, Math.round(nextX))),
          y: Math.max(-320, Math.min(320, Math.round(nextY)))
        };
      } else {
        safeLayout.legend = {
          x: Math.max(-520, Math.min(520, Math.round(nextX))),
          y: Math.max(-320, Math.min(320, Math.round(nextY)))
        };
      }
      widget.polarLayout = safeLayout;
      redraw(highlightedIndex);
    }
    tooltip.classList.add("hidden");
    canvas.style.cursor = "grabbing";
    dragState.moved = Math.abs(x - dragState.startX) > 1 || Math.abs(y - dragState.startY) > 1;
  });

  canvas.addEventListener("click", (event) => {
    if (Date.now() < suppressClickUntil) {
      return;
    }
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const index = resolveBiHoverIndex(model, x, y);
    if (index < 0 || index >= snapshot.rows.length) {
      return;
    }
    if (typeof onSelect === "function") {
      onSelect(index, event);
    }
  });
}

function renderBiWidgetCustomContent(target, snapshot, widget, project) {
  if (!(target instanceof HTMLElement) || !snapshot || !widget) {
    return;
  }
  const chartType = normalizeBiChartType(widget.chartType || "bar");
  const typeConfig = getBiSpecificChartConfigForWidget(widget, chartType);
  if (chartType === "table") {
    const groupLabel = getBiGroupLabel(snapshot.groupBy, project);
    const metricLabel = getBiMetricLabel(snapshot.metric);
    const rowLimit = sanitizeBiInteger(typeConfig.rowLimit, BI_WIDGET_TABLE_LIMIT, 1, 200);
    const rows = snapshot.rows.slice(0, rowLimit);
    const showIndex = normalizeBiToggle(typeConfig.showIndex, true);
    const compactClass = normalizeBiToggle(typeConfig.compact, false) ? " compact" : "";
    const body = rows.length === 0
      ? `<tr><td colspan="3" class="muted">Sin datos.</td></tr>`
      : rows.map((row, index) => `<tr data-bi-row-index="${index}" data-bi-label="${escapeAttribute(row.label)}">
          <td>${showIndex ? (index + 1) : ""}</td>
          <td>${escapeHtml(row.label)}</td>
          <td>${escapeHtml(row.valueText)}</td>
        </tr>`).join("");
    target.innerHTML = `<div class="bi-widget-table-view${compactClass}">
      <table>
        <thead>
          <tr>
            <th>${showIndex ? "#" : ""}</th>
            <th>${escapeHtml(groupLabel)}</th>
            <th>${escapeHtml(metricLabel)}</th>
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>`;
    return;
  }

  if (chartType === "scorecard") {
    const total = snapshot.rows.reduce((sum, row) => sum + (row.value || 0), 0);
    const valueText = `${trimOrFallback(typeConfig.prefix, "")}${formatBiMetricValue(snapshot.metric, total)}${trimOrFallback(typeConfig.suffix, "")}`;
    const showSource = normalizeBiToggle(typeConfig.showSource, true);
    const showGroups = normalizeBiToggle(typeConfig.showGroups, true);
    const subTokens = [];
    if (showSource) {
      subTokens.push(`Fuente: ${getBiSourceLabel(snapshot.source)}`);
    }
    if (showGroups) {
      subTokens.push(`grupos: ${snapshot.groupCount}`);
    }
    target.innerHTML = `<div class="bi-scorecard-view" data-bi-row-index="0">
      <div class="bi-scorecard-value">${escapeHtml(valueText)}</div>
      <div class="bi-scorecard-label">${escapeHtml(getBiMetricLabel(snapshot.metric))}</div>
      <div class="bi-scorecard-sub">${escapeHtml(subTokens.join(" | "))}</div>
    </div>`;
    return;
  }

  if (chartType === "pivot") {
    const rowLabel = getBiGroupLabel(snapshot.pivot?.rowDimension || snapshot.groupBy, project);
    const colLabels = Array.isArray(snapshot.pivot?.columns) ? snapshot.pivot.columns : [];
    const rowLimit = sanitizeBiInteger(typeConfig.rowLimit, BI_WIDGET_TABLE_LIMIT, 1, 200);
    const showTotals = normalizeBiToggle(typeConfig.showTotals, true);
    const compactClass = normalizeBiToggle(typeConfig.compact, false) ? " compact" : "";
    const rows = (Array.isArray(snapshot.pivot?.rows) ? snapshot.pivot.rows : []).slice(0, rowLimit);
    const headerCols = colLabels.map((label) => `<th>${escapeHtml(label)}</th>`).join("");
    const body = rows.length === 0
      ? `<tr><td colspan="${2 + colLabels.length}" class="muted">Sin datos.</td></tr>`
      : rows.map((row, rowIndex) => {
        const cells = (row.cells || []).map((cell) => `<td>${escapeHtml(formatBiMetricValue(snapshot.metric, cell.value || 0))}</td>`).join("");
        const totalCell = showTotals ? `<td><strong>${escapeHtml(formatBiMetricValue(snapshot.metric, row.total || 0))}</strong></td>` : "";
        return `<tr data-bi-row-index="${rowIndex}" data-bi-label="${escapeAttribute(row.label)}">
          <td>${rowIndex + 1}</td>
          <td>${escapeHtml(row.label)}</td>
          ${cells}
          ${totalCell}
        </tr>`;
      }).join("");
    target.innerHTML = `<div class="bi-widget-table-view${compactClass}">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>${escapeHtml(rowLabel)}</th>
            ${headerCols}
            ${showTotals ? "<th>Total</th>" : ""}
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>`;
    return;
  }

  if (chartType === "sankey") {
    const maxLinks = sanitizeBiInteger(typeConfig.maxLinks, 12, 1, 200);
    const minLinkPercent = sanitizeBiDecimal(typeConfig.minLinkPercent, 2, 0, 50);
    const showValues = normalizeBiToggle(typeConfig.showValues, true);
    const links = (Array.isArray(snapshot.sankey?.links) ? snapshot.sankey.links : []).slice(0, maxLinks);
    const maxValue = links.length > 0 ? Math.max(...links.map((item) => Number(item.value) || 0), 1) : 1;
    const body = links.length === 0
      ? `<div class="bi-empty">Sin flujos para mostrar.</div>`
      : links.map((link, index) => {
        const ratio = Math.max(minLinkPercent / 100, Math.min(1, (Number(link.value) || 0) / maxValue));
        return `<div class="bi-sankey-row" data-bi-row-index="${index}" data-bi-label="${escapeAttribute(link.from)}">
          <div class="bi-sankey-labels">
            <span>${escapeHtml(link.from)}</span>
            <span class="muted">-></span>
            <span>${escapeHtml(link.to)}</span>
          </div>
          <div class="bi-sankey-bar">
            <div class="bi-sankey-bar-fill" style="width:${Math.round(ratio * 100)}%"></div>
          </div>
          ${showValues ? `<strong>${escapeHtml(formatBiMetricValue(snapshot.metric, link.value || 0))}</strong>` : ""}
        </div>`;
      }).join("");
    target.innerHTML = `<div class="bi-sankey-view">${body}</div>`;
    return;
  }

  if (chartType === "timeline") {
    const rowLimit = sanitizeBiInteger(typeConfig.rowLimit, 20, 1, 200);
    const showDates = normalizeBiToggle(typeConfig.showDates, true);
    const minBarPercent = sanitizeBiDecimal(typeConfig.minBarPercent, 1, 0.2, 50);
    const items = (Array.isArray(snapshot.timeline?.items) ? snapshot.timeline.items : []).slice(0, rowLimit);
    const minDay = Number.isFinite(snapshot.timeline?.minDay) ? snapshot.timeline.minDay : 0;
    const maxDay = Number.isFinite(snapshot.timeline?.maxDay) ? snapshot.timeline.maxDay : (minDay + 1);
    const daySpan = Math.max(1, maxDay - minDay + 1);
    const body = items.length === 0
      ? `<div class="bi-empty">Sin rango de fechas para el cronograma.</div>`
      : items.map((item, index) => {
        const leftPct = Math.max(0, Math.min(100, ((item.startDay - minDay) / daySpan) * 100));
        const widthPct = Math.max(minBarPercent, Math.min(100, (((item.endDay - item.startDay + 1) / daySpan) * 100)));
        return `<div class="bi-timeline-row" data-bi-row-index="${index}" data-bi-label="${escapeAttribute(item.label)}">
          <div class="bi-timeline-meta">
            <strong>${escapeHtml(item.label)}</strong>
            ${showDates ? `<span class="muted">${escapeHtml(formatDateFromInput(item.startDate))} - ${escapeHtml(formatDateFromInput(item.endDate))}</span>` : ""}
          </div>
          <div class="bi-timeline-track">
            <div class="bi-timeline-bar" style="left:${leftPct}%;width:${widthPct}%"></div>
          </div>
        </div>`;
      }).join("");
    target.innerHTML = `<div class="bi-timeline-view">${body}</div>`;
    return;
  }

  target.innerHTML = "";
}

function normalizeBiDateDimension(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "enddate" || token === "fin" || token === "fechafin") {
    return "endDate";
  }
  if (token === "createdat" || token === "creado") {
    return "createdAt";
  }
  return "startDate";
}

function getBiRowMetricValue(row, metric) {
  const safeMetric = normalizeBiMetric(metric || "count");
  if (safeMetric === "count") {
    return 1;
  }
  if (safeMetric === "baseunits") {
    return Number.isFinite(row?.baseUnits) ? row.baseUnits : 0;
  }
  if (safeMetric === "baseavg" || safeMetric === "basemax" || safeMetric === "basemin") {
    return Number.isFinite(row?.baseUnits) ? row.baseUnits : 0;
  }
  if (safeMetric === "realavg" || safeMetric === "realmax" || safeMetric === "realmin") {
    return Number.isFinite(row?.realProgress) ? row.realProgress : 0;
  }
  if (safeMetric === "programmedavg" || safeMetric === "programmedmax" || safeMetric === "programmedmin") {
    return Number.isFinite(row?.programmedPercent) ? row.programmedPercent : 0;
  }
  if (safeMetric === "weightedreal") {
    return Number.isFinite(row?.weightedReal) ? row.weightedReal : 0;
  }
  if (safeMetric === "weightedprogrammed") {
    return Number.isFinite(row?.weightedProgrammed) ? row.weightedProgrammed : 0;
  }
  if (safeMetric === "weightedgap") {
    const real = Number.isFinite(row?.weightedReal) ? row.weightedReal : 0;
    const programmed = Number.isFinite(row?.weightedProgrammed) ? row.weightedProgrammed : 0;
    return real - programmed;
  }
  if (safeMetric === "invaliddates") {
    return row?.invalidDates ? 1 : 0;
  }
  return 0;
}

function computeBiQuantile(values, q) {
  if (!Array.isArray(values) || values.length === 0) {
    return 0;
  }
  const safe = values
    .map((item) => Number(item))
    .filter((item) => Number.isFinite(item))
    .sort((a, b) => a - b);
  if (safe.length === 0) {
    return 0;
  }
  const pos = (safe.length - 1) * Math.max(0, Math.min(1, q));
  const base = Math.floor(pos);
  const rest = pos - base;
  if (safe[base + 1] === undefined) {
    return safe[base];
  }
  return safe[base] + rest * (safe[base + 1] - safe[base]);
}

function sortBiGroupedRows(grouped, sortMode) {
  return grouped.sort((a, b) => {
    if (sortMode === "label_asc") {
      return a.label.localeCompare(b.label, "es");
    }
    if (sortMode === "label_desc") {
      return b.label.localeCompare(a.label, "es");
    }
    if (sortMode === "value_asc") {
      if (a.value !== b.value) {
        return a.value - b.value;
      }
      return a.label.localeCompare(b.label, "es");
    }
    if (b.value !== a.value) {
      return b.value - a.value;
    }
    return a.label.localeCompare(b.label, "es");
  });
}

function buildBiWidgetSnapshot(rows, widget) {
  const source = normalizeBiSource(widget?.source || "all");
  const chartType = normalizeBiChartType(widget?.chartType || "bar");
  const sortMode = normalizeBiSortMode(widget?.sortMode || "value_desc");
  const topN = sanitizeBiTopN(widget?.topN ?? 10);
  const dataRoles = normalizeBiDataRoles(widget?.dataRoles, widget);
  const groupBy = normalizeBiGroupBy(dataRoles.dimensions?.[0] || widget?.groupBy || "disciplina");
  const metric = normalizeBiMetric(dataRoles.metrics?.[0] || widget?.metric || "count");
  const optionalMetric = normalizeBiMetric(dataRoles.optionalMetrics?.[0] || "baseunits");
  const breakdownDimension = normalizeBiOptionalGroupBy(dataRoles.breakdownDimension || "");
  const effectiveBreakdownDimension = breakdownDimension || "fuente";
  const dateDimension = normalizeBiDateDimension(dataRoles.dateDimension || "startDate");
  const sourceRows = source === "all"
    ? rows.slice()
    : rows.filter((row) => row.source === source);
  const canUseBreakdownSeries = !!breakdownDimension && new Set([
    "bar", "bar_horizontal", "bar_stacked", "line", "timeseries", "area", "combo", "pie", "donut",
    "scatter", "bubble", "radar", "waterfall", "pareto", "treemap", "funnel"
  ]).has(chartType);
  const rowsBySnapshotLabel = new Map();
  sourceRows.forEach((row) => {
    const groupLabel = getBiRowGroupValue(row, groupBy);
    const breakdownLabel = canUseBreakdownSeries ? getBiRowGroupValue(row, breakdownDimension) : "";
    const label = canUseBreakdownSeries
      ? `${groupLabel} | ${breakdownLabel}`
      : groupLabel;
    if (!rowsBySnapshotLabel.has(label)) {
      rowsBySnapshotLabel.set(label, {
        rows: [],
        groupLabel,
        breakdownLabel
      });
    }
    rowsBySnapshotLabel.get(label).rows.push(row);
  });
  const groupedRaw = Array.from(rowsBySnapshotLabel.entries()).map(([label, bucket]) => {
    const groupedBucket = buildBiGroupedRows(bucket.rows, groupBy);
    const aggregate = groupedBucket[0] || {
      label,
      rows: bucket.rows.slice(),
      count: bucket.rows.length,
      baseUnits: 0,
      realAvg: 0,
      programmedAvg: 0,
      weightedReal: 0,
      weightedProgrammed: 0,
      invalidDates: 0
    };
    return {
      ...aggregate,
      label,
      groupLabel: bucket.groupLabel,
      breakdownLabel: bucket.breakdownLabel,
      colorKey: canUseBreakdownSeries ? bucket.breakdownLabel : bucket.groupLabel
    };
  });

  const grouped = groupedRaw
    .map((group) => {
      const value = getBiMetricValue(group, metric);
      const safeValue = Number.isFinite(value) ? value : 0;
      const sourceGroupRows = rowsBySnapshotLabel.get(group.label)?.rows || [];
      const sourceGroupMetricValues = sourceGroupRows
        .map((row) => getBiRowMetricValue(row, metric))
        .filter((item) => Number.isFinite(item));
      const optionalValue = getBiMetricValue(group, optionalMetric);
      const xValue = Number.isFinite(optionalValue) ? optionalValue : group.count;
      const q1 = computeBiQuantile(sourceGroupMetricValues, 0.25);
      const median = computeBiQuantile(sourceGroupMetricValues, 0.5);
      const q3 = computeBiQuantile(sourceGroupMetricValues, 0.75);
      const min = sourceGroupMetricValues.length > 0 ? Math.min(...sourceGroupMetricValues) : 0;
      const max = sourceGroupMetricValues.length > 0 ? Math.max(...sourceGroupMetricValues) : 0;
      const orderedByDate = sourceGroupRows
        .slice()
        .sort((left, right) => {
          const a = sanitizeDateInput(left?.[dateDimension] || left?.startDate || "");
          const b = sanitizeDateInput(right?.[dateDimension] || right?.startDate || "");
          return a.localeCompare(b);
        });
      const open = orderedByDate.length > 0 ? getBiRowMetricValue(orderedByDate[0], metric) : safeValue;
      const close = orderedByDate.length > 0 ? getBiRowMetricValue(orderedByDate[orderedByDate.length - 1], metric) : safeValue;
      const rowKey = buildBiRowKey(
        source,
        canUseBreakdownSeries ? `${groupBy}_${breakdownDimension}` : groupBy,
        group.label
      );
      return {
        label: group.label,
        groupLabel: trimOrFallback(group.groupLabel, group.label),
        breakdownLabel: trimOrFallback(group.breakdownLabel, ""),
        colorKey: trimOrFallback(group.colorKey, group.label),
        rowKey,
        value: safeValue,
        valueText: formatBiMetricValue(metric, safeValue),
        count: group.count,
        baseUnits: group.baseUnits,
        realAvg: group.realAvg,
        programmedAvg: group.programmedAvg,
        weightedReal: group.weightedReal,
        weightedProgrammed: group.weightedProgrammed,
        invalidDates: group.invalidDates,
        xValue,
        bubbleSize: Math.max(6, Math.min(36, Math.round((group.baseUnits || group.count || 1) * 0.4))),
        box: { min, q1, median, q3, max },
        candle: { open, high: max, low: min, close }
      };
    });
  sortBiGroupedRows(grouped, sortMode);
  const trimmed = grouped.slice(0, topN);

  const pivotColumnsSet = new Set();
  const pivotMap = new Map();
  sourceRows.forEach((row) => {
    const rowLabel = getBiRowGroupValue(row, groupBy);
    const colLabel = getBiRowGroupValue(row, effectiveBreakdownDimension);
    const cellValue = getBiRowMetricValue(row, metric);
    pivotColumnsSet.add(colLabel);
    if (!pivotMap.has(rowLabel)) {
      pivotMap.set(rowLabel, new Map());
    }
    const rowBucket = pivotMap.get(rowLabel);
    rowBucket.set(colLabel, (rowBucket.get(colLabel) || 0) + cellValue);
  });
  const pivotColumns = Array.from(pivotColumnsSet.values()).sort((a, b) => a.localeCompare(b, "es"));
  const pivotRows = Array.from(pivotMap.entries()).map(([rowLabel, rowBucket]) => {
    const cells = pivotColumns.map((columnLabel) => {
      const value = rowBucket.get(columnLabel) || 0;
      return {
        key: buildBiRowKey(source, `${groupBy}_${effectiveBreakdownDimension}`, `${rowLabel}|${columnLabel}`),
        label: columnLabel,
        value
      };
    });
    const total = cells.reduce((sum, cell) => sum + (cell.value || 0), 0);
    return {
      rowKey: buildBiRowKey(source, groupBy, rowLabel),
      label: rowLabel,
      cells,
      total
    };
  }).sort((a, b) => b.total - a.total).slice(0, topN);

  const sankeyMap = new Map();
  sourceRows.forEach((row) => {
    const fromLabel = getBiRowGroupValue(row, groupBy);
    const toLabel = getBiRowGroupValue(row, effectiveBreakdownDimension);
    const key = `${fromLabel}__${toLabel}`;
    const value = getBiRowMetricValue(row, metric);
    sankeyMap.set(key, (sankeyMap.get(key) || 0) + value);
  });
  const sankeyLinks = Array.from(sankeyMap.entries())
    .map(([key, value]) => {
      const [from, to] = key.split("__");
      return {
        rowKey: buildBiRowKey(source, `${groupBy}_${effectiveBreakdownDimension}`, `${from}->${to}`),
        from,
        to,
        value
      };
    })
    .sort((a, b) => b.value - a.value)
    .slice(0, topN);

  const timelineItems = sourceRows
    .map((row, index) => {
      const label = getBiRowGroupValue(row, groupBy);
      const startDate = sanitizeDateInput(row?.[dateDimension] || row?.startDate || "");
      const endDate = sanitizeDateInput(row?.endDate || row?.startDate || "");
      if (!startDate || !endDate) {
        return null;
      }
      const startDay = dateInputToDayIndex(startDate);
      const endDay = dateInputToDayIndex(endDate);
      if (!Number.isFinite(startDay) || !Number.isFinite(endDay)) {
        return null;
      }
      const safeStartDay = Math.min(startDay, endDay);
      const safeEndDay = Math.max(startDay, endDay);
      return {
        rowKey: buildBiRowKey(source, groupBy, `${label}_${startDate}_${endDate}_${index}`),
        label,
        startDate,
        endDate,
        startDay: safeStartDay,
        endDay: safeEndDay,
        value: getBiRowMetricValue(row, metric)
      };
    })
    .filter((item) => !!item)
    .sort((a, b) => {
      if (a.startDay !== b.startDay) {
        return a.startDay - b.startDay;
      }
      return a.label.localeCompare(b.label, "es");
    })
    .slice(0, topN);

  const timelineMin = timelineItems.length > 0 ? Math.min(...timelineItems.map((item) => item.startDay)) : 0;
  const timelineMax = timelineItems.length > 0 ? Math.max(...timelineItems.map((item) => item.endDay)) : 1;

  return {
    source,
    groupBy,
    metric,
    chartType,
    sortMode,
    topN,
    sourceRowsCount: sourceRows.length,
    groupCount: grouped.length,
    dataRoles,
    breakdownDimension: effectiveBreakdownDimension,
    hasBreakdownSeries: canUseBreakdownSeries,
    colorByDimension: canUseBreakdownSeries ? breakdownDimension : groupBy,
    dateDimension,
    rows: trimmed,
    pivot: {
      rowDimension: groupBy,
      colDimension: effectiveBreakdownDimension,
      columns: pivotColumns,
      rows: pivotRows
    },
    sankey: {
      fromDimension: groupBy,
      toDimension: effectiveBreakdownDimension,
      links: sankeyLinks
    },
    timeline: {
      items: timelineItems,
      minDay: timelineMin,
      maxDay: timelineMax
    }
  };
}

function buildBiWidgetMetaLabel(widget, snapshot, project) {
  const kind = normalizeBiWidgetKind(widget?.kind || "chart");
  if (kind === "text") {
    return "Cuadro de texto";
  }
  if (kind === "image") {
    return "Imagen";
  }
  if (!snapshot) {
    return "Widget";
  }
  const groupLabel = getBiGroupLabel(snapshot.groupBy, project);
  const metricLabel = getBiMetricLabel(snapshot.metric);
  const sourceLabel = getBiSourceLabel(snapshot.source);
  return `${sourceLabel} | ${metricLabel} por ${groupLabel}`;
}

function renderBiTextWidgetContent(target, widget) {
  if (!(target instanceof HTMLElement) || !widget) {
    return;
  }
  const textConfig = normalizeBiTextConfig(widget.textConfig);
  const text = typeof widget.textContent === "string"
    ? widget.textContent.slice(0, 4000)
    : "";
  const style = [
    `font-family:${textConfig.fontFamily}`,
    `font-size:${textConfig.fontSize}px`,
    `font-weight:${textConfig.fontWeight}`,
    `font-style:${textConfig.fontStyle}`,
    `text-align:${textConfig.textAlign}`,
    `line-height:${textConfig.lineHeight}`,
    `color:${textConfig.color}`,
    `padding:${textConfig.padding}px`
  ].join(";");
  target.innerHTML = `<div class="bi-text-widget-box" data-bi-text-widget="${escapeAttribute(widget.id)}" data-placeholder="Escribe texto..." contenteditable="true" spellcheck="true" style="${escapeAttribute(style)}">${escapeHtml(text)}</div>`;
}

function renderBiImageWidgetContent(target, widget) {
  if (!(target instanceof HTMLElement) || !widget) {
    return;
  }
  const imageConfig = normalizeBiImageConfig(widget.imageConfig);
  const src = trimOrFallback(widget.imageSrc, "");
  const alt = trimOrFallback(widget.imageAlt || widget.name || "Imagen", "Imagen");
  const frameStyle = `background:${imageConfig.background};border-radius:${imageConfig.borderRadius}px;border-color:${imageConfig.frameColor};`;
  const imageStyle = `object-fit:${imageConfig.fit};opacity:${imageConfig.opacity};border-radius:${imageConfig.borderRadius}px;background:${imageConfig.background};`;
  const preview = src
    ? `<img src="${escapeAttribute(src)}" alt="${escapeAttribute(alt)}" class="bi-image-widget-preview" style="${escapeAttribute(imageStyle)}">`
    : `<div class="bi-image-widget-empty">Sin imagen</div>`;
  target.innerHTML = `<div class="bi-image-widget">
    <div class="bi-image-widget-frame" style="${escapeAttribute(frameStyle)}">${preview}</div>
    <div class="bi-image-widget-actions">
      <button type="button" class="secondary" data-bi-select-image="${escapeAttribute(widget.id)}">Subir</button>
      <input type="url" class="bi-image-widget-url" data-bi-image-url="${escapeAttribute(widget.id)}" placeholder="https://... o pega data URL" value="${escapeAttribute(src)}">
      <button type="button" class="secondary" data-bi-apply-image-url="${escapeAttribute(widget.id)}">Aplicar URL</button>
    </div>
  </div>`;
}

function renderBiWidgets(project, rows) {
  if (!project || !els.biWidgetsGrid) {
    return;
  }
  const renderStartedAt = typeof performance !== "undefined" && typeof performance.now === "function"
    ? performance.now()
    : Date.now();

  const widgets = Array.isArray(project.biWidgets) ? project.biWidgets : [];
  const snapshots = {};
  biWidgetSnapshotsByProject[project.id] = snapshots;

  if (els.biWidgetsMetaText) {
    els.biWidgetsMetaText.textContent = `Widgets: ${widgets.length}/${BI_MAX_WIDGETS} | Filas: ${rows.length}`;
  }

  const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
  const surfaceWidth = canvasSize.width;
  const surfaceHeight = canvasSize.height;
  const zoomPercent = sanitizeBiCanvasZoom(project?.biConfig?.canvasZoom, BI_CANVAS_ZOOM_DEFAULT);
  const zoomScale = getBiCanvasZoomScale(project?.biConfig);
  const scaledSurfaceWidth = Math.max(1, Math.round(surfaceWidth * zoomScale));
  const scaledSurfaceHeight = Math.max(1, Math.round(surfaceHeight * zoomScale));
  const showCanvasGrid = normalizeBiToggle(project?.biConfig?.showCanvasGrid, false);
  const gridSize = sanitizeBiInteger(project?.biConfig?.gridSnapSize, 12, 4, 120);
  const canvasClass = showCanvasGrid ? "bi-canvas-surface show-grid" : "bi-canvas-surface";
  const canvasStyleSuffix = showCanvasGrid ? `;--bi-canvas-grid-size:${gridSize}px` : "";
  const stageStyle = `--bi-canvas-zoom:${zoomScale};`;
  const zoomWrapStyle = `width:${scaledSurfaceWidth}px;height:${scaledSurfaceHeight}px;`;
  const surfaceStyle = `position:absolute;left:0;top:0;width:${surfaceWidth}px;height:${surfaceHeight}px;min-width:${surfaceWidth}px;min-height:${surfaceHeight}px;transform:scale(${zoomScale});transform-origin:top left;${canvasStyleSuffix}`;

  if (widgets.length === 0) {
    els.biWidgetsGrid.innerHTML = `<div class="bi-canvas-stage" style="${escapeAttribute(stageStyle)}">
      <div class="bi-canvas-zoom-wrap" style="${escapeAttribute(zoomWrapStyle)}">
        <div class="${canvasClass}" data-bi-canvas-surface data-bi-canvas-width="${surfaceWidth}" data-bi-canvas-height="${surfaceHeight}" data-bi-canvas-zoom="${zoomPercent}" style="${escapeAttribute(surfaceStyle)}">
          <div class="bi-empty">Lienzo vacio. Configura Dimension/Metrica y usa "Agregar widget".</div>
        </div>
      </div>
    </div>`;
    const renderEndedAt = typeof performance !== "undefined" && typeof performance.now === "function"
      ? performance.now()
      : Date.now();
    updateBiRuntimeStats(project.id, {
      lastRenderMs: Math.max(0, renderEndedAt - renderStartedAt),
      widgets: 0,
      rows: Array.isArray(rows) ? rows.length : 0,
      points: 0
    });
    return;
  }

  const widgetsHtml = widgets
    .map((widget, index) => {
      const normalizedWidget = normalizeBiWidget(widget, index);
      widget.layout = normalizeBiWidgetLayout(normalizedWidget.layout, index, normalizedWidget);
      widget.kind = normalizedWidget.kind;
      widget.labelOffsets = normalizeBiLabelOffsets(normalizedWidget.labelOffsets);
      widget.labelLayoutV2 = normalizeBiLabelLayoutV2(normalizedWidget.labelLayoutV2);
      widget.chartTypeConfig = normalizeBiChartTypeConfigMap(normalizedWidget.chartTypeConfig);
      widget.textConfig = normalizeBiTextConfig(normalizedWidget.textConfig);
      widget.imageConfig = normalizeBiImageConfig(normalizedWidget.imageConfig);
      widget.polarLayout = normalizeBiCircularLayout(normalizedWidget.polarLayout);
      widget.showTitle = normalizedWidget.showTitle !== false;
      widget.showBorder = normalizedWidget.showBorder !== false;
      widget.showBackground = normalizedWidget.showBackground !== false;
      widget.layout.x = Math.max(0, Math.min(surfaceWidth - widget.layout.w, widget.layout.x));
      widget.layout.y = Math.max(0, Math.min(surfaceHeight - widget.layout.h, widget.layout.y));
      const widgetKind = normalizeBiWidgetKind(normalizedWidget.kind || "chart");
      const isChartWidget = widgetKind === "chart";
      const snapshot = isChartWidget ? buildBiWidgetSnapshot(rows, normalizedWidget) : null;
      const effectiveVisual = isChartWidget
        ? getBiEffectiveVisualSettings(project, normalizedWidget)
        : createDefaultBiVisualSettings();
      if (snapshot) {
        snapshots[normalizedWidget.id] = snapshot;
      }

      const sourceLabel = buildBiWidgetMetaLabel(normalizedWidget, snapshot, project);
      const chartCanvasId = `biWidgetCanvas_${normalizedWidget.id}`;
      const contentTargetId = `biWidgetContent_${normalizedWidget.id}`;
      const tooltipId = `biWidgetTooltip_${normalizedWidget.id}`;
      const useCanvas = !!(snapshot && !BI_CUSTOM_CONTENT_TYPES.has(snapshot.chartType));
      const selectedClass = selectedBiWidgetId === normalizedWidget.id ? " selected" : "";
      const widgetKindClass = widgetKind === "text"
        ? " is-text-widget"
        : (widgetKind === "image" ? " is-image-widget" : "");
      const showTitle = normalizedWidget.showTitle !== false;
      const showBorder = normalizedWidget.showBorder !== false;
      const showBackground = normalizedWidget.showBackground !== false;
      const locked = normalizeBiToggle(normalizedWidget.locked, false);
      const borderClass = showBorder ? "" : " no-border";
      const titleClass = showTitle ? "" : " no-title";
      const backgroundClass = showBackground ? "" : " no-background";
      const lockedClass = locked ? " locked" : "";
      const titleSize = getBiFontSizeByLayer(effectiveVisual, "title", 16);
      const titleFamily = getBiFontFamilyByLayer(effectiveVisual, "title");
      const metaSize = Math.max(9, Math.round(titleSize * 0.68));
      const labelFamily = getBiFontFamilyByLayer(effectiveVisual, "label");
      const titleStyle = `font-size:${titleSize}px;font-family:${escapeAttribute(titleFamily)};`;
      const metaStyle = `font-size:${metaSize}px;font-family:${escapeAttribute(labelFamily)};`;
      const lockLabel = locked ? "Desbloquear" : "Bloquear";
      const lockClass = locked ? "warning" : "secondary";

      return `<article class="bi-canvas-widget${selectedClass}${widgetKindClass}${borderClass}${titleClass}${backgroundClass}${lockedClass}" data-bi-widget-id="${escapeAttribute(normalizedWidget.id)}" style="left:${widget.layout.x}px;top:${widget.layout.y}px;width:${widget.layout.w}px;height:${widget.layout.h}px;">
        <div class="bi-canvas-widget-header" data-bi-drag-widget="${escapeAttribute(normalizedWidget.id)}">
          <div class="bi-canvas-widget-title-wrap${showTitle ? "" : " hidden"}">
            <div class="bi-widget-title" style="${titleStyle}">${escapeHtml(normalizedWidget.name)}</div>
            <div class="bi-widget-meta" style="${metaStyle}">${escapeHtml(sourceLabel)}</div>
          </div>
          <div class="inline-actions">
            <button type="button" class="secondary" data-bi-duplicate-widget="${escapeAttribute(normalizedWidget.id)}">Clonar</button>
            <button type="button" class="${lockClass}" data-bi-toggle-lock-widget="${escapeAttribute(normalizedWidget.id)}">${lockLabel}</button>
            <button type="button" class="secondary" data-bi-export-widget-png="${escapeAttribute(normalizedWidget.id)}">PNG</button>
            <button type="button" class="secondary" data-bi-export-widget="${escapeAttribute(normalizedWidget.id)}">Exportar</button>
            <button type="button" class="danger" data-bi-remove-widget="${escapeAttribute(normalizedWidget.id)}">Eliminar</button>
          </div>
        </div>
        <div class="bi-canvas-widget-body">
          <canvas id="${escapeAttribute(chartCanvasId)}" class="bi-widget-chart${useCanvas ? "" : " hidden"}"></canvas>
          <div id="${escapeAttribute(contentTargetId)}" class="bi-widget-content${useCanvas ? " hidden" : ""}" data-bi-widget-kind="${escapeAttribute(widgetKind)}"></div>
          <div id="${escapeAttribute(tooltipId)}" class="bi-chart-tooltip hidden"></div>
        </div>
        <div class="bi-resize-handle" data-bi-resize-widget="${escapeAttribute(normalizedWidget.id)}"></div>
      </article>`;
    })
    .join("");

  els.biWidgetsGrid.innerHTML = `<div class="bi-canvas-stage" style="${escapeAttribute(stageStyle)}">
    <div class="bi-canvas-zoom-wrap" style="${escapeAttribute(zoomWrapStyle)}">
      <div class="${canvasClass}" data-bi-canvas-surface data-bi-canvas-width="${surfaceWidth}" data-bi-canvas-height="${surfaceHeight}" data-bi-canvas-zoom="${zoomPercent}" style="${escapeAttribute(surfaceStyle)}">${widgetsHtml}</div>
    </div>
  </div>`;

  widgets.forEach((widget, index) => {
    const normalizedWidget = normalizeBiWidget(widget, index);
    widget.kind = normalizedWidget.kind;
    widget.labelOffsets = normalizeBiLabelOffsets(normalizedWidget.labelOffsets);
    widget.labelLayoutV2 = normalizeBiLabelLayoutV2(normalizedWidget.labelLayoutV2);
    widget.chartTypeConfig = normalizeBiChartTypeConfigMap(normalizedWidget.chartTypeConfig);
    widget.textConfig = normalizeBiTextConfig(normalizedWidget.textConfig);
    widget.imageConfig = normalizeBiImageConfig(normalizedWidget.imageConfig);
    widget.polarLayout = normalizeBiCircularLayout(normalizedWidget.polarLayout);
    widget.showTitle = normalizedWidget.showTitle !== false;
    widget.showBorder = normalizedWidget.showBorder !== false;
    widget.showBackground = normalizedWidget.showBackground !== false;
    widget.locked = normalizeBiToggle(normalizedWidget.locked, false);
    widget.sortMode = normalizeBiSortMode(normalizedWidget.sortMode || "value_desc");
    const widgetKind = normalizeBiWidgetKind(normalizedWidget.kind || "chart");
    const snapshot = widgetKind === "chart" ? snapshots[normalizedWidget.id] : null;
    const canvas = document.getElementById(`biWidgetCanvas_${normalizedWidget.id}`);
    const contentTarget = document.getElementById(`biWidgetContent_${normalizedWidget.id}`);
    const tooltip = document.getElementById(`biWidgetTooltip_${normalizedWidget.id}`);
    if (widgetKind === "text") {
      if (tooltip instanceof HTMLElement) {
        tooltip.classList.add("hidden");
      }
      renderBiTextWidgetContent(contentTarget, normalizedWidget);
      return;
    }
    if (widgetKind === "image") {
      if (tooltip instanceof HTMLElement) {
        tooltip.classList.add("hidden");
      }
      renderBiImageWidgetContent(contentTarget, normalizedWidget);
      return;
    }
    if (!snapshot) {
      return;
    }
    if (BI_CUSTOM_CONTENT_TYPES.has(snapshot.chartType)) {
      renderBiWidgetCustomContent(contentTarget, snapshot, normalizedWidget, project);
      if (tooltip instanceof HTMLElement) {
        tooltip.classList.add("hidden");
      }
      if (contentTarget instanceof HTMLElement) {
        contentTarget.addEventListener("click", (event) => {
          const target = event.target;
          if (!(target instanceof HTMLElement)) {
            return;
          }
          const rowNode = target.closest("[data-bi-row-index]");
          if (!(rowNode instanceof HTMLElement)) {
            return;
          }
          const label = trimOrFallback(rowNode.dataset.biLabel || "", "");
          if (label) {
            const appendMode = !!(event && (event.ctrlKey || event.metaKey));
            const changed = toggleBiCrossFilter(project, snapshot.groupBy, label, snapshot.source, appendMode);
            if (changed) {
              saveState();
              renderBiPanel(project);
              const hasFilter = normalizeBiCrossFilters(project.biConfig.crossFilters).length > 0;
              if (hasFilter) {
                const modeLabel = appendMode ? " (multi)" : "";
                setStatus(`Filtro cruzado aplicado${modeLabel}: ${getBiGroupLabel(snapshot.groupBy, project)} = ${label}`);
              } else {
                setStatus("Filtro cruzado removido.");
              }
            }
            return;
          }
          const rowIndex = parseInt(rowNode.dataset.biRowIndex || "", 10);
          if (Number.isNaN(rowIndex)) {
            return;
          }
          applyBiCrossFilterFromWidget(project, snapshot, rowIndex, event);
        });
      }
    } else {
      if (contentTarget instanceof HTMLElement) {
        contentTarget.innerHTML = "";
      }
      const seriesColors = getBiSnapshotSeriesColors(project, snapshot);
      const effectiveVisual = getBiEffectiveVisualSettings(project, normalizedWidget);
      const hoverModel = drawBiWidgetChart(
        canvas,
        snapshot.chartType,
        snapshot.rows.map((row) => row.label),
        snapshot.rows.map((row) => row.value),
        -1,
        seriesColors,
        effectiveVisual,
        normalizedWidget.labelOffsets,
        normalizedWidget.polarLayout,
        snapshot.rows,
        normalizedWidget.chartConfig,
        normalizedWidget.labelLayoutV2,
        normalizedWidget.chartTypeConfig
      );
      bindBiCanvasHover(canvas, tooltip, snapshot, snapshot.chartType, (index, event) => {
        applyBiCrossFilterFromWidget(project, snapshot, index, event);
      }, hoverModel, seriesColors, effectiveVisual, widget);
    }
  });

  const pointCount = Object.values(snapshots).reduce((sum, snapshot) => {
    const count = Array.isArray(snapshot?.rows) ? snapshot.rows.length : 0;
    return sum + count;
  }, 0);
  const renderEndedAt = typeof performance !== "undefined" && typeof performance.now === "function"
    ? performance.now()
    : Date.now();
  updateBiRuntimeStats(project.id, {
    lastRenderMs: Math.max(0, renderEndedAt - renderStartedAt),
    widgets: widgets.length,
    rows: Array.isArray(rows) ? rows.length : 0,
    points: pointCount
  });
}

function buildBiDuplicateName(baseName, existingNames) {
  const sourceName = trimOrFallback(baseName, "Widget").slice(0, 60);
  const normalizedExisting = new Set((existingNames || []).map((name) => normalizeLookup(name || "")));
  const baseCandidate = `${sourceName} copia`.slice(0, 60);
  if (!normalizedExisting.has(normalizeLookup(baseCandidate))) {
    return baseCandidate;
  }
  for (let index = 2; index <= 99; index += 1) {
    const candidate = `${sourceName} copia ${index}`.slice(0, 60);
    if (!normalizedExisting.has(normalizeLookup(candidate))) {
      return candidate;
    }
  }
  return `${sourceName} copia`.slice(0, 60);
}

function clampBiWidgetLayoutToCanvas(layout, canvasWidth, canvasHeight, minSize) {
  const minWidth = Math.max(80, Math.round(minSize?.width || BI_CANVAS_MIN_WIDTH));
  const minHeight = Math.max(52, Math.round(minSize?.height || BI_CANVAS_MIN_HEIGHT));
  const width = Math.max(minWidth, Math.round(canvasWidth || BI_CANVAS_SURFACE_MIN_WIDTH));
  const height = Math.max(minHeight, Math.round(canvasHeight || BI_CANVAS_SURFACE_HEIGHT));
  const w = Math.max(minWidth, Math.min(width, Math.round(layout?.w || minWidth)));
  const h = Math.max(minHeight, Math.min(height, Math.round(layout?.h || minHeight)));
  const x = Math.max(0, Math.min(width - w, Math.round(layout?.x || 0)));
  const y = Math.max(0, Math.min(height - h, Math.round(layout?.y || 0)));
  return { x, y, w, h };
}

function duplicateBiWidget(project, widgetId) {
  if (!project || !widgetId) {
    return { ok: false, message: "Widget no encontrado." };
  }
  ensureBiState(project);
  if (!Array.isArray(project.biWidgets) || project.biWidgets.length >= BI_MAX_WIDGETS) {
    return { ok: false, message: `Maximo de ${BI_MAX_WIDGETS} widgets alcanzado.` };
  }
  const sourceWidget = getBiWidgetById(project, widgetId);
  if (!sourceWidget) {
    return { ok: false, message: "Widget no encontrado." };
  }
  const normalized = normalizeBiWidget(sourceWidget, project.biWidgets.length);
  const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
  const minSize = getBiWidgetMinSize(normalized);
  const duplicated = {
    ...normalized,
    id: uid(),
    name: buildBiDuplicateName(normalized.name, project.biWidgets.map((item) => item?.name)),
    labelOffsets: normalizeBiLabelOffsets(normalized.labelOffsets),
    labelLayoutV2: normalizeBiLabelLayoutV2(normalized.labelLayoutV2),
    chartTypeConfig: normalizeBiChartTypeConfigMap(normalized.chartTypeConfig),
    textConfig: normalizeBiTextConfig(normalized.textConfig),
    imageConfig: normalizeBiImageConfig(normalized.imageConfig),
    polarLayout: normalizeBiCircularLayout(normalized.polarLayout),
    visualOverride: normalized.visualOverride ? normalizeBiVisualSettings(normalized.visualOverride) : null,
    locked: false
  };
  duplicated.layout = clampBiWidgetLayoutToCanvas({
    x: (normalized.layout?.x || 0) + 26,
    y: (normalized.layout?.y || 0) + 26,
    w: normalized.layout?.w || minSize.width,
    h: normalized.layout?.h || minSize.height
  }, canvasSize.width, canvasSize.height, minSize);
  project.biWidgets.push(duplicated);
  return { ok: true, newWidgetId: duplicated.id };
}

function toggleBiWidgetLock(project, widgetId) {
  const widget = getBiWidgetById(project, widgetId);
  if (!widget) {
    return { ok: false, locked: false };
  }
  widget.locked = !normalizeBiToggle(widget.locked, false);
  return { ok: true, locked: widget.locked };
}

function autoArrangeBiWidgets(project, options = {}) {
  if (!project || !Array.isArray(project.biWidgets) || project.biWidgets.length === 0) {
    return { arranged: 0 };
  }
  ensureBiState(project);
  const expandCanvas = options.expandCanvas !== false;
  const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
  const margin = 20;
  const gap = 14;
  const availableWidth = Math.max(280, canvasSize.width - (margin * 2));
  let cursorX = margin;
  let cursorY = margin;
  let rowHeight = 0;

  project.biWidgets.forEach((widget, index) => {
    const normalized = normalizeBiWidget(widget, index);
    const minSize = getBiWidgetMinSize(normalized);
    const fallbackWidth = normalized.kind === "text"
      ? 520
      : (normalized.kind === "image" ? 420 : 420);
    const fallbackHeight = normalized.kind === "text"
      ? 130
      : (normalized.kind === "image" ? 290 : 300);
    const safeWidth = Math.max(minSize.width, Math.round(normalized.layout?.w || fallbackWidth));
    const safeHeight = Math.max(minSize.height, Math.round(normalized.layout?.h || fallbackHeight));
    const widgetWidth = Math.min(availableWidth, safeWidth);

    if (cursorX > margin && cursorX + widgetWidth > (canvasSize.width - margin)) {
      cursorX = margin;
      cursorY += rowHeight + gap;
      rowHeight = 0;
    }

    widget.layout = clampBiWidgetLayoutToCanvas({
      x: cursorX,
      y: cursorY,
      w: widgetWidth,
      h: safeHeight
    }, canvasSize.width, canvasSize.height, minSize);

    cursorX = widget.layout.x + widget.layout.w + gap;
    rowHeight = Math.max(rowHeight, widget.layout.h);
  });

  const requiredHeight = cursorY + rowHeight + margin;
  if (expandCanvas && requiredHeight > canvasSize.height) {
    project.biConfig.canvasHeight = sanitizeBiCanvasDimension(
      requiredHeight,
      BI_CANVAS_SURFACE_HEIGHT,
      BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT,
      BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
    );
  }

  return { arranged: project.biWidgets.length };
}

function generateBiStarterDashboard(project) {
  if (!project) {
    return { added: 0 };
  }
  ensureBiState(project);
  const templates = [
    { name: "KPI Total de filas", source: "all", groupBy: "proyecto", metric: "count", chartType: "scorecard", topN: 1, sortMode: "value_desc" },
    { name: "Avance real por Disciplina", source: "all", groupBy: "disciplina", metric: "weightedreal", chartType: "bar", topN: 12, sortMode: "value_desc" },
    { name: "Programado por Mes de inicio", source: "deliverable", groupBy: "mes_inicio", metric: "programmedavg", chartType: "line", topN: 12, sortMode: "label_asc" },
    { name: "Pareto UP por Sistema", source: "all", groupBy: "sistema", metric: "baseunits", chartType: "pareto", topN: 12, sortMode: "value_desc" },
    { name: "Distribucion UP por Disciplina", source: "all", groupBy: "disciplina", metric: "baseunits", chartType: "donut", topN: 10, sortMode: "value_desc" },
    { name: "Mapa por Paquete", source: "package", groupBy: "paquete", metric: "baseunits", chartType: "treemap", topN: 12, sortMode: "value_desc" },
    { name: "Radar avance por Fase", source: "all", groupBy: "fase", metric: "realavg", chartType: "radar", topN: 10, sortMode: "label_asc" }
  ];

  const existingSignatures = new Set(
    project.biWidgets.map((widget) => {
      const normalized = normalizeBiWidget(widget, 0);
      return [
        normalized.source,
        normalized.groupBy,
        normalized.metric,
        normalized.chartType,
        normalizeBiSortMode(normalized.sortMode || "value_desc")
      ].join("|");
    })
  );

  let added = 0;
  templates.forEach((template) => {
    if (project.biWidgets.length >= BI_MAX_WIDGETS) {
      return;
    }
    const signature = [
      template.source,
      template.groupBy,
      template.metric,
      template.chartType,
      normalizeBiSortMode(template.sortMode || "value_desc")
    ].join("|");
    if (existingSignatures.has(signature)) {
      return;
    }
    existingSignatures.add(signature);
    const widget = {
      id: uid(),
      kind: "chart",
      name: template.name.slice(0, 60),
      source: normalizeBiSource(template.source),
      groupBy: normalizeBiGroupBy(template.groupBy),
      metric: normalizeBiMetric(template.metric),
      chartType: normalizeBiChartType(template.chartType),
      sortMode: normalizeBiSortMode(template.sortMode || "value_desc"),
      topN: sanitizeBiTopN(template.topN),
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: template.groupBy, metric: template.metric, chartType: template.chartType }),
      chartConfig: createDefaultBiChartConfig(template.chartType),
      chartTypeConfig: {
        [normalizeBiChartType(template.chartType)]: normalizeBiChartTypeSpecificConfig({}, template.chartType)
      },
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      layout: getDefaultBiWidgetLayout(project.biWidgets.length),
      visualOverride: null
    };
    project.biWidgets.push(widget);
    selectedBiWidgetId = widget.id;
    added += 1;
  });

  if (added > 0) {
    autoArrangeBiWidgets(project, { expandCanvas: true });
  }
  return { added };
}

function getBiCanvasSurface() {
  return els.biWidgetsGrid?.querySelector("[data-bi-canvas-surface]") || null;
}

function getBiWidgetById(project, widgetId) {
  if (!project || !Array.isArray(project.biWidgets) || !widgetId) {
    return null;
  }
  return project.biWidgets.find((item) => item.id === widgetId) || null;
}

function findBiWidgetElement(widgetId) {
  if (!els.biWidgetsGrid || !widgetId) {
    return null;
  }
  const items = els.biWidgetsGrid.querySelectorAll("[data-bi-widget-id]");
  for (let index = 0; index < items.length; index += 1) {
    const node = items[index];
    if (node instanceof HTMLElement && trimOrFallback(node.dataset.biWidgetId, "") === widgetId) {
      return node;
    }
  }
  return null;
}

function startBiWidgetInteraction(project, widget, widgetElement, mode, pointerEvent) {
  if (!project || !widget || !widgetElement || !(pointerEvent instanceof PointerEvent)) {
    return;
  }
  if (widget.locked) {
    return;
  }
  const surface = getBiCanvasSurface();
  if (!(surface instanceof HTMLElement)) {
    return;
  }

  const layout = normalizeBiWidgetLayout(widget.layout, 0, widget);
  const zoomScale = Math.max(
    0.25,
    Math.min(2, sanitizeBiCanvasZoom(surface.dataset.biCanvasZoom, BI_CANVAS_ZOOM_DEFAULT) / 100)
  );
  const kindHint = widgetElement.classList.contains("is-text-widget")
    ? "text"
    : (widgetElement.classList.contains("is-image-widget") ? "image" : "chart");
  widget.kind = normalizeBiWidgetKind(widget.kind || kindHint);
  const minSize = getBiWidgetMinSize(widget);
  widget.layout = layout;
  biCanvasInteraction = {
    projectId: project.id,
    widgetId: widget.id,
    mode,
    pointerId: pointerEvent.pointerId,
    startX: pointerEvent.clientX,
    startY: pointerEvent.clientY,
    layoutStart: { ...layout },
    layoutLive: { ...layout },
    minSize,
    zoomScale,
    element: widgetElement,
    surface
  };
  widgetElement.classList.add("moving");
  pointerEvent.preventDefault();
}

function clampBiWidgetLayout(layout, surface, minSize) {
  const configuredWidth = parseInt(surface.dataset.biCanvasWidth || "", 10);
  const configuredHeight = parseInt(surface.dataset.biCanvasHeight || "", 10);
  const widthFallback = Number.isFinite(configuredWidth) ? configuredWidth : BI_CANVAS_SURFACE_MIN_WIDTH;
  const heightFallback = Number.isFinite(configuredHeight) ? configuredHeight : BI_CANVAS_SURFACE_HEIGHT;
  const width = Math.round(surface.clientWidth || widthFallback);
  const height = Math.round(surface.clientHeight || heightFallback);
  return clampBiWidgetLayoutToCanvas(layout, width, height, minSize);
}

function updateBiSurfaceDimensionsInDom(surface, width, height) {
  if (!(surface instanceof HTMLElement)) {
    return;
  }
  const safeWidth = sanitizeBiCanvasDimension(width, BI_CANVAS_SURFACE_MIN_WIDTH, BI_CANVAS_SURFACE_MIN_EDIT_WIDTH, BI_CANVAS_SURFACE_MAX_EDIT_WIDTH);
  const safeHeight = sanitizeBiCanvasDimension(height, BI_CANVAS_SURFACE_HEIGHT, BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT, BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT);
  surface.dataset.biCanvasWidth = String(safeWidth);
  surface.dataset.biCanvasHeight = String(safeHeight);
  surface.style.width = `${safeWidth}px`;
  surface.style.height = `${safeHeight}px`;
  surface.style.minWidth = `${safeWidth}px`;
  surface.style.minHeight = `${safeHeight}px`;
  const zoomPercent = sanitizeBiCanvasZoom(surface.dataset.biCanvasZoom, BI_CANVAS_ZOOM_DEFAULT);
  const zoomScale = Math.max(0.25, Math.min(2, zoomPercent / 100));
  const zoomWrap = surface.parentElement;
  if (zoomWrap instanceof HTMLElement && zoomWrap.classList.contains("bi-canvas-zoom-wrap")) {
    zoomWrap.style.width = `${Math.max(1, Math.round(safeWidth * zoomScale))}px`;
    zoomWrap.style.height = `${Math.max(1, Math.round(safeHeight * zoomScale))}px`;
  }
}

function ensureBiCanvasHasDragSpace(project, interaction, layoutCandidate) {
  if (!project || !interaction || !(interaction.surface instanceof HTMLElement) || !layoutCandidate) {
    return;
  }
  const rawWidth = Number(interaction.surface.dataset.biCanvasWidth || 0);
  const rawHeight = Number(interaction.surface.dataset.biCanvasHeight || 0);
  const currentWidth = sanitizeBiCanvasDimension(rawWidth, BI_CANVAS_SURFACE_MIN_WIDTH, BI_CANVAS_SURFACE_MIN_EDIT_WIDTH, BI_CANVAS_SURFACE_MAX_EDIT_WIDTH);
  const currentHeight = sanitizeBiCanvasDimension(rawHeight, BI_CANVAS_SURFACE_HEIGHT, BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT, BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT);
  const targetRight = Math.max(0, Math.round(layoutCandidate.x + layoutCandidate.w + BI_CANVAS_AUTO_EXPAND_MARGIN));
  const targetBottom = Math.max(0, Math.round(layoutCandidate.y + layoutCandidate.h + BI_CANVAS_AUTO_EXPAND_MARGIN));
  const requiredWidth = Math.max(currentWidth, Math.ceil(targetRight / BI_CANVAS_AUTO_EXPAND_STEP) * BI_CANVAS_AUTO_EXPAND_STEP);
  const requiredHeight = Math.max(currentHeight, Math.ceil(targetBottom / BI_CANVAS_AUTO_EXPAND_STEP) * BI_CANVAS_AUTO_EXPAND_STEP);
  const nextWidth = sanitizeBiCanvasDimension(requiredWidth, currentWidth, BI_CANVAS_SURFACE_MIN_EDIT_WIDTH, BI_CANVAS_SURFACE_MAX_EDIT_WIDTH);
  const nextHeight = sanitizeBiCanvasDimension(requiredHeight, currentHeight, BI_CANVAS_SURFACE_MIN_EDIT_HEIGHT, BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT);
  if (nextWidth === currentWidth && nextHeight === currentHeight) {
    return;
  }
  project.biConfig.canvasWidth = nextWidth;
  project.biConfig.canvasHeight = nextHeight;
  updateBiSurfaceDimensionsInDom(interaction.surface, nextWidth, nextHeight);
  if (els.biCanvasWidthInput instanceof HTMLInputElement) {
    els.biCanvasWidthInput.value = String(nextWidth);
  }
  if (els.biCanvasHeightInput instanceof HTMLInputElement) {
    els.biCanvasHeightInput.value = String(nextHeight);
  }
}

function getBiGridSettings(project) {
  return {
    enabled: normalizeBiToggle(project?.biConfig?.snapToGrid, true),
    size: sanitizeBiInteger(project?.biConfig?.gridSnapSize, 12, 4, 120)
  };
}

function snapBiCoordinate(value, size) {
  if (!Number.isFinite(value) || !Number.isFinite(size) || size <= 1) {
    return Math.round(Number.isFinite(value) ? value : 0);
  }
  return Math.round(value / size) * size;
}

function snapBiWidgetLayout(layout, mode, gridSize, minSize) {
  if (!layout || !Number.isFinite(gridSize) || gridSize <= 1) {
    return layout;
  }
  const minWidth = Math.max(80, sanitizeBiInteger(minSize?.width, 120, 80, BI_CANVAS_SURFACE_MAX_EDIT_WIDTH));
  const minHeight = Math.max(52, sanitizeBiInteger(minSize?.height, 64, 52, BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT));
  const snapped = {
    ...layout,
    x: snapBiCoordinate(layout.x, gridSize),
    y: snapBiCoordinate(layout.y, gridSize)
  };
  if (mode === "resize") {
    snapped.w = Math.max(minWidth, snapBiCoordinate(layout.w, gridSize));
    snapped.h = Math.max(minHeight, snapBiCoordinate(layout.h, gridSize));
  }
  return snapped;
}

function updateBiWidgetInteraction(pointerEvent) {
  if (!biCanvasInteraction || !(pointerEvent instanceof PointerEvent)) {
    return;
  }
  if (pointerEvent.pointerId !== biCanvasInteraction.pointerId) {
    return;
  }

  const interactionScale = Math.max(0.25, Math.min(2, Number(biCanvasInteraction.zoomScale) || 1));
  const deltaX = (pointerEvent.clientX - biCanvasInteraction.startX) / interactionScale;
  const deltaY = (pointerEvent.clientY - biCanvasInteraction.startY) / interactionScale;
  const next = { ...biCanvasInteraction.layoutStart };
  if (biCanvasInteraction.mode === "drag") {
    next.x += deltaX;
    next.y += deltaY;
  } else {
    next.w += deltaX;
    next.h += deltaY;
  }

  const project = state.projects.find((item) => item.id === biCanvasInteraction.projectId);
  const gridSettings = getBiGridSettings(project);
  const snapped = gridSettings.enabled
    ? snapBiWidgetLayout(next, biCanvasInteraction.mode, gridSettings.size, biCanvasInteraction.minSize)
    : next;
  ensureBiCanvasHasDragSpace(project, biCanvasInteraction, snapped);
  const clamped = clampBiWidgetLayout(snapped, biCanvasInteraction.surface, biCanvasInteraction.minSize);
  biCanvasInteraction.layoutLive = clamped;
  biCanvasInteraction.element.style.left = `${clamped.x}px`;
  biCanvasInteraction.element.style.top = `${clamped.y}px`;
  biCanvasInteraction.element.style.width = `${clamped.w}px`;
  biCanvasInteraction.element.style.height = `${clamped.h}px`;
}

function endBiWidgetInteraction(pointerEvent) {
  if (!biCanvasInteraction || !(pointerEvent instanceof PointerEvent)) {
    return;
  }
  if (pointerEvent.pointerId !== biCanvasInteraction.pointerId) {
    return;
  }

  const interaction = biCanvasInteraction;
  biCanvasInteraction = null;
  interaction.element.classList.remove("moving");

  const project = state.projects.find((item) => item.id === interaction.projectId);
  const widget = getBiWidgetById(project, interaction.widgetId);
  if (!project || !widget) {
    return;
  }

  widget.layout = { ...interaction.layoutLive };
  saveState();
  renderBiPanel(project);
}

function renderBiRowsTable(project, rows) {
  if (!els.biRowsHeader || !els.biRowsBody) {
    return;
  }

  const customFields = getBiCustomFieldDefs(project);
  const columns = [
    "#",
    "Fuente",
    "Item",
    "Proyecto",
    "Disciplina",
    "Sistema",
    "Paquete",
    "Hito",
    "Creador",
    "Fase",
    "Sector",
    "Nivel",
    "Tipo",
    ...customFields.map((field) => field.label),
    "Fecha inicio",
    "Fecha fin",
    "UP Base",
    "Programado",
    "Avance Real",
    "Prog. Proyecto",
    "Real Proyecto",
    "Estado",
    "Creado"
  ];
  els.biRowsHeader.innerHTML = `<tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;

  const totalRows = rows.length;
  const visibleRows = rows.slice(0, BI_DETAIL_ROW_LIMIT);

  const sourceCounts = new Map();
  rows.forEach((row) => {
    const sourceLabel = row.sourceLabel || getBiSourceLabel(row.source);
    sourceCounts.set(sourceLabel, (sourceCounts.get(sourceLabel) || 0) + 1);
  });
  const sourceSummary = Array.from(sourceCounts.entries())
    .map(([label, count]) => `${label}: ${count}`)
    .join(" | ");

  if (els.biRowsMetaText) {
    const capLabel = totalRows > BI_DETAIL_ROW_LIMIT
      ? ` | mostrando ${visibleRows.length} de ${totalRows}`
      : ` | ${totalRows} filas`;
    const sourceLabel = sourceSummary ? ` | ${sourceSummary}` : "";
    els.biRowsMetaText.textContent = `Detalle BI${capLabel}${sourceLabel}`;
  }

  if (totalRows === 0) {
    els.biRowsBody.innerHTML = `<tr><td colspan="${columns.length}" class="muted">Sin datos para los filtros actuales.</td></tr>`;
    return;
  }

  els.biRowsBody.innerHTML = visibleRows
    .map((row, index) => {
      const invalidClass = row.invalidDates ? " overdue" : "";
      const programmedText = row.programmedPercent === null ? "" : formatPercent(row.programmedPercent, 2);
      const realText = row.realProgress === null ? "" : formatPercent(row.realProgress, 2);
      const weightedProgrammedText = row.weightedProgrammed === null ? "" : formatPercent(row.weightedProgrammed, 2);
      const weightedRealText = row.weightedReal === null ? "" : formatPercent(row.weightedReal, 2);
      return `<tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.sourceLabel)}</td>
        <td>${escapeHtml(row.itemLabel)}</td>
        <td>${escapeHtml(row.projectLabel)}</td>
        <td>${escapeHtml(row.disciplineLabel)}</td>
        <td>${escapeHtml(row.systemLabel)}</td>
        <td>${escapeHtml(row.packageLabel)}</td>
        <td>${escapeHtml(row.milestoneLabel)}</td>
        <td>${escapeHtml(row.creatorLabel)}</td>
        <td>${escapeHtml(row.phaseLabel)}</td>
        <td>${escapeHtml(row.sectorLabel)}</td>
        <td>${escapeHtml(row.levelLabel)}</td>
        <td>${escapeHtml(row.typeLabel)}</td>
        ${customFields.map((field) => `<td>${escapeHtml(row.customValues?.[field.id] || "")}</td>`).join("")}
        <td>${escapeHtml(formatDateFromInput(row.startDate) || "")}</td>
        <td>${escapeHtml(formatDateFromInput(row.endDate) || "")}</td>
        <td>${escapeHtml(formatNumberForInput(row.baseUnits) || "0")}</td>
        <td>${programmedText ? `<span class="project-progress-chip">${escapeHtml(programmedText)}</span>` : ""}</td>
        <td>${realText ? `<span class="project-progress-chip">${escapeHtml(realText)}</span>` : ""}</td>
        <td>${weightedProgrammedText ? `<span class="project-progress-chip">${escapeHtml(weightedProgrammedText)}</span>` : ""}</td>
        <td>${weightedRealText ? `<span class="project-progress-chip">${escapeHtml(weightedRealText)}</span>` : ""}</td>
        <td><span class="progress-chip${invalidClass}">${escapeHtml(row.statusLabel)}</span></td>
        <td>${escapeHtml(formatDate(row.createdAt) || "")}</td>
      </tr>`;
    })
    .join("");
}

function escapeCsvCell(value) {
  const raw = value === null || value === undefined ? "" : String(value);
  const compact = raw.replace(/\r?\n/g, " ");
  return `"${compact.replace(/"/g, "\"\"")}"`;
}

function buildCsvText(headers, rows) {
  const headLine = headers.map((item) => escapeCsvCell(item)).join(",");
  const bodyLines = rows.map((row) => row.map((cell) => escapeCsvCell(cell)).join(","));
  return [headLine, ...bodyLines].join("\r\n");
}

function downloadCsv(filename, headers, rows) {
  const safeName = trimOrFallback(filename, "reporte.csv");
  const csvText = buildCsvText(headers, rows);
  const blob = new Blob([`\uFEFF${csvText}`], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = safeName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function downloadDataUrl(filename, dataUrl) {
  if (!dataUrl) {
    return;
  }
  const safeName = trimOrFallback(filename, "archivo");
  const anchor = document.createElement("a");
  anchor.href = dataUrl;
  anchor.download = safeName;
  anchor.click();
}

function buildBiExportFileName(prefix, project, extension = "csv") {
  const projectToken = toCodeToken(project?.name || project?.title || "proyecto").toLowerCase() || "proyecto";
  const stamp = new Date().toISOString().slice(0, 19).replace("T", "_").replace(/:/g, "-");
  const ext = trimOrFallback(extension, "csv").replace(/^\.+/, "");
  return `${prefix}_${projectToken}_${stamp}.${ext}`;
}

function exportBiRowsCsv(project) {
  if (!project) {
    return;
  }

  ensureBiState(project);
  const rows = queryBiProjectRows(project, project.biConfig);
  const customFields = getBiCustomFieldDefs(project);
  if (rows.length === 0) {
    setStatus("No hay filas BI para exportar.");
    return;
  }

  const headers = [
    "Fuente",
    "Item",
    "Proyecto",
    "Disciplina",
    "Sistema",
    "Paquete",
    "Hito",
    "Creador",
    "Fase",
    "Sector",
    "Nivel",
    "Tipo",
    ...customFields.map((field) => field.label),
    "Fecha inicio",
    "Fecha fin",
    "UP Base",
    "Programado %",
    "Avance real %",
    "Programado proyecto %",
    "Real proyecto %",
    "Estado",
    "Creado"
  ];
  const csvRows = rows.map((row) => [
    row.sourceLabel,
    row.itemLabel,
    row.projectLabel,
    row.disciplineLabel,
    row.systemLabel,
    row.packageLabel,
    row.milestoneLabel,
    row.creatorLabel,
    row.phaseLabel,
    row.sectorLabel,
    row.levelLabel,
    row.typeLabel,
    ...customFields.map((field) => row.customValues?.[field.id] || ""),
    formatDateFromInput(row.startDate) || "",
    formatDateFromInput(row.endDate) || "",
    formatNumberForInput(row.baseUnits) || "0",
    row.programmedPercent === null ? "" : formatPercent(row.programmedPercent, 2),
    row.realProgress === null ? "" : formatPercent(row.realProgress, 2),
    row.weightedProgrammed === null ? "" : formatPercent(row.weightedProgrammed, 2),
    row.weightedReal === null ? "" : formatPercent(row.weightedReal, 2),
    row.statusLabel,
    formatDate(row.createdAt) || ""
  ]);

  downloadCsv(buildBiExportFileName("dechini_bi_detalle", project), headers, csvRows);
  setStatus(`Detalle BI exportado (${rows.length} filas).`);
}

function exportBiWidgetCsv(project, widgetId) {
  if (!project || !widgetId) {
    return;
  }

  ensureBiState(project);
  const widget = project.biWidgets.find((item) => item.id === widgetId);
  if (!widget) {
    setStatus("No se encontro el widget solicitado.");
    return;
  }
  const kind = normalizeBiWidgetKind(widget.kind || "chart");
  if (kind === "text") {
    const headers = ["Tipo", "Nombre", "Texto"];
    const rows = [["Texto", widget.name || "Caja de texto", trimOrFallback(widget.textContent, "")]];
    downloadCsv(buildBiExportFileName("dechini_bi_widget_texto", project), headers, rows);
    setStatus(`Widget de texto exportado: ${widget.name}.`);
    return;
  }
  if (kind === "image") {
    const headers = ["Tipo", "Nombre", "Imagen URL/Data"];
    const rows = [["Imagen", widget.name || "Imagen", trimOrFallback(widget.imageSrc, "")]];
    downloadCsv(buildBiExportFileName("dechini_bi_widget_imagen", project), headers, rows);
    setStatus(`Widget de imagen exportado: ${widget.name}.`);
    return;
  }

  let snapshot = biWidgetSnapshotsByProject[project.id]?.[widget.id];
  if (!snapshot) {
    const rows = queryBiProjectRows(project, project.biConfig);
    snapshot = buildBiWidgetSnapshot(rows, widget);
  }

  if (!snapshot || snapshot.rows.length === 0) {
    setStatus("Este widget no tiene datos para exportar.");
    return;
  }

  const groupLabel = getBiGroupLabel(snapshot.groupBy, project);
  const metricLabel = getBiMetricLabel(snapshot.metric);
  const headers = [
    "#",
    groupLabel,
    metricLabel,
    "Filas",
    "UP Base",
    "Real promedio %",
    "Programado promedio %",
    "Real proyecto %",
    "Programado proyecto %",
    "Fechas invertidas"
  ];
  const csvRows = snapshot.rows.map((row, index) => [
    index + 1,
    row.label,
    row.valueText,
    row.count,
    formatNumberForInput(row.baseUnits) || "0",
    row.realAvg === null ? "" : formatPercent(row.realAvg, 2),
    row.programmedAvg === null ? "" : formatPercent(row.programmedAvg, 2),
    row.weightedReal === null ? "" : formatPercent(row.weightedReal, 2),
    row.weightedProgrammed === null ? "" : formatPercent(row.weightedProgrammed, 2),
    row.invalidDates
  ]);

  downloadCsv(buildBiExportFileName("dechini_bi_widget", project), headers, csvRows);
  setStatus(`Widget exportado: ${widget.name}.`);
}

function drawBiWidgetFallbackPreview(ctx, snapshot, width, height, drawBackground = true) {
  if (!ctx) {
    return;
  }
  if (drawBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  ctx.fillStyle = "#1f2f44";
  ctx.font = "600 12px Segoe UI";
  ctx.fillText("Vista previa", 12, 18);
  ctx.font = "11px Segoe UI";
  ctx.fillStyle = "#516377";
  const rows = Array.isArray(snapshot?.rows) ? snapshot.rows.slice(0, 8) : [];
  if (rows.length === 0) {
    ctx.fillText("Sin datos", 12, 36);
    return;
  }
  rows.forEach((row, index) => {
    const y = 38 + (index * 18);
    ctx.fillStyle = "#3d4f63";
    ctx.fillText(`${index + 1}. ${trimOrFallback(row.label, "").slice(0, 26)}`, 12, y);
    ctx.fillStyle = "#1f5f9f";
    ctx.fillText(trimOrFallback(row.valueText, ""), Math.max(140, width - 90), y);
  });
}

function drawBiTableWidgetPreview(ctx, snapshot, width, height, settings, drawBackground = true, typeConfig = null, chartType = "table") {
  if (!ctx) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const safeType = normalizeBiChartType(chartType || "table");
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, safeType === "pivot" ? "pivot" : "table");
  const compact = normalizeBiToggle(specific.compact, false);
  const rowLimit = sanitizeBiInteger(specific.rowLimit, BI_WIDGET_TABLE_LIMIT, 1, 200);
  const showIndex = safeType === "pivot"
    ? true
    : normalizeBiToggle(specific.showIndex, true);
  if (drawBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const headerHeight = compact ? 20 : 24;
  const rowHeight = compact ? 16 : 20;
  const left = 8;
  const top = 8;
  const tableWidth = Math.max(80, width - (left * 2));
  const tableHeight = Math.max(42, height - (top * 2));
  const indexWidth = showIndex ? 36 : 0;
  const valueWidth = Math.max(84, Math.floor(tableWidth * 0.36));
  const labelWidth = Math.max(70, tableWidth - indexWidth - valueWidth);
  const maxRowsByHeight = Math.max(0, Math.floor((tableHeight - headerHeight) / rowHeight));
  const maxRows = Math.max(0, Math.min(rowLimit, maxRowsByHeight));
  const metricLabel = getBiMetricLabel(snapshot?.metric || "baseunits");
  let rows = Array.isArray(snapshot?.rows) ? snapshot.rows.slice(0, maxRows) : [];
  let groupLabel = getBiGroupLabel(snapshot?.groupBy || "disciplina");
  let valueHeader = metricLabel;
  if (safeType === "pivot") {
    const rowLabel = getBiGroupLabel(snapshot?.pivot?.rowDimension || snapshot?.groupBy || "disciplina");
    const colLabels = Array.isArray(snapshot?.pivot?.columns) ? snapshot.pivot.columns : [];
    const showTotals = normalizeBiToggle(specific.showTotals, true);
    groupLabel = rowLabel;
    valueHeader = showTotals ? "Total" : trimOrFallback(colLabels[0], metricLabel);
    const pivotRows = Array.isArray(snapshot?.pivot?.rows) ? snapshot.pivot.rows : [];
    rows = pivotRows.slice(0, maxRows).map((row) => {
      const firstCellValue = Number(row?.cells?.[0]?.value) || 0;
      const value = showTotals ? (Number(row?.total) || 0) : firstCellValue;
      return {
        label: trimOrFallback(row?.label, ""),
        valueText: formatBiMetricValue(snapshot?.metric || "baseunits", value)
      };
    });
  }

  ctx.strokeStyle = "#d3dce8";
  ctx.lineWidth = 1;
  ctx.strokeRect(left + 0.5, top + 0.5, tableWidth - 1, tableHeight - 1);
  ctx.fillStyle = "#eef3fa";
  ctx.fillRect(left + 1, top + 1, tableWidth - 2, headerHeight - 1);
  ctx.strokeStyle = "#d8e1ec";
  ctx.beginPath();
  if (showIndex) {
    ctx.moveTo(left + indexWidth + 0.5, top);
    ctx.lineTo(left + indexWidth + 0.5, top + tableHeight);
  }
  ctx.moveTo(left + indexWidth + labelWidth + 0.5, top);
  ctx.lineTo(left + indexWidth + labelWidth + 0.5, top + tableHeight);
  ctx.moveTo(left, top + headerHeight + 0.5);
  ctx.lineTo(left + tableWidth, top + headerHeight + 0.5);
  ctx.stroke();

  applyBiCanvasFont(ctx, visual, "label", { bold: true, fallbackSize: Math.max(10, visual.fontSizeLabels) });
  ctx.fillStyle = "#21374f";
  ctx.textAlign = "left";
  const headerY = compact ? top + 14 : top + 16;
  if (showIndex) {
    ctx.fillText("#", left + 8, headerY);
  }
  ctx.fillText(truncateBiLabel(groupLabel, visual, 16), left + indexWidth + 6, headerY);
  ctx.fillText(truncateBiLabel(valueHeader, visual, 14), left + indexWidth + labelWidth + 6, headerY);

  if (rows.length === 0) {
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(10, visual.fontSizeLabels) });
    ctx.fillStyle = "#5b6f84";
    ctx.fillText("Sin datos", left + 8, top + headerHeight + 16);
    return;
  }

  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: Math.max(10, visual.fontSizeLabels) });
  rows.forEach((row, index) => {
    const y = top + headerHeight + (index * rowHeight);
    ctx.strokeStyle = "#edf2f8";
    ctx.beginPath();
    ctx.moveTo(left, y + rowHeight + 0.5);
    ctx.lineTo(left + tableWidth, y + rowHeight + 0.5);
    ctx.stroke();
    ctx.fillStyle = "#2b3f57";
    const rowY = compact ? y + 12 : y + 14;
    if (showIndex) {
      ctx.fillText(String(index + 1), left + 8, rowY);
    }
    ctx.fillText(truncateBiLabel(row.label, visual, 24), left + indexWidth + 6, rowY);
    ctx.fillStyle = "#1d5b95";
    ctx.fillText(trimOrFallback(row.valueText, "0"), left + indexWidth + labelWidth + 6, rowY);
  });
}

function drawBiScorecardWidgetPreview(ctx, snapshot, width, height, settings, drawBackground = true, typeConfig = null) {
  if (!ctx) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "scorecard");
  if (drawBackground) {
    const gradient = ctx.createLinearGradient(0, 0, width, height);
    gradient.addColorStop(0, "#f6fbff");
    gradient.addColorStop(1, "#ffffff");
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  ctx.strokeStyle = "#d7e1ec";
  ctx.strokeRect(0.5, 0.5, width - 1, height - 1);
  const total = Array.isArray(snapshot?.rows)
    ? snapshot.rows.reduce((sum, row) => sum + (Number.isFinite(row?.value) ? row.value : 0), 0)
    : 0;
  const valueText = `${trimOrFallback(specific.prefix, "")}${formatBiMetricValue(snapshot?.metric || "count", total)}${trimOrFallback(specific.suffix, "")}`;
  const metricLabel = getBiMetricLabel(snapshot?.metric || "count");
  const subTokens = [];
  if (normalizeBiToggle(specific.showSource, true)) {
    subTokens.push(`Fuente: ${getBiSourceLabel(snapshot?.source || "all")}`);
  }
  if (normalizeBiToggle(specific.showGroups, true)) {
    subTokens.push(`grupos: ${snapshot?.groupCount || 0}`);
  }
  const subLabel = subTokens.join(" | ");
  ctx.textAlign = "center";
  ctx.fillStyle = "#20466e";
  applyBiCanvasFont(ctx, visual, "title", { bold: true, fallbackSize: 30 });
  ctx.fillText(valueText, Math.floor(width / 2), Math.floor(height * 0.48));
  ctx.fillStyle = "#2f4861";
  applyBiCanvasFont(ctx, visual, "label", { bold: true, fallbackSize: 13 });
  ctx.fillText(metricLabel, Math.floor(width / 2), Math.floor(height * 0.64));
  if (subLabel) {
    ctx.fillStyle = "#5f7389";
    applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 10 });
    ctx.fillText(truncateBiLabel(subLabel, visual, 44), Math.floor(width / 2), Math.floor(height * 0.75));
  }
}

function drawBiSankeyWidgetPreview(ctx, snapshot, width, height, settings, drawBackground = true, typeConfig = null) {
  if (!ctx) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "sankey");
  if (drawBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const maxLinks = sanitizeBiInteger(specific.maxLinks, 12, 1, 200);
  const minPercent = sanitizeBiDecimal(specific.minLinkPercent, 2, 0, 50);
  const showValues = normalizeBiToggle(specific.showValues, true);
  const links = Array.isArray(snapshot?.sankey?.links) ? snapshot.sankey.links.slice(0, maxLinks) : [];
  if (links.length === 0) {
    drawBiWidgetFallbackPreview(ctx, snapshot, width, height, false);
    return;
  }
  const left = 10;
  const top = 8;
  const rowGap = 20;
  const maxRowsByHeight = Math.max(1, Math.floor((height - 16) / rowGap));
  const rows = links.slice(0, maxRowsByHeight);
  const maxValue = Math.max(...rows.map((item) => Number(item.value) || 0), 1);
  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 10 });
  rows.forEach((link, index) => {
    const y = top + (index * rowGap);
    const ratio = Math.max(minPercent / 100, Math.min(1, (Number(link.value) || 0) / maxValue));
    const barX = left + 140;
    const barY = y + 2;
    const barWidth = Math.max(40, Math.floor((width - barX - 12) * ratio));
    ctx.fillStyle = "#264f80";
    ctx.fillText(`${trimOrFallback(link.from, "").slice(0, 12)} -> ${trimOrFallback(link.to, "").slice(0, 12)}`, left, y + 10);
    ctx.fillStyle = "#dbe8f7";
    ctx.fillRect(barX, barY, Math.max(24, width - barX - 12), 10);
    ctx.fillStyle = "#3c79bf";
    ctx.fillRect(barX, barY, barWidth, 10);
    if (showValues) {
      ctx.fillStyle = "#1f2f44";
      ctx.fillText(formatBiMetricValue(snapshot?.metric || "baseunits", Number(link.value) || 0), barX + 4, y + 10);
    }
  });
}

function drawBiTimelineWidgetPreview(ctx, snapshot, width, height, settings, drawBackground = true, typeConfig = null) {
  if (!ctx) {
    return;
  }
  const visual = normalizeBiVisualSettings(settings);
  const specific = normalizeBiChartTypeSpecificConfig(typeConfig, "timeline");
  if (drawBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const rowLimit = sanitizeBiInteger(specific.rowLimit, 20, 1, 200);
  const showDates = normalizeBiToggle(specific.showDates, true);
  const minBarPercent = sanitizeBiDecimal(specific.minBarPercent, 1, 0.2, 50);
  const items = Array.isArray(snapshot?.timeline?.items) ? snapshot.timeline.items.slice(0, rowLimit) : [];
  if (items.length === 0) {
    drawBiWidgetFallbackPreview(ctx, snapshot, width, height, false);
    return;
  }
  const minDay = Number.isFinite(snapshot?.timeline?.minDay) ? snapshot.timeline.minDay : 0;
  const maxDay = Number.isFinite(snapshot?.timeline?.maxDay) ? snapshot.timeline.maxDay : (minDay + 1);
  const span = Math.max(1, maxDay - minDay + 1);
  const left = 8;
  const labelWidth = Math.max(70, Math.floor(width * 0.3));
  const trackLeft = left + labelWidth + 8;
  const trackWidth = Math.max(60, width - trackLeft - 10);
  const rowHeight = 22;
  const maxRowsByHeight = Math.max(1, Math.floor((height - 10) / rowHeight));
  const rows = items.slice(0, maxRowsByHeight);
  applyBiCanvasFont(ctx, visual, "label", { fallbackSize: 10 });
  rows.forEach((item, index) => {
    const y = 10 + (index * rowHeight);
    const leftPct = Math.max(0, Math.min(100, ((item.startDay - minDay) / span) * 100));
    const widthPct = Math.max(minBarPercent, Math.min(100, (((item.endDay - item.startDay + 1) / span) * 100)));
    const barX = trackLeft + Math.round((leftPct / 100) * trackWidth);
    const barWidth = Math.max(2, Math.round((widthPct / 100) * trackWidth));
    ctx.fillStyle = "#264f80";
    ctx.fillText(trimOrFallback(item.label, "").slice(0, 24), left, y + 10);
    ctx.fillStyle = "#e4edf8";
    ctx.fillRect(trackLeft, y + 2, trackWidth, 10);
    ctx.fillStyle = "#5d8fcf";
    ctx.fillRect(barX, y + 2, barWidth, 10);
    if (showDates) {
      ctx.fillStyle = "#62778e";
      ctx.fillText(`${formatDateFromInput(item.startDate)} - ${formatDateFromInput(item.endDate)}`, left, y + 20);
    }
  });
}

function drawBiWrappedText(ctx, text, x, y, maxWidth, maxHeight, lineHeight) {
  if (!ctx || !Number.isFinite(maxWidth) || maxWidth <= 0 || !Number.isFinite(maxHeight) || maxHeight <= 0) {
    return;
  }
  const safeLineHeight = Number.isFinite(lineHeight) && lineHeight > 0 ? lineHeight : 16;
  const paragraphs = String(text || "").replace(/\r\n/g, "\n").split("\n");
  let currentY = y;
  const bottom = y + maxHeight;
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    const paragraph = paragraphs[paragraphIndex];
    const words = paragraph.split(/\s+/).filter(Boolean);
    if (words.length === 0) {
      if (currentY + safeLineHeight > bottom) {
        break;
      }
      currentY += safeLineHeight;
      continue;
    }
    let currentLine = words[0];
    for (let wordIndex = 1; wordIndex < words.length; wordIndex += 1) {
      const candidate = `${currentLine} ${words[wordIndex]}`;
      if (ctx.measureText(candidate).width <= maxWidth) {
        currentLine = candidate;
        continue;
      }
      if (currentY + safeLineHeight > bottom) {
        return;
      }
      ctx.fillText(currentLine, x, currentY);
      currentY += safeLineHeight;
      currentLine = words[wordIndex];
    }
    if (currentY + safeLineHeight > bottom) {
      return;
    }
    ctx.fillText(currentLine, x, currentY);
    currentY += safeLineHeight;
  }
}

function drawBiTextWidgetPreview(ctx, widget, width, height, drawBackground = true) {
  if (!ctx) {
    return;
  }
  const textConfig = normalizeBiTextConfig(widget?.textConfig);
  if (drawBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const text = trimOrFallback(widget?.textContent, "");
  if (!text) {
    ctx.fillStyle = "#63768a";
    ctx.font = "600 13px Segoe UI";
    ctx.fillText("Sin texto", 14, 26);
    return;
  }
  ctx.fillStyle = textConfig.color;
  ctx.textAlign = textConfig.textAlign === "center" ? "center" : (textConfig.textAlign === "right" ? "right" : "left");
  const safeFamily = textConfig.fontFamily.includes(" ") ? `'${textConfig.fontFamily}'` : textConfig.fontFamily;
  ctx.font = `${textConfig.fontStyle} ${textConfig.fontWeight} ${Math.max(12, textConfig.fontSize)}px ${safeFamily}`;
  const xBase = textConfig.textAlign === "center"
    ? Math.floor(width / 2)
    : (textConfig.textAlign === "right" ? Math.max(12, width - textConfig.padding - 2) : (12 + textConfig.padding));
  const maxWidth = Math.max(20, width - ((textConfig.padding + 12) * 2));
  drawBiWrappedText(
    ctx,
    text,
    xBase,
    24 + textConfig.padding,
    maxWidth,
    Math.max(24, height - 30 - (textConfig.padding * 2)),
    Math.max(12, Math.round(textConfig.fontSize * textConfig.lineHeight))
  );
  ctx.textAlign = "left";
}

function loadBiImage(src) {
  return new Promise((resolve) => {
    const safeSrc = trimOrFallback(src, "");
    if (!safeSrc) {
      resolve(null);
      return;
    }
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => resolve(null);
    image.src = safeSrc;
  });
}

async function drawBiImageWidgetPreview(ctx, widget, width, height, drawBackground = true) {
  if (!ctx) {
    return false;
  }
  const imageConfig = normalizeBiImageConfig(widget?.imageConfig);
  if (drawBackground) {
    ctx.fillStyle = imageConfig.background;
    ctx.fillRect(0, 0, width, height);
    ctx.strokeStyle = imageConfig.frameColor;
    ctx.strokeRect(0.5, 0.5, width - 1, height - 1);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const image = await loadBiImage(widget?.imageSrc);
  if (!image) {
    ctx.fillStyle = "#63768a";
    ctx.font = "600 13px Segoe UI";
    ctx.fillText("Sin imagen", 14, 26);
    return false;
  }
  const fit = imageConfig.fit;
  const ratioContain = Math.min(width / image.width, height / image.height);
  const ratioCover = Math.max(width / image.width, height / image.height);
  const ratio = fit === "cover" ? ratioCover : (fit === "fill" ? null : ratioContain);
  const drawWidth = fit === "fill" ? width : Math.max(1, Math.round(image.width * ratio));
  const drawHeight = fit === "fill" ? height : Math.max(1, Math.round(image.height * ratio));
  const offsetX = fit === "fill" ? 0 : Math.round((width - drawWidth) / 2);
  const offsetY = fit === "fill" ? 0 : Math.round((height - drawHeight) / 2);
  ctx.save();
  ctx.globalAlpha = imageConfig.opacity;
  ctx.drawImage(image, offsetX, offsetY, drawWidth, drawHeight);
  ctx.restore();
  return true;
}

async function exportBiWidgetPng(project, widgetId) {
  if (!project || !widgetId) {
    return;
  }
  ensureBiState(project);
  const widget = project.biWidgets.find((item) => item.id === widgetId);
  if (!widget) {
    setStatus("No se encontro el widget para PNG.");
    return;
  }
  let snapshot = biWidgetSnapshotsByProject[project.id]?.[widget.id];
  const widgetKind = normalizeBiWidgetKind(widget.kind || "chart");
  if (widgetKind === "chart" && !snapshot) {
    const rows = queryBiProjectRows(project, project.biConfig);
    snapshot = buildBiWidgetSnapshot(rows, widget);
  }
  const width = Math.max(320, Math.round(widget.layout?.w || 360));
  const height = Math.max(220, Math.round(widget.layout?.h || 280));
  const canvas = document.createElement("canvas");
  const exportScale = 3;
  canvas.width = width * exportScale;
  canvas.height = height * exportScale;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    setStatus("No se pudo preparar la exportacion PNG.");
    return;
  }
  ctx.scale(exportScale, exportScale);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  const showBackground = widget.showBackground !== false;
  if (showBackground) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }
  const showBorder = widget.showBorder !== false;
  const showTitle = widget.showTitle !== false;
  if (showBorder) {
    ctx.strokeStyle = "#d1dbe7";
    ctx.strokeRect(0.5, 0.5, width - 1, height - 1);
  }
  if (showTitle) {
    ctx.fillStyle = "#1f2f44";
    ctx.font = "600 14px Segoe UI";
    ctx.fillText(trimOrFallback(widget.name, "Widget"), 12, 22);
    ctx.fillStyle = "#5c6f84";
    ctx.font = "11px Segoe UI";
    ctx.fillText(buildBiWidgetMetaLabel(widget, snapshot, project), 12, 38);
  }

  const chartType = normalizeBiChartType(widget.chartType || "bar");
  const chartTop = showTitle ? 46 : 10;
  const chartWidth = width - 24;
  const chartHeight = Math.max(40, height - chartTop - 10);
  const renderWidth = Math.max(1, Math.round(chartWidth));
  const renderHeight = Math.max(1, Math.round(chartHeight));
  const chartCanvas = document.createElement("canvas");
  chartCanvas.width = Math.max(1, Math.round(renderWidth * exportScale));
  chartCanvas.height = Math.max(1, Math.round(renderHeight * exportScale));
  chartCanvas.dataset.biRenderWidth = String(renderWidth);
  chartCanvas.dataset.biRenderHeight = String(renderHeight);
  chartCanvas.dataset.biDpr = String(exportScale);
  const chartCtx = chartCanvas.getContext("2d");
  if (!chartCtx) {
    setStatus("No se pudo preparar el contenido del widget.");
    return;
  }
  chartCtx.setTransform(exportScale, 0, 0, exportScale, 0, 0);
  chartCtx.imageSmoothingEnabled = true;
  chartCtx.imageSmoothingQuality = "high";
  const visual = getBiEffectiveVisualSettings(project, widget);
  if (widgetKind === "text") {
    drawBiTextWidgetPreview(chartCtx, widget, renderWidth, renderHeight, showBackground);
  } else if (widgetKind === "image") {
    await drawBiImageWidgetPreview(chartCtx, widget, renderWidth, renderHeight, showBackground);
  } else if (chartType === "table" || chartType === "pivot") {
    drawBiTableWidgetPreview(
      chartCtx,
      snapshot,
      renderWidth,
      renderHeight,
      visual,
      showBackground,
      getBiSpecificChartConfigForWidget(widget, chartType),
      chartType
    );
  } else if (chartType === "scorecard") {
    drawBiScorecardWidgetPreview(
      chartCtx,
      snapshot,
      renderWidth,
      renderHeight,
      visual,
      showBackground,
      getBiSpecificChartConfigForWidget(widget, chartType)
    );
  } else if (chartType === "sankey") {
    drawBiSankeyWidgetPreview(
      chartCtx,
      snapshot,
      renderWidth,
      renderHeight,
      visual,
      showBackground,
      getBiSpecificChartConfigForWidget(widget, chartType)
    );
  } else if (chartType === "timeline") {
    drawBiTimelineWidgetPreview(
      chartCtx,
      snapshot,
      renderWidth,
      renderHeight,
      visual,
      showBackground,
      getBiSpecificChartConfigForWidget(widget, chartType)
    );
  } else {
    if (snapshot && snapshot.rows.length > 0) {
      const seriesColors = getBiSnapshotSeriesColors(project, snapshot);
      drawBiWidgetChart(
        chartCanvas,
        chartType,
        snapshot.rows.map((row) => row.label),
        snapshot.rows.map((row) => row.value),
        -1,
        seriesColors,
        visual,
        widget.labelOffsets,
        widget.polarLayout,
        snapshot.rows,
        widget.chartConfig,
        widget.labelLayoutV2,
        widget.chartTypeConfig
      );
    } else {
      drawBiWidgetFallbackPreview(chartCtx, snapshot, renderWidth, renderHeight, showBackground);
    }
  }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(chartCanvas, 12, chartTop, chartWidth, chartHeight);
  downloadDataUrl(buildBiExportFileName("dechini_bi_widget", project, "png"), canvas.toDataURL("image/png"));
  setStatus(`Widget PNG exportado: ${widget.name}.`);
}

async function exportBiBoardPng(project) {
  if (!project) {
    return;
  }
  ensureBiState(project);
  const widgets = Array.isArray(project.biWidgets) ? project.biWidgets : [];
  if (widgets.length === 0) {
    setStatus("No hay widgets para exportar en la pizarra.");
    return;
  }
  const canvasSize = getBiCanvasSizeFromConfig(project.biConfig);
  const boardCanvas = document.createElement("canvas");
  const exportScale = 3;
  boardCanvas.width = canvasSize.width * exportScale;
  boardCanvas.height = canvasSize.height * exportScale;
  const ctx = boardCanvas.getContext("2d");
  if (!ctx) {
    setStatus("No se pudo generar PNG de pizarra.");
    return;
  }
  ctx.scale(exportScale, exportScale);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvasSize.width, canvasSize.height);

  const rows = queryBiProjectRows(project, project.biConfig);
  for (let index = 0; index < widgets.length; index += 1) {
    const rawWidget = widgets[index];
    const widget = normalizeBiWidget(rawWidget, index);
    const layout = normalizeBiWidgetLayout(widget.layout, index, widget);
    const x = Math.max(0, Math.round(layout.x));
    const y = Math.max(0, Math.round(layout.y));
    const w = Math.max(220, Math.round(layout.w));
    const h = Math.max(160, Math.round(layout.h));
    if (x >= canvasSize.width || y >= canvasSize.height) {
      continue;
    }
    const widgetKind = normalizeBiWidgetKind(widget.kind || "chart");
    const snapshot = widgetKind === "chart" ? buildBiWidgetSnapshot(rows, widget) : null;
    const showBackground = widget.showBackground !== false;
    const showBorder = widget.showBorder !== false;
    const showTitle = widget.showTitle !== false;
    if (showBackground) {
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(x, y, w, h);
    }
    if (showBorder) {
      ctx.strokeStyle = "#c7d4e4";
      ctx.lineWidth = 1;
      ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    }
    if (showTitle) {
      ctx.fillStyle = "#1f2f44";
      ctx.font = "600 13px Segoe UI";
      ctx.fillText(trimOrFallback(widget.name, "Widget"), x + 8, y + 18);
      ctx.fillStyle = "#5c6f84";
      ctx.font = "11px Segoe UI";
      ctx.fillText(buildBiWidgetMetaLabel(widget, snapshot, project), x + 8, y + 33);
    }

    const chartType = normalizeBiChartType(widget.chartType || "bar");
    const chartTop = showTitle ? 38 : 8;
    const chartW = Math.max(120, w - 16);
    const chartH = Math.max(90, h - chartTop - 8);
    const renderWidth = Math.max(1, Math.round(chartW));
    const renderHeight = Math.max(1, Math.round(chartH));
    const chartCanvas = document.createElement("canvas");
    chartCanvas.width = Math.max(1, Math.round(renderWidth * exportScale));
    chartCanvas.height = Math.max(1, Math.round(renderHeight * exportScale));
    chartCanvas.dataset.biRenderWidth = String(renderWidth);
    chartCanvas.dataset.biRenderHeight = String(renderHeight);
    chartCanvas.dataset.biDpr = String(exportScale);
    const chartCtx = chartCanvas.getContext("2d");
    if (!chartCtx) {
      continue;
    }
    chartCtx.setTransform(exportScale, 0, 0, exportScale, 0, 0);
    chartCtx.imageSmoothingEnabled = true;
    chartCtx.imageSmoothingQuality = "high";
    const visual = getBiEffectiveVisualSettings(project, widget);
    if (widgetKind === "text") {
      drawBiTextWidgetPreview(chartCtx, widget, renderWidth, renderHeight, showBackground);
    } else if (widgetKind === "image") {
      await drawBiImageWidgetPreview(chartCtx, widget, renderWidth, renderHeight, showBackground);
    } else if (chartType === "table" || chartType === "pivot") {
      drawBiTableWidgetPreview(
        chartCtx,
        snapshot,
        renderWidth,
        renderHeight,
        visual,
        showBackground,
        getBiSpecificChartConfigForWidget(widget, chartType),
        chartType
      );
    } else if (chartType === "scorecard") {
      drawBiScorecardWidgetPreview(
        chartCtx,
        snapshot,
        renderWidth,
        renderHeight,
        visual,
        showBackground,
        getBiSpecificChartConfigForWidget(widget, chartType)
      );
    } else if (chartType === "sankey") {
      drawBiSankeyWidgetPreview(
        chartCtx,
        snapshot,
        renderWidth,
        renderHeight,
        visual,
        showBackground,
        getBiSpecificChartConfigForWidget(widget, chartType)
      );
    } else if (chartType === "timeline") {
      drawBiTimelineWidgetPreview(
        chartCtx,
        snapshot,
        renderWidth,
        renderHeight,
        visual,
        showBackground,
        getBiSpecificChartConfigForWidget(widget, chartType)
      );
    } else if (snapshot && snapshot.rows.length > 0) {
      const seriesColors = getBiSnapshotSeriesColors(project, snapshot);
      drawBiWidgetChart(
        chartCanvas,
        chartType,
        snapshot.rows.map((row) => row.label),
        snapshot.rows.map((row) => row.value),
        -1,
        seriesColors,
        visual,
        widget.labelOffsets,
        widget.polarLayout,
        snapshot.rows,
        widget.chartConfig,
        widget.labelLayoutV2,
        widget.chartTypeConfig
      );
    } else {
      drawBiWidgetFallbackPreview(chartCtx, snapshot, renderWidth, renderHeight, showBackground);
    }
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(chartCanvas, x + 8, y + chartTop, chartW, chartH);
  }
  downloadDataUrl(buildBiExportFileName("dechini_bi_pizarra", project, "png"), boardCanvas.toDataURL("image/png"));
  setStatus("Pizarra PNG exportada.");
}

function ensurePackageControls(project) {
  if (!project) {
    return;
  }
  project.packageControls = normalizePackageControls(project.packageControls);
}

function ensureReviewMilestones(project) {
  if (!project) {
    return;
  }
  project.reviewMilestones = normalizeReviewMilestones(project.reviewMilestones);
}

function ensureReviewControls(project) {
  if (!project) {
    return;
  }
  project.reviewControls = normalizeReviewControls(project.reviewControls);
}

function insertReviewMilestone(project, name, baseUnits) {
  if (!project) {
    return;
  }
  ensureReviewMilestones(project);

  const safeName = trimOrFallback(name, "");
  const safeBaseUnits = sanitizeBaseUnits(baseUnits);
  project.reviewMilestones.push({
    id: uid(),
    name: safeName,
    baseUnits: safeBaseUnits === "" ? null : safeBaseUnits,
    createdAt: new Date().toISOString()
  });
}

function getReviewControlFieldIds(fields) {
  return {
    projectFieldId: findFieldIdByAlias(fields, ["proyecto"]),
    disciplineFieldId: findFieldIdByAlias(fields, ["disciplina"]),
    systemFieldId: findFieldIdByAlias(fields, ["sistema"])
  };
}

function collectReviewControlBaseCandidates(project, fieldIds) {
  const candidates = new Map();
  if (!project || !fieldIds?.projectFieldId || !fieldIds?.disciplineFieldId || !fieldIds?.systemFieldId) {
    return candidates;
  }

  const addFromDeliverables = () => {
    if (!Array.isArray(project.deliverables)) {
      return;
    }
    project.deliverables.forEach((deliverable) => {
      const projectRowId = trimOrFallback(deliverable?.rowRefs?.[fieldIds.projectFieldId], "");
      const disciplineRowId = trimOrFallback(deliverable?.rowRefs?.[fieldIds.disciplineFieldId], "");
      const systemRowId = trimOrFallback(deliverable?.rowRefs?.[fieldIds.systemFieldId], "");
      const key = buildReviewControlBaseKey(projectRowId, disciplineRowId, systemRowId);
      if (!key || candidates.has(key)) {
        return;
      }
      candidates.set(key, { projectRowId, disciplineRowId, systemRowId });
    });
  };

  const addFromPackageControls = () => {
    if (!Array.isArray(project.packageControls)) {
      return;
    }
    project.packageControls.forEach((item) => {
      const projectRowId = trimOrFallback(item?.projectRowId, "");
      const disciplineRowId = trimOrFallback(item?.disciplineRowId, "");
      const systemRowId = trimOrFallback(item?.systemRowId, "");
      const key = buildReviewControlBaseKey(projectRowId, disciplineRowId, systemRowId);
      if (!key || candidates.has(key)) {
        return;
      }
      candidates.set(key, { projectRowId, disciplineRowId, systemRowId });
    });
  };

  if (getProjectProgressControlMode(project) === "package") {
    addFromPackageControls();
    if (candidates.size === 0) {
      addFromDeliverables();
    }
  } else {
    addFromDeliverables();
  }

  return candidates;
}

function buildReviewControlBaseKey(projectRowId, disciplineRowId, systemRowId) {
  const projectRef = trimOrFallback(projectRowId, "");
  const disciplineRef = trimOrFallback(disciplineRowId, "");
  const systemRef = trimOrFallback(systemRowId, "");
  if (!projectRef || !disciplineRef || !systemRef) {
    return "";
  }
  return `${projectRef}|${disciplineRef}|${systemRef}`;
}

function buildReviewControlKeyFromRowIds(projectRowId, disciplineRowId, systemRowId, milestoneId) {
  const baseKey = buildReviewControlBaseKey(projectRowId, disciplineRowId, systemRowId);
  const milestoneRef = trimOrFallback(milestoneId, "");
  if (!baseKey || !milestoneRef) {
    return "";
  }
  return `${baseKey}|${milestoneRef}`;
}

function buildReviewControlDesiredEntries(baseCandidates, milestones) {
  const desired = new Map();
  if (!(baseCandidates instanceof Map) || !Array.isArray(milestones) || milestones.length === 0) {
    return desired;
  }

  baseCandidates.forEach((candidate) => {
    milestones.forEach((milestone) => {
      const key = buildReviewControlKeyFromRowIds(
        candidate.projectRowId,
        candidate.disciplineRowId,
        candidate.systemRowId,
        milestone.id
      );
      if (!key || desired.has(key)) {
        return;
      }
      desired.set(key, {
        key,
        projectRowId: candidate.projectRowId,
        disciplineRowId: candidate.disciplineRowId,
        systemRowId: candidate.systemRowId,
        milestoneId: milestone.id
      });
    });
  });

  return desired;
}

function syncReviewControls(project) {
  ensureReviewMilestones(project);
  ensureReviewControls(project);
  const fieldIds = getReviewControlFieldIds(project.fields);
  if (!fieldIds.projectFieldId || !fieldIds.disciplineFieldId || !fieldIds.systemFieldId) {
    return { added: 0, removed: 0, missingRequired: true, missingMilestones: false };
  }

  const milestones = project.reviewMilestones;
  if (milestones.length === 0) {
    return { added: 0, removed: 0, missingRequired: false, missingMilestones: true };
  }

  const baseCandidates = collectReviewControlBaseCandidates(project, fieldIds);
  const desiredEntries = buildReviewControlDesiredEntries(baseCandidates, milestones);
  const validKeys = new Set(desiredEntries.keys());

  const beforeCount = project.reviewControls.length;
  project.reviewControls = project.reviewControls.filter((item) => validKeys.has(item.key));
  const removed = Math.max(0, beforeCount - project.reviewControls.length);

  const milestoneMap = new Map(milestones.map((item) => [item.id, item]));
  project.reviewControls.forEach((item) => {
    const milestone = milestoneMap.get(item.milestoneId) || null;
    const baseUnitsValue = sanitizeBaseUnits(milestone?.baseUnits);
    item.baseUnits = baseUnitsValue === "" ? null : baseUnitsValue;
  });

  const existingKeys = new Set(project.reviewControls.map((item) => item.key));
  let added = 0;
  desiredEntries.forEach((candidate) => {
    if (existingKeys.has(candidate.key)) {
      return;
    }

    existingKeys.add(candidate.key);
    const milestone = milestoneMap.get(candidate.milestoneId) || null;
    const baseUnitsValue = sanitizeBaseUnits(milestone?.baseUnits);
    project.reviewControls.push({
      id: uid(),
      key: candidate.key,
      projectRowId: candidate.projectRowId,
      disciplineRowId: candidate.disciplineRowId,
      systemRowId: candidate.systemRowId,
      milestoneId: candidate.milestoneId,
      startDate: "",
      endDate: "",
      baseUnits: baseUnitsValue === "" ? null : baseUnitsValue,
      realProgress: null,
      realAdvances: [],
      consumedHours: [],
      createdAt: new Date().toISOString()
    });
    added += 1;
  });

  return { added, removed, missingRequired: false, missingMilestones: false };
}

function getPackageControlFieldIds(fields) {
  return {
    projectFieldId: findFieldIdByAlias(fields, ["proyecto"]),
    disciplineFieldId: findFieldIdByAlias(fields, ["disciplina"]),
    systemFieldId: findFieldIdByAlias(fields, ["sistema"]),
    packageFieldId: findFieldIdByAlias(fields, ["paquete"])
  };
}

function syncPackageControlsFromDeliverables(project) {
  ensurePackageControls(project);
  const fieldIds = getPackageControlFieldIds(project.fields);
  if (!fieldIds.projectFieldId || !fieldIds.disciplineFieldId || !fieldIds.systemFieldId || !fieldIds.packageFieldId) {
    return { added: 0, removed: 0, missingRequired: true };
  }

  const candidatesByKey = new Map();
  project.deliverables.forEach((deliverable) => {
    const candidate = buildPackageControlCandidateFromDeliverable(deliverable, fieldIds);
    if (!candidate || candidatesByKey.has(candidate.key)) {
      return;
    }
    candidatesByKey.set(candidate.key, candidate);
  });

  const validKeys = new Set(candidatesByKey.keys());
  const beforeCount = project.packageControls.length;
  project.packageControls = project.packageControls.filter((item) => validKeys.has(item.key));
  const removed = Math.max(0, beforeCount - project.packageControls.length);

  const existingKeys = new Set(project.packageControls.map((item) => item.key));
  let added = 0;
  candidatesByKey.forEach((candidate) => {
    if (existingKeys.has(candidate.key)) {
      return;
    }

    existingKeys.add(candidate.key);
    project.packageControls.push({
      id: uid(),
      key: candidate.key,
      projectRowId: candidate.projectRowId,
      disciplineRowId: candidate.disciplineRowId,
      systemRowId: candidate.systemRowId,
      packageRowId: candidate.packageRowId,
      startDate: "",
      endDate: "",
      realProgress: null,
      realAdvances: [],
      consumedHours: [],
      createdAt: new Date().toISOString()
    });
    added += 1;
  });

  return { added, removed, missingRequired: false };
}

function buildPackageControlCandidateFromDeliverable(deliverable, fieldIds) {
  if (!deliverable || !deliverable.rowRefs) {
    return null;
  }

  const projectRowId = trimOrFallback(deliverable.rowRefs[fieldIds.projectFieldId], "");
  const disciplineRowId = trimOrFallback(deliverable.rowRefs[fieldIds.disciplineFieldId], "");
  const systemRowId = trimOrFallback(deliverable.rowRefs[fieldIds.systemFieldId], "");
  const packageRowId = trimOrFallback(deliverable.rowRefs[fieldIds.packageFieldId], "");
  const key = buildPackageControlKeyFromRowIds(projectRowId, disciplineRowId, systemRowId, packageRowId);
  if (!key) {
    return null;
  }

  return { key, projectRowId, disciplineRowId, systemRowId, packageRowId };
}

function buildPackageControlKeyFromRowIds(projectRowId, disciplineRowId, systemRowId, packageRowId) {
  if (!projectRowId || !disciplineRowId || !systemRowId || !packageRowId) {
    return "";
  }
  return `${projectRowId}|${disciplineRowId}|${systemRowId}|${packageRowId}`;
}

function computePackageControlMetrics(deliverables, fieldIds) {
  const grouped = new Map();
  if (!Array.isArray(deliverables)) {
    return grouped;
  }

  deliverables.forEach((deliverable) => {
    const candidate = buildPackageControlCandidateFromDeliverable(deliverable, fieldIds);
    if (!candidate) {
      return;
    }

    let summary = grouped.get(candidate.key);
    if (!summary) {
      summary = {
        projectRowId: candidate.projectRowId,
        disciplineRowId: candidate.disciplineRowId,
        systemRowId: candidate.systemRowId,
        packageRowId: candidate.packageRowId,
        deliverablesCount: 0,
        baseUnitsTotal: 0,
        startDate: "",
        endDate: "",
        hasInvalidDates: false,
        realWeighted: 0,
        realBase: 0
      };
      grouped.set(candidate.key, summary);
    }

    const baseUnitsRaw = sanitizeBaseUnits(deliverable.baseUnits);
    const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
    const startDate = sanitizeDateInput(deliverable.startDate || "");
    const endDate = sanitizeDateInput(deliverable.endDate || "");
    if (isDateRangeInvalid(startDate, endDate)) {
      summary.hasInvalidDates = true;
    }
    if (startDate && (!summary.startDate || startDate < summary.startDate)) {
      summary.startDate = startDate;
    }
    if (endDate && (!summary.endDate || endDate > summary.endDate)) {
      summary.endDate = endDate;
    }

    const realProgressRaw = sanitizeRealProgress(deliverable.realProgress);
    const realProgress = realProgressRaw === "" ? 0 : realProgressRaw;

    summary.deliverablesCount += 1;
    summary.baseUnitsTotal += baseUnits;

    if (baseUnits > 0) {
      summary.realWeighted += (baseUnits * realProgress);
      summary.realBase += baseUnits;
    }
  });

  const result = new Map();
  grouped.forEach((summary, key) => {
    const realPercent = summary.realBase > 0
      ? Math.round((summary.realWeighted / summary.realBase) * 100) / 100
      : null;

    result.set(key, {
      projectRowId: summary.projectRowId,
      disciplineRowId: summary.disciplineRowId,
      systemRowId: summary.systemRowId,
      packageRowId: summary.packageRowId,
      deliverablesCount: summary.deliverablesCount,
      baseUnitsTotal: Math.round(summary.baseUnitsTotal * 100) / 100,
      startDate: summary.startDate,
      endDate: summary.endDate,
      hasInvalidDates: summary.hasInvalidDates,
      realPercent
    });
  });
  return result;
}

function computePackageControlGroupedTotals(metricsByKey) {
  const disciplineTotals = new Map();
  const systemTotals = new Map();
  if (!(metricsByKey instanceof Map)) {
    return { disciplineTotals, systemTotals };
  }

  metricsByKey.forEach((metrics) => {
    const baseUnitsRaw = sanitizeBaseUnits(metrics?.baseUnitsTotal);
    const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
    if (baseUnits <= 0) {
      return;
    }

    const disciplineRowId = trimOrFallback(metrics?.disciplineRowId, "");
    if (disciplineRowId) {
      const nextDiscipline = (disciplineTotals.get(disciplineRowId) || 0) + baseUnits;
      disciplineTotals.set(disciplineRowId, Math.round(nextDiscipline * 100) / 100);
    }

    const systemGroupKey = buildPackageControlSystemGroupKey(metrics?.disciplineRowId, metrics?.systemRowId);
    if (systemGroupKey) {
      const nextSystem = (systemTotals.get(systemGroupKey) || 0) + baseUnits;
      systemTotals.set(systemGroupKey, Math.round(nextSystem * 100) / 100);
    }
  });

  return { disciplineTotals, systemTotals };
}

function buildPackageControlSystemGroupKey(disciplineRowId, systemRowId) {
  const discipline = trimOrFallback(disciplineRowId, "");
  const system = trimOrFallback(systemRowId, "");
  if (!discipline || !system) {
    return "";
  }
  return `${discipline}|${system}`;
}

function getFieldById(fields, fieldId) {
  if (!Array.isArray(fields) || !fieldId) {
    return null;
  }
  return fields.find((field) => field.id === fieldId) || null;
}

function buildPackageControlCode(project, rows) {
  if (!Array.isArray(rows)) {
    return "";
  }

  const separator = sanitizeSeparator(project?.codeSeparator || "-");
  const tokens = rows
    .map((row) => trimOrFallback(row?.code, "") || toCodeToken(row?.name || ""))
    .filter((token) => !!token);
  return tokens.join(separator);
}

function renderProjectChooser(project) {
  els.projectChooserList.innerHTML = state.projects
    .map((item) => {
      const activeClass = project && item.id === project.id ? "active" : "";
      return `<div class="project-item ${activeClass}">
        <p class="project-name">${escapeHtml(item.name)}</p>
        <p class="muted">${escapeHtml(item.title)}</p>
        <button type="button" data-select-project="${escapeAttribute(item.id)}">Entrar</button>
      </div>`;
    })
    .join("");

  els.closeChooserButton.style.visibility = chooserLocked ? "hidden" : "visible";
}

function getBiCommandDefinitions(project) {
  const definitions = [
    { id: "add_chart", label: "Agregar widget grafico", keywords: "agregar widget grafico chart", run: () => els.biAddWidgetButton?.click() },
    { id: "add_text", label: "Agregar cuadro de texto", keywords: "agregar texto cuadro", run: () => els.biAddTextWidgetButton?.click() },
    { id: "add_image", label: "Agregar imagen", keywords: "agregar imagen logo", run: () => els.biAddImageWidgetButton?.click() },
    { id: "update_widget", label: "Actualizar widget seleccionado", keywords: "actualizar widget seleccionado", run: () => els.biUpdateWidgetButton?.click() },
    { id: "auto_layout", label: "Auto organizar widgets", keywords: "auto organizar layout", run: () => els.biAutoLayoutButton?.click() },
    { id: "starter_dashboard", label: "Generar dashboard base", keywords: "dashboard base plantilla", run: () => els.biGenerateDashboardButton?.click() },
    { id: "refresh_dashboard", label: "Actualizar dashboard", keywords: "actualizar refresh", run: () => els.biRefreshButton?.click() },
    { id: "export_board_png", label: "Exportar pizarra PNG", keywords: "exportar pizarra png", run: () => els.biExportBoardPngButton?.click() },
    { id: "export_detail_csv", label: "Exportar detalle CSV", keywords: "exportar csv detalle", run: () => els.biExportRowsButton?.click() },
    { id: "clear_cross_filter", label: "Limpiar filtro cruzado", keywords: "limpiar filtro cruzado", run: () => els.biClearCrossFilterButton?.click() },
    { id: "open_data", label: "Panel Data", keywords: "panel data", run: () => applyBiRailMode("data", { forceOpen: true }) },
    { id: "open_properties", label: "Panel Propiedades", keywords: "panel propiedades", run: () => applyBiRailMode("properties", { forceOpen: true }) },
    { id: "open_settings", label: "Panel Configuracion", keywords: "panel configuracion", run: () => applyBiRailMode("settings", { forceOpen: true }) },
    { id: "open_board", label: "Panel Pizarra", keywords: "panel pizarra", run: () => applyBiRailMode("board", { forceOpen: true }) },
    { id: "open_color", label: "Panel Colorimetria", keywords: "panel colorimetria", run: () => applyBiRailMode("colorimetry", { forceOpen: true }) },
    { id: "open_filter", label: "Panel Filtro", keywords: "panel filtro", run: () => applyBiRailMode("filter", { forceOpen: true }) },
    { id: "open_performance", label: "Panel Rendimiento", keywords: "panel rendimiento performance", run: () => applyBiRailMode("performance", { forceOpen: true }) },
    { id: "open_others", label: "Panel Otros", keywords: "panel otros resumen kpi", run: () => applyBiRailMode("others", { forceOpen: true }) },
    { id: "optimize_now", label: "Optimizar rendimiento ahora", keywords: "optimizar rendimiento turbo", run: () => els.biOptimizeNowButton?.click() }
  ];
  if (project) {
    definitions.push({
      id: "mode_basic",
      label: "Modo Basico (guiado)",
      keywords: "modo basico guiado simple",
      run: () => {
        ensureBiState(project);
        project.biConfig.uiMode = "basic";
        saveState();
        renderBiPanel(project);
        setStatus("Modo BI basico activado.");
      }
    });
    definitions.push({
      id: "mode_advanced",
      label: "Modo Avanzado (pro)",
      keywords: "modo avanzado experto pro",
      run: () => {
        ensureBiState(project);
        project.biConfig.uiMode = "advanced";
        saveState();
        renderBiPanel(project);
        setStatus("Modo BI avanzado activado.");
      }
    });
  }
  return definitions;
}

function renderBiCommandPalette(project) {
  if (!(els.biCommandPaletteList instanceof HTMLElement)) {
    return;
  }
  const query = normalizeLookup(els.biCommandPaletteInput?.value || "");
  const defs = getBiCommandDefinitions(project);
  const filtered = defs.filter((item) => {
    if (!query) {
      return true;
    }
    const blob = normalizeLookup(`${item.label} ${item.keywords || ""}`);
    return blob.includes(query);
  });
  biCommandVisibleIds = filtered.map((item) => item.id);
  if (biCommandVisibleIds.length === 0) {
    biCommandPaletteSelection = 0;
    els.biCommandPaletteList.innerHTML = "<div class=\"bi-empty\">No hay comandos para esa busqueda.</div>";
    return;
  }
  biCommandPaletteSelection = Math.max(0, Math.min(biCommandPaletteSelection, biCommandVisibleIds.length - 1));
  els.biCommandPaletteList.innerHTML = filtered
    .map((item, index) => {
      const active = index === biCommandPaletteSelection ? " active" : "";
      return `<button type="button" class="bi-command-item${active}" data-bi-command-action="${escapeAttribute(item.id)}">
        <strong>${escapeHtml(item.label)}</strong>
        <span>${escapeHtml(item.keywords || "")}</span>
      </button>`;
    })
    .join("");
}

function openBiCommandPalette() {
  if (activeTab !== "bi") {
    return;
  }
  const project = getActiveProject();
  biCommandPaletteOpen = true;
  biCommandPaletteSelection = 0;
  if (els.biCommandPaletteOverlay instanceof HTMLElement) {
    els.biCommandPaletteOverlay.classList.remove("hidden");
  }
  if (els.biCommandPaletteInput instanceof HTMLInputElement) {
    els.biCommandPaletteInput.value = "";
  }
  renderBiCommandPalette(project);
  window.setTimeout(() => {
    els.biCommandPaletteInput?.focus();
    els.biCommandPaletteInput?.select();
  }, 0);
  setStatus("Comandos BI abierto.");
}

function closeBiCommandPalette() {
  biCommandPaletteOpen = false;
  biCommandPaletteSelection = 0;
  biCommandVisibleIds = [];
  if (els.biCommandPaletteOverlay instanceof HTMLElement) {
    els.biCommandPaletteOverlay.classList.add("hidden");
  }
}

function toggleBiCommandPalette() {
  if (biCommandPaletteOpen) {
    closeBiCommandPalette();
    return;
  }
  openBiCommandPalette();
}

function runBiCommandById(commandId) {
  const project = getActiveProject();
  const defs = getBiCommandDefinitions(project);
  const command = defs.find((item) => item.id === commandId);
  if (!command || typeof command.run !== "function") {
    return false;
  }
  closeBiCommandPalette();
  command.run();
  return true;
}

function handleBiCommandPaletteKeydown(event) {
  if (!biCommandPaletteOpen) {
    return false;
  }
  const key = trimOrFallback(event.key, "");
  if (key === "Escape") {
    event.preventDefault();
    closeBiCommandPalette();
    return true;
  }
  if (key === "ArrowDown") {
    event.preventDefault();
    if (biCommandVisibleIds.length > 0) {
      biCommandPaletteSelection = (biCommandPaletteSelection + 1) % biCommandVisibleIds.length;
      renderBiCommandPalette(getActiveProject());
    }
    return true;
  }
  if (key === "ArrowUp") {
    event.preventDefault();
    if (biCommandVisibleIds.length > 0) {
      biCommandPaletteSelection = (biCommandPaletteSelection - 1 + biCommandVisibleIds.length) % biCommandVisibleIds.length;
      renderBiCommandPalette(getActiveProject());
    }
    return true;
  }
  if (key === "Enter") {
    event.preventDefault();
    const commandId = biCommandVisibleIds[biCommandPaletteSelection] || "";
    if (commandId) {
      runBiCommandById(commandId);
    }
    return true;
  }
  return false;
}

function normalizeBiRailMode(mode) {
  if (mode === "properties" || mode === "filter" || mode === "board" || mode === "colorimetry" || mode === "settings" || mode === "performance" || mode === "others") {
    return mode;
  }
  return "data";
}

function focusBiInspectorSection(mode) {
  const normalizedMode = normalizeBiRailMode(mode);
  let section = null;
  if (normalizedMode === "properties") {
    if (els.biSourceSection instanceof HTMLElement && !els.biSourceSection.classList.contains("hidden")) {
      section = els.biSourceSection;
    } else {
      section = els.biBuilderSection;
    }
  } else if (normalizedMode === "filter") {
    section = els.biFilterSection;
  } else if (normalizedMode === "board") {
    section = els.biBoardSection;
  } else if (normalizedMode === "colorimetry") {
    section = els.biColorSection;
  } else if (normalizedMode === "settings") {
    section = els.biSettingsSection;
  } else if (normalizedMode === "performance") {
    section = els.biPerformanceSection;
  }
  if (!(section instanceof HTMLElement)) {
    return;
  }
  if (section.classList.contains("hidden")) {
    return;
  }
  section.classList.add("bi-focus");
  if (biFilterFocusTimer !== null) {
    clearTimeout(biFilterFocusTimer);
  }
  biFilterFocusTimer = window.setTimeout(() => {
    section.classList.remove("bi-focus");
    biFilterFocusTimer = null;
  }, 950);
}

function refreshBiStudioPanelsUi() {
  if (els.biStudioShell) {
    els.biStudioShell.classList.toggle("bi-panel-data-open", !!biDataPanelOpen);
    els.biStudioShell.classList.toggle("bi-panel-props-open", !!biInspectorPanelOpen);
    els.biStudioShell.classList.toggle("bi-panel-other-open", !!biOtherPanelOpen);
  }

  if (els.biStudioOtherPanel instanceof HTMLElement) {
    els.biStudioOtherPanel.classList.toggle("hidden", !biOtherPanelOpen);
  }

  if (els.biRailButtons && els.biRailButtons.length) {
    els.biRailButtons.forEach((button) => {
      if (!(button instanceof HTMLElement)) {
        return;
      }
      const buttonMode = normalizeBiRailMode(button.dataset.biRailMode || "data");
      let active = false;
      if (buttonMode === "data") {
        active = !!biDataPanelOpen;
      } else if (buttonMode === "others") {
        active = !!biOtherPanelOpen;
      } else {
        active = !!biInspectorPanelOpen && biRailMode === buttonMode;
      }
      button.classList.toggle("active", active);
    });
  }
  updateBiWorkspaceInsets();
}

function updateBiWorkspaceInsets() {
  const tabPanel = document.getElementById("tab-bi");
  if (!(tabPanel instanceof HTMLElement)) {
    return;
  }
  applyBiStudioPanelWidths();
  if (!(els.biStudioShell instanceof HTMLElement) || activeTab !== "bi") {
    tabPanel.style.setProperty("--bi-shell-reserve-right", "0px");
    return;
  }
  const desktopMode = window.innerWidth > 1180;
  if (!desktopMode) {
    tabPanel.style.setProperty("--bi-shell-reserve-right", "0px");
    return;
  }
  const shellRect = els.biStudioShell.getBoundingClientRect();
  const reserve = Math.max(72, Math.round(shellRect.width + 12));
  tabPanel.style.setProperty("--bi-shell-reserve-right", `${reserve}px`);
}

function updateBiStudioOverlayOffset() {
  if (!(els.biStudioShell instanceof HTMLElement)) {
    return;
  }
  applyBiStudioPanelWidths();
  const tabPanel = document.getElementById("tab-bi");
  if (!(tabPanel instanceof HTMLElement)) {
    return;
  }
  const header = tabPanel.querySelector(".panel-header");
  const contextBar = document.querySelector(".workspace-main .context-bar");
  const panelRect = tabPanel.getBoundingClientRect();
  const headerBottom = header instanceof HTMLElement
    ? header.getBoundingClientRect().bottom
    : panelRect.top + 82;
  const contextBottom = contextBar instanceof HTMLElement
    ? contextBar.getBoundingClientRect().bottom
    : 0;
  const topOffset = Math.max(56, Math.round(Math.max(headerBottom, contextBottom) + 8));
  const rightOffset = Math.max(10, Math.round(window.innerWidth - panelRect.right + 10));
  els.biStudioShell.style.top = `${topOffset}px`;
  els.biStudioShell.style.right = `${rightOffset}px`;

  if (els.biStudioOtherPanel instanceof HTMLElement) {
    const shellRect = els.biStudioShell.getBoundingClientRect();
    const otherTop = Math.max(topOffset + 8, Math.round(shellRect.bottom + 10));
    els.biStudioOtherPanel.style.top = `${otherTop}px`;
    els.biStudioOtherPanel.style.right = `${rightOffset}px`;
  }
  updateBiWorkspaceInsets();
}

function applyBiRailMode(mode, options = {}) {
  const normalizedMode = normalizeBiRailMode(mode);
  const forceOpen = options.forceOpen === true;
  const forceClose = options.forceClose === true;

  if (normalizedMode === "data") {
    const wasDataMode = biRailMode === "data";
    biRailMode = "data";
    biOtherPanelOpen = false;
    if (forceClose) {
      biDataPanelOpen = false;
    } else if (forceOpen) {
      biDataPanelOpen = true;
    } else if (wasDataMode && biDataPanelOpen) {
      biDataPanelOpen = false;
    } else {
      biDataPanelOpen = true;
    }
    if (biDataPanelOpen) {
      biInspectorPanelOpen = false;
    }
  } else if (normalizedMode === "others") {
    const wasOthersMode = biRailMode === "others";
    biRailMode = "others";
    biDataPanelOpen = false;
    biInspectorPanelOpen = false;
    if (forceClose) {
      biOtherPanelOpen = false;
    } else if (forceOpen) {
      biOtherPanelOpen = true;
    } else if (wasOthersMode && biOtherPanelOpen) {
      biOtherPanelOpen = false;
    } else {
      biOtherPanelOpen = true;
    }
  } else {
    const wasSameInspectorMode = biRailMode === normalizedMode;
    biRailMode = normalizedMode;
    biDataPanelOpen = false;
    biOtherPanelOpen = false;
    if (forceClose) {
      biInspectorPanelOpen = false;
    } else if (forceOpen) {
      biInspectorPanelOpen = true;
    } else if (wasSameInspectorMode && biInspectorPanelOpen) {
      biInspectorPanelOpen = false;
    } else {
      biInspectorPanelOpen = true;
    }
  }

  const activeProject = getActiveProject();
  syncBiInspectorByWidgetType(activeProject, activeProject ? getBiSelectedWidget(activeProject) : null);
  refreshBiStudioPanelsUi();
  updateBiStudioOverlayOffset();
  saveBiStudioUiPreferences();

  if (biInspectorPanelOpen && normalizedMode !== "data") {
    focusBiInspectorSection(normalizedMode);
  }

  if (options.announce === false) {
    return;
  }
  if (normalizedMode === "properties") {
    setStatus(biInspectorPanelOpen ? "Panel de propiedades mostrado." : "Panel de propiedades oculto.");
  } else if (normalizedMode === "board") {
    setStatus(biInspectorPanelOpen ? "Panel de pizarra mostrado." : "Panel de pizarra oculto.");
  } else if (normalizedMode === "colorimetry") {
    setStatus(biInspectorPanelOpen ? "Panel de colorimetria mostrado." : "Panel de colorimetria oculto.");
  } else if (normalizedMode === "settings") {
    setStatus(biInspectorPanelOpen ? "Panel de configuracion mostrado." : "Panel de configuracion oculto.");
  } else if (normalizedMode === "performance") {
    setStatus(biInspectorPanelOpen ? "Panel de rendimiento mostrado." : "Panel de rendimiento oculto.");
  } else if (normalizedMode === "filter") {
    setStatus(biInspectorPanelOpen ? "Panel de filtros mostrado." : "Panel de filtros oculto.");
  } else if (normalizedMode === "others") {
    setStatus(biOtherPanelOpen ? "Panel Otros mostrado." : "Panel Otros oculto.");
  } else {
    setStatus(biDataPanelOpen ? "Panel de datos mostrado." : "Panel de datos oculto.");
  }
}

function renderTabState() {
  const project = getActiveProject();
  const packageTabEnabled = getProjectProgressControlMode(project) === "package";
  if (!packageTabEnabled && activeTab === "packages") {
    activeTab = "deliverables";
  }

  const fieldPanel = document.getElementById("tab-fields");
  const deliverablesPanel = document.getElementById("tab-deliverables");
  const packagesPanel = document.getElementById("tab-packages");
  const reviewFlowPanel = document.getElementById("tab-review-flow");
  const reviewControlsPanel = document.getElementById("tab-review-controls");
  const biPanel = document.getElementById("tab-bi");
  const quickSightPanel = document.getElementById("tab-quicksight");
  const tabButtons = document.querySelectorAll(".tab-button");
  const packagesTabButton = document.querySelector(".tab-button[data-tab=\"packages\"]");

  if (packagesTabButton) {
    packagesTabButton.classList.toggle("hidden", !packageTabEnabled);
  }

  fieldPanel.classList.toggle("hidden", activeTab !== "fields");
  deliverablesPanel.classList.toggle("hidden", activeTab !== "deliverables");
  packagesPanel.classList.toggle("hidden", activeTab !== "packages" || !packageTabEnabled);
  reviewFlowPanel.classList.toggle("hidden", activeTab !== "review-flow");
  reviewControlsPanel.classList.toggle("hidden", activeTab !== "review-controls");
  biPanel.classList.toggle("hidden", activeTab !== "bi");
  quickSightPanel.classList.toggle("hidden", activeTab !== "quicksight");
  if (activeTab !== "quicksight" && quickSightEditMenuOpen) {
    setQuickSightEditMenuOpen(false);
  }
  if (activeTab !== "quicksight" && quickSightZoomMenuOpen) {
    setQuickSightZoomMenuOpen(false);
  }
  if (activeTab !== "quicksight") {
    hideQuickSightDrillThroughDrawerOverlay();
  }
  syncQuickSightViewportModeUi();

  tabButtons.forEach((button) => {
    const isActive = button.dataset.tab === activeTab;
    button.classList.toggle("active", isActive);
  });

  document.body.classList.add("bi-hide-context");

  if (els.currentViewLabel) {
    if (activeTab === "fields") {
      els.currentViewLabel.textContent = "LISTAS";
    } else if (activeTab === "packages") {
      els.currentViewLabel.textContent = "CONTROL PAQUETES";
    } else if (activeTab === "review-flow") {
      els.currentViewLabel.textContent = "HITOS FLUJO REVISION";
    } else if (activeTab === "review-controls") {
      els.currentViewLabel.textContent = "CONTROL FLUJOS REVISION";
    } else if (activeTab === "bi") {
      els.currentViewLabel.textContent = "DECHINI BI";
    } else if (activeTab === "quicksight") {
      els.currentViewLabel.textContent = "DECHINI QUICKSIGTH";
    } else {
      els.currentViewLabel.textContent = "MIDP";
    }
  }

  if (els.globalSearchInput) {
    if (activeTab === "fields") {
      els.globalSearchInput.placeholder = "Buscar en listas";
    } else if (activeTab === "packages") {
      els.globalSearchInput.placeholder = "Buscar en control de paquetes";
    } else if (activeTab === "review-flow") {
      els.globalSearchInput.placeholder = "Buscar en hitos de revision";
    } else if (activeTab === "review-controls") {
      els.globalSearchInput.placeholder = "Buscar en control de flujos de revision";
    } else if (activeTab === "bi") {
      els.globalSearchInput.placeholder = "Buscar en Dechini BI";
    } else if (activeTab === "quicksight") {
      els.globalSearchInput.placeholder = "Buscar en Dechini Quicksigth";
    } else {
      els.globalSearchInput.placeholder = "Buscar en entregables";
    }
  }

  updateMidpEditUi();
  updateFieldsEditUi();
  updatePackageEditUi();
  updateReviewFlowEditUi();
  updateReviewControlsEditUi();
}

function switchTab(tab) {
  const project = getActiveProject();
  const packageTabEnabled = getProjectProgressControlMode(project) === "package";
  if (tab === "packages") {
    activeTab = packageTabEnabled ? "packages" : "deliverables";
  } else if (tab === "deliverables" || tab === "fields" || tab === "review-flow" || tab === "review-controls" || tab === "bi" || tab === "quicksight") {
    activeTab = tab;
  } else {
    activeTab = "fields";
  }
  if (activeTab !== "bi") {
    closeBiCommandPalette();
  }
  const tabMatchesTracking = (activeTab === "deliverables" && trackingPanelTargetType === "deliverable")
    || (activeTab === "packages" && trackingPanelTargetType === "package")
    || (activeTab === "review-controls" && trackingPanelTargetType === "review-control");
  if (!tabMatchesTracking) {
    closeTrackingPanel();
  }
  if (activeTab !== "review-flow") {
    closeReviewMilestoneDrawer();
  }
  renderTabState();
  if (activeTab === "packages") {
    renderPackageControlsPanel(getActiveProject());
  } else if (activeTab === "review-flow") {
    renderReviewFlowPanel(getActiveProject());
  } else if (activeTab === "review-controls") {
    renderReviewControlsPanel(getActiveProject());
  } else if (activeTab === "bi") {
    if (biPendingRenderRafId) {
      window.cancelAnimationFrame(biPendingRenderRafId);
      biPendingRenderRafId = 0;
    }
    biPendingRenderRafId = window.requestAnimationFrame(() => {
      biPendingRenderRafId = 0;
      renderBiPanel(getActiveProject());
    });
  } else if (activeTab === "quicksight") {
    renderQuickSightPanel(getActiveProject());
  }
}

function updateMidpEditUi() {
  if (!els.toggleMidpEditButton) {
    return;
  }

  const editEnabled = midpEditMode;
  els.toggleMidpEditButton.textContent = editEnabled ? "Bloquear edicion" : "Editar";
  els.toggleMidpEditButton.classList.toggle("active", editEnabled);

  if (els.openAddDeliverableButton) {
    els.openAddDeliverableButton.disabled = false;
  }
  if (els.addBlankDeliverableButton) {
    els.addBlankDeliverableButton.disabled = false;
  }
  if (els.recomputeCodesButton) {
    els.recomputeCodesButton.disabled = !editEnabled;
  }
  if (els.toggleMidpAdvanceDetailsButton) {
    const project = getActiveProject();
    const showProgressColumns = getProjectProgressControlMode(project) !== "package";
    const showDetails = getProjectMidpAdvanceDetailsVisible(project);
    els.toggleMidpAdvanceDetailsButton.textContent = showDetails ? "Ocultar avances" : "Mostrar avances";
    els.toggleMidpAdvanceDetailsButton.disabled = !showProgressColumns;
    els.toggleMidpAdvanceDetailsButton.classList.toggle("hidden", !showProgressColumns);
  }
}

function updateFieldsEditUi() {
  if (!els.toggleFieldsEditButton) {
    return;
  }

  els.toggleFieldsEditButton.textContent = fieldsEditMode ? "Bloquear edicion" : "Editar";
  els.toggleFieldsEditButton.classList.toggle("active", fieldsEditMode);

  if (els.projectTitleInput) {
    els.projectTitleInput.disabled = !fieldsEditMode;
  }
  if (els.fieldSeparatorInput) {
    els.fieldSeparatorInput.disabled = !fieldsEditMode;
  }
  if (els.progressControlModeSelect) {
    els.progressControlModeSelect.disabled = !fieldsEditMode;
  }
  if (els.addFieldButton) {
    els.addFieldButton.disabled = !fieldsEditMode;
  }
  if (els.addFieldRowButton) {
    els.addFieldRowButton.disabled = !fieldsEditMode;
  }
}

function updatePackageEditUi() {
  if (!els.togglePackageEditButton) {
    return;
  }

  els.togglePackageEditButton.textContent = packageEditMode ? "Bloquear edicion" : "Editar";
  els.togglePackageEditButton.classList.toggle("active", packageEditMode);
}

function updateReviewFlowEditUi() {
  if (!els.toggleReviewFlowEditButton) {
    return;
  }

  els.toggleReviewFlowEditButton.textContent = reviewFlowEditMode ? "Bloquear edicion" : "Editar";
  els.toggleReviewFlowEditButton.classList.toggle("active", reviewFlowEditMode);

  if (els.openAddReviewMilestoneButton) {
    els.openAddReviewMilestoneButton.disabled = !reviewFlowEditMode;
  }
  if (els.addBlankReviewMilestoneButton) {
    els.addBlankReviewMilestoneButton.disabled = !reviewFlowEditMode;
  }
}

function updateReviewControlsEditUi() {
  if (!els.toggleReviewControlsEditButton) {
    return;
  }

  els.toggleReviewControlsEditButton.textContent = reviewControlsEditMode ? "Bloquear edicion" : "Editar";
  els.toggleReviewControlsEditButton.classList.toggle("active", reviewControlsEditMode);
}

function openProjectChooser(lockSelection) {
  chooserLocked = !!lockSelection;
  renderProjectChooser(getActiveProject());
  els.projectChooserOverlay.classList.remove("hidden");
}

function closeProjectChooser(forceClose) {
  if (chooserLocked && !forceClose) {
    return;
  }
  els.projectChooserOverlay.classList.add("hidden");
}

function openDeliverableDrawer() {
  if (drawerCloseTimer !== null) {
    clearTimeout(drawerCloseTimer);
    drawerCloseTimer = null;
  }

  els.deliverableDrawerOverlay.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.deliverableDrawerOverlay.classList.add("open");
  });
}

function closeDeliverableDrawer() {
  if (els.deliverableDrawerOverlay.classList.contains("hidden")) {
    return;
  }

  els.deliverableDrawerOverlay.classList.remove("open");
  if (drawerCloseTimer !== null) {
    clearTimeout(drawerCloseTimer);
  }
  drawerCloseTimer = window.setTimeout(() => {
    els.deliverableDrawerOverlay.classList.add("hidden");
    drawerCloseTimer = null;
  }, 220);
}

function openReviewMilestoneDrawer() {
  const project = getActiveProject();
  if (!project) {
    return;
  }

  if (els.reviewMilestoneNameInput) {
    els.reviewMilestoneNameInput.value = "";
  }
  if (els.reviewMilestoneWeightInput) {
    els.reviewMilestoneWeightInput.value = "";
  }

  if (reviewMilestoneDrawerCloseTimer !== null) {
    clearTimeout(reviewMilestoneDrawerCloseTimer);
    reviewMilestoneDrawerCloseTimer = null;
  }

  els.reviewMilestoneDrawerOverlay.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.reviewMilestoneDrawerOverlay.classList.add("open");
    els.reviewMilestoneNameInput?.focus();
  });
}

function closeReviewMilestoneDrawer() {
  if (els.reviewMilestoneDrawerOverlay.classList.contains("hidden")) {
    return;
  }

  els.reviewMilestoneDrawerOverlay.classList.remove("open");
  if (reviewMilestoneDrawerCloseTimer !== null) {
    clearTimeout(reviewMilestoneDrawerCloseTimer);
  }
  reviewMilestoneDrawerCloseTimer = window.setTimeout(() => {
    els.reviewMilestoneDrawerOverlay.classList.add("hidden");
    reviewMilestoneDrawerCloseTimer = null;
  }, 220);
}

function getTrackingContext() {
  if (!trackingPanelTargetType || !trackingPanelTargetId) {
    return null;
  }
  const project = getActiveProject();
  if (!project) {
    return null;
  }

  let entity = null;
  if (trackingPanelTargetType === "deliverable") {
    entity = project.deliverables.find((item) => item.id === trackingPanelTargetId) || null;
  } else if (trackingPanelTargetType === "package") {
    ensurePackageControls(project);
    entity = project.packageControls.find((item) => item.id === trackingPanelTargetId) || null;
  } else if (trackingPanelTargetType === "review-control") {
    ensureReviewControls(project);
    entity = project.reviewControls.find((item) => item.id === trackingPanelTargetId) || null;
  }
  if (!entity) {
    return null;
  }

  ensureTrackingCollections(entity);
  return { project, entity, entityType: trackingPanelTargetType };
}

function openTrackingPanel(deliverableId) {
  const project = getActiveProject();
  if (!project) {
    return;
  }

  const deliverable = project.deliverables.find((item) => item.id === deliverableId);
  if (!deliverable) {
    return;
  }

  trackingPanelTargetType = "deliverable";
  trackingPanelTargetId = deliverable.id;
  ensureTrackingCollections(deliverable);
  renderTrackingPanel(project, deliverable, "deliverable");
  openTrackingPanelOverlay();
}

function openPackageTrackingPanel(packageControlId) {
  const project = getActiveProject();
  if (!project) {
    return;
  }

  ensurePackageControls(project);
  const packageControl = project.packageControls.find((item) => item.id === packageControlId);
  if (!packageControl) {
    return;
  }

  trackingPanelTargetType = "package";
  trackingPanelTargetId = packageControl.id;
  ensureTrackingCollections(packageControl);
  renderTrackingPanel(project, packageControl, "package");
  openTrackingPanelOverlay();
}

function openReviewControlTrackingPanel(reviewControlId) {
  const project = getActiveProject();
  if (!project) {
    return;
  }

  ensureReviewControls(project);
  const reviewControl = project.reviewControls.find((item) => item.id === reviewControlId);
  if (!reviewControl) {
    return;
  }

  trackingPanelTargetType = "review-control";
  trackingPanelTargetId = reviewControl.id;
  ensureTrackingCollections(reviewControl);
  renderTrackingPanel(project, reviewControl, "review-control");
  openTrackingPanelOverlay();
}

function openTrackingPanelOverlay() {

  if (trackingCloseTimer !== null) {
    clearTimeout(trackingCloseTimer);
    trackingCloseTimer = null;
  }
  els.trackingPanelOverlay.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.trackingPanelOverlay.classList.add("open");
  });
}

function closeTrackingPanel() {
  trackingPanelTargetType = "";
  trackingPanelTargetId = "";
  if (els.trackingPanelOverlay.classList.contains("hidden")) {
    return;
  }

  els.trackingPanelOverlay.classList.remove("open");
  if (trackingCloseTimer !== null) {
    clearTimeout(trackingCloseTimer);
  }
  trackingCloseTimer = window.setTimeout(() => {
    els.trackingPanelOverlay.classList.add("hidden");
    trackingCloseTimer = null;
  }, 220);
}

function getTrackingStatusLabel(context) {
  if (!context || !context.project || !context.entity) {
    return "Fila";
  }

  if (context.entityType === "deliverable") {
    const ref = trimOrFallback(context.entity.ref, "");
    return ref ? `Fila ${ref}` : "Fila";
  }

  if (context.entityType === "package") {
    const packageCode = buildPackageControlTrackingCode(context.project, context.entity);
    return packageCode ? `Paquete ${packageCode}` : "Paquete";
  }

  const reviewCode = buildReviewControlTrackingCode(context.project, context.entity);
  return reviewCode ? `Flujo ${reviewCode}` : "Flujo";
}

function buildPackageControlTrackingCode(project, packageControl) {
  if (!project || !packageControl) {
    return "";
  }

  const fieldIds = getPackageControlFieldIds(project.fields);
  const projectField = getFieldById(project.fields, fieldIds.projectFieldId);
  const disciplineField = getFieldById(project.fields, fieldIds.disciplineFieldId);
  const systemField = getFieldById(project.fields, fieldIds.systemFieldId);
  const packageField = getFieldById(project.fields, fieldIds.packageFieldId);
  const projectRow = getFieldRowById(projectField, packageControl.projectRowId);
  const disciplineRow = getFieldRowById(disciplineField, packageControl.disciplineRowId);
  const systemRow = getFieldRowById(systemField, packageControl.systemRowId);
  const packageRow = getFieldRowById(packageField, packageControl.packageRowId);
  return buildPackageControlCode(project, [projectRow, disciplineRow, systemRow, packageRow])
    || trimOrFallback(packageControl.key || "", "");
}

function buildReviewControlTrackingCode(project, reviewControl) {
  if (!project || !reviewControl) {
    return "";
  }

  const fieldIds = getReviewControlFieldIds(project.fields);
  const projectField = getFieldById(project.fields, fieldIds.projectFieldId);
  const disciplineField = getFieldById(project.fields, fieldIds.disciplineFieldId);
  const systemField = getFieldById(project.fields, fieldIds.systemFieldId);
  const projectRow = getFieldRowById(projectField, reviewControl.projectRowId);
  const disciplineRow = getFieldRowById(disciplineField, reviewControl.disciplineRowId);
  const systemRow = getFieldRowById(systemField, reviewControl.systemRowId);
  const combination = buildPackageControlCode(project, [projectRow, disciplineRow, systemRow]);
  ensureReviewMilestones(project);
  const milestone = project.reviewMilestones.find((item) => item.id === reviewControl.milestoneId) || null;
  const milestoneName = trimOrFallback(milestone?.name, "");
  const detail = [combination, milestoneName].filter((token) => !!token).join(" | ");
  return detail || trimOrFallback(reviewControl.key || "", "");
}

function renderTrackingPanel(project, entity, entityType) {
  ensureTrackingCollections(entity);
  const type = entityType === "package"
    ? "package"
    : (entityType === "review-control" ? "review-control" : "deliverable");
  const packageCode = type === "package" ? buildPackageControlTrackingCode(project, entity) : "";
  const reviewCode = type === "review-control" ? buildReviewControlTrackingCode(project, entity) : "";
  if (els.trackingPanelTitle) {
    if (type === "package") {
      els.trackingPanelTitle.textContent = `Detalle paquete ${packageCode || ""}`.trim();
    } else if (type === "review-control") {
      els.trackingPanelTitle.textContent = `Detalle flujo ${reviewCode || ""}`.trim();
    } else {
      els.trackingPanelTitle.textContent = `Detalle ${entity.ref || ""}`.trim();
    }
  }
  if (els.trackingPanelSubtitle) {
    if (type === "package") {
      els.trackingPanelSubtitle.textContent = packageCode || project.title || "";
    } else if (type === "review-control") {
      els.trackingPanelSubtitle.textContent = reviewCode || project.title || "";
    } else {
      els.trackingPanelSubtitle.textContent = entity.code || project.title || "";
    }
  }

  if (els.realAdvancesBody) {
    if (entity.realAdvances.length === 0) {
      els.realAdvancesBody.innerHTML = "<tr><td class=\"tracking-empty\" colspan=\"6\">Sin registros</td></tr>";
    } else {
      els.realAdvancesBody.innerHTML = entity.realAdvances
        .map((item, index) => {
          const stamp = trimOrFallback(item.createdAt || "", "");
          const dateText = stamp ? formatDate(stamp) : formatDateFromInput(item.date || "");
          const userText = trimOrFallback(item.author || "", "");
          const addedText = formatSignedPercent(item.addedPercent, 2);
          const valueText = formatNumberForInput(item.value);
          return `<tr>
            <td>${index + 1}</td>
            <td>${escapeHtml(dateText)}</td>
            <td>${escapeHtml(userText)}</td>
            <td>${escapeHtml(addedText)}</td>
            <td><input type="number" min="0" max="100" step="0.01" class="tracking-input" data-real-advance-row="${escapeAttribute(item.id)}" data-real-advance-key="value" value="${escapeAttribute(valueText)}"></td>
            <td><button type="button" class="danger mini-button" data-remove-real-advance="${escapeAttribute(item.id)}">Quitar</button></td>
          </tr>`;
        })
        .join("");
    }
  }

  if (els.consumedHoursBody) {
    if (entity.consumedHours.length === 0) {
      els.consumedHoursBody.innerHTML = "<tr><td class=\"tracking-empty\" colspan=\"5\">Sin registros</td></tr>";
    } else {
      els.consumedHoursBody.innerHTML = entity.consumedHours
        .map((item, index) => {
          const hoursText = formatNumberForInput(item.hours);
          const noteText = trimOrFallback(item.note || "", "");
          return `<tr>
            <td>${index + 1}</td>
            <td><input type="date" class="tracking-input" data-consumed-hour-row="${escapeAttribute(item.id)}" data-consumed-hour-key="date" value="${escapeAttribute(sanitizeDateInput(item.date || ""))}"></td>
            <td><input type="number" min="0" step="0.01" class="tracking-input" data-consumed-hour-row="${escapeAttribute(item.id)}" data-consumed-hour-key="hours" value="${escapeAttribute(hoursText)}"></td>
            <td><input type="text" maxlength="${TRACKING_MAX_NOTE}" class="tracking-input tracking-note-input" data-consumed-hour-row="${escapeAttribute(item.id)}" data-consumed-hour-key="note" value="${escapeAttribute(noteText)}"></td>
            <td><button type="button" class="danger mini-button" data-remove-consumed-hour="${escapeAttribute(item.id)}">Quitar</button></td>
          </tr>`;
        })
        .join("");
    }
  }
}

function ensureActiveProject() {
  if (!Array.isArray(state.projects) || state.projects.length === 0) {
    state = createDefaultState();
  }

  const exists = state.projects.some((project) => project.id === state.activeProjectId);
  if (!exists) {
    state.activeProjectId = state.projects[0].id;
  }

  let shouldPersist = false;
  state.projects.forEach((project) => {
    ensureNomenclatureConfig(project);
    const demoResult = ensureProjectDemoData(project);
    if (demoResult.changed) {
      shouldPersist = true;
    }
  });

  if (shouldPersist) {
    saveState(true);
  }
}

function getActiveProject() {
  return state.projects.find((project) => project.id === state.activeProjectId) || null;
}

function ensureProjectDemoData(project) {
  if (!project) {
    return { changed: false, added: 0 };
  }

  ensurePackageControls(project);
  const wasDemoSeeded = !!project.demoSeeded;
  if (wasDemoSeeded) {
    return { changed: false, added: 0 };
  }

  if (Array.isArray(project.deliverables) && project.deliverables.length > 0) {
    const missingToEight = Math.max(0, 8 - project.deliverables.length);
    const added = missingToEight > 0
      ? populateProjectWithDemoDeliverables(project, missingToEight)
      : 0;
    const syncResult = syncPackageControlsFromDeliverables(project);
    project.demoSeeded = true;
    return { changed: added > 0 || syncResult.added > 0 || !wasDemoSeeded, added };
  }

  const added = populateProjectWithDemoDeliverables(project);
  const syncResult = syncPackageControlsFromDeliverables(project);
  project.demoSeeded = true;
  return { changed: added > 0 || syncResult.added > 0 || !wasDemoSeeded, added };
}

function populateProjectWithDemoDeliverables(project, maxItems) {
  const fieldIds = getPackageControlFieldIds(project.fields);
  if (!fieldIds.projectFieldId || !fieldIds.disciplineFieldId || !fieldIds.systemFieldId || !fieldIds.packageFieldId) {
    return 0;
  }

  const creatorFieldId = findFieldIdByAlias(project.fields, ["creador"]);
  const phaseFieldId = findFieldIdByAlias(project.fields, ["fase"]);
  const sectorFieldId = findFieldIdByAlias(project.fields, ["sector"]);
  const levelFieldId = findFieldIdByAlias(project.fields, ["nivel"]);
  const typeFieldId = findFieldIdByAlias(project.fields, ["tipo"]);
  const numberFieldId = findFieldIdByAlias(project.fields, ["numero"]);
  const scaleFieldId = findFieldIdByAlias(project.fields, ["escala"]);

  const scenarios = [
    { creator: 0, phase: 0, sector: 0, level: 0, type: 0, discipline: 0, system: 0, package: 0, number: 0, scale: 0, startDate: "2026-02-01", endDate: "2026-03-20", baseUnits: 50, realProgress: 42 },
    { creator: 1, phase: 1, sector: 1, level: 1, type: 0, discipline: 0, system: 0, package: 0, number: 1, scale: 1, startDate: "2026-02-10", endDate: "2026-03-10", baseUnits: 30, realProgress: 25 },
    { creator: 0, phase: 1, sector: 0, level: 2, type: 1, discipline: 0, system: 1, package: 1, number: 2, scale: 0, startDate: "2026-01-15", endDate: "2026-02-20", baseUnits: 40, realProgress: 100 },
    { creator: 1, phase: 2, sector: 2, level: 1, type: 2, discipline: 1, system: 1, package: 1, number: 0, scale: 2, startDate: "2026-02-20", endDate: "2026-04-05", baseUnits: 22, realProgress: 12 },
    { creator: 0, phase: 0, sector: 1, level: 0, type: 1, discipline: 1, system: 2, package: 2, number: 1, scale: 0, startDate: "2026-03-05", endDate: "2026-03-25", baseUnits: 28, realProgress: 0 },
    { creator: 1, phase: 1, sector: 2, level: 2, type: 2, discipline: 2, system: 2, package: 2, number: 2, scale: 1, startDate: "2026-02-01", endDate: "2026-02-28", baseUnits: 35, realProgress: 68 },
    { creator: 0, phase: 2, sector: 0, level: 0, type: 0, discipline: 2, system: 0, package: 0, number: 0, scale: 0, startDate: "2026-02-14", endDate: "2026-03-02", baseUnits: 18, realProgress: 40 },
    { creator: 1, phase: 0, sector: 1, level: 1, type: 0, discipline: 1, system: 0, package: 2, number: 1, scale: 1, startDate: "2026-01-20", endDate: "2026-03-30", baseUnits: 26, realProgress: 30 },
    { creator: 0, phase: 2, sector: 2, level: 0, type: 2, discipline: 0, system: 2, package: 1, number: 2, scale: 2, startDate: "2026-02-27", endDate: "2026-03-15", baseUnits: 14, realProgress: 5 },
    { creator: 1, phase: 1, sector: 1, level: 2, type: 1, discipline: 2, system: 1, package: 0, number: 0, scale: 0, startDate: "2026-01-05", endDate: "2026-01-25", baseUnits: 12, realProgress: 100 },
    { creator: 0, phase: 1, sector: 0, level: 1, type: 2, discipline: 1, system: 1, package: 0, number: 2, scale: 1, startDate: "2026-03-08", endDate: "2026-02-28", baseUnits: 10, realProgress: 0 },
    { creator: 1, phase: 0, sector: 2, level: 0, type: 1, discipline: 0, system: 1, package: 2, number: 1, scale: 0, startDate: "2026-02-05", endDate: "2026-03-28", baseUnits: 16, realProgress: 24 }
  ];

  let created = 0;
  const limit = Number.isFinite(maxItems) && maxItems > 0
    ? Math.min(scenarios.length, Math.floor(maxItems))
    : scenarios.length;
  scenarios.slice(0, limit).forEach((scenario, index) => {
    const rowRefs = createEmptyRowRefs(project);
    assignDemoRowRef(project, rowRefs, fieldIds.projectFieldId, 0);
    assignDemoRowRef(project, rowRefs, creatorFieldId, scenario.creator);
    assignDemoRowRef(project, rowRefs, phaseFieldId, scenario.phase);
    assignDemoRowRef(project, rowRefs, sectorFieldId, scenario.sector);
    assignDemoRowRef(project, rowRefs, levelFieldId, scenario.level);
    assignDemoRowRef(project, rowRefs, typeFieldId, scenario.type);
    assignDemoRowRef(project, rowRefs, fieldIds.disciplineFieldId, scenario.discipline);
    assignDemoRowRef(project, rowRefs, fieldIds.systemFieldId, scenario.system);
    assignDemoRowRef(project, rowRefs, fieldIds.packageFieldId, scenario.package);
    assignDemoRowRef(project, rowRefs, numberFieldId, scenario.number);
    assignDemoRowRef(project, rowRefs, scaleFieldId, scenario.scale);

    const createdAt = new Date(Date.now() - ((index + 1) * 3600000)).toISOString();
    const deliverable = createSeedDeliverable(project, rowRefs, {
      startDate: scenario.startDate,
      endDate: scenario.endDate,
      baseUnits: scenario.baseUnits,
      realProgress: scenario.realProgress
    }, createdAt);
    if (!deliverable) {
      return;
    }

    if (scenario.realProgress > 0) {
      deliverable.realAdvances.unshift({
        id: uid(),
        createdAt,
        date: sanitizeDateInput(scenario.startDate),
        value: scenario.realProgress,
        addedPercent: scenario.realProgress,
        author: "DEMO"
      });
    }

    deliverable.consumedHours.unshift({
      id: uid(),
      date: sanitizeDateInput(scenario.startDate),
      hours: Math.max(1, Math.round((scenario.baseUnits || 0) * 0.5)),
      note: "Carga inicial demo"
    });

    project.deliverables.unshift(deliverable);
    created += 1;
  });

  return created;
}

function createSeedDeliverable(project, rowRefsInput, progressMetaInput, createdAtOverride) {
  if (!project) {
    return null;
  }

  const rowRefs = {};
  project.fields.forEach((field) => {
    rowRefs[field.id] = trimOrFallback(rowRefsInput?.[field.id], "");
  });

  const progressMeta = normalizeProgressMeta(progressMetaInput);
  const deliverable = {
    id: uid(),
    ref: buildDeliverableRef(project),
    code: "",
    startDate: progressMeta.startDate,
    endDate: progressMeta.endDate,
    baseUnits: progressMeta.baseUnits,
    realProgress: progressMeta.realProgress,
    realAdvances: [],
    consumedHours: [],
    rowRefs,
    values: {},
    codes: {},
    createdAt: trimOrFallback(createdAtOverride, new Date().toISOString())
  };

  applyDeliverableRows(project, deliverable, false);
  return deliverable;
}

function assignDemoRowRef(project, rowRefs, fieldId, desiredIndex) {
  if (!fieldId || !rowRefs) {
    return;
  }

  const field = getFieldById(project.fields, fieldId);
  if (!field) {
    return;
  }

  const options = getFieldOptions(field);
  if (options.length === 0) {
    rowRefs[fieldId] = "";
    return;
  }

  const safeIndex = Math.max(0, desiredIndex || 0) % options.length;
  rowRefs[fieldId] = options[safeIndex].id;
}

function insertDeliverable(project, rowRefsInput, useFallbackCode, progressMetaInput) {
  const rowRefs = {};
  project.fields.forEach((field) => {
    rowRefs[field.id] = trimOrFallback(rowRefsInput?.[field.id] || "", "");
  });

  const progressMeta = normalizeProgressMeta(progressMetaInput);
  const deliverable = {
    id: uid(),
    ref: buildDeliverableRef(project),
    code: "",
    startDate: progressMeta.startDate,
    endDate: progressMeta.endDate,
    baseUnits: progressMeta.baseUnits,
    realProgress: progressMeta.realProgress,
    realAdvances: [],
    consumedHours: [],
    rowRefs,
    values: {},
    codes: {},
    createdAt: new Date().toISOString()
  };

  applyDeliverableRows(project, deliverable, !!useFallbackCode);
  project.deliverables.unshift(deliverable);
  saveState();
}

function applyDeliverableRows(project, deliverable, useFallbackCode) {
  const selectedRows = {};
  project.fields.forEach((field) => {
    const rowId = trimOrFallback(deliverable.rowRefs?.[field.id] || "", "");
    const row = getFieldRowById(field, rowId);
    selectedRows[field.id] = row;
    deliverable.values[field.id] = row ? row.name : "";
    deliverable.codes[field.id] = row ? row.code : "";
  });

  deliverable.code = buildNomenclature(project, selectedRows, { fallbackWhenEmpty: !!useFallbackCode });
}

function buildNomenclature(project, selectedRows, options) {
  const fallbackWhenEmpty = options?.fallbackWhenEmpty ?? true;
  const parts = [];
  const orderedFields = getOrderedNomenclatureFields(project);
  orderedFields.forEach((field) => {

    const row = selectedRows[field.id];
    let source = trimOrFallback(row?.code || "", "");
    if (!source) {
      source = trimOrFallback(row?.name || "", "");
    }
    if (!source && isProjectLikeField(field.label)) {
      source = project.name;
    }

    const token = toCodeToken(source);
    if (token) {
      parts.push(token);
    }
  });

  if (parts.length === 0) {
    if (!fallbackWhenEmpty) {
      return "";
    }

    const nextNumber = String(project.deliverables.length + 1).padStart(3, "0");
    parts.push(`ENT${nextNumber}`);
  }

  return parts.join(project.codeSeparator || "-");
}

function buildDeliverableRef(project) {
  let max = 0;
  project.deliverables.forEach((item) => {
    const text = trimOrFallback(item.ref || "", "");
    const match = /^REF-(\d+)$/i.exec(text);
    if (!match) {
      return;
    }

    const parsed = parseInt(match[1], 10);
    if (!Number.isNaN(parsed) && parsed > max) {
      max = parsed;
    }
  });

  const next = max + 1;
  return `REF-${String(next).padStart(3, "0")}`;
}

function createEmptyRowRefs(project) {
  const refs = {};
  project.fields.forEach((field) => {
    refs[field.id] = "";
  });
  return refs;
}

function createEmptyProgressMeta() {
  return {
    startDate: "",
    endDate: "",
    baseUnits: null,
    realProgress: null
  };
}

function ensureTrackingCollections(deliverable) {
  if (!Array.isArray(deliverable.realAdvances)) {
    deliverable.realAdvances = [];
  }
  if (!Array.isArray(deliverable.consumedHours)) {
    deliverable.consumedHours = [];
  }
}

function getCurrentUserLabel() {
  const raw = trimOrFallback(els.currentUserBadge?.textContent || "", "");
  return raw || "Usuario";
}

function registerRealAdvanceLog(projectId, deliverableId, previousValueInput, nextValueInput) {
  const nextValue = sanitizeRealProgress(nextValueInput);
  if (nextValue === "") {
    return;
  }

  const previousValue = sanitizeRealProgress(previousValueInput);
  const addedPercent = previousValue === ""
    ? nextValue
    : Math.round((nextValue - previousValue) * 100) / 100;

  pendingRealAdvanceLogs.push({
    projectId,
    deliverableId,
    entry: {
      id: uid(),
      createdAt: new Date().toISOString(),
      date: "",
      value: nextValue,
      addedPercent,
      author: getCurrentUserLabel()
    }
  });
}

function registerPackageRealAdvanceLog(projectId, packageControlId, previousValueInput, nextValueInput) {
  const nextValue = sanitizeRealProgress(nextValueInput);
  if (nextValue === "") {
    return;
  }

  const previousValue = sanitizeRealProgress(previousValueInput);
  const addedPercent = previousValue === ""
    ? nextValue
    : Math.round((nextValue - previousValue) * 100) / 100;

  pendingPackageRealAdvanceLogs.push({
    projectId,
    packageControlId,
    entry: {
      id: uid(),
      createdAt: new Date().toISOString(),
      date: "",
      value: nextValue,
      addedPercent,
      author: getCurrentUserLabel()
    }
  });
}

function registerReviewControlRealAdvanceLog(projectId, reviewControlId, previousValueInput, nextValueInput) {
  const nextValue = sanitizeRealProgress(nextValueInput);
  if (nextValue === "") {
    return;
  }

  const previousValue = sanitizeRealProgress(previousValueInput);
  const addedPercent = previousValue === ""
    ? nextValue
    : Math.round((nextValue - previousValue) * 100) / 100;

  pendingReviewControlRealAdvanceLogs.push({
    projectId,
    reviewControlId,
    entry: {
      id: uid(),
      createdAt: new Date().toISOString(),
      date: "",
      value: nextValue,
      addedPercent,
      author: getCurrentUserLabel()
    }
  });
}

function applyPendingRealAdvanceLogs() {
  if (pendingRealAdvanceLogs.length === 0) {
    return 0;
  }

  let applied = 0;
  pendingRealAdvanceLogs.forEach((log) => {
    const project = state.projects.find((item) => item.id === log.projectId);
    if (!project) {
      return;
    }

    const deliverable = project.deliverables.find((item) => item.id === log.deliverableId);
    if (!deliverable) {
      return;
    }

    ensureTrackingCollections(deliverable);
    deliverable.realAdvances.unshift(log.entry);
    applied += 1;
  });
  pendingRealAdvanceLogs.length = 0;
  return applied;
}

function applyPendingPackageRealAdvanceLogs() {
  if (pendingPackageRealAdvanceLogs.length === 0) {
    return 0;
  }

  let applied = 0;
  pendingPackageRealAdvanceLogs.forEach((log) => {
    const project = state.projects.find((item) => item.id === log.projectId);
    if (!project) {
      return;
    }

    const packageControl = project.packageControls.find((item) => item.id === log.packageControlId);
    if (!packageControl) {
      return;
    }

    ensureTrackingCollections(packageControl);
    packageControl.realAdvances.unshift(log.entry);
    applied += 1;
  });
  pendingPackageRealAdvanceLogs.length = 0;
  return applied;
}

function applyPendingReviewControlRealAdvanceLogs() {
  if (pendingReviewControlRealAdvanceLogs.length === 0) {
    return 0;
  }

  let applied = 0;
  pendingReviewControlRealAdvanceLogs.forEach((log) => {
    const project = state.projects.find((item) => item.id === log.projectId);
    if (!project) {
      return;
    }

    ensureReviewControls(project);
    const reviewControl = project.reviewControls.find((item) => item.id === log.reviewControlId);
    if (!reviewControl) {
      return;
    }

    ensureTrackingCollections(reviewControl);
    reviewControl.realAdvances.unshift(log.entry);
    applied += 1;
  });
  pendingReviewControlRealAdvanceLogs.length = 0;
  return applied;
}

function recomputeAllDeliverableCodes(project, useFallbackCode) {
  project.deliverables.forEach((deliverable) => {
    applyDeliverableRows(project, deliverable, !!useFallbackCode);
  });
}

function ensureNomenclatureConfig(project) {
  if (!Array.isArray(project.nomenclatureOrder)) {
    project.nomenclatureOrder = [];
  }

  const fieldIdSet = new Set(project.fields.map((field) => field.id));
  project.nomenclatureOrder = project.nomenclatureOrder.filter((fieldId) => fieldIdSet.has(fieldId));

  project.fields.forEach((field) => {
    if (!field.includeInCode) {
      return;
    }

    if (!project.nomenclatureOrder.includes(field.id)) {
      project.nomenclatureOrder.push(field.id);
    }
  });
}

function getOrderedNomenclatureFields(project) {
  ensureNomenclatureConfig(project);
  const fieldMap = new Map(project.fields.map((field) => [field.id, field]));
  const ordered = [];

  project.nomenclatureOrder.forEach((fieldId) => {
    const field = fieldMap.get(fieldId);
    if (field && field.includeInCode) {
      ordered.push(field);
    }
  });

  return ordered;
}

function addToNomenclatureOrder(project, fieldId) {
  ensureNomenclatureConfig(project);
  if (!project.nomenclatureOrder.includes(fieldId)) {
    project.nomenclatureOrder.push(fieldId);
  }
}

function removeFromNomenclatureOrder(project, fieldId) {
  ensureNomenclatureConfig(project);
  project.nomenclatureOrder = project.nomenclatureOrder.filter((id) => id !== fieldId);
}

function moveNomenclatureField(project, fieldId, direction) {
  ensureNomenclatureConfig(project);
  const index = project.nomenclatureOrder.indexOf(fieldId);
  if (index < 0) {
    return;
  }

  const target = index + direction;
  if (target < 0 || target >= project.nomenclatureOrder.length) {
    return;
  }

  const aux = project.nomenclatureOrder[target];
  project.nomenclatureOrder[target] = project.nomenclatureOrder[index];
  project.nomenclatureOrder[index] = aux;
}

function moveNomenclatureFieldToTarget(project, draggedFieldId, targetFieldId, placeAfter) {
  ensureNomenclatureConfig(project);
  const fromIndex = project.nomenclatureOrder.indexOf(draggedFieldId);
  const targetIndex = project.nomenclatureOrder.indexOf(targetFieldId);
  if (fromIndex < 0 || targetIndex < 0 || draggedFieldId === targetFieldId) {
    return false;
  }

  const nextOrder = project.nomenclatureOrder.slice();
  nextOrder.splice(fromIndex, 1);
  let insertIndex = nextOrder.indexOf(targetFieldId);
  if (insertIndex < 0) {
    return false;
  }
  if (placeAfter) {
    insertIndex += 1;
  }

  nextOrder.splice(insertIndex, 0, draggedFieldId);
  const changed = nextOrder.join("|") !== project.nomenclatureOrder.join("|");
  if (changed) {
    project.nomenclatureOrder = nextOrder;
  }
  return changed;
}

function clearNomenclatureDropMarkers() {
  if (!els.nomenclatureDragArea) {
    return;
  }

  els.nomenclatureDragArea
    .querySelectorAll(".nomenclature-item.drop-before, .nomenclature-item.drop-after")
    .forEach((item) => {
      item.classList.remove("drop-before", "drop-after");
    });
}

function renderNomenclaturePreview(project) {
  if (!els.nomenclaturePreviewText) {
    return;
  }

  if (!project) {
    els.nomenclaturePreviewText.textContent = "(sin configuracion)";
    return;
  }

  const orderedFields = getOrderedNomenclatureFields(project);
  if (orderedFields.length === 0) {
    els.nomenclaturePreviewText.textContent = "(sin apartados marcados)";
    return;
  }

  const selectedRows = {};
  orderedFields.forEach((field) => {
    selectedRows[field.id] = getFieldOptions(field)[0] || null;
  });

  const preview = buildNomenclature(project, selectedRows, { fallbackWhenEmpty: false });
  els.nomenclaturePreviewText.textContent = preview || "(sin datos en listas)";
}

function getFilteredFieldIndexes(fields, normalizedQuery) {
  const result = new Set();
  if (!normalizedQuery) {
    return result;
  }

  fields.forEach((field) => {
    field.rows.forEach((row, rowIndex) => {
      const text = normalizeLookup(`${row.name} ${row.code}`);
      if (text.includes(normalizedQuery)) {
        result.add(rowIndex);
      }
    });
  });

  return result;
}

function filterDeliverablesBySearch(deliverables, project, normalizedQuery) {
  if (!normalizedQuery) {
    return deliverables;
  }

  const totalBaseUnits = computeProjectBaseUnitsTotal(deliverables);
  const disciplineFieldId = findFieldIdByAlias(project.fields, ["disciplina"]);
  const systemFieldId = findFieldIdByAlias(project.fields, ["sistema"]);
  const disciplineTotals = computeGroupedBaseUnitsTotals(deliverables, disciplineFieldId);
  const systemGroupFieldIds = disciplineFieldId && systemFieldId ? [disciplineFieldId, systemFieldId] : [];
  const systemTotals = computeGroupedBaseUnitsTotals(deliverables, systemGroupFieldIds);
  return deliverables.filter((deliverable) => {
    const baseUnits = sanitizeBaseUnits(deliverable.baseUnits);
    const incidenceRatio = computeProjectIncidenceRatio(baseUnits, totalBaseUnits);
    const disciplineRowId = disciplineFieldId ? trimOrFallback(deliverable.rowRefs?.[disciplineFieldId], "") : "";
    const systemGroupKey = buildDeliverableGroupKey(deliverable, systemGroupFieldIds);
    const disciplineTotal = disciplineRowId ? (disciplineTotals.get(disciplineRowId) ?? "") : "";
    const systemTotal = systemGroupKey ? (systemTotals.get(systemGroupKey) ?? "") : "";
    const disciplineIncidenceRatio = computeProjectIncidenceRatio(baseUnits, disciplineTotal);
    const systemIncidenceRatio = computeProjectIncidenceRatio(baseUnits, systemTotal);
    const realProgress = sanitizeRealProgress(deliverable.realProgress);
    const realProgressText = formatPercent(realProgress, 2);
    const incidenceText = incidenceRatio === null ? "" : formatPercent(incidenceRatio * 100, 2);
    const disciplineIncidenceText = disciplineIncidenceRatio === null ? "" : formatPercent(disciplineIncidenceRatio * 100, 2);
    const systemIncidenceText = systemIncidenceRatio === null ? "" : formatPercent(systemIncidenceRatio * 100, 2);
    const progressSnapshot = buildProgressSnapshot(
      sanitizeDateInput(deliverable.startDate || ""),
      sanitizeDateInput(deliverable.endDate || "")
    );
    const programmedProjectText = formatPercent(
      computeWeightedProjectProgress(incidenceRatio, progressSnapshot.percent),
      2
    );
    const programmedDisciplineText = formatPercent(
      computeWeightedProjectProgress(disciplineIncidenceRatio, progressSnapshot.percent),
      2
    );
    const programmedSystemText = formatPercent(
      computeWeightedProjectProgress(systemIncidenceRatio, progressSnapshot.percent),
      2
    );
    const realProjectText = formatPercent(
      computeWeightedProjectProgress(incidenceRatio, realProgress),
      2
    );
    const realDisciplineText = formatPercent(
      computeWeightedProjectProgress(disciplineIncidenceRatio, realProgress),
      2
    );
    const realSystemText = formatPercent(
      computeWeightedProjectProgress(systemIncidenceRatio, realProgress),
      2
    );
    const baseText = normalizeLookup(
      `${deliverable.ref} ${deliverable.code} ${deliverable.startDate || ""} ${deliverable.endDate || ""} ${formatNumberForInput(baseUnits)} ${incidenceText} ${disciplineIncidenceText} ${systemIncidenceText} ${progressSnapshot.label} ${programmedProjectText} ${programmedDisciplineText} ${programmedSystemText} ${realProgressText} ${realProjectText} ${realDisciplineText} ${realSystemText}`
    );
    if (baseText.includes(normalizedQuery)) {
      return true;
    }

    return project.fields.some((field) => {
      const rowId = trimOrFallback(deliverable.rowRefs?.[field.id], "");
      const row = getFieldRowById(field, rowId);
      if (!row) {
        return false;
      }
      const rowText = normalizeLookup(`${field.label} ${row.name} ${row.code}`);
      return rowText.includes(normalizedQuery);
    });
  });
}

function visibleIndexLabel(allDeliverables, deliverable, fallbackIndex) {
  const realIndex = allDeliverables.findIndex((item) => item.id === deliverable.id);
  if (realIndex < 0) {
    return String(fallbackIndex + 1);
  }
  return String(allDeliverables.length - realIndex);
}

function syncDraftValues(project) {
  const draft = getDraftValues(project.id);
  const validFieldIds = new Set(project.fields.map((field) => field.id));
  validFieldIds.add(DRAFT_META_START);
  validFieldIds.add(DRAFT_META_END);
  validFieldIds.add(DRAFT_META_BASE_UNITS);
  validFieldIds.add(DRAFT_META_REAL_PROGRESS);

  Object.keys(draft).forEach((fieldId) => {
    if (!validFieldIds.has(fieldId)) {
      delete draft[fieldId];
    }
  });

  project.fields.forEach((field) => {
    if (!(field.id in draft)) {
      draft[field.id] = "";
      return;
    }

    const selectedId = trimOrFallback(draft[field.id], "");
    if (!selectedId) {
      return;
    }

    const exists = field.rows.some((row) => row.id === selectedId && hasRowContent(row));
    if (!exists) {
      draft[field.id] = "";
    }
  });

  reconcileSystemSelectionForRowRefs(project, draft);

  const startDate = sanitizeDateInput(draft[DRAFT_META_START]);
  const endDate = sanitizeDateInput(draft[DRAFT_META_END]);
  const baseUnits = sanitizeBaseUnits(draft[DRAFT_META_BASE_UNITS]);
  const realProgress = sanitizeRealProgress(draft[DRAFT_META_REAL_PROGRESS]);

  draft[DRAFT_META_START] = startDate;
  draft[DRAFT_META_END] = endDate;
  draft[DRAFT_META_BASE_UNITS] = formatNumberForInput(baseUnits);
  draft[DRAFT_META_REAL_PROGRESS] = formatNumberForInput(realProgress);
}

function primeDraftSelections(project) {
  const draft = getDraftValues(project.id);
  project.fields.forEach((field) => {
    const current = trimOrFallback(draft[field.id], "");
    if (current) {
      return;
    }

    const first = getFieldOptions(field, { project, rowRefs: draft })[0];
    draft[field.id] = first ? first.id : "";
  });

  reconcileSystemSelectionForRowRefs(project, draft);
}

function getDraftValues(projectId) {
  if (!draftValuesByProject[projectId]) {
    draftValuesByProject[projectId] = {};
  }
  return draftValuesByProject[projectId];
}

function clearDraft(projectId, fields) {
  draftValuesByProject[projectId] = {};
  fields.forEach((field) => {
    draftValuesByProject[projectId][field.id] = "";
  });
  draftValuesByProject[projectId][DRAFT_META_START] = "";
  draftValuesByProject[projectId][DRAFT_META_END] = "";
  draftValuesByProject[projectId][DRAFT_META_BASE_UNITS] = "";
  draftValuesByProject[projectId][DRAFT_META_REAL_PROGRESS] = "";
}

function getDisciplineSystemFields(fields) {
  const disciplineFieldId = findFieldIdByAlias(fields, ["disciplina"]);
  const systemFieldId = findFieldIdByAlias(fields, ["sistema"]);
  const disciplineField = getFieldById(fields, disciplineFieldId);
  const systemField = getFieldById(fields, systemFieldId);
  return { disciplineFieldId, systemFieldId, disciplineField, systemField };
}

function getSystemParentDisciplineRowId(systemRow) {
  if (!systemRow) {
    return "";
  }
  return trimOrFallback(systemRow.parentDisciplineRowId, "");
}

function isSystemRowAllowedForDiscipline(systemRow, disciplineRowId) {
  const parentDisciplineRowId = getSystemParentDisciplineRowId(systemRow);
  if (!parentDisciplineRowId) {
    return false;
  }
  if (!disciplineRowId) {
    return false;
  }
  return parentDisciplineRowId === disciplineRowId;
}

function getFieldOptions(field, context) {
  const baseOptions = field.rows.filter((row) => hasRowContent(row));
  const project = context?.project || null;
  const rowRefs = context?.rowRefs || null;
  if (!project || !rowRefs) {
    return baseOptions;
  }

  const pair = getDisciplineSystemFields(project.fields);
  if (!pair.systemFieldId || !pair.disciplineFieldId || field.id !== pair.systemFieldId) {
    return baseOptions;
  }

  const disciplineRowId = trimOrFallback(rowRefs[pair.disciplineFieldId], "");
  const validDisciplineIds = new Set(getFieldOptions(pair.disciplineField).map((row) => row.id));
  return baseOptions.filter((row) => {
    const parentDisciplineRowId = getSystemParentDisciplineRowId(row);
    if (!parentDisciplineRowId || !validDisciplineIds.has(parentDisciplineRowId)) {
      return false;
    }
    return isSystemRowAllowedForDiscipline(row, disciplineRowId);
  });
}

function reconcileSystemSelectionForRowRefs(project, rowRefs) {
  if (!project || !rowRefs) {
    return false;
  }

  const pair = getDisciplineSystemFields(project.fields);
  if (!pair.systemField || !pair.systemFieldId || !pair.disciplineFieldId) {
    return false;
  }

  const currentSystemRowId = trimOrFallback(rowRefs[pair.systemFieldId], "");
  if (!currentSystemRowId) {
    return false;
  }

  const allowedOptions = getFieldOptions(pair.systemField, { project, rowRefs });
  const isValid = allowedOptions.some((row) => row.id === currentSystemRowId);
  if (isValid) {
    return false;
  }

  rowRefs[pair.systemFieldId] = "";
  return true;
}

function reconcileAllDeliverableSystemSelections(project) {
  if (!project || !Array.isArray(project.deliverables)) {
    return 0;
  }

  let adjusted = 0;
  project.deliverables.forEach((deliverable) => {
    if (!deliverable?.rowRefs) {
      return;
    }
    const changed = reconcileSystemSelectionForRowRefs(project, deliverable.rowRefs);
    if (!changed) {
      return;
    }
    applyDeliverableRows(project, deliverable, false);
    adjusted += 1;
  });
  return adjusted;
}

function getFieldRowById(field, rowId) {
  if (!rowId) {
    return null;
  }
  return field.rows.find((row) => row.id === rowId) || null;
}

function inferRowForField(field, value, code) {
  const normalizedValue = normalizeLookup(value);
  const normalizedCode = normalizeLookup(code);

  if (normalizedCode) {
    const byCode = field.rows.find((row) => normalizeLookup(row.code) === normalizedCode);
    if (byCode) {
      return byCode;
    }
  }

  if (normalizedValue) {
    const byName = field.rows.find((row) => normalizeLookup(row.name) === normalizedValue);
    if (byName) {
      return byName;
    }
  }

  return null;
}

function ensureFieldRow(field, rowIndex) {
  while (field.rows.length <= rowIndex) {
    field.rows.push(createFieldRow("", ""));
  }
  return field.rows[rowIndex];
}

function computeVisualRowCount(fields) {
  let count = MIN_MATRIX_ROWS;
  fields.forEach((field) => {
    const lastFilled = findLastFilledRowIndex(field.rows);
    count = Math.max(count, field.rows.length, lastFilled + 3);
  });
  return count;
}

function findLastFilledRowIndex(rows) {
  for (let i = rows.length - 1; i >= 0; i -= 1) {
    if (hasRowContent(rows[i])) {
      return i;
    }
  }
  return -1;
}

function hasRowContent(row) {
  const name = trimOrFallback(row?.name || "", "");
  const code = trimOrFallback(row?.code || "", "");
  return !!name || !!code;
}

function loadState() {
  const parsed = storageAdapter.getJson(STORAGE_KEY, null);
  if (!parsed) {
    return createDefaultState();
  }

  return normalizeState(parsed);
}

function scheduleSaveState() {
  if (pendingSaveTimer !== null) {
    clearTimeout(pendingSaveTimer);
  }
  setSyncIndicator("Cambios pendientes", "pending");
  pendingSaveTimer = window.setTimeout(() => {
    pendingSaveTimer = null;
    saveState(true);
  }, SAVE_DEBOUNCE_MS);
}

function flushPendingSave() {
  if (pendingSaveTimer === null) {
    return;
  }

  clearTimeout(pendingSaveTimer);
  pendingSaveTimer = null;
  saveState(true);
}

function saveState(forcePersist) {
  const shouldPersist = !!forcePersist;
  if (!shouldPersist) {
    scheduleSaveState();
    return;
  }

  try {
    applyPendingRealAdvanceLogs();
    applyPendingPackageRealAdvanceLogs();
    applyPendingReviewControlRealAdvanceLogs();
    if (!storageAdapter.setJson(STORAGE_KEY, state)) {
      throw new Error("storage_write_failed");
    }
    const stamp = new Date().toLocaleTimeString();
    setSyncIndicator(`Guardado ${stamp}`, "");
  } catch {
    setSyncIndicator("Error al guardar", "error");
  }
}

function createDefaultState() {
  return {
    activeProjectId: "proyecto_norte",
    projects: [
      createDefaultProject("proyecto_norte", "Proyecto Norte"),
      createDefaultProject("proyecto_sur", "Proyecto Sur")
    ]
  };
}

function normalizeState(raw) {
  if (!raw || !Array.isArray(raw.projects) || raw.projects.length === 0) {
    return createDefaultState();
  }

  const projects = raw.projects.map((project, index) => normalizeProject(project, index));
  const normalized = {
    projects,
    activeProjectId: typeof raw.activeProjectId === "string" ? raw.activeProjectId : projects[0].id
  };

  const hasActive = projects.some((project) => project.id === normalized.activeProjectId);
  if (!hasActive) {
    normalized.activeProjectId = projects[0].id;
  }

  return normalized;
}

function normalizeProject(rawProject, index) {
  const fallbackName = `Proyecto ${index + 1}`;
  const id = typeof rawProject?.id === "string" && rawProject.id.trim()
    ? rawProject.id.trim()
    : `proyecto_${index + 1}`;
  const name = trimOrFallback(rawProject?.name, fallbackName);
  const title = trimOrFallback(rawProject?.title, `MIDP - ${name}`);
  const codeSeparator = sanitizeSeparator(rawProject?.codeSeparator || "-");

  const sourceFields = Array.isArray(rawProject?.fields) && rawProject.fields.length > 0
    ? rawProject.fields.slice(0, MAX_FIELDS)
    : DEFAULT_FIELD_TEMPLATES;
  const normalizedSourceFields = sourceFields.map((field, fieldIndex) => normalizeField(field, fieldIndex));
  const fields = ensureRequiredFixedFields(normalizedSourceFields);
  applyDefaultSystemDisciplineMapping(fields);
  const nomenclatureOrder = normalizeNomenclatureOrder(rawProject?.nomenclatureOrder, fields);
  const deliverables = Array.isArray(rawProject?.deliverables)
    ? rawProject.deliverables.map((item, itemIndex) => normalizeDeliverable(item, fields, itemIndex))
    : [];
  const reviewMilestones = normalizeReviewMilestones(
    rawProject?.reviewMilestones ?? rawProject?.hitosFlujoRevision ?? rawProject?.hitosRevision);
  const reviewControls = normalizeReviewControls(
    rawProject?.reviewControls ?? rawProject?.controlFlujosRevision ?? rawProject?.flujosRevisionControl);
  const packageControls = normalizePackageControls(rawProject?.packageControls ?? rawProject?.controlPaquetes);
  const biConfig = normalizeBiConfig(rawProject?.biConfig ?? rawProject?.dechiniBIConfig);
  const biWidgets = normalizeBiWidgets(rawProject?.biWidgets ?? rawProject?.dechiniBIWidgets);
  const quickSightConfig = normalizeQuickSightConfig(
    rawProject?.quickSightConfig
    ?? rawProject?.dechiniQuickSightConfig
    ?? rawProject?.dechiniQuicksigthConfig);
  const quickSightVisuals = normalizeQuickSightVisuals(
    rawProject?.quickSightVisuals
    ?? rawProject?.dechiniQuickSightVisuals
    ?? rawProject?.dechiniQuicksigthVisuals);
  const demoSeeded = !!rawProject?.demoSeeded;
  const progressControlMode = normalizeProgressControlMode(
    rawProject?.progressControlMode ?? rawProject?.controlAvance ?? rawProject?.modoControlAvance);
  const showMidpAdvanceDetails = rawProject?.showMidpAdvanceDetails !== false;

  return {
    id,
    name,
    title,
    codeSeparator,
    fields,
    nomenclatureOrder,
    deliverables,
    reviewMilestones,
    reviewControls,
    packageControls,
    biConfig,
    biWidgets,
    quickSightConfig,
    quickSightVisuals,
    demoSeeded,
    progressControlMode,
    showMidpAdvanceDetails
  };
}

function normalizeField(rawField, index) {
  if (typeof rawField === "string") {
    return createField(rawField, index < 3, index < 3, []);
  }

  const id = typeof rawField?.id === "string" && rawField.id.trim()
    ? rawField.id.trim()
    : uid();
  const label = trimOrFallback(rawField?.label, `Campo ${index + 1}`);
  const includeInCode = !!rawField?.includeInCode;
  const locked = !!rawField?.locked;
  const rows = normalizeFieldRows(rawField?.rows);

  return { id, label, includeInCode, locked, rows };
}

function ensureRequiredFixedFields(fields) {
  if (!Array.isArray(fields)) {
    return [];
  }

  const nextFields = fields.slice();
  ensureFixedFieldTemplate(nextFields, "Paquete", "Sistema");
  return nextFields;
}

function ensureFixedFieldTemplate(fields, templateLabel, insertAfterLabel) {
  const template = DEFAULT_FIELD_TEMPLATES.find(
    (item) => normalizeLookup(item.label) === normalizeLookup(templateLabel));
  if (!template) {
    return;
  }

  const existingIndex = findFieldIndexByLabel(fields, template.label);
  if (existingIndex >= 0) {
    const existingField = fields[existingIndex];
    existingField.locked = true;
    moveFieldAfterLabel(fields, existingIndex, insertAfterLabel);
    return;
  }

  if (fields.length >= MAX_FIELDS) {
    return;
  }

  const newField = createField(template.label, template.includeInCode, true, template.rows);
  const insertAfterIndex = findFieldIndexByLabel(fields, insertAfterLabel);
  if (insertAfterIndex >= 0) {
    fields.splice(insertAfterIndex + 1, 0, newField);
  } else {
    fields.push(newField);
  }
}

function moveFieldAfterLabel(fields, fieldIndex, targetAfterLabel) {
  if (!Array.isArray(fields) || fieldIndex < 0 || fieldIndex >= fields.length) {
    return;
  }

  const targetIndex = findFieldIndexByLabel(fields, targetAfterLabel);
  if (targetIndex < 0 || targetIndex === fieldIndex) {
    return;
  }

  const [movedField] = fields.splice(fieldIndex, 1);
  const adjustedTargetIndex = findFieldIndexByLabel(fields, targetAfterLabel);
  if (adjustedTargetIndex < 0) {
    fields.push(movedField);
    return;
  }

  fields.splice(adjustedTargetIndex + 1, 0, movedField);
}

function findFieldIndexByLabel(fields, label) {
  if (!Array.isArray(fields)) {
    return -1;
  }

  const labelKey = normalizeLookup(label || "");
  if (!labelKey) {
    return -1;
  }

  return fields.findIndex((field) => normalizeLookup(field?.label || "") === labelKey);
}

function normalizeNomenclatureOrder(rawOrder, fields) {
  const validFieldIds = new Set(fields.map((field) => field.id));
  const result = [];

  if (Array.isArray(rawOrder)) {
    rawOrder.forEach((fieldId) => {
      if (typeof fieldId !== "string") {
        return;
      }
      if (!validFieldIds.has(fieldId)) {
        return;
      }
      if (result.includes(fieldId)) {
        return;
      }

      const field = fields.find((item) => item.id === fieldId);
      if (field?.includeInCode) {
        result.push(fieldId);
      }
    });
  }

  fields.forEach((field) => {
    if (!field.includeInCode) {
      return;
    }
    if (!result.includes(field.id)) {
      result.push(field.id);
    }
  });

  return result;
}

function normalizeFieldRows(rawRows) {
  if (!Array.isArray(rawRows)) {
    return [];
  }

  return rawRows.map((row) => normalizeFieldRow(row));
}

function normalizeFieldRow(rawRow) {
  if (typeof rawRow === "string") {
    return createFieldRow(rawRow, "");
  }

  const id = typeof rawRow?.id === "string" && rawRow.id.trim()
    ? rawRow.id.trim()
    : uid();
  const name = trimOrFallback(rawRow?.name, "");
  const code = trimOrFallback(rawRow?.code, "");
  const parentDisciplineRowId = trimOrFallback(
    rawRow?.parentDisciplineRowId ?? rawRow?.disciplineRowId ?? rawRow?.parentRowId,
    "");
  return { id, name, code, parentDisciplineRowId };
}

function normalizeBiSource(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "deliverable" || token === "entregable") {
    return "deliverable";
  }
  if (token === "package" || token === "paquete") {
    return "package";
  }
  if (token === "review-control" || token === "reviewcontrol" || token === "flujo" || token === "revision") {
    return "review-control";
  }
  return "all";
}

function normalizeBiUiMode(value) {
  return trimOrFallback(value, "").toLowerCase() === "advanced" ? "advanced" : "basic";
}

function normalizeBiPerformanceMode(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "quality" || token === "balanced" || token === "turbo") {
    return token;
  }
  return "balanced";
}

function normalizeBiGroupBy(value) {
  const raw = trimOrFallback(value, "");
  if (/^field:/i.test(raw)) {
    const fieldId = trimOrFallback(raw.slice(6), "");
    if (fieldId) {
      return `field:${fieldId}`;
    }
  }
  const token = raw.toLowerCase();
  const allowed = new Set([
    "proyecto",
    "disciplina",
    "sistema",
    "paquete",
    "creador",
    "fase",
    "sector",
    "nivel",
    "tipo",
    "hito",
    "fuente",
    "mes_inicio",
    "mes_fin"
  ]);
  return allowed.has(token) ? token : "disciplina";
}

function normalizeBiMetric(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  const allowed = new Set([
    "count",
    "baseunits",
    "baseavg",
    "basemax",
    "basemin",
    "realavg",
    "realmax",
    "realmin",
    "programmedavg",
    "programmedmax",
    "programmedmin",
    "weightedreal",
    "weightedprogrammed",
    "weightedgap",
    "invaliddates"
  ]);
  return allowed.has(token) ? token : "baseunits";
}

function normalizeBiSortMode(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  const allowed = new Set([
    "value_desc",
    "value_asc",
    "label_asc",
    "label_desc"
  ]);
  return allowed.has(token) ? token : "value_desc";
}

function getQuickSightVisualOptionalMetric(visual, fallbackConfig = null) {
  const fallback = normalizeQuickSightConfig(fallbackConfig || {});
  const chartType = normalizeBiChartType(
    visual?.chartType
    || fallback.chartType
    || "bar"
  );
  const capabilities = getBiChartCapabilities(chartType);
  if (!capabilities.supportsOptionalMetric) {
    return "";
  }
  if (!visual || typeof visual !== "object") {
    return normalizeBiOptionalMetric(fallback.optionalMetric || "");
  }
  const visualOptionalMetric = getOwnObjectValue(
    visual,
    "optionalMetric",
    visual.dataRoles?.optionalMetrics?.[0] ?? fallback.optionalMetric ?? ""
  );
  return normalizeBiOptionalMetric(
    visualOptionalMetric || ""
  );
}

function getQuickSightRoleUiLabels(chartType) {
  const type = normalizeBiChartType(chartType || "bar");
  if (BI_HORIZONTAL_BAR_TYPES.has(type)) {
    return {
      dimension: "EJE Y",
      metric: "EJE X / VALOR",
      optionalMetric: "EJE X",
      breakdown: "GRUPO/COLOR"
    };
  }
  if (type === "scatter" || type === "bubble") {
    return {
      dimension: "ETIQUETA",
      metric: "EJE Y",
      optionalMetric: "EJE X",
      breakdown: "GRUPO/COLOR"
    };
  }
  if (BI_BAR_FAMILY_TYPES.has(type) || new Set(["line", "timeseries", "area", "combo", "waterfall", "pareto"]).has(type)) {
    return {
      dimension: "EJE X",
      metric: "EJE Y / VALOR",
      optionalMetric: "EJE X",
      breakdown: "GRUPO/COLOR"
    };
  }
  return {
    dimension: "CATEGORIA",
    metric: "VALOR",
    optionalMetric: "EJE X",
    breakdown: "GRUPO/COLOR"
  };
}

function createDefaultQuickSightConfig() {
  return {
    source: "all",
    groupBy: "disciplina",
    breakdownDimension: "",
    metric: "count",
    optionalMetric: "",
    dateDimension: "",
    chartType: "bar",
    topN: 12,
    sortMode: "value_desc",
    fieldsSearch: "",
    interactionFilters: [],
    drillThroughSelection: null,
    canvasWidth: QUICKSIGHT_CANVAS_DEFAULT_WIDTH,
    canvasHeight: QUICKSIGHT_CANVAS_DEFAULT_HEIGHT,
    canvasZoom: QUICKSIGHT_CANVAS_ZOOM_DEFAULT
  };
}

function createDefaultQuickSightInteractionSettings() {
  return {
    clickAction: "cross_filter",
    filterBy: "group",
    scope: "all",
    appendWithModifier: true,
    drillPath: []
  };
}

function createDefaultQuickSightDisplaySettings() {
  return {
    showTitle: true,
    showSubtitle: true,
    subtitle: "",
    altText: "",
    titleFontFamily: "Segoe UI",
    titleFontSize: 12,
    titleColor: "#1f2f44",
    subtitleFontFamily: "Segoe UI",
    subtitleFontSize: 10,
    subtitleColor: "#6c819d",
    showBackground: true,
    backgroundColor: "#ffffff",
    backgroundOpacity: 100,
    showBorder: true,
    borderColor: "#b7c9df",
    borderOpacity: 100,
    borderWidth: 1,
    selectionColor: "#1495bf",
    selectionOpacity: 35,
    loadAnimation: true,
    padding: 8
  };
}

function normalizeQuickSightDisplaySettings(rawSettings) {
  const defaults = createDefaultQuickSightDisplaySettings();
  const source = rawSettings && typeof rawSettings === "object" && !Array.isArray(rawSettings)
    ? rawSettings
    : {};
  return {
    showTitle: normalizeBiToggle(source.showTitle, defaults.showTitle),
    showSubtitle: normalizeBiToggle(source.showSubtitle, defaults.showSubtitle),
    subtitle: trimOrFallback(source.subtitle || source.subtitulo, defaults.subtitle).slice(0, 140),
    altText: trimOrFallback(source.altText || source.textoAlternativo, defaults.altText).slice(0, 260),
    titleFontFamily: sanitizeBiFontFamily(source.titleFontFamily || source.fuenteTitulo || "", defaults.titleFontFamily),
    titleFontSize: sanitizeBiInteger(source.titleFontSize ?? source.tamanoTitulo, defaults.titleFontSize, 10, 24),
    titleColor: normalizeBiColorHex(source.titleColor || source.colorTitulo, defaults.titleColor),
    subtitleFontFamily: sanitizeBiFontFamily(source.subtitleFontFamily || source.fuenteSubtitulo || "", defaults.subtitleFontFamily),
    subtitleFontSize: sanitizeBiInteger(source.subtitleFontSize ?? source.tamanoSubtitulo, defaults.subtitleFontSize, 8, 18),
    subtitleColor: normalizeBiColorHex(source.subtitleColor || source.colorSubtitulo, defaults.subtitleColor),
    showBackground: normalizeBiToggle(source.showBackground ?? source.fondo, defaults.showBackground),
    backgroundColor: normalizeBiColorHex(source.backgroundColor || source.colorFondo, defaults.backgroundColor),
    backgroundOpacity: sanitizeBiInteger(source.backgroundOpacity ?? source.opacidadFondo, defaults.backgroundOpacity, 0, 100),
    showBorder: normalizeBiToggle(source.showBorder ?? source.borde, defaults.showBorder),
    borderColor: normalizeBiColorHex(source.borderColor || source.colorBorde, defaults.borderColor),
    borderOpacity: sanitizeBiInteger(source.borderOpacity ?? source.opacidadBorde, defaults.borderOpacity, 0, 100),
    borderWidth: sanitizeBiInteger(source.borderWidth ?? source.grosorBorde, defaults.borderWidth, 0, 6),
    selectionColor: normalizeBiColorHex(source.selectionColor || source.colorSeleccion, defaults.selectionColor),
    selectionOpacity: sanitizeBiInteger(source.selectionOpacity ?? source.opacidadSeleccion, defaults.selectionOpacity, 0, 100),
    loadAnimation: normalizeBiToggle(source.loadAnimation ?? source.animacionCarga, defaults.loadAnimation),
    padding: sanitizeBiInteger(source.padding ?? source.relleno, defaults.padding, 0, 24)
  };
}

function normalizeQuickSightConfig(rawConfig) {
  const base = rawConfig || {};
  const defaults = createDefaultQuickSightConfig();
  return {
    source: normalizeBiSource(base.source || base.fuente || defaults.source),
    groupBy: normalizeBiGroupBy(base.groupBy || base.dimension || defaults.groupBy),
    breakdownDimension: normalizeBiOptionalGroupBy(base.breakdownDimension || base.grupoColor || base.dimensionColor || defaults.breakdownDimension),
    metric: normalizeBiMetric(base.metric || base.metrica || defaults.metric),
    optionalMetric: normalizeBiOptionalMetric(base.optionalMetric || base.metricaOpcional || defaults.optionalMetric),
    dateDimension: normalizeBiDateRole(base.dateDimension || base.dimensionFecha || defaults.dateDimension),
    chartType: normalizeBiChartType(base.chartType || base.chart || defaults.chartType),
    topN: sanitizeBiTopN(base.topN ?? base.top ?? defaults.topN),
    sortMode: normalizeBiSortMode(base.sortMode || base.orden || defaults.sortMode),
    fieldsSearch: trimOrFallback(base.fieldsSearch || base.busquedaCampos || "", "").slice(0, 80),
    interactionFilters: normalizeQuickSightInteractionFilters(base.interactionFilters || base.filtrosInteraccion || []),
    drillThroughSelection: normalizeQuickSightDrillThroughSelection(base.drillThroughSelection || base.detalleDrill || null),
    canvasWidth: sanitizeBiCanvasDimension(
      base.canvasWidth ?? base.anchoPizarra ?? defaults.canvasWidth,
      defaults.canvasWidth,
      QUICKSIGHT_CANVAS_MIN_WIDTH,
      BI_CANVAS_SURFACE_MAX_EDIT_WIDTH
    ),
    canvasHeight: sanitizeBiCanvasDimension(
      base.canvasHeight ?? base.altoPizarra ?? defaults.canvasHeight,
      defaults.canvasHeight,
      QUICKSIGHT_CANVAS_MIN_HEIGHT,
      BI_CANVAS_SURFACE_MAX_EDIT_HEIGHT
    ),
    canvasZoom: sanitizeQuickSightCanvasZoom(base.canvasZoom ?? base.zoomPizarra ?? defaults.canvasZoom, defaults.canvasZoom)
  };
}

function normalizeQuickSightInteractionSettings(rawSettings, visualLike = null) {
  const defaults = createDefaultQuickSightInteractionSettings();
  const source = rawSettings && typeof rawSettings === "object" && !Array.isArray(rawSettings)
    ? rawSettings
    : {};
  const clickActionToken = trimOrFallback(
    source.clickAction || source.action || source.click || "",
    defaults.clickAction
  ).toLowerCase();
  const filterByToken = trimOrFallback(
    source.filterBy || source.filtrarPor || "",
    defaults.filterBy
  ).toLowerCase();
  const resolvedBreakdown = normalizeBiOptionalGroupBy(
    getOwnObjectValue(
      visualLike,
      "breakdownDimension",
      visualLike?.dataRoles?.breakdownDimension ?? ""
    ) || ""
  );
  const safeFilterBy = filterByToken === "breakdown" && resolvedBreakdown
    ? "breakdown"
    : "group";
  const drillPath = normalizeQuickSightDrillPath(
    source.drillPath || source.rutaDrill || source.drillFields || source.nivelesDrill || [],
    visualLike?.groupBy || visualLike?.dataRoles?.dimensions?.[0] || "disciplina"
  );
  return {
    clickAction: clickActionToken === "none"
      ? "none"
      : (new Set(["drill", "drill_down", "drilldown"]).has(clickActionToken)
        ? "drill_down"
        : (new Set(["drill_through", "drillthrough", "detail"]).has(clickActionToken) ? "drill_through" : "cross_filter")),
    filterBy: safeFilterBy,
    scope: normalizeBiCrossFilterScope(source.scope || source.alcance || defaults.scope),
    appendWithModifier: normalizeBiToggle(source.appendWithModifier ?? source.multiSelect ?? defaults.appendWithModifier, defaults.appendWithModifier),
    drillPath
  };
}

function normalizeQuickSightDrillPath(rawPath, baseGroupBy = "") {
  const source = Array.isArray(rawPath)
    ? rawPath
    : (typeof rawPath === "string" ? rawPath.split(/[>,|]/g) : []);
  const safeBase = normalizeBiOptionalGroupBy(baseGroupBy);
  const seen = new Set(safeBase ? [safeBase] : []);
  const result = [];
  source.forEach((item) => {
    if (result.length >= QUICKSIGHT_DRILL_PATH_SLOTS) {
      return;
    }
    const normalized = normalizeBiOptionalGroupBy(item);
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    result.push(normalized);
  });
  return result;
}

function createDefaultQuickSightDrillState() {
  return {
    filters: []
  };
}

function normalizeQuickSightDrillFilterEntry(rawEntry, fallbackGroupBy = "") {
  const source = rawEntry && typeof rawEntry === "object" && !Array.isArray(rawEntry)
    ? rawEntry
    : {};
  const groupBy = normalizeBiOptionalGroupBy(
    source.groupBy
    || source.dimension
    || source.campo
    || fallbackGroupBy
    || ""
  );
  const label = trimOrFallback(
    source.label
    || source.value
    || source.groupLabel
    || source.nombre
    || "",
    ""
  ).slice(0, 120);
  if (!groupBy || !label) {
    return null;
  }
  return {
    groupBy,
    label
  };
}

function getQuickSightDrillSequence(visualLike) {
  const safeVisual = visualLike && typeof visualLike === "object" ? visualLike : {};
  const baseGroupBy = normalizeBiGroupBy(
    safeVisual.groupBy
    || safeVisual.dataRoles?.dimensions?.[0]
    || "disciplina"
  );
  const interactions = normalizeQuickSightInteractionSettings(safeVisual.interactions, safeVisual);
  return [baseGroupBy, ...normalizeQuickSightDrillPath(interactions.drillPath, baseGroupBy)];
}

function normalizeQuickSightDrillState(rawState, visualLike = null) {
  const source = rawState && typeof rawState === "object" && !Array.isArray(rawState)
    ? rawState
    : {};
  const sequence = getQuickSightDrillSequence(visualLike);
  const rawFilters = Array.isArray(source.filters)
    ? source.filters
    : (Array.isArray(source.breadcrumbs) ? source.breadcrumbs : []);
  const filters = [];
  for (const item of rawFilters) {
    const expectedGroupBy = sequence[filters.length];
    if (!expectedGroupBy || filters.length >= (sequence.length - 1)) {
      break;
    }
    const normalized = normalizeQuickSightDrillFilterEntry(item, expectedGroupBy);
    if (!normalized || normalized.groupBy !== expectedGroupBy) {
      break;
    }
    filters.push(normalized);
  }
  return {
    filters
  };
}

function normalizeQuickSightDrillThroughSelection(rawSelection) {
  const source = rawSelection && typeof rawSelection === "object" && !Array.isArray(rawSelection)
    ? rawSelection
    : {};
  const visualId = trimOrFallback(source.visualId || source.visual || "", "");
  const groupBy = normalizeBiOptionalGroupBy(source.groupBy || source.dimension || source.campo || "");
  const label = trimOrFallback(source.label || source.value || source.valor || "", "").slice(0, 120);
  if (!visualId || !groupBy || !label) {
    return null;
  }
  return {
    visualId,
    groupBy,
    label,
    source: normalizeBiSource(source.source || source.fuente || "all"),
    openedAt: trimOrFallback(source.openedAt, new Date().toISOString())
  };
}

function ensureQuickSightVisualDrillState(visual) {
  if (!visual || typeof visual !== "object") {
    return;
  }
  visual.interactions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
  visual.drillState = normalizeQuickSightDrillState(visual.drillState, visual);
}

function getQuickSightVisualDrillMeta(visual, project = null) {
  const safeVisual = visual && typeof visual === "object" ? visual : {};
  if (visual && typeof visual === "object") {
    ensureQuickSightVisualDrillState(visual);
  }
  const interactions = normalizeQuickSightInteractionSettings(safeVisual.interactions, safeVisual);
  const sequence = getQuickSightDrillSequence(safeVisual);
  const storedState = normalizeQuickSightDrillState(safeVisual.drillState, safeVisual);
  const enabled = interactions.clickAction === "drill_down" && sequence.length > 1;
  const filters = enabled ? storedState.filters : [];
  const currentIndex = Math.max(0, Math.min(filters.length, Math.max(0, sequence.length - 1)));
  const currentGroupBy = sequence[currentIndex] || sequence[0] || "disciplina";
  return {
    enabled,
    filters,
    sequence,
    currentGroupBy,
    currentLevel: currentIndex + 1,
    totalLevels: sequence.length,
    canDrillDown: enabled && currentIndex < (sequence.length - 1),
    canDrillUp: enabled && filters.length > 0,
    sequenceText: sequence.map((groupBy) => getBiGroupLabel(groupBy, project)).join(" > "),
    breadcrumbText: filters.length > 0
      ? filters.map((item) => `${getBiGroupLabel(item.groupBy, project)} = ${item.label}`).join(" > ")
      : "Sin drill activo."
  };
}

function rowMatchesQuickSightDrillFilter(row, filter) {
  if (!row || !filter) {
    return false;
  }
  const rowValue = getBiRowGroupValue(row, filter.groupBy);
  return normalizeLookup(rowValue) === normalizeLookup(filter.label);
}

function getQuickSightPanelBaseRows(project) {
  const rows = queryQuickSightProjectRows(project);
  if (!currentSearchQuery) {
    return rows;
  }
  return rows.filter((row) => row.searchBlob.includes(currentSearchQuery));
}

function filterQuickSightRowsByDrillState(rows, visual) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const drillMeta = getQuickSightVisualDrillMeta(visual);
  if (!drillMeta.enabled || drillMeta.filters.length === 0) {
    return safeRows;
  }
  return safeRows.filter((row) => drillMeta.filters.every((filter) => rowMatchesQuickSightDrillFilter(row, filter)));
}

function buildQuickSightVisualSnapshot(project, visual, rows = null) {
  if (!project || !visual) {
    return buildBiWidgetSnapshot([], visual || {});
  }
  ensureQuickSightVisualDrillState(visual);
  const baseRows = Array.isArray(rows) ? rows : getQuickSightPanelBaseRows(project);
  const scopedRows = filterQuickSightRowsByDrillState(baseRows, visual);
  const drillMeta = getQuickSightVisualDrillMeta(visual, project);
  const currentGroupBy = drillMeta.currentGroupBy;
  const metric = normalizeBiMetric(visual.metric || visual.dataRoles?.metrics?.[0] || "count");
  const breakdownDimension = normalizeBiOptionalGroupBy(
    getOwnObjectValue(
      visual,
      "breakdownDimension",
      visual.dataRoles?.breakdownDimension ?? ""
    ) || ""
  );
  const effectiveVisual = {
    ...visual,
    groupBy: currentGroupBy,
    dataRoles: normalizeBiDataRoles({
      ...(visual.dataRoles && typeof visual.dataRoles === "object" ? visual.dataRoles : {}),
      dimensions: [currentGroupBy],
      metrics: [metric],
      breakdownDimension: breakdownDimension || null
    }, {
      groupBy: currentGroupBy,
      metric,
      chartType: visual.chartType || "bar"
    })
  };
  const snapshot = buildBiWidgetSnapshot(scopedRows, effectiveVisual);
  snapshot.quickSightDrill = drillMeta;
  return snapshot;
}

function getQuickSightEChartsAdapter() {
  if (typeof window === "undefined" || !window.MIDPEChartsAdapter || typeof window.MIDPEChartsAdapter !== "object") {
    return null;
  }
  return window.MIDPEChartsAdapter;
}

function disposeQuickSightChartSurface(surface) {
  const adapter = getQuickSightEChartsAdapter();
  if (!adapter || typeof adapter.disposeQuickSightChart !== "function" || !(surface instanceof HTMLElement)) {
    return false;
  }
  return adapter.disposeQuickSightChart(surface);
}

function disposeQuickSightChartsInContainer(container) {
  if (!(container instanceof HTMLElement)) {
    return;
  }
  container.querySelectorAll("[data-qs-echart-surface]").forEach((node) => {
    if (node instanceof HTMLElement) {
      disposeQuickSightChartSurface(node);
    }
  });
}

function hasCustomQuickSightLabelOffsets(labelOffsets) {
  const safeOffsets = normalizeBiLabelOffsets(labelOffsets);
  return Object.keys(safeOffsets).length > 0;
}

function normalizeQuickSightRenderEngine(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "echarts" || token === "canvas") {
    return token;
  }
  return "auto";
}

function getQuickSightRenderEngineMeta(visual) {
  const preference = normalizeQuickSightRenderEngine(visual?.renderEngine);
  const adapter = getQuickSightEChartsAdapter();
  const chartType = normalizeBiChartType(visual?.chartType || "bar");
  const supportsECharts = !!(adapter && typeof adapter.canRenderChartType === "function" && adapter.canRenderChartType(chartType));
  const customLabelOffsets = hasCustomQuickSightLabelOffsets(visual?.labelOffsets);
  const requiresECharts = QUICKSIGHT_ECHARTS_REQUIRED_TYPES.has(chartType);
  let resolved = "canvas";
  let note = "Render clasico canvas.";
  if (requiresECharts && supportsECharts) {
    resolved = "echarts";
    note = "Este tipo requiere ECharts en QuickSight.";
    return { preference, resolved, supportsECharts, customLabelOffsets, note };
  }
  if (preference === "canvas") {
    return { preference, resolved, supportsECharts, customLabelOffsets, note };
  }
  if (supportsECharts && !customLabelOffsets) {
    resolved = "echarts";
    note = "Render moderno ECharts activo.";
    return { preference, resolved, supportsECharts, customLabelOffsets, note };
  }
  if (!supportsECharts) {
    note = "Este tipo aun no usa ECharts en QuickSight; se mantiene canvas.";
  } else if (customLabelOffsets) {
    note = "Este visual tiene labels movidos manualmente; se mantiene canvas.";
  }
  return { preference, resolved, supportsECharts, customLabelOffsets, note };
}

function getQuickSightChartCapabilities(visualLike = null, fallbackConfig = null) {
  const safeVisual = visualLike && typeof visualLike === "object" ? visualLike : {};
  const chartType = normalizeBiChartType(
    safeVisual.chartType
    || fallbackConfig?.chartType
    || "bar"
  );
  const base = getBiChartCapabilities(chartType);
  const renderMeta = getQuickSightRenderEngineMeta({
    chartType,
    renderEngine: safeVisual.renderEngine ?? fallbackConfig?.renderEngine,
    labelOffsets: safeVisual.labelOffsets
  });
  const echartsLegendTypes = new Set([
    "bar",
    "bar_horizontal",
    "bar_stacked",
    "line",
    "area",
    "combo",
    "pie",
    "donut",
    "funnel",
    "scatter",
    "bubble",
    "waterfall",
    "pareto"
  ]);
  return {
    ...base,
    renderMeta,
    supportsLegend: base.supportsLegend || (renderMeta.resolved === "echarts" && echartsLegendTypes.has(chartType))
  };
}

function shouldUseQuickSightECharts(visual, snapshot) {
  const renderMeta = getQuickSightRenderEngineMeta(visual);
  if (renderMeta.preference === "canvas") {
    return false;
  }
  if (renderMeta.resolved !== "echarts") {
    return false;
  }
  if (!snapshot || !Array.isArray(snapshot.rows) || snapshot.rows.length === 0) {
    return false;
  }
  return true;
}

function bindQuickSightEChartsInteractions(surface, chart, project, visual, snapshot) {
  if (!(surface instanceof HTMLElement) || !chart || !project || !visual || !snapshot || !Array.isArray(snapshot.rows)) {
    return;
  }
  chart.off("click");
  chart.on("click", (params) => {
    if (Date.now() < quickSightSuppressClickUntil) {
      return;
    }
    const rawIndex = Number(params?.data?.rawIndex);
    const rowIndex = Number.isInteger(rawIndex) && rawIndex >= 0
      ? rawIndex
      : (Number.isInteger(params?.dataIndex) ? params.dataIndex : -1);
    if (rowIndex < 0 || rowIndex >= snapshot.rows.length) {
      return;
    }
    applyQuickSightInteractionFromVisual(project, visual, { rows: snapshot.rows }, rowIndex, params?.event?.event || null);
  });
  const interactions = normalizeQuickSightInteractionSettings(visual.interactions, visual);
  surface.style.cursor = interactions.clickAction === "none" ? "default" : "pointer";
}

function normalizeQuickSightInteractionFilterEntry(rawEntry) {
  const base = normalizeBiCrossFilterEntry(rawEntry);
  if (!base) {
    return null;
  }
  return {
    ...base,
    scope: normalizeBiCrossFilterScope(rawEntry?.scope || rawEntry?.alcance || "all"),
    visualId: trimOrFallback(rawEntry?.visualId || "", "")
  };
}

function normalizeQuickSightInteractionFilters(rawFilters) {
  const source = Array.isArray(rawFilters)
    ? rawFilters
    : (rawFilters ? [rawFilters] : []);
  const result = [];
  const seen = new Set();
  source.forEach((item) => {
    const normalized = normalizeQuickSightInteractionFilterEntry(item);
    if (!normalized) {
      return;
    }
    const key = [
      normalized.groupBy,
      normalizeLookup(normalized.label),
      normalized.source,
      normalized.scope
    ].join("|");
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    result.push(normalized);
  });
  return result;
}

function syncQuickSightVisualDataRoles(visual, configLike = null) {
  if (!visual || typeof visual !== "object") {
    return;
  }
  const fallback = configLike && typeof configLike === "object" ? configLike : {};
  const groupBy = normalizeBiGroupBy(visual.groupBy || fallback.groupBy || "disciplina");
  const metric = normalizeBiMetric(visual.metric || fallback.metric || "count");
  const chartType = normalizeBiChartType(visual.chartType || fallback.chartType || "bar");
  const capabilities = getBiChartCapabilities(chartType);
  const breakdownDimension = normalizeBiOptionalGroupBy(
    getOwnObjectValue(
      visual,
      "breakdownDimension",
      visual.dataRoles?.breakdownDimension ?? fallback.breakdownDimension ?? ""
    ) || ""
  );
  const optionalMetric = normalizeBiOptionalMetric(
    getOwnObjectValue(
      visual,
      "optionalMetric",
      visual.dataRoles?.optionalMetrics?.[0] ?? fallback.optionalMetric ?? ""
    ) || ""
  );
  const dateDimension = normalizeBiDateRole(
    getOwnObjectValue(
      visual,
      "dateDimension",
      visual.dataRoles?.dateDimension ?? fallback.dateDimension ?? ""
    ) || ""
  );
  const safeOptionalMetric = capabilities.supportsOptionalMetric ? optionalMetric : "";
  const safeDateDimension = capabilities.supportsDateDimension
    ? (dateDimension || (capabilities.requiresDateDimension ? "startDate" : ""))
    : "";
  visual.groupBy = groupBy;
  visual.metric = metric;
  visual.chartType = chartType;
  visual.breakdownDimension = capabilities.supportsBreakdown ? breakdownDimension : "";
  visual.optionalMetric = safeOptionalMetric;
  visual.dateDimension = safeDateDimension;
  visual.dataRoles = normalizeBiDataRoles({
    ...(visual.dataRoles && typeof visual.dataRoles === "object" ? visual.dataRoles : {}),
    dimensions: [groupBy],
    metrics: [metric],
    optionalMetrics: safeOptionalMetric ? [safeOptionalMetric] : [],
    breakdownDimension: capabilities.supportsBreakdown && breakdownDimension ? breakdownDimension : null,
    dateDimension: safeDateDimension || null
  }, { groupBy, metric, chartType });
  if (!capabilities.supportsBreakdown || !breakdownDimension) {
    visual.dataRoles.breakdownDimension = null;
  }
  if (!capabilities.supportsOptionalMetric) {
    visual.dataRoles.optionalMetrics = [];
  }
  if (!capabilities.supportsDateDimension) {
    visual.dataRoles.dateDimension = null;
  }
}

function normalizeQuickSightVisual(rawVisual, index = 0) {
  const source = rawVisual || {};
  const chartType = normalizeBiChartType(source.chartType || source.chart || "bar");
  const chartTypeConfig = normalizeBiChartTypeConfigMap(source.chartTypeConfig ?? source.configTipoGrafico);
  if (!chartTypeConfig[chartType]) {
    chartTypeConfig[chartType] = normalizeBiChartTypeSpecificConfig({}, chartType);
  }
  const displaySource = {
    showTitle: source.showTitle,
    showSubtitle: source.showSubtitle,
    subtitle: source.subtitle ?? source.subtitulo,
    altText: source.altText ?? source.textoAlternativo,
    titleFontFamily: source.titleFontFamily ?? source.fuenteTitulo,
    titleFontSize: source.titleFontSize ?? source.tamanoTitulo,
    titleColor: source.titleColor ?? source.colorTitulo,
    subtitleFontFamily: source.subtitleFontFamily ?? source.fuenteSubtitulo,
    subtitleFontSize: source.subtitleFontSize ?? source.tamanoSubtitulo,
    subtitleColor: source.subtitleColor ?? source.colorSubtitulo,
    showBackground: source.showBackground,
    backgroundColor: source.backgroundColor,
    backgroundOpacity: source.backgroundOpacity,
    showBorder: source.showBorder,
    borderColor: source.borderColor,
    borderOpacity: source.borderOpacity,
    borderWidth: source.borderWidth,
    selectionColor: source.selectionColor,
    selectionOpacity: source.selectionOpacity,
    loadAnimation: source.loadAnimation,
    padding: source.padding,
    ...(source.display ?? source.estiloTarjeta ?? {})
  };
  const normalizedDisplay = normalizeQuickSightDisplaySettings(displaySource);
  const normalizedInteractions = normalizeQuickSightInteractionSettings(source.interactions ?? source.interaccion, source);
  const normalizedVisual = {
    id: typeof source.id === "string" && source.id.trim() ? source.id.trim() : uid(),
    name: trimOrFallback(source.name, `Visual ${index + 1}`).slice(0, 120),
    source: normalizeBiSource(source.source || "all"),
    groupBy: normalizeBiGroupBy(source.groupBy || source.dimension || "disciplina"),
    breakdownDimension: normalizeBiOptionalGroupBy(source.breakdownDimension || source.grupoColor || source.dimensionColor || source.dataRoles?.breakdownDimension || ""),
    metric: normalizeBiMetric(source.metric || source.metrica || "count"),
    optionalMetric: normalizeBiOptionalMetric(source.optionalMetric || source.metricaOpcional || source.dataRoles?.optionalMetrics?.[0] || ""),
    dateDimension: normalizeBiDateRole(source.dateDimension || source.dimensionFecha || source.dataRoles?.dateDimension || ""),
    chartType,
    renderEngine: normalizeQuickSightRenderEngine(source.renderEngine ?? source.motorRender ?? source.engineRender),
    topN: sanitizeBiTopN(source.topN ?? source.top ?? 12),
    sortMode: normalizeBiSortMode(source.sortMode || source.orden || "value_desc"),
    subtitle: normalizedDisplay.subtitle,
    altText: normalizedDisplay.altText,
    display: normalizedDisplay,
    visualSettings: applyBiChartTypeVisualDefaults(
      normalizeBiVisualSettings(source.visualSettings ?? source.visual ?? source.estiloVisual),
      chartType
    ),
    chartConfig: normalizeBiChartConfig(source.chartConfig ?? source.configGrafico, chartType),
    chartTypeConfig,
    interactions: normalizedInteractions,
    drillState: normalizeQuickSightDrillState(source.drillState ?? source.estadoDrill, {
      ...source,
      interactions: normalizedInteractions
    }),
    dataRoles: normalizeBiDataRoles(source.dataRoles ?? source.rolesDatos, {
      groupBy: source.groupBy || source.dimension || "disciplina",
      metric: source.metric || source.metrica || "count",
      chartType
    }),
    labelOffsets: normalizeBiLabelOffsets(source.labelOffsets ?? source.posicionLabels ?? source.desplazamientoLabels),
    labelLayoutV2: normalizeBiLabelLayoutV2(source.labelLayoutV2 ?? source.layoutLabels),
    polarLayout: normalizeBiCircularLayout(source.polarLayout ?? source.layoutPolar),
    createdAt: trimOrFallback(source.createdAt, new Date().toISOString()),
    layout: normalizeQuickSightVisualLayout(source.layout ?? source.posicion, index)
  };
  syncQuickSightVisualDataRoles(normalizedVisual);
  ensureQuickSightVisualDrillState(normalizedVisual);
  return normalizedVisual;
}

function normalizeQuickSightVisuals(rawVisuals) {
  if (!Array.isArray(rawVisuals)) {
    return [];
  }
  return rawVisuals.map((item, index) => normalizeQuickSightVisual(item, index));
}

function ensureQuickSightState(project) {
  if (!project) {
    return;
  }
  project.quickSightConfig = normalizeQuickSightConfig(project.quickSightConfig);
  project.quickSightVisuals = normalizeQuickSightVisuals(project.quickSightVisuals).map((visual, index) => ({
    ...visual,
    layout: clampQuickSightVisualLayoutToCanvas(
      normalizeQuickSightVisualLayout(visual.layout, index),
      project.quickSightConfig
    )
  }));
  const drillThroughSelection = normalizeQuickSightDrillThroughSelection(project.quickSightConfig.drillThroughSelection);
  project.quickSightConfig.drillThroughSelection = drillThroughSelection
    && project.quickSightVisuals.some((visual) => visual.id === drillThroughSelection.visualId)
      ? drillThroughSelection
      : null;
}

function createQuickSightVisualFromConfig(config, index = 0) {
  const safe = normalizeQuickSightConfig(config);
  const visual = normalizeQuickSightVisual({
    id: uid(),
    name: `${getBiGroupLabel(safe.groupBy)} | ${getBiMetricLabel(safe.metric)}`,
    source: safe.source,
    groupBy: safe.groupBy,
    breakdownDimension: safe.breakdownDimension,
    metric: safe.metric,
    chartType: safe.chartType,
    topN: safe.topN,
    sortMode: safe.sortMode
  }, index);
  syncQuickSightVisualDataRoles(visual, safe);
  return visual;
}

function createQuickSightConfigFromVisual(visual, fallbackConfig = null) {
  const fallback = normalizeQuickSightConfig(fallbackConfig || {});
  if (!visual || typeof visual !== "object") {
    return fallback;
  }
  return normalizeQuickSightConfig({
    ...fallback,
    source: visual.source,
    groupBy: visual.groupBy,
    breakdownDimension: visual.breakdownDimension,
    metric: visual.metric,
    optionalMetric: getQuickSightVisualOptionalMetric(visual, fallback),
    dateDimension: normalizeBiDateRole(visual.dataRoles?.dateDimension || fallback.dateDimension || ""),
    chartType: visual.chartType,
    topN: visual.topN,
    sortMode: visual.sortMode
  });
}

function normalizeBiChartType(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (BI_CHART_TYPES.has(token)) {
    return token;
  }
  return "bar";
}

function normalizeBiWidgetKind(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "text" || token === "texto") {
    return "text";
  }
  if (token === "image" || token === "imagen") {
    return "image";
  }
  return "chart";
}

function getBiWidgetMinSize(kindOrWidget) {
  const rawKind = typeof kindOrWidget === "string"
    ? kindOrWidget
    : kindOrWidget?.kind;
  const kind = normalizeBiWidgetKind(rawKind || "chart");
  if (kind === "text") {
    return { width: 120, height: 64 };
  }
  if (kind === "image") {
    return { width: 180, height: 120 };
  }
  return { width: BI_CANVAS_MIN_WIDTH, height: BI_CANVAS_MIN_HEIGHT };
}

function biStableHash(value) {
  const text = String(value || "");
  let hash = 5381;
  for (let index = 0; index < text.length; index += 1) {
    hash = ((hash << 5) + hash) + text.charCodeAt(index);
    hash |= 0;
  }
  return Math.abs(hash).toString(36);
}

function buildBiRowKey(source, groupBy, label) {
  const sourceToken = normalizeLookup(source || "all") || "all";
  const groupToken = normalizeLookup(groupBy || "group") || "group";
  const labelToken = normalizeLookup(label || "") || "sin_dato";
  return `${sourceToken}_${groupToken}_${labelToken}_${biStableHash(`${sourceToken}|${groupToken}|${label || ""}`)}`;
}

function buildBiDefaultDataRoles(widget) {
  const groupBy = normalizeBiGroupBy(widget?.groupBy || "disciplina");
  const metric = normalizeBiMetric(widget?.metric || "count");
  const chartType = normalizeBiChartType(widget?.chartType || "bar");
  const dateDimension = chartType === "timeseries" || chartType === "timeline"
    ? (groupBy === "mes_fin" ? "endDate" : "startDate")
    : null;
  return {
    dimensions: [groupBy],
    metrics: [metric],
    optionalMetrics: [],
    breakdownDimension: chartType === "pivot" || chartType === "sankey" ? "fuente" : null,
    dateDimension
  };
}

function normalizeBiDataRoles(rawRoles, widget) {
  const defaults = buildBiDefaultDataRoles(widget);
  const source = rawRoles && typeof rawRoles === "object" && !Array.isArray(rawRoles)
    ? rawRoles
    : {};
  const dimensions = Array.isArray(source.dimensions)
    ? source.dimensions.map((item) => normalizeBiGroupBy(item)).filter((item, index, arr) => !!item && arr.indexOf(item) === index)
    : [];
  const metrics = Array.isArray(source.metrics)
    ? source.metrics.map((item) => normalizeBiMetric(item)).filter((item, index, arr) => !!item && arr.indexOf(item) === index)
    : [];
  const optionalMetrics = Array.isArray(source.optionalMetrics)
    ? source.optionalMetrics.map((item) => normalizeBiMetric(item)).filter((item, index, arr) => !!item && arr.indexOf(item) === index)
    : [];
  const breakdownDimension = trimOrFallback(source.breakdownDimension, "")
    ? normalizeBiGroupBy(source.breakdownDimension)
    : defaults.breakdownDimension;
  const dateDimensionToken = trimOrFallback(source.dateDimension, "").toLowerCase();
  const dateDimension = new Set(["startdate", "enddate", "createdat"]).has(dateDimensionToken)
    ? (dateDimensionToken === "createdat" ? "createdAt" : (dateDimensionToken === "enddate" ? "endDate" : "startDate"))
    : defaults.dateDimension;
  return {
    dimensions: dimensions.length > 0 ? dimensions : defaults.dimensions.slice(),
    metrics: metrics.length > 0 ? metrics : defaults.metrics.slice(),
    optionalMetrics,
    breakdownDimension,
    dateDimension
  };
}

function createDefaultBiTextConfig() {
  return {
    fontFamily: "Segoe UI",
    fontSize: 56,
    fontWeight: 700,
    fontStyle: "normal",
    textAlign: "left",
    lineHeight: 1.12,
    color: "#1f2f44",
    padding: 6
  };
}

function normalizeBiTextConfig(rawConfig) {
  const defaults = createDefaultBiTextConfig();
  const source = rawConfig && typeof rawConfig === "object" && !Array.isArray(rawConfig)
    ? rawConfig
    : {};
  const fontFamily = trimOrFallback(source.fontFamily, defaults.fontFamily);
  const safeFamily = BI_ALLOWED_FONT_FAMILIES.has(fontFamily) ? fontFamily : defaults.fontFamily;
  const styleToken = trimOrFallback(source.fontStyle, "").toLowerCase();
  const alignToken = trimOrFallback(source.textAlign, "").toLowerCase();
  return {
    fontFamily: safeFamily,
    fontSize: sanitizeBiInteger(source.fontSize, defaults.fontSize, 12, 180),
    fontWeight: sanitizeBiInteger(source.fontWeight, defaults.fontWeight, 300, 900),
    fontStyle: styleToken === "italic" ? "italic" : "normal",
    textAlign: new Set(["left", "center", "right"]).has(alignToken) ? alignToken : defaults.textAlign,
    lineHeight: sanitizeBiDecimal(source.lineHeight, defaults.lineHeight, 0.9, 2.4),
    color: normalizeBiColorHex(source.color, defaults.color),
    padding: sanitizeBiInteger(source.padding, defaults.padding, 0, 48)
  };
}

function createDefaultBiImageConfig() {
  return {
    fit: "contain",
    opacity: 1,
    borderRadius: 8,
    frameColor: "#ffffff",
    background: "#ffffff"
  };
}

function normalizeBiImageConfig(rawConfig) {
  const defaults = createDefaultBiImageConfig();
  const source = rawConfig && typeof rawConfig === "object" && !Array.isArray(rawConfig)
    ? rawConfig
    : {};
  const fitToken = trimOrFallback(source.fit, "").toLowerCase();
  return {
    fit: new Set(["contain", "cover", "fill"]).has(fitToken) ? fitToken : defaults.fit,
    opacity: sanitizeBiDecimal(source.opacity, defaults.opacity, 0.05, 1),
    borderRadius: sanitizeBiInteger(source.borderRadius, defaults.borderRadius, 0, 40),
    frameColor: normalizeBiColorHex(source.frameColor, defaults.frameColor),
    background: normalizeBiColorHex(source.background, defaults.background)
  };
}

function getBiChartTypeDefaultSpecificConfig(chartType) {
  const type = normalizeBiChartType(chartType || "bar");
  if (type === "pie") {
    return { labelMinPercent: 4, startAngle: -90 };
  }
  if (type === "donut") {
    return { labelMinPercent: 4, innerRadiusRatio: 0.58, centerTitle: "Total" };
  }
  if (type === "treemap") {
    return { minTilePercent: 1, showValueText: true };
  }
  if (type === "funnel") {
    return { sortMode: "desc", minSegmentPercent: 2 };
  }
  if (type === "gauge") {
    return { minValue: 0, maxValue: 100, targetValue: 80 };
  }
  if (type === "table") {
    return { rowLimit: BI_WIDGET_TABLE_LIMIT, showIndex: true, compact: false };
  }
  if (type === "pivot") {
    return { rowLimit: BI_WIDGET_TABLE_LIMIT, showTotals: true, compact: false };
  }
  if (type === "scorecard") {
    return { showSource: true, showGroups: true, prefix: "", suffix: "" };
  }
  if (type === "waterfall") {
    return { showTotalBar: true, totalLabel: "Total" };
  }
  if (type === "scatter") {
    return { pointAlpha: 1, minPointRadius: 7 };
  }
  if (type === "bubble") {
    return { pointAlpha: 0.95, minBubbleRadius: 6, maxBubbleRadius: 40 };
  }
  if (type === "bullet") {
    return { targetValue: 80, lowBand: 40, midBand: 70 };
  }
  if (type === "boxplot") {
    return { showMedianValue: true, whiskerMultiplier: 1.5 };
  }
  if (type === "candlestick") {
    return { upColor: "#25a368", downColor: "#d7565b", wickWidth: 1 };
  }
  if (type === "sankey") {
    return { maxLinks: 12, minLinkPercent: 2, showValues: true };
  }
  if (type === "timeline") {
    return { rowLimit: 20, showDates: true, minBarPercent: 1 };
  }
  if (type === "radar") {
    return { maxScale: 100, showPointValues: true };
  }
  if (type === "pareto") {
    return { targetPercent: 80, showCumulativeLine: true };
  }
  if (type === "line" || type === "timeseries" || type === "area" || type === "combo" || type === "bar") {
    return { labelMinValue: 0 };
  }
  return {};
}

function normalizeBiChartTypeSpecificConfig(rawConfig, chartType) {
  const defaults = getBiChartTypeDefaultSpecificConfig(chartType);
  const source = rawConfig && typeof rawConfig === "object" && !Array.isArray(rawConfig)
    ? rawConfig
    : {};
  const type = normalizeBiChartType(chartType || "bar");
  const next = { ...defaults };
  if (Object.prototype.hasOwnProperty.call(defaults, "labelMinPercent")) {
    next.labelMinPercent = sanitizeBiDecimal(source.labelMinPercent, defaults.labelMinPercent, 0, 100);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "startAngle")) {
    next.startAngle = sanitizeBiDecimal(source.startAngle, defaults.startAngle, -360, 360);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "innerRadiusRatio")) {
    next.innerRadiusRatio = sanitizeBiDecimal(source.innerRadiusRatio, defaults.innerRadiusRatio, 0.35, 0.8);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "centerTitle")) {
    next.centerTitle = trimOrFallback(source.centerTitle, defaults.centerTitle).slice(0, 20);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minTilePercent")) {
    next.minTilePercent = sanitizeBiDecimal(source.minTilePercent, defaults.minTilePercent, 0, 25);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showValueText")) {
    next.showValueText = normalizeBiToggle(source.showValueText, defaults.showValueText);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "sortMode")) {
    const token = trimOrFallback(source.sortMode, defaults.sortMode).toLowerCase();
    next.sortMode = token === "asc" ? "asc" : "desc";
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minSegmentPercent")) {
    next.minSegmentPercent = sanitizeBiDecimal(source.minSegmentPercent, defaults.minSegmentPercent, 0, 60);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minValue")) {
    next.minValue = sanitizeBiDecimal(source.minValue, defaults.minValue, -1000000, 1000000);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "maxValue")) {
    const rawMax = sanitizeBiDecimal(source.maxValue, defaults.maxValue, -1000000, 1000000);
    const minRef = Number(next.minValue || 0);
    next.maxValue = rawMax < minRef ? minRef : rawMax;
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "targetValue")) {
    next.targetValue = sanitizeBiDecimal(source.targetValue, defaults.targetValue, -1000000, 1000000);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "rowLimit")) {
    next.rowLimit = sanitizeBiInteger(source.rowLimit, defaults.rowLimit, 1, 200);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showIndex")) {
    next.showIndex = normalizeBiToggle(source.showIndex, defaults.showIndex);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "compact")) {
    next.compact = normalizeBiToggle(source.compact, defaults.compact);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showTotals")) {
    next.showTotals = normalizeBiToggle(source.showTotals, defaults.showTotals);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showSource")) {
    next.showSource = normalizeBiToggle(source.showSource, defaults.showSource);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showGroups")) {
    next.showGroups = normalizeBiToggle(source.showGroups, defaults.showGroups);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "prefix")) {
    next.prefix = trimOrFallback(source.prefix, defaults.prefix).slice(0, 10);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "suffix")) {
    next.suffix = trimOrFallback(source.suffix, defaults.suffix).slice(0, 10);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showTotalBar")) {
    next.showTotalBar = normalizeBiToggle(source.showTotalBar, defaults.showTotalBar);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "totalLabel")) {
    next.totalLabel = trimOrFallback(source.totalLabel, defaults.totalLabel).slice(0, 18);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "pointAlpha")) {
    next.pointAlpha = sanitizeBiDecimal(source.pointAlpha, defaults.pointAlpha, 0.15, 1);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minPointRadius")) {
    next.minPointRadius = sanitizeBiInteger(source.minPointRadius, defaults.minPointRadius, 3, 40);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minBubbleRadius")) {
    next.minBubbleRadius = sanitizeBiInteger(source.minBubbleRadius, defaults.minBubbleRadius, 3, 60);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "maxBubbleRadius")) {
    const minRef = sanitizeBiInteger(next.minBubbleRadius, defaults.minBubbleRadius, 3, 60);
    next.maxBubbleRadius = sanitizeBiInteger(source.maxBubbleRadius, defaults.maxBubbleRadius, minRef, 80);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "lowBand")) {
    next.lowBand = sanitizeBiDecimal(source.lowBand, defaults.lowBand, 0, 100);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "midBand")) {
    const minBand = Number(next.lowBand || 0);
    next.midBand = sanitizeBiDecimal(source.midBand, defaults.midBand, minBand, 100);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showMedianValue")) {
    next.showMedianValue = normalizeBiToggle(source.showMedianValue, defaults.showMedianValue);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "whiskerMultiplier")) {
    next.whiskerMultiplier = sanitizeBiDecimal(source.whiskerMultiplier, defaults.whiskerMultiplier, 0.5, 3.5);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "upColor")) {
    next.upColor = normalizeBiColorHex(source.upColor, defaults.upColor);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "downColor")) {
    next.downColor = normalizeBiColorHex(source.downColor, defaults.downColor);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "wickWidth")) {
    next.wickWidth = sanitizeBiInteger(source.wickWidth, defaults.wickWidth, 1, 4);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "maxLinks")) {
    next.maxLinks = sanitizeBiInteger(source.maxLinks, defaults.maxLinks, 1, 200);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minLinkPercent")) {
    next.minLinkPercent = sanitizeBiDecimal(source.minLinkPercent, defaults.minLinkPercent, 0, 50);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showValues")) {
    next.showValues = normalizeBiToggle(source.showValues, defaults.showValues);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showDates")) {
    next.showDates = normalizeBiToggle(source.showDates, defaults.showDates);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "minBarPercent")) {
    next.minBarPercent = sanitizeBiDecimal(source.minBarPercent, defaults.minBarPercent, 0.2, 50);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "maxScale")) {
    next.maxScale = sanitizeBiDecimal(source.maxScale, defaults.maxScale, 1, 10000);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showPointValues")) {
    next.showPointValues = normalizeBiToggle(source.showPointValues, defaults.showPointValues);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "targetPercent")) {
    next.targetPercent = sanitizeBiDecimal(source.targetPercent, defaults.targetPercent, 1, 100);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "showCumulativeLine")) {
    next.showCumulativeLine = normalizeBiToggle(source.showCumulativeLine, defaults.showCumulativeLine);
  }
  if (Object.prototype.hasOwnProperty.call(defaults, "labelMinValue")) {
    next.labelMinValue = sanitizeBiDecimal(source.labelMinValue, defaults.labelMinValue, -1000000, 1000000);
  }
  return next;
}

function normalizeBiChartTypeConfigMap(rawMap) {
  const source = rawMap && typeof rawMap === "object" && !Array.isArray(rawMap)
    ? rawMap
    : {};
  const map = {};
  BI_CHART_TYPES.forEach((type) => {
    if (source[type] && typeof source[type] === "object" && !Array.isArray(source[type])) {
      map[type] = normalizeBiChartTypeSpecificConfig(source[type], type);
    }
  });
  return map;
}

function getBiSpecificChartConfigForWidget(widget, chartType) {
  const type = normalizeBiChartType(chartType || widget?.chartType || "bar");
  const map = normalizeBiChartTypeConfigMap(widget?.chartTypeConfig);
  return normalizeBiChartTypeSpecificConfig(map[type], type);
}

function setBiSpecificChartConfigForWidget(widget, chartType, nextConfig) {
  if (!widget) {
    return;
  }
  const type = normalizeBiChartType(chartType || widget.chartType || "bar");
  const map = normalizeBiChartTypeConfigMap(widget.chartTypeConfig);
  map[type] = normalizeBiChartTypeSpecificConfig(nextConfig, type);
  widget.chartTypeConfig = map;
}

function createDefaultBiChartConfig(chartType) {
  const safeType = normalizeBiChartType(chartType || "bar");
  return {
    common: {
      smooth: false,
      stacked: "none",
      showMedian: true
    },
    [safeType]: {
      markerSize: 4
    }
  };
}

function normalizeBiChartConfig(rawConfig, chartType) {
  const defaults = createDefaultBiChartConfig(chartType);
  const source = rawConfig && typeof rawConfig === "object" && !Array.isArray(rawConfig)
    ? rawConfig
    : {};
  const common = source.common && typeof source.common === "object" && !Array.isArray(source.common)
    ? source.common
    : {};
  const safeType = normalizeBiChartType(chartType || "bar");
  const typeConfig = source[safeType] && typeof source[safeType] === "object" && !Array.isArray(source[safeType])
    ? source[safeType]
    : {};
  return {
    common: {
      smooth: normalizeBiToggle(common.smooth, defaults.common.smooth),
      stacked: new Set(["none", "normal", "percent"]).has(trimOrFallback(common.stacked, "").toLowerCase())
        ? trimOrFallback(common.stacked, "").toLowerCase()
        : defaults.common.stacked,
      showMedian: normalizeBiToggle(common.showMedian, defaults.common.showMedian)
    },
    [safeType]: {
      markerSize: sanitizeBiInteger(typeConfig.markerSize, defaults[safeType].markerSize, 1, 24)
    }
  };
}

function normalizeBiLabelOffsets(rawOffsets) {
  if (!rawOffsets || typeof rawOffsets !== "object" || Array.isArray(rawOffsets)) {
    return {};
  }
  const next = {};
  Object.entries(rawOffsets).forEach(([rawIndex, rawOffset]) => {
    const index = parseInt(String(rawIndex || "").trim(), 10);
    if (!Number.isInteger(index) || index < 0) {
      return;
    }
    const offset = rawOffset || {};
    const x = Number(offset.x);
    const y = Number(offset.y);
    if (!Number.isFinite(x) && !Number.isFinite(y)) {
      return;
    }
    next[String(index)] = {
      x: Number.isFinite(x) ? Math.max(-400, Math.min(400, Math.round(x))) : 0,
      y: Number.isFinite(y) ? Math.max(-220, Math.min(220, Math.round(y))) : 0
    };
  });
  return next;
}

function normalizeBiLabelLayoutV2(rawLayout) {
  const source = rawLayout && typeof rawLayout === "object" && !Array.isArray(rawLayout)
    ? rawLayout
    : {};
  const rawByKey = source.byKey && typeof source.byKey === "object" && !Array.isArray(source.byKey)
    ? source.byKey
    : {};
  const byKey = {};
  Object.entries(rawByKey).forEach(([rawKey, rawPlacement]) => {
    const key = trimOrFallback(rawKey, "");
    if (!key) {
      return;
    }
    const placement = rawPlacement && typeof rawPlacement === "object" ? rawPlacement : {};
    const dxPct = Number(placement.dxPct);
    const dyPct = Number(placement.dyPct);
    const anchorXPct = Number(placement.anchorXPct);
    const anchorYPct = Number(placement.anchorYPct);
    if (!Number.isFinite(dxPct) && !Number.isFinite(dyPct) && !Number.isFinite(anchorXPct) && !Number.isFinite(anchorYPct)) {
      return;
    }
    const scopeToken = trimOrFallback(placement.scope, "").toLowerCase();
    byKey[key] = {
      anchorXPct: Number.isFinite(anchorXPct) ? Math.max(0, Math.min(1, anchorXPct)) : 0.5,
      anchorYPct: Number.isFinite(anchorYPct) ? Math.max(0, Math.min(1, anchorYPct)) : 0.5,
      dxPct: Number.isFinite(dxPct) ? Math.max(-1.5, Math.min(1.5, dxPct)) : 0,
      dyPct: Number.isFinite(dyPct) ? Math.max(-1.5, Math.min(1.5, dyPct)) : 0,
      scope: new Set(["plot", "legend", "canvas"]).has(scopeToken) ? scopeToken : "plot",
      updatedAt: trimOrFallback(placement.updatedAt, "")
    };
  });
  return {
    version: 2,
    byKey,
    cachePolicy: "keep_hidden"
  };
}

function buildBiLabelFallbackKey(label) {
  const token = normalizeLookup(label || "");
  return token ? `label_${token}` : "";
}

function buildBiRowKeyAlias(rowKey) {
  const key = trimOrFallback(rowKey, "");
  if (!key) {
    return "";
  }
  const splitAt = key.lastIndexOf("_");
  if (splitAt <= 0) {
    return key;
  }
  return key.slice(0, splitAt);
}

function getBiLabelPlacementFromLayout(layoutInput, rowKey, rowLabel = "") {
  const layout = normalizeBiLabelLayoutV2(layoutInput);
  const key = trimOrFallback(rowKey, "");
  if (key && layout.byKey[key]) {
    return layout.byKey[key];
  }

  const aliasKey = buildBiRowKeyAlias(key);
  if (aliasKey && layout.byKey[aliasKey]) {
    return layout.byKey[aliasKey];
  }

  if (aliasKey) {
    const prefix = `${aliasKey}_`;
    const candidates = Object.entries(layout.byKey)
      .filter(([candidateKey]) => candidateKey.startsWith(prefix))
      .map(([, placement]) => placement)
      .filter((placement) => placement && typeof placement === "object");
    if (candidates.length > 0) {
      candidates.sort((left, right) => {
        const a = Date.parse(trimOrFallback(left.updatedAt, ""));
        const b = Date.parse(trimOrFallback(right.updatedAt, ""));
        const safeA = Number.isFinite(a) ? a : 0;
        const safeB = Number.isFinite(b) ? b : 0;
        return safeB - safeA;
      });
      return candidates[0] || null;
    }
  }

  const fallbackKey = buildBiLabelFallbackKey(rowLabel);
  if (fallbackKey && layout.byKey[fallbackKey]) {
    return layout.byKey[fallbackKey];
  }
  return null;
}

function getBiLabelPlacement(widget, rowKey, rowLabel = "") {
  if (!widget) {
    return null;
  }
  return getBiLabelPlacementFromLayout(widget.labelLayoutV2, rowKey, rowLabel);
}

function upsertBiLabelPlacement(widget, rowKey, nextPlacement, rowLabel = "") {
  if (!widget) {
    return;
  }
  const key = trimOrFallback(rowKey, "");
  if (!nextPlacement || typeof nextPlacement !== "object") {
    return;
  }
  const layout = normalizeBiLabelLayoutV2(widget.labelLayoutV2);
  const normalizedPlacement = {
    anchorXPct: Number.isFinite(nextPlacement.anchorXPct) ? Math.max(0, Math.min(1, nextPlacement.anchorXPct)) : 0.5,
    anchorYPct: Number.isFinite(nextPlacement.anchorYPct) ? Math.max(0, Math.min(1, nextPlacement.anchorYPct)) : 0.5,
    dxPct: Number.isFinite(nextPlacement.dxPct) ? Math.max(-1.5, Math.min(1.5, nextPlacement.dxPct)) : 0,
    dyPct: Number.isFinite(nextPlacement.dyPct) ? Math.max(-1.5, Math.min(1.5, nextPlacement.dyPct)) : 0,
    scope: new Set(["plot", "legend", "canvas"]).has(trimOrFallback(nextPlacement.scope, "").toLowerCase())
      ? trimOrFallback(nextPlacement.scope, "").toLowerCase()
      : "plot",
    updatedAt: new Date().toISOString()
  };
  if (key) {
    layout.byKey[key] = normalizedPlacement;
  }
  const aliasKey = buildBiRowKeyAlias(key);
  if (aliasKey && aliasKey !== key) {
    layout.byKey[aliasKey] = {
      ...normalizedPlacement,
      scope: "plot"
    };
  }
  const fallbackKey = buildBiLabelFallbackKey(rowLabel);
  if (fallbackKey) {
    layout.byKey[fallbackKey] = {
      ...normalizedPlacement,
      // Helps keep label position when row keys change by regroup/sort.
      scope: "plot"
    };
  }
  widget.labelLayoutV2 = layout;
}

function normalizeBiCircularLayout(rawLayout) {
  const layout = rawLayout && typeof rawLayout === "object" && !Array.isArray(rawLayout) ? rawLayout : {};
  const chartSource = layout.chart && typeof layout.chart === "object"
    ? layout.chart
    : {
      x: layout.chartX ?? layout.graficoX ?? layout.dxChart,
      y: layout.chartY ?? layout.graficoY ?? layout.dyChart
    };
  const legendSource = layout.legend && typeof layout.legend === "object"
    ? layout.legend
    : {
      x: layout.legendX ?? layout.leyendaX ?? layout.dxLegend,
      y: layout.legendY ?? layout.leyendaY ?? layout.dyLegend
    };
  const chartX = Number(chartSource?.x);
  const chartY = Number(chartSource?.y);
  const legendX = Number(legendSource?.x);
  const legendY = Number(legendSource?.y);
  return {
    chart: {
      x: Number.isFinite(chartX) ? Math.max(-520, Math.min(520, Math.round(chartX))) : 0,
      y: Number.isFinite(chartY) ? Math.max(-320, Math.min(320, Math.round(chartY))) : 0
    },
    legend: {
      x: Number.isFinite(legendX) ? Math.max(-520, Math.min(520, Math.round(legendX))) : 0,
      y: Number.isFinite(legendY) ? Math.max(-320, Math.min(320, Math.round(legendY))) : 0
    }
  };
}

function normalizeBiToggle(value, fallback = true) {
  if (typeof value === "boolean") {
    return value;
  }
  if (value === null || value === undefined || value === "") {
    return !!fallback;
  }
  const token = trimOrFallback(value, "").toLowerCase();
  if (["false", "0", "off", "no"].includes(token)) {
    return false;
  }
  if (["true", "1", "on", "si", "yes"].includes(token)) {
    return true;
  }
  return !!fallback;
}

function sanitizeBiTopN(value) {
  const parsed = parseInt(String(value ?? "").trim(), 10);
  if (Number.isNaN(parsed)) {
    return 10;
  }
  return Math.max(1, Math.min(30, parsed));
}

function normalizeBiCrossFilterEntry(rawEntry) {
  const entry = rawEntry || {};
  const groupBy = normalizeBiGroupBy(entry.groupBy || entry.dimension || "");
  const label = trimOrFallback(entry.label || entry.value || "", "");
  const source = normalizeBiSource(entry.source || entry.fuente || "all");
  if (!groupBy || !label) {
    return null;
  }
  return {
    groupBy,
    label: label.slice(0, 120),
    source
  };
}

function normalizeBiCrossFilterScope(value) {
  return trimOrFallback(value, "").toLowerCase() === "same_source" ? "same_source" : "all";
}

function normalizeBiCrossFilters(rawFilters) {
  const source = Array.isArray(rawFilters)
    ? rawFilters
    : (rawFilters ? [rawFilters] : []);
  const result = [];
  const seen = new Set();
  source.forEach((item) => {
    const normalized = normalizeBiCrossFilterEntry(item);
    if (!normalized) {
      return;
    }
    const key = `${normalized.groupBy}|${normalizeLookup(normalized.label)}|${normalized.source}`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    result.push(normalized);
  });
  return result;
}

function getDefaultBiWidgetLayout(index, kindOrWidget = "chart") {
  const safeIndex = Math.max(0, Number.isInteger(index) ? index : 0);
  const kind = normalizeBiWidgetKind(
    typeof kindOrWidget === "string"
      ? kindOrWidget
      : kindOrWidget?.kind
  );
  const columns = 3;
  const slotWidth = kind === "text" ? 520 : (kind === "image" ? 420 : 360);
  const slotHeight = kind === "text" ? 108 : (kind === "image" ? 260 : 290);
  const col = safeIndex % columns;
  const row = Math.floor(safeIndex / columns);
  return {
    x: 26 + (col * (slotWidth + 18)),
    y: 22 + (row * (slotHeight + 18)),
    w: slotWidth,
    h: slotHeight
  };
}

function normalizeBiWidgetLayout(rawLayout, index, kindOrWidget = "chart") {
  const fallback = getDefaultBiWidgetLayout(index, kindOrWidget);
  const minSize = getBiWidgetMinSize(kindOrWidget);
  const minWidth = Math.max(80, Math.round(minSize.width));
  const minHeight = Math.max(52, Math.round(minSize.height));
  const layout = rawLayout || {};
  const x = Number.isFinite(layout.x) ? Math.max(0, Math.round(layout.x)) : fallback.x;
  const y = Number.isFinite(layout.y) ? Math.max(0, Math.round(layout.y)) : fallback.y;
  const w = Number.isFinite(layout.w)
    ? Math.max(minWidth, Math.round(layout.w))
    : Math.max(minWidth, fallback.w);
  const h = Number.isFinite(layout.h)
    ? Math.max(minHeight, Math.round(layout.h))
    : Math.max(minHeight, fallback.h);
  return { x, y, w, h };
}

function sanitizeBiInteger(value, fallback, min, max) {
  const parsed = parseInt(String(value ?? "").trim(), 10);
  if (Number.isNaN(parsed)) {
    return Math.max(min, Math.min(max, fallback));
  }
  return Math.max(min, Math.min(max, parsed));
}

function sanitizeBiDecimal(value, fallback, min, max) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return Math.max(min, Math.min(max, fallback));
  }
  return Math.max(min, Math.min(max, Math.round(parsed * 100) / 100));
}

function sanitizeBiNullableDecimal(value, min, max) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  return Math.max(min, Math.min(max, Math.round(parsed * 100) / 100));
}

function sanitizeBiFontFamily(value, fallback = "Segoe UI") {
  const token = trimOrFallback(value, "");
  if (BI_ALLOWED_FONT_FAMILIES.has(token)) {
    return token;
  }
  return BI_ALLOWED_FONT_FAMILIES.has(fallback) ? fallback : "Segoe UI";
}

function createDefaultBiVisualSettings() {
  return {
    showLegend: true,
    legendPosition: "right",
    legendMaxItems: 8,
    showGrid: true,
    showAxisLabels: true,
    showDataLabels: false,
    axisXLabel: "",
    axisYLabel: "",
    axisScale: "linear",
    axisMin: null,
    axisMax: null,
    showTargetLine: false,
    targetLineValue: null,
    targetLineLabel: "Meta",
    targetLineColor: "#b03831",
    labelMaxChars: 12,
    valueDecimals: 1,
    numberFormat: "auto",
    valuePrefix: "",
    valueSuffix: "",
    fontSize: 10,
    fontFamilyTitle: "Segoe UI",
    fontFamilyAxis: "Segoe UI",
    fontFamilyLabel: "Segoe UI",
    fontFamilyTooltip: "Segoe UI",
    fontSizeTitle: 16,
    fontSizeAxis: 10,
    fontSizeLabels: 10,
    fontSizeTooltip: 11,
    labelColorMode: "auto",
    labelColor: "#1f2f44",
    lineWidth: 2,
    seriesLineStyle: "solid",
    seriesMarkerStyle: "circle",
    seriesMarkerSize: 4,
    areaOpacity: 0.2,
    smoothLines: false,
    barWidthRatio: 0.64,
    gridOpacity: 0.35,
    gridDash: 4,
    stackMode: "none",
    tooltipMode: "full"
  };
}

function normalizeBiVisualSettings(rawSettings) {
  const defaults = createDefaultBiVisualSettings();
  const source = rawSettings || {};
  const legendPositionToken = trimOrFallback(source.legendPosition || source.posicionLeyenda || "", "").toLowerCase();
  const legendPosition = new Set(["top", "right", "bottom", "left"]).has(legendPositionToken)
    ? legendPositionToken
    : defaults.legendPosition;
  const axisScaleToken = trimOrFallback(source.axisScale || source.escalaEje || "", "").toLowerCase();
  const axisScale = axisScaleToken === "log" ? "log" : "linear";
  const numberFormatToken = trimOrFallback(source.numberFormat || source.formatoNumero || "", "").toLowerCase();
  const numberFormat = new Set(["auto", "number", "percent", "currency_pen", "hours"]).has(numberFormatToken)
    ? numberFormatToken
    : defaults.numberFormat;
  const tooltipModeToken = trimOrFallback(source.tooltipMode || source.modoTooltip || "", "").toLowerCase();
  const tooltipMode = new Set(["full", "compact"]).has(tooltipModeToken)
    ? tooltipModeToken
    : defaults.tooltipMode;
  const lineStyleToken = trimOrFallback(source.seriesLineStyle || source.estiloLineaSerie || "", "").toLowerCase();
  const seriesLineStyle = new Set(["solid", "dashed", "dotted"]).has(lineStyleToken)
    ? lineStyleToken
    : defaults.seriesLineStyle;
  const markerStyleToken = trimOrFallback(source.seriesMarkerStyle || source.estiloMarcadorSerie || "", "").toLowerCase();
  const seriesMarkerStyle = new Set(["circle", "square", "diamond", "triangle", "none"]).has(markerStyleToken)
    ? markerStyleToken
    : defaults.seriesMarkerStyle;
  const stackModeToken = trimOrFallback(source.stackMode || source.modoStacking || source.stacking || "", "").toLowerCase();
  const stackMode = new Set(["none", "normal", "percent"]).has(stackModeToken)
    ? stackModeToken
    : defaults.stackMode;
  const labelColorModeToken = trimOrFallback(
    source.labelColorMode || source.modoColorLabels || source.labelColorAutoMode || "",
    ""
  ).toLowerCase();
  const labelColorMode = labelColorModeToken === "manual" ? "manual" : "auto";
  return {
    showLegend: source.showLegend !== false,
    legendPosition,
    legendMaxItems: sanitizeBiInteger(source.legendMaxItems, defaults.legendMaxItems, 3, 20),
    showGrid: source.showGrid !== false,
    showAxisLabels: source.showAxisLabels !== false,
    showDataLabels: !!source.showDataLabels,
    axisXLabel: trimOrFallback(source.axisXLabel || source.ejeX || "", "").slice(0, 40),
    axisYLabel: trimOrFallback(source.axisYLabel || source.ejeY || "", "").slice(0, 40),
    axisScale,
    axisMin: sanitizeBiNullableDecimal(source.axisMin ?? source.minEjeY, -1000000000, 1000000000),
    axisMax: sanitizeBiNullableDecimal(source.axisMax ?? source.maxEjeY, -1000000000, 1000000000),
    showTargetLine: normalizeBiToggle(source.showTargetLine ?? source.mostrarLineaObjetivo, defaults.showTargetLine),
    targetLineValue: sanitizeBiNullableDecimal(source.targetLineValue ?? source.valorLineaObjetivo, -1000000000, 1000000000),
    targetLineLabel: trimOrFallback(source.targetLineLabel || source.etiquetaLineaObjetivo || defaults.targetLineLabel, defaults.targetLineLabel).slice(0, 24),
    targetLineColor: normalizeBiColorHex(source.targetLineColor || source.colorLineaObjetivo || "", defaults.targetLineColor),
    labelMaxChars: sanitizeBiInteger(source.labelMaxChars, defaults.labelMaxChars, 4, 40),
    valueDecimals: sanitizeBiInteger(source.valueDecimals, defaults.valueDecimals, 0, 4),
    numberFormat,
    valuePrefix: trimOrFallback(source.valuePrefix || source.prefijoValor || "", "").slice(0, 8),
    valueSuffix: trimOrFallback(source.valueSuffix || source.sufijoValor || "", "").slice(0, 8),
    fontSize: sanitizeBiInteger(source.fontSize, defaults.fontSize, 9, 18),
    fontFamilyTitle: sanitizeBiFontFamily(source.fontFamilyTitle || source.fuenteTitulo, defaults.fontFamilyTitle),
    fontFamilyAxis: sanitizeBiFontFamily(source.fontFamilyAxis || source.fuenteEjes, defaults.fontFamilyAxis),
    fontFamilyLabel: sanitizeBiFontFamily(source.fontFamilyLabel || source.fuenteLabels, defaults.fontFamilyLabel),
    fontFamilyTooltip: sanitizeBiFontFamily(source.fontFamilyTooltip || source.fuenteTooltip, defaults.fontFamilyTooltip),
    fontSizeTitle: sanitizeBiInteger(source.fontSizeTitle, defaults.fontSizeTitle, 10, 32),
    fontSizeAxis: sanitizeBiInteger(source.fontSizeAxis, defaults.fontSizeAxis, 8, 20),
    fontSizeLabels: sanitizeBiInteger(source.fontSizeLabels, defaults.fontSizeLabels, 8, 22),
    fontSizeTooltip: sanitizeBiInteger(source.fontSizeTooltip, defaults.fontSizeTooltip, 8, 20),
    labelColorMode,
    labelColor: normalizeBiColorHex(source.labelColor || source.colorLabels || "", defaults.labelColor),
    lineWidth: sanitizeBiInteger(source.lineWidth, defaults.lineWidth, 1, 6),
    seriesLineStyle,
    seriesMarkerStyle,
    seriesMarkerSize: sanitizeBiInteger(
      source.seriesMarkerSize ?? source.markerSize ?? source.tamanoMarcadorSerie,
      defaults.seriesMarkerSize,
      0,
      16
    ),
    areaOpacity: sanitizeBiDecimal(source.areaOpacity, defaults.areaOpacity, 0.05, 0.8),
    smoothLines: !!source.smoothLines,
    barWidthRatio: sanitizeBiDecimal(source.barWidthRatio, defaults.barWidthRatio, 0.2, 0.9),
    gridOpacity: sanitizeBiDecimal(source.gridOpacity, defaults.gridOpacity, 0.1, 1),
    gridDash: sanitizeBiInteger(source.gridDash, defaults.gridDash, 0, 12),
    stackMode,
    tooltipMode
  };
}

function applyBiChartTypeVisualDefaults(settings, chartType) {
  const safeType = normalizeBiChartType(chartType || "bar");
  const visual = normalizeBiVisualSettings(settings || {});
  if (BI_STACKED_BAR_TYPES.has(safeType) && visual.stackMode === "none") {
    return normalizeBiVisualSettings({
      ...visual,
      stackMode: "normal"
    });
  }
  return visual;
}

function normalizeBiVisualPresets(rawPresets) {
  if (!rawPresets || typeof rawPresets !== "object" || Array.isArray(rawPresets)) {
    return {};
  }
  const next = {};
  Object.entries(rawPresets).forEach(([rawId, rawPreset]) => {
    const id = trimOrFallback(rawId, "");
    if (!id) {
      return;
    }
    const preset = rawPreset || {};
    const name = trimOrFallback(preset.name || preset.nombre || id, id).slice(0, 30);
    next[id] = {
      name,
      settings: normalizeBiVisualSettings(preset.settings || preset.visual || preset)
    };
  });
  return next;
}

function createDefaultBiConfig() {
  return {
    startDate: "",
    endDate: "",
    textFilter: "",
    source: "all",
    invalidOnly: false,
    uiMode: "basic",
    canvasWidth: BI_CANVAS_SURFACE_MIN_WIDTH,
    canvasHeight: BI_CANVAS_SURFACE_HEIGHT,
    canvasZoom: BI_CANVAS_ZOOM_DEFAULT,
    showCanvasGrid: false,
    snapToGrid: true,
    gridSnapSize: 12,
    performanceMode: "balanced",
    colorSource: "all",
    colorGroupBy: "disciplina",
    colorMap: {},
    crossFilterScope: "all",
    crossFilters: [],
    visual: createDefaultBiVisualSettings(),
    visualPresets: {},
    activeVisualPreset: "corporativo"
  };
}

function normalizeBiConfig(rawConfig) {
  const base = rawConfig || {};
  const canvasSize = getBiCanvasSizeFromConfig(base);
  return {
    startDate: sanitizeDateInput(base.startDate || base.fechaInicio || ""),
    endDate: sanitizeDateInput(base.endDate || base.fechaFin || ""),
    textFilter: trimOrFallback(base.textFilter || base.filtroTexto || "", "").slice(0, 120),
    source: normalizeBiSource(base.source || base.fuente || "all"),
    invalidOnly: !!(base.invalidOnly || base.soloFechasInvertidas),
    uiMode: normalizeBiUiMode(base.uiMode || base.modoInterfaz || "basic"),
    canvasWidth: canvasSize.width,
    canvasHeight: canvasSize.height,
    canvasZoom: sanitizeBiCanvasZoom(base.canvasZoom ?? base.zoomPizarra ?? base.boardZoom, BI_CANVAS_ZOOM_DEFAULT),
    showCanvasGrid: normalizeBiToggle(base.showCanvasGrid ?? base.mostrarRejillaPizarra, false),
    snapToGrid: normalizeBiToggle(base.snapToGrid ?? base.ajustarARejilla, true),
    gridSnapSize: sanitizeBiInteger(base.gridSnapSize ?? base.pasoRejilla, 12, 4, 120),
    performanceMode: normalizeBiPerformanceMode(base.performanceMode || base.modoRendimiento || "balanced"),
    colorSource: normalizeBiSource(base.colorSource || base.fuenteColor || "all"),
    colorGroupBy: normalizeBiGroupBy(base.colorGroupBy || base.dimensionColor || "disciplina"),
    colorMap: normalizeBiColorMap(base.colorMap || base.mapaColor),
    crossFilterScope: normalizeBiCrossFilterScope(base.crossFilterScope || base.alcanceFiltroCruzado || "all"),
    crossFilters: normalizeBiCrossFilters(base.crossFilters || base.filtrosCruzados),
    visual: normalizeBiVisualSettings(base.visual || base.configuracionVisual || base.estiloGrafico),
    visualPresets: normalizeBiVisualPresets(base.visualPresets || base.presetsVisuales),
    activeVisualPreset: trimOrFallback(base.activeVisualPreset || base.presetVisualActivo || "corporativo", "corporativo")
  };
}

function createDefaultBiWidgets() {
  return [
    {
      id: uid(),
      kind: "chart",
      name: "Avance real por Disciplina",
      source: "all",
      groupBy: "disciplina",
      metric: "weightedreal",
      chartType: "bar",
      sortMode: "value_desc",
      topN: 10,
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: "disciplina", metric: "weightedreal", chartType: "bar" }),
      chartConfig: createDefaultBiChartConfig("bar"),
      chartTypeConfig: {
        bar: normalizeBiChartTypeSpecificConfig({}, "bar")
      },
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      visualOverride: null
    },
    {
      id: uid(),
      kind: "chart",
      name: "UP Base por Sistema",
      source: "all",
      groupBy: "sistema",
      metric: "baseunits",
      chartType: "donut",
      sortMode: "value_desc",
      topN: 10,
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: "sistema", metric: "baseunits", chartType: "donut" }),
      chartConfig: createDefaultBiChartConfig("donut"),
      chartTypeConfig: {
        donut: normalizeBiChartTypeSpecificConfig({}, "donut")
      },
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      visualOverride: null
    },
    {
      id: uid(),
      kind: "chart",
      name: "Programado por Mes de inicio",
      source: "deliverable",
      groupBy: "mes_inicio",
      metric: "programmedavg",
      chartType: "line",
      sortMode: "value_desc",
      topN: 12,
      textContent: "",
      imageSrc: "",
      imageAlt: "",
      showTitle: true,
      showBorder: true,
      showBackground: true,
      locked: false,
      dataRoles: buildBiDefaultDataRoles({ groupBy: "mes_inicio", metric: "programmedavg", chartType: "line" }),
      chartConfig: createDefaultBiChartConfig("line"),
      chartTypeConfig: {
        line: normalizeBiChartTypeSpecificConfig({}, "line")
      },
      textConfig: createDefaultBiTextConfig(),
      imageConfig: createDefaultBiImageConfig(),
      labelOffsets: {},
      labelLayoutV2: normalizeBiLabelLayoutV2({}),
      polarLayout: normalizeBiCircularLayout({}),
      visualOverride: null
    }
  ];
}

function normalizeBiWidget(rawWidget, index) {
  const widget = rawWidget || {};
  const kind = normalizeBiWidgetKind(widget.kind || widget.tipoWidget || widget.tipo || "chart");
  const textContent = trimOrFallback(widget.textContent ?? widget.texto ?? widget.contenidoTexto ?? "", "").slice(0, 4000);
  const imageSrc = trimOrFallback(widget.imageSrc ?? widget.imagenSrc ?? widget.imagen ?? widget.image ?? widget.urlImagen ?? "", "").slice(0, 2000000);
  const imageAlt = trimOrFallback(widget.imageAlt ?? widget.imagenAlt ?? widget.altImagen ?? "", "").slice(0, 120);
  const labelOffsets = normalizeBiLabelOffsets(widget.labelOffsets ?? widget.posicionLabels ?? widget.desplazamientoLabels);
  const labelLayoutV2 = normalizeBiLabelLayoutV2(
    widget.labelLayoutV2
    ?? widget.labelLayout
    ?? widget.etiquetasLayout
  );
  const polarLayout = normalizeBiCircularLayout(
    widget.polarLayout
    ?? widget.circularLayout
    ?? widget.layoutCircular
    ?? widget.desplazamientoCircular
  );
  const defaultShowTitle = kind === "text" ? false : true;
  const defaultShowBorder = kind === "text" ? false : true;
  const defaultShowBackground = kind === "text" ? false : true;
  const showTitle = normalizeBiToggle(
    widget.showTitle ?? widget.mostrarTitulo ?? (widget.hideTitle === true ? false : undefined),
    defaultShowTitle
  );
  const showBorder = normalizeBiToggle(
    widget.showBorder ?? widget.mostrarBorde ?? (widget.hideBorder === true ? false : undefined),
    defaultShowBorder
  );
  const showBackground = normalizeBiToggle(
    widget.showBackground ?? widget.mostrarFondo ?? (widget.hideBackground === true ? false : undefined),
    defaultShowBackground
  );
  const locked = normalizeBiToggle(
    widget.locked ?? widget.bloqueado ?? widget.widgetLocked,
    false
  );
  const source = normalizeBiSource(widget.source || widget.fuente || "all");
  const groupByInput = normalizeBiGroupBy(widget.groupBy || widget.agruparPor || "disciplina");
  const metricInput = normalizeBiMetric(widget.metric || widget.metrica || "baseunits");
  const chartType = normalizeBiChartType(widget.chartType || widget.tipoGrafico || "bar");
  const baseForRoles = { source, groupBy: groupByInput, metric: metricInput, chartType };
  const dataRoles = normalizeBiDataRoles(widget.dataRoles ?? widget.rolesDatos, baseForRoles);
  const groupBy = normalizeBiGroupBy(dataRoles.dimensions?.[0] || groupByInput);
  const metric = normalizeBiMetric(dataRoles.metrics?.[0] || metricInput);
  const textConfig = normalizeBiTextConfig(widget.textConfig ?? widget.configTexto);
  const imageConfig = normalizeBiImageConfig(widget.imageConfig ?? widget.configImagen);
  const chartTypeConfig = normalizeBiChartTypeConfigMap(widget.chartTypeConfig ?? widget.configTipoGrafico);
  if (!chartTypeConfig[chartType]) {
    chartTypeConfig[chartType] = normalizeBiChartTypeSpecificConfig({}, chartType);
  }
  return {
    id: typeof widget.id === "string" && widget.id.trim() ? widget.id.trim() : uid(),
    kind,
    name: trimOrFallback(widget.name || `Widget ${index + 1}`, `Widget ${index + 1}`).slice(0, 60),
    source,
    groupBy,
    metric,
    chartType,
    sortMode: normalizeBiSortMode(widget.sortMode || widget.orden || widget.sort || "value_desc"),
    topN: sanitizeBiTopN(widget.topN),
    textContent,
    imageSrc,
    imageAlt,
    showTitle,
    showBorder,
    showBackground,
    locked,
    dataRoles,
    chartConfig: normalizeBiChartConfig(widget.chartConfig ?? widget.configGrafico, chartType),
    chartTypeConfig,
    textConfig,
    imageConfig,
    labelOffsets,
    labelLayoutV2,
    polarLayout,
    layout: normalizeBiWidgetLayout(widget.layout ?? widget.posicion, index, kind),
    visualOverride: widget.visualOverride ? normalizeBiVisualSettings(widget.visualOverride) : null
  };
}

function normalizeBiWidgets(rawWidgets) {
  if (!Array.isArray(rawWidgets) || rawWidgets.length === 0) {
    return createDefaultBiWidgets();
  }

  const unique = [];
  const seen = new Set();
  rawWidgets.forEach((item, index) => {
    const normalized = normalizeBiWidget(item, index);
    if (seen.has(normalized.id)) {
      return;
    }
    seen.add(normalized.id);
    unique.push(normalized);
  });
  return unique.length > 0 ? unique : createDefaultBiWidgets();
}

function normalizeRealAdvances(rawEntries) {
  if (!Array.isArray(rawEntries)) {
    return [];
  }

  return rawEntries.map((entry) => {
    const id = typeof entry?.id === "string" && entry.id.trim() ? entry.id.trim() : uid();
    const createdAt = trimOrFallback(entry?.createdAt, "");
    const date = sanitizeDateInput(entry?.date || "");
    const value = sanitizeRealProgress(entry?.value);
    const addedPercent = sanitizeSignedPercent(entry?.addedPercent);
    const author = trimOrFallback(entry?.author, "");
    return {
      id,
      createdAt: createdAt || (date ? `${date}T00:00:00` : ""),
      date,
      value: value === "" ? null : value,
      addedPercent: addedPercent === "" ? null : addedPercent,
      author
    };
  });
}

function normalizeConsumedHours(rawEntries) {
  if (!Array.isArray(rawEntries)) {
    return [];
  }

  return rawEntries.map((entry) => {
    const id = typeof entry?.id === "string" && entry.id.trim() ? entry.id.trim() : uid();
    const date = sanitizeDateInput(entry?.date || "");
    const hours = sanitizeBaseUnits(entry?.hours);
    const note = trimOrFallback(entry?.note, "").slice(0, TRACKING_MAX_NOTE);
    return {
      id,
      date,
      hours: hours === "" ? null : hours,
      note
    };
  });
}

function normalizeReviewMilestones(rawEntries) {
  if (!Array.isArray(rawEntries)) {
    return [];
  }

  const unique = [];
  const seenIds = new Set();
  rawEntries.forEach((entry, index) => {
    const normalized = normalizeReviewMilestone(entry, index);
    if (seenIds.has(normalized.id)) {
      return;
    }
    seenIds.add(normalized.id);
    unique.push(normalized);
  });
  return unique;
}

function normalizeReviewMilestone(rawEntry, index) {
  const id = typeof rawEntry?.id === "string" && rawEntry.id.trim()
    ? rawEntry.id.trim()
    : uid();
  const rawName = rawEntry?.name ?? rawEntry?.hito;
  const name = typeof rawName === "string"
    ? rawName.trim()
    : `Hito ${index + 1}`;
  const baseUnitsValue = sanitizeBaseUnits(rawEntry?.baseUnits ?? rawEntry?.pesoUP ?? rawEntry?.unidadProductiva);

  return {
    id,
    name,
    baseUnits: baseUnitsValue === "" ? null : baseUnitsValue,
    createdAt: trimOrFallback(rawEntry?.createdAt, new Date().toISOString())
  };
}

function normalizeReviewControls(rawEntries) {
  if (!Array.isArray(rawEntries)) {
    return [];
  }

  const unique = [];
  const seenKeys = new Set();
  rawEntries.forEach((entry) => {
    const normalized = normalizeReviewControl(entry);
    if (!normalized.key || seenKeys.has(normalized.key)) {
      return;
    }
    seenKeys.add(normalized.key);
    unique.push(normalized);
  });
  return unique;
}

function normalizeReviewControl(rawEntry) {
  const projectRowId = trimOrFallback(rawEntry?.projectRowId, "");
  const disciplineRowId = trimOrFallback(rawEntry?.disciplineRowId, "");
  const systemRowId = trimOrFallback(rawEntry?.systemRowId, "");
  const milestoneId = trimOrFallback(rawEntry?.milestoneId, "");
  const derivedKey = buildReviewControlKeyFromRowIds(projectRowId, disciplineRowId, systemRowId, milestoneId);
  const key = trimOrFallback(rawEntry?.key, derivedKey);
  const startDate = sanitizeDateInput(rawEntry?.startDate || "");
  const endDate = sanitizeDateInput(rawEntry?.endDate || "");
  const baseUnitsRaw = sanitizeBaseUnits(rawEntry?.baseUnits);
  const realProgressRaw = sanitizeRealProgress(rawEntry?.realProgress);
  const realAdvances = normalizeRealAdvances(rawEntry?.realAdvances ?? rawEntry?.avancesReales);
  const consumedHours = normalizeConsumedHours(rawEntry?.consumedHours ?? rawEntry?.horasConsumidas);

  return {
    id: typeof rawEntry?.id === "string" && rawEntry.id.trim() ? rawEntry.id.trim() : uid(),
    key,
    projectRowId,
    disciplineRowId,
    systemRowId,
    milestoneId,
    startDate,
    endDate,
    baseUnits: baseUnitsRaw === "" ? null : baseUnitsRaw,
    realProgress: realProgressRaw === "" ? null : realProgressRaw,
    realAdvances,
    consumedHours,
    createdAt: trimOrFallback(rawEntry?.createdAt, new Date().toISOString())
  };
}

function normalizePackageControls(rawEntries) {
  if (!Array.isArray(rawEntries)) {
    return [];
  }

  const unique = [];
  const seenKeys = new Set();
  rawEntries.forEach((entry) => {
    const normalized = normalizePackageControl(entry);
    if (!normalized.key || seenKeys.has(normalized.key)) {
      return;
    }
    seenKeys.add(normalized.key);
    unique.push(normalized);
  });
  return unique;
}

function normalizePackageControl(rawEntry) {
  const projectRowId = trimOrFallback(rawEntry?.projectRowId, "");
  const disciplineRowId = trimOrFallback(rawEntry?.disciplineRowId, "");
  const systemRowId = trimOrFallback(rawEntry?.systemRowId, "");
  const packageRowId = trimOrFallback(rawEntry?.packageRowId, "");
  const startDate = sanitizeDateInput(rawEntry?.startDate || "");
  const endDate = sanitizeDateInput(rawEntry?.endDate || "");
  const realProgressRaw = sanitizeRealProgress(rawEntry?.realProgress);
  const realAdvances = normalizeRealAdvances(rawEntry?.realAdvances ?? rawEntry?.avancesReales);
  const consumedHours = normalizeConsumedHours(rawEntry?.consumedHours ?? rawEntry?.horasConsumidas);
  const derivedKey = buildPackageControlKeyFromRowIds(projectRowId, disciplineRowId, systemRowId, packageRowId);
  const key = trimOrFallback(rawEntry?.key, derivedKey);
  return {
    id: typeof rawEntry?.id === "string" && rawEntry.id.trim() ? rawEntry.id.trim() : uid(),
    key,
    projectRowId,
    disciplineRowId,
    systemRowId,
    packageRowId,
    startDate,
    endDate,
    realProgress: realProgressRaw === "" ? null : realProgressRaw,
    realAdvances,
    consumedHours,
    createdAt: trimOrFallback(rawEntry?.createdAt, new Date().toISOString())
  };
}

function normalizeDeliverable(rawDeliverable, fields, index) {
  const rowRefs = {};
  const values = {};
  const codes = {};

  fields.forEach((field) => {
    const existingRowRef = trimOrFallback(rawDeliverable?.rowRefs?.[field.id], "");
    const rowByRef = getFieldRowById(field, existingRowRef);
    const legacyValue = trimOrFallback(rawDeliverable?.values?.[field.id], "");
    const legacyCode = trimOrFallback(rawDeliverable?.codes?.[field.id], "");
    const inferredRow = rowByRef || inferRowForField(field, legacyValue, legacyCode);

    rowRefs[field.id] = inferredRow ? inferredRow.id : "";
    values[field.id] = inferredRow ? inferredRow.name : legacyValue;
    codes[field.id] = inferredRow ? inferredRow.code : legacyCode;
  });

  const ref = trimOrFallback(rawDeliverable?.ref, `REF-${String(index + 1).padStart(3, "0")}`);
  const progressMeta = normalizeProgressMeta({
    startDate: rawDeliverable?.startDate ?? rawDeliverable?.fechaInicio ?? rawDeliverable?.control?.startDate,
    endDate: rawDeliverable?.endDate ?? rawDeliverable?.fechaFin ?? rawDeliverable?.control?.endDate,
    baseUnits: rawDeliverable?.baseUnits ?? rawDeliverable?.unidadesProductivasBase ?? rawDeliverable?.control?.baseUnits,
    realProgress: rawDeliverable?.realProgress ?? rawDeliverable?.avanceReal ?? rawDeliverable?.control?.realProgress
  });
  const realAdvances = normalizeRealAdvances(rawDeliverable?.realAdvances ?? rawDeliverable?.avancesReales);
  const consumedHours = normalizeConsumedHours(rawDeliverable?.consumedHours ?? rawDeliverable?.horasConsumidas);

  return {
    id: typeof rawDeliverable?.id === "string" && rawDeliverable.id.trim() ? rawDeliverable.id.trim() : uid(),
    ref,
    code: trimOrFallback(rawDeliverable?.code, ""),
    startDate: progressMeta.startDate,
    endDate: progressMeta.endDate,
    baseUnits: progressMeta.baseUnits,
    realProgress: progressMeta.realProgress,
    realAdvances,
    consumedHours,
    rowRefs,
    values,
    codes,
    createdAt: trimOrFallback(rawDeliverable?.createdAt, new Date().toISOString())
  };
}

function createDefaultProject(id, name) {
  const fields = DEFAULT_FIELD_TEMPLATES.map((template) =>
    createField(template.label, template.includeInCode, template.locked, template.rows));
  applyDefaultSystemDisciplineMapping(fields);

  return {
    id,
    name,
    title: `MIDP - ${name}`,
    codeSeparator: "-",
    fields,
    nomenclatureOrder: fields.filter((field) => field.includeInCode).map((field) => field.id),
    deliverables: [],
    reviewMilestones: [],
    reviewControls: [],
    packageControls: [],
    biConfig: createDefaultBiConfig(),
    biWidgets: createDefaultBiWidgets(),
    quickSightConfig: createDefaultQuickSightConfig(),
    quickSightVisuals: [],
    demoSeeded: false,
    progressControlMode: "deliverable",
    showMidpAdvanceDetails: true
  };
}

function applyDefaultSystemDisciplineMapping(fields) {
  const pair = getDisciplineSystemFields(fields);
  if (!pair.disciplineField || !pair.systemField) {
    return;
  }

  const disciplineRows = pair.disciplineField.rows.filter((row) => hasRowContent(row));
  const systemRows = pair.systemField.rows.filter((row) => hasRowContent(row));
  if (disciplineRows.length === 0 || systemRows.length === 0) {
    return;
  }

  systemRows.forEach((systemRow, index) => {
    if (trimOrFallback(systemRow.parentDisciplineRowId, "")) {
      return;
    }
    const disciplineRow = disciplineRows[Math.min(index, disciplineRows.length - 1)];
    if (!disciplineRow) {
      return;
    }
    systemRow.parentDisciplineRowId = disciplineRow.id;
  });
}

function createField(label, includeInCode, locked, rowTemplates) {
  const rows = Array.isArray(rowTemplates)
    ? rowTemplates.map((row) => createFieldRow(
      row.name || "",
      row.code || "",
      row.parentDisciplineRowId || row.disciplineRowId || row.parentRowId || ""))
    : [];

  return {
    id: uid(),
    label,
    includeInCode,
    locked,
    rows
  };
}

function createFieldRow(name, code, parentDisciplineRowId) {
  return {
    id: uid(),
    name: trimOrFallback(name, ""),
    code: trimOrFallback(code, ""),
    parentDisciplineRowId: trimOrFallback(parentDisciplineRowId, "")
  };
}

function buildUniqueProjectId(rawName) {
  const base = slugify(rawName) || "proyecto";
  let candidate = base;
  let index = 2;

  while (state.projects.some((project) => project.id === candidate)) {
    candidate = `${base}_${index}`;
    index += 1;
  }
  return candidate;
}

function slugify(value) {
  return removeDiacritics(trimOrFallback(value, ""))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function isProjectLikeField(label) {
  const token = removeDiacritics(label || "").toLowerCase();
  return token.includes("proyecto");
}

function toCodeToken(value) {
  return removeDiacritics(trimOrFallback(value, ""))
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

function sanitizeSeparator(value) {
  const clean = trimOrFallback(value, "-");
  return clean.slice(0, 1);
}

function normalizeProgressControlMode(value) {
  const token = trimOrFallback(value, "").toLowerCase();
  if (token === "package" || token === "paquete" || token === "porpaquete" || token === "por_paquete") {
    return "package";
  }
  return "deliverable";
}

function getProjectProgressControlMode(project) {
  return normalizeProgressControlMode(project?.progressControlMode);
}

function getProjectMidpAdvanceDetailsVisible(project) {
  return project?.showMidpAdvanceDetails !== false;
}

function formatDate(value) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }
  return parsed.toLocaleString();
}

function formatDateFromInput(value) {
  const safe = sanitizeDateInput(value);
  if (!safe) {
    return "";
  }
  const parsed = new Date(`${safe}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return safe;
  }
  return parsed.toLocaleDateString();
}

function normalizeProgressMeta(rawMeta) {
  const startDate = sanitizeDateInput(rawMeta?.startDate);
  const endDate = sanitizeDateInput(rawMeta?.endDate);
  const baseUnitsValue = sanitizeBaseUnits(rawMeta?.baseUnits);
  const realProgressValue = sanitizeRealProgress(rawMeta?.realProgress);

  return {
    startDate,
    endDate,
    baseUnits: baseUnitsValue === "" ? null : baseUnitsValue,
    realProgress: realProgressValue === "" ? null : realProgressValue
  };
}

function sanitizeDateInput(value) {
  if (typeof value !== "string") {
    return "";
  }

  const text = value.trim();
  if (!text) {
    return "";
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return "";
  }

  const parsed = new Date(`${text}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  return text;
}

function sanitizeBaseUnits(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 0) {
      return "";
    }
    return Math.round(value * 100) / 100;
  }

  if (typeof value !== "string") {
    return "";
  }

  const text = value.trim().replace(",", ".");
  if (!text) {
    return "";
  }

  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return "";
  }
  return Math.round(parsed * 100) / 100;
}

function sanitizeRealProgress(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 0 || value > 100) {
      return "";
    }
    return Math.round(value * 100) / 100;
  }

  if (typeof value !== "string") {
    return "";
  }

  const text = value.trim().replace(",", ".");
  if (!text) {
    return "";
  }

  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 100) {
    return "";
  }
  return Math.round(parsed * 100) / 100;
}

function sanitizeSignedPercent(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < -100 || value > 100) {
      return "";
    }
    return Math.round(value * 100) / 100;
  }

  if (typeof value !== "string") {
    return "";
  }

  const text = value.trim().replace(",", ".");
  if (!text) {
    return "";
  }

  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed) || parsed < -100 || parsed > 100) {
    return "";
  }
  return Math.round(parsed * 100) / 100;
}

function formatNumberForInput(value) {
  const parsed = sanitizeBaseUnits(value);
  if (parsed === "") {
    return "";
  }
  const rounded = Math.round(parsed * 100) / 100;
  return String(rounded).replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

function formatPercent(value, decimals) {
  if (!Number.isFinite(value)) {
    return "";
  }
  const digits = Number.isInteger(decimals) ? Math.max(0, decimals) : 2;
  const fixed = value.toFixed(digits).replace(/\.?0+$/, "");
  return `${fixed}%`;
}

function formatSignedPercent(value, decimals) {
  if (!Number.isFinite(value)) {
    return "";
  }

  const base = formatPercent(value, decimals);
  if (!base) {
    return "";
  }
  if (value > 0) {
    return `+${base}`;
  }
  return base;
}

function computeProjectBaseUnitsTotal(deliverables) {
  if (!Array.isArray(deliverables)) {
    return 0;
  }

  let total = 0;
  deliverables.forEach((item) => {
    const value = sanitizeBaseUnits(item?.baseUnits);
    if (value === "") {
      return;
    }
    total += value;
  });

  return Math.round(total * 100) / 100;
}

function findFieldIdByAlias(fields, aliases) {
  if (!Array.isArray(fields) || !Array.isArray(aliases) || aliases.length === 0) {
    return null;
  }

  const normalizedAliases = aliases
    .map((alias) => normalizeLookup(alias))
    .filter((alias) => !!alias);

  if (normalizedAliases.length === 0) {
    return null;
  }

  for (const field of fields) {
    const labelLookup = normalizeLookup(field?.label || "");
    if (!labelLookup) {
      continue;
    }

    const matches = normalizedAliases.some((alias) => labelLookup.includes(alias));
    if (matches) {
      return field.id;
    }
  }

  return null;
}

function buildDeliverableGroupKey(deliverable, fieldIds) {
  if (!deliverable || !deliverable.rowRefs || !Array.isArray(fieldIds) || fieldIds.length === 0) {
    return "";
  }

  const parts = [];
  for (const fieldId of fieldIds) {
    if (typeof fieldId !== "string" || !fieldId) {
      return "";
    }
    const rowId = trimOrFallback(deliverable.rowRefs[fieldId], "");
    if (!rowId) {
      return "";
    }
    parts.push(rowId);
  }

  return parts.join("|");
}

function computeGroupedBaseUnitsTotals(deliverables, fieldIdOrIds) {
  const totals = new Map();
  const fieldIds = Array.isArray(fieldIdOrIds)
    ? fieldIdOrIds.filter((fieldId) => typeof fieldId === "string" && !!fieldId)
    : (typeof fieldIdOrIds === "string" && !!fieldIdOrIds ? [fieldIdOrIds] : []);

  if (!Array.isArray(deliverables) || fieldIds.length === 0) {
    return totals;
  }

  deliverables.forEach((item) => {
    const groupKey = buildDeliverableGroupKey(item, fieldIds);
    if (!groupKey) {
      return;
    }

    const value = sanitizeBaseUnits(item?.baseUnits);
    if (value === "") {
      return;
    }

    const current = totals.get(groupKey) || 0;
    totals.set(groupKey, Math.round((current + value) * 100) / 100);
  });

  return totals;
}

function computeProjectIncidenceRatio(baseUnits, totalBaseUnits) {
  const safeBase = sanitizeBaseUnits(baseUnits);
  const safeTotal = sanitizeBaseUnits(totalBaseUnits);
  if (safeBase === "" || safeTotal === "" || safeTotal <= 0) {
    return null;
  }
  return safeBase / safeTotal;
}

function formatProjectIncidence(baseUnits, totalBaseUnits) {
  const ratio = computeProjectIncidenceRatio(baseUnits, totalBaseUnits);
  if (ratio === null) {
    return "";
  }
  return formatPercent(ratio * 100, 2);
}

function computeWeightedProjectProgress(incidenceRatio, progressPercent) {
  if (typeof incidenceRatio !== "number" || !Number.isFinite(incidenceRatio) || incidenceRatio < 0) {
    return null;
  }

  const safeProgress = typeof progressPercent === "number" ? progressPercent : NaN;
  if (!Number.isFinite(safeProgress) || safeProgress < 0 || safeProgress > 100) {
    return null;
  }

  return Math.round(incidenceRatio * safeProgress * 100) / 100;
}

function isDateRangeInvalid(startDate, endDate) {
  const safeStart = sanitizeDateInput(startDate || "");
  const safeEnd = sanitizeDateInput(endDate || "");
  return !!safeStart && !!safeEnd && safeEnd < safeStart;
}

function buildProgressSnapshot(startDate, endDate) {
  const safeStart = sanitizeDateInput(startDate || "");
  const safeEnd = sanitizeDateInput(endDate || "");

  if (!safeStart || !safeEnd) {
    return { label: "Sin rango", tone: "", percent: null };
  }

  if (isDateRangeInvalid(safeStart, safeEnd)) {
    return { label: "Fechas invertidas", tone: "overdue", percent: null };
  }

  const startDay = dateInputToDayIndex(safeStart);
  const endDay = dateInputToDayIndex(safeEnd);
  const today = new Date();
  const todayDate = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
  const todayDay = dateInputToDayIndex(todayDate);

  if (todayDay < startDay) {
    return { label: "0%", tone: "", percent: 0 };
  }

  if (todayDay > endDay) {
    return { label: "100%", tone: "complete", percent: 100 };
  }

  // Avance programado inclusivo: el dia de inicio y el dia de fin cuentan.
  const totalDays = Math.max(1, endDay - startDay + 1);
  const elapsedDays = Math.max(1, todayDay - startDay + 1);
  const percent = Math.max(0, Math.min(100, Math.round((elapsedDays / totalDays) * 100)));

  if (percent >= 100) {
    return { label: "100%", tone: "complete", percent: 100 };
  }
  return { label: `${percent}%`, tone: "running", percent };
}

function dateInputToDayIndex(value) {
  const safe = sanitizeDateInput(value);
  if (!safe) {
    return 0;
  }
  const parts = safe.split("-").map((item) => Number.parseInt(item, 10));
  const utc = Date.UTC(parts[0], parts[1] - 1, parts[2]);
  return Math.floor(utc / 86400000);
}

function formatValueWithCode(value, code) {
  if (value && code) {
    return `${value} [${code}]`;
  }
  return value || code || "";
}

function setStatus(message) {
  els.statusMessage.textContent = message;
}

function setSyncIndicator(text, mode) {
  if (!els.syncIndicator) {
    return;
  }

  els.syncIndicator.textContent = text;
  els.syncIndicator.classList.remove("pending", "error");
  if (mode === "pending") {
    els.syncIndicator.classList.add("pending");
  } else if (mode === "error") {
    els.syncIndicator.classList.add("error");
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function removeDiacritics(value) {
  if (!value || typeof value.normalize !== "function") {
    return value || "";
  }
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeLookup(value) {
  return removeDiacritics(trimOrFallback(value, ""))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function trimOrFallback(value, fallback) {
  if (typeof value !== "string") {
    return fallback;
  }
  const clean = value.trim();
  return clean || fallback;
}

function uid() {
  return `${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}











