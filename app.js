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

let state = loadState();
let activeTab = "fields";
let midpEditMode = false;
let packageEditMode = false;
let chooserLocked = false;
let drawerCloseTimer = null;
let trackingCloseTimer = null;
let trackingPanelTargetId = "";
let trackingPanelTargetType = "";
let pendingSaveTimer = null;
let pendingSearchTimer = null;
let currentSearchQuery = "";
let draggedNomenclatureFieldId = "";
const pendingRealAdvanceLogs = [];
const pendingPackageRealAdvanceLogs = [];
const draftValuesByProject = {};
const els = {};

document.addEventListener("DOMContentLoaded", () => {
  bindElements();
  wireEvents();
  ensureActiveProject();
  renderAll();

  window.addEventListener("beforeunload", () => {
    flushPendingSave();
  });

  if (state.projects.length > 1) {
    openProjectChooser(true);
  }
});

function bindElements() {
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

function wireEvents() {
  els.syncButton.addEventListener("click", () => {
    flushPendingSave();
    saveState(true);
    renderDeliverablesPanel(getActiveProject());
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

  els.switchProjectButton.addEventListener("click", () => {
    if (state.projects.length === 0) {
      setStatus("No hay proyectos configurados.");
      return;
    }
    openProjectChooser(false);
  });

  els.projectTitleInput.addEventListener("input", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }
    project.title = trimOrFallback(els.projectTitleInput.value, `MIDP - ${project.name}`);
    scheduleSaveState();
    renderHeader(project);
  });

  els.fieldSeparatorInput.addEventListener("input", () => {
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
    renderTabState();
    setStatus(project.progressControlMode === "package"
      ? "Control de avance configurado por paquete."
      : "Control de avance configurado por entregable.");
  });

  els.addFieldButton.addEventListener("click", () => {
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
    const project = getActiveProject();
    if (!project) {
      return;
    }

    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
  });

  els.fieldMatrixBody.addEventListener("input", (event) => {
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

  els.fieldMatrixBody.addEventListener("change", () => {
    const project = getActiveProject();
    if (!project) {
      return;
    }

    syncDraftValues(project);
    recomputeAllDeliverableCodes(project, false);
    renderFieldsPanel(project);
    renderDeliverablesPanel(project);
  });

  els.fieldsTableBody.addEventListener("change", (event) => {
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

  els.openAddDeliverableButton.addEventListener("click", () => {
    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      return;
    }

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
    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      return;
    }

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
    if (!midpEditMode) {
      setStatus("Activa Editar para modificar campos en MIDP.");
      return;
    }

    const project = getActiveProject();
    if (!project) {
      return;
    }

    const draft = getDraftValues(project.id);
    const rowRefs = {};
    project.fields.forEach((field) => {
      rowRefs[field.id] = trimOrFallback(draft[field.id], "");
    });

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
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      flushPendingSave();
      saveState(true);
      renderDeliverablesPanel(getActiveProject());
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

    if (event.key !== "Escape") {
      return;
    }

    closeDeliverableDrawer();
    closeTrackingPanel();
  });
}

function renderAll() {
  const project = getActiveProject();
  renderHeader(project);
  renderFieldsPanel(project);
  renderDeliverablesPanel(project);
  renderProjectChooser(project);
  renderTabState();
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
  els.projectTitleInput.value = project.title;
  els.fieldSeparatorInput.value = project.codeSeparator;
  if (els.progressControlModeSelect) {
    els.progressControlModeSelect.value = getProjectProgressControlMode(project);
  }
  els.fieldCounter.textContent = `${project.fields.length} / ${MAX_FIELDS} campos`;

  const rowsToRender = computeVisualRowCount(project.fields);
  const filteredFieldIndexes = getFilteredFieldIndexes(project.fields, currentSearchQuery);
  els.fieldMatrixHead.innerHTML = `<tr>
    <th class="row-index-head">#</th>
    ${project.fields.map((field) => `
      <th class="matrix-field-head">
        <input
          type="text"
          class="matrix-head-input"
          data-field-label="${escapeAttribute(field.id)}"
          value="${escapeAttribute(field.label)}"
          maxlength="40"
          ${field.locked ? "readonly" : ""}>
      </th>
      <th class="matrix-code-head">codigo</th>
    `).join("")}
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
      return `
        <td>
          <input
            type="text"
            class="matrix-cell-input"
            data-matrix-field="${escapeAttribute(field.id)}"
            data-matrix-row="${rowIndex}"
            data-matrix-key="name"
            value="${escapeAttribute(nameValue)}"
            placeholder="Nombre">
        </td>
        <td>
          <input
            type="text"
            class="matrix-cell-input matrix-code-input"
            data-matrix-field="${escapeAttribute(field.id)}"
            data-matrix-row="${rowIndex}"
            data-matrix-key="code"
            value="${escapeAttribute(codeValue)}"
            placeholder="Cod.">
        </td>
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
        .map((field, index) => `<div class="nomenclature-item" draggable="true" data-nomenclature-item="${escapeAttribute(field.id)}" title="Arrastra para ordenar">
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
        : `<button type="button" class="danger" data-remove-field="${escapeAttribute(field.id)}">Quitar</button>`;

      const orderIndex = project.nomenclatureOrder.indexOf(field.id);
      const isIncluded = !!field.includeInCode;
      const orderChip = isIncluded && orderIndex >= 0
        ? `<span class="order-chip">${orderIndex + 1}</span>`
        : "<span class=\"muted\">-</span>";

      return `<tr>
        <td>${index + 1}. ${escapeHtml(field.label)}</td>
        <td><input type="checkbox" data-field-code="${escapeAttribute(field.id)}" ${field.includeInCode ? "checked" : ""}></td>
        <td>${orderChip}</td>
        <td>${removeButton}</td>
      </tr>`;
    })
    .join("");

  renderNomenclaturePreview(project);
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
    renderPackageControlsPanel(project);
    closeDeliverableDrawer();
    closeTrackingPanel();
    return;
  }

  syncDraftValues(project);
  const draft = getDraftValues(project.id);
  const editEnabled = midpEditMode;
  const progressMode = getProjectProgressControlMode(project);
  const showDeliverableProgressColumns = progressMode !== "package";
  const drawerDisabledAttr = editEnabled ? "" : "disabled";

  els.deliverableInputs.innerHTML = project.fields
    .map((field) => {
      const options = getFieldOptions(field);
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
        <select data-deliverable-input="${escapeAttribute(field.id)}" ${drawerDisabledAttr}>
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
    ? `<th>Incidencia Proyecto</th>
      <th>Incidencia Disciplina</th>
      <th>Incidencia Sistema</th>
      <th>Programado</th>
      <th>Avance Programado Proyecto</th>
      <th>Avance Programado Disciplina</th>
      <th>Avance Programado Sistema</th>
      <th>Avance Real</th>
      <th>Avance Real Proyecto</th>
      <th>Avance Real Disciplina</th>
      <th>Avance Real Sistema</th>`
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

          const options = getFieldOptions(field);
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
            <select class="cell-select" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-field="${escapeAttribute(field.id)}">
              <option value=""></option>
              ${optionHtml}
            </select>
          </td>`;
        })
        .join("");

      const startDateCell = editEnabled
        ? `<input type="date" class="cell-date-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="startDate" value="${escapeAttribute(startDate)}">`
        : `<span class="cell-readonly">${escapeHtml(formatDateFromInput(startDate) || "")}</span>`;
      const endDateCell = editEnabled
        ? `<input type="date" class="cell-date-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="endDate" value="${escapeAttribute(endDate)}">`
        : `<span class="cell-readonly">${escapeHtml(formatDateFromInput(endDate) || "")}</span>`;
      const baseUnitsCell = editEnabled
        ? `<input type="number" class="cell-number-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="baseUnits" min="0" step="0.01" value="${escapeAttribute(baseUnitsText)}" placeholder="0">`
        : `<span class="cell-readonly">${escapeHtml(baseUnitsText || "")}</span>`;
      const realProgressCell = editEnabled
        ? `<input type="number" class="cell-number-input" data-edit-deliverable="${escapeAttribute(deliverable.id)}" data-edit-meta="realProgress" min="0" max="100" step="0.01" value="${escapeAttribute(realProgressText)}" placeholder="0">`
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
        ? `<td>${incidenceText ? `<span class="incidence-chip" title="${escapeAttribute(incidenceDetail)}">${escapeHtml(incidenceText)}</span>` : ""}</td>
        <td>${disciplineIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(disciplineIncidenceDetail)}">${escapeHtml(disciplineIncidenceText)}</span>` : ""}</td>
        <td>${systemIncidenceText ? `<span class="incidence-chip" title="${escapeAttribute(systemIncidenceDetail)}">${escapeHtml(systemIncidenceText)}</span>` : ""}</td>
        <td><span class="progress-chip${progressClass}">${escapeHtml(progress.label)}</span></td>
        <td>${programmedProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedProjectDetail)}">${escapeHtml(programmedProjectText)}</span>` : ""}</td>
        <td>${programmedDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedDisciplineDetail)}">${escapeHtml(programmedDisciplineText)}</span>` : ""}</td>
        <td>${programmedSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(programmedSystemDetail)}">${escapeHtml(programmedSystemText)}</span>` : ""}</td>
        <td>${realProgressCell}</td>
        <td>${realProjectText ? `<span class="project-progress-chip" title="${escapeAttribute(realProjectDetail)}">${escapeHtml(realProjectText)}</span>` : ""}</td>
        <td>${realDisciplineText ? `<span class="project-progress-chip" title="${escapeAttribute(realDisciplineDetail)}">${escapeHtml(realDisciplineText)}</span>` : ""}</td>
        <td>${realSystemText ? `<span class="project-progress-chip" title="${escapeAttribute(realSystemDetail)}">${escapeHtml(realSystemText)}</span>` : ""}</td>`
        : "";

      return `<tr data-deliverable-row="${escapeAttribute(deliverable.id)}"${rowClass}>
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

  renderPackageControlsPanel(project);

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

function ensurePackageControls(project) {
  if (!project) {
    return;
  }
  project.packageControls = normalizePackageControls(project.packageControls);
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

function renderTabState() {
  const project = getActiveProject();
  const packageTabEnabled = getProjectProgressControlMode(project) === "package";
  if (!packageTabEnabled && activeTab === "packages") {
    activeTab = "deliverables";
  }

  const fieldPanel = document.getElementById("tab-fields");
  const deliverablesPanel = document.getElementById("tab-deliverables");
  const packagesPanel = document.getElementById("tab-packages");
  const tabButtons = document.querySelectorAll(".tab-button");
  const packagesTabButton = document.querySelector(".tab-button[data-tab=\"packages\"]");

  if (packagesTabButton) {
    packagesTabButton.classList.toggle("hidden", !packageTabEnabled);
  }

  fieldPanel.classList.toggle("hidden", activeTab !== "fields");
  deliverablesPanel.classList.toggle("hidden", activeTab !== "deliverables");
  packagesPanel.classList.toggle("hidden", activeTab !== "packages" || !packageTabEnabled);

  tabButtons.forEach((button) => {
    const isActive = button.dataset.tab === activeTab;
    button.classList.toggle("active", isActive);
  });

  if (els.currentViewLabel) {
    if (activeTab === "fields") {
      els.currentViewLabel.textContent = "LISTAS";
    } else if (activeTab === "packages") {
      els.currentViewLabel.textContent = "CONTROL PAQUETES";
    } else {
      els.currentViewLabel.textContent = "MIDP";
    }
  }

  if (els.globalSearchInput) {
    if (activeTab === "fields") {
      els.globalSearchInput.placeholder = "Buscar en listas";
    } else if (activeTab === "packages") {
      els.globalSearchInput.placeholder = "Buscar en control de paquetes";
    } else {
      els.globalSearchInput.placeholder = "Buscar en entregables";
    }
  }

  updateMidpEditUi();
  updatePackageEditUi();
}

function switchTab(tab) {
  const project = getActiveProject();
  const packageTabEnabled = getProjectProgressControlMode(project) === "package";
  if (tab === "packages") {
    activeTab = packageTabEnabled ? "packages" : "deliverables";
  } else if (tab === "deliverables" || tab === "fields") {
    activeTab = tab;
  } else {
    activeTab = "fields";
  }
  const tabMatchesTracking = (activeTab === "deliverables" && trackingPanelTargetType === "deliverable")
    || (activeTab === "packages" && trackingPanelTargetType === "package");
  if (!tabMatchesTracking) {
    closeTrackingPanel();
  }
  if (activeTab === "packages") {
    renderPackageControlsPanel(getActiveProject());
  }
  renderTabState();
}

function updateMidpEditUi() {
  if (!els.toggleMidpEditButton) {
    return;
  }

  const editEnabled = midpEditMode;
  els.toggleMidpEditButton.textContent = editEnabled ? "Bloquear edicion" : "Editar";
  els.toggleMidpEditButton.classList.toggle("active", editEnabled);

  if (els.openAddDeliverableButton) {
    els.openAddDeliverableButton.disabled = !editEnabled;
  }
  if (els.addBlankDeliverableButton) {
    els.addBlankDeliverableButton.disabled = !editEnabled;
  }
  if (els.recomputeCodesButton) {
    els.recomputeCodesButton.disabled = !editEnabled;
  }
}

function updatePackageEditUi() {
  if (!els.togglePackageEditButton) {
    return;
  }

  els.togglePackageEditButton.textContent = packageEditMode ? "Bloquear edicion" : "Editar";
  els.togglePackageEditButton.classList.toggle("active", packageEditMode);
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
  if (!midpEditMode) {
    return;
  }

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

  const code = buildPackageControlTrackingCode(context.project, context.entity);
  return code ? `Paquete ${code}` : "Paquete";
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

function renderTrackingPanel(project, entity, entityType) {
  ensureTrackingCollections(entity);
  const type = entityType === "package" ? "package" : "deliverable";
  const packageCode = type === "package" ? buildPackageControlTrackingCode(project, entity) : "";
  if (els.trackingPanelTitle) {
    els.trackingPanelTitle.textContent = type === "package"
      ? `Detalle paquete ${packageCode || ""}`.trim()
      : `Detalle ${entity.ref || ""}`.trim();
  }
  if (els.trackingPanelSubtitle) {
    els.trackingPanelSubtitle.textContent = type === "package"
      ? (packageCode || project.title || "")
      : (entity.code || project.title || "");
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

    const first = getFieldOptions(field)[0];
    draft[field.id] = first ? first.id : "";
  });
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

function getFieldOptions(field) {
  return field.rows.filter((row) => hasRowContent(row));
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
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return createDefaultState();
  }

  try {
    const parsed = JSON.parse(raw);
    return normalizeState(parsed);
  } catch (error) {
    return createDefaultState();
  }
}

function scheduleSaveState() {
  saveState();
}

function flushPendingSave() {
  if (pendingSaveTimer === null) {
    return;
  }

  clearTimeout(pendingSaveTimer);
  pendingSaveTimer = null;
}

function saveState(forcePersist) {
  const shouldPersist = !!forcePersist;
  if (!shouldPersist) {
    setSyncIndicator("Cambios pendientes", "pending");
    return;
  }

  try {
    applyPendingRealAdvanceLogs();
    applyPendingPackageRealAdvanceLogs();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
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
  const nomenclatureOrder = normalizeNomenclatureOrder(rawProject?.nomenclatureOrder, fields);
  const deliverables = Array.isArray(rawProject?.deliverables)
    ? rawProject.deliverables.map((item, itemIndex) => normalizeDeliverable(item, fields, itemIndex))
    : [];
  const packageControls = normalizePackageControls(rawProject?.packageControls ?? rawProject?.controlPaquetes);
  const demoSeeded = !!rawProject?.demoSeeded;
  const progressControlMode = normalizeProgressControlMode(
    rawProject?.progressControlMode ?? rawProject?.controlAvance ?? rawProject?.modoControlAvance);

  return {
    id,
    name,
    title,
    codeSeparator,
    fields,
    nomenclatureOrder,
    deliverables,
    packageControls,
    demoSeeded,
    progressControlMode
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
  return { id, name, code };
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

  return {
    id,
    name,
    title: `MIDP - ${name}`,
    codeSeparator: "-",
    fields,
    nomenclatureOrder: fields.filter((field) => field.includeInCode).map((field) => field.id),
    deliverables: [],
    packageControls: [],
    demoSeeded: false,
    progressControlMode: "deliverable"
  };
}

function createField(label, includeInCode, locked, rowTemplates) {
  const rows = Array.isArray(rowTemplates)
    ? rowTemplates.map((row) => createFieldRow(row.name || "", row.code || ""))
    : [];

  return {
    id: uid(),
    label,
    includeInCode,
    locked,
    rows
  };
}

function createFieldRow(name, code) {
  return {
    id: uid(),
    name: trimOrFallback(name, ""),
    code: trimOrFallback(code, "")
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



