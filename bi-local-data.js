"use strict";

(function attachMidpBiLocalData(globalScope) {
  function trimString(value, fallback = "") {
    if (typeof value !== "string") {
      return fallback;
    }
    const clean = value.trim();
    return clean || fallback;
  }

  function normalizePlainObject(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return {};
    }
    return { ...value };
  }

  function createLocalProjectDataService(options) {
    const safeOptions = options && typeof options === "object" ? options : {};
    const project = safeOptions.project || null;
    const ensurePackageControls = typeof safeOptions.ensurePackageControls === "function"
      ? safeOptions.ensurePackageControls
      : () => {};
    const ensureReviewControls = typeof safeOptions.ensureReviewControls === "function"
      ? safeOptions.ensureReviewControls
      : () => {};
    const ensureReviewMilestones = typeof safeOptions.ensureReviewMilestones === "function"
      ? safeOptions.ensureReviewMilestones
      : () => {};
    const getPackageControlFieldIds = typeof safeOptions.getPackageControlFieldIds === "function"
      ? safeOptions.getPackageControlFieldIds
      : () => ({});
    const collectDeliverableMetricsByPackageKey = typeof safeOptions.collectDeliverableMetricsByPackageKey === "function"
      ? safeOptions.collectDeliverableMetricsByPackageKey
      : () => new Map();
    const getFieldRowData = typeof safeOptions.getFieldRowData === "function"
      ? safeOptions.getFieldRowData
      : () => ({ label: "", code: "", name: "" });
    const findFieldIdByAlias = typeof safeOptions.findFieldIdByAlias === "function"
      ? safeOptions.findFieldIdByAlias
      : () => "";
    const getCustomValuesFromRefs = typeof safeOptions.getCustomValuesFromRefs === "function"
      ? safeOptions.getCustomValuesFromRefs
      : () => ({});
    const sanitizeDateInput = typeof safeOptions.sanitizeDateInput === "function"
      ? safeOptions.sanitizeDateInput
      : (value) => trimString(value, "");
    const isDateRangeInvalid = typeof safeOptions.isDateRangeInvalid === "function"
      ? safeOptions.isDateRangeInvalid
      : () => false;
    const buildProgressSnapshot = typeof safeOptions.buildProgressSnapshot === "function"
      ? safeOptions.buildProgressSnapshot
      : () => ({ label: "", percent: null });
    const sanitizeBaseUnits = typeof safeOptions.sanitizeBaseUnits === "function"
      ? safeOptions.sanitizeBaseUnits
      : () => "";
    const sanitizeRealProgress = typeof safeOptions.sanitizeRealProgress === "function"
      ? safeOptions.sanitizeRealProgress
      : () => "";
    const getSourceLabel = typeof safeOptions.getSourceLabel === "function"
      ? safeOptions.getSourceLabel
      : (value) => trimString(value, "");
    const trimOrFallback = typeof safeOptions.trimOrFallback === "function"
      ? safeOptions.trimOrFallback
      : trimString;
    const buildPackageControlCode = typeof safeOptions.buildPackageControlCode === "function"
      ? safeOptions.buildPackageControlCode
      : () => "";
    const computeProjectIncidenceRatio = typeof safeOptions.computeProjectIncidenceRatio === "function"
      ? safeOptions.computeProjectIncidenceRatio
      : () => 0;
    const computeWeightedProjectProgress = typeof safeOptions.computeWeightedProjectProgress === "function"
      ? safeOptions.computeWeightedProjectProgress
      : () => 0;
    const normalizeLookup = typeof safeOptions.normalizeLookup === "function"
      ? safeOptions.normalizeLookup
      : (value) => trimString(value, "").toLowerCase();
    const formatNumberForInput = typeof safeOptions.formatNumberForInput === "function"
      ? safeOptions.formatNumberForInput
      : (value) => trimString(String(value ?? ""), "");
    const formatPercent = typeof safeOptions.formatPercent === "function"
      ? safeOptions.formatPercent
      : (value) => trimString(String(value ?? ""), "");
    const normalizeSource = typeof safeOptions.normalizeSource === "function"
      ? safeOptions.normalizeSource
      : (value) => trimString(value, "all");
    const normalizeCrossFilterScope = typeof safeOptions.normalizeCrossFilterScope === "function"
      ? safeOptions.normalizeCrossFilterScope
      : (value) => trimString(value, "all");
    const normalizeCrossFilters = typeof safeOptions.normalizeCrossFilters === "function"
      ? safeOptions.normalizeCrossFilters
      : () => [];
    const rowMatchesCrossFilter = typeof safeOptions.rowMatchesCrossFilter === "function"
      ? safeOptions.rowMatchesCrossFilter
      : () => true;
    const getGlobalSearchQuery = typeof safeOptions.getGlobalSearchQuery === "function"
      ? safeOptions.getGlobalSearchQuery
      : () => "";

    function buildSearchBlob(row) {
      return normalizeLookup([
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
        row.startDate,
        row.endDate,
        formatNumberForInput(row.baseUnits),
        formatPercent(row.programmedPercent, 2),
        formatPercent(row.realProgress, 2),
        formatPercent(row.weightedProgrammed, 2),
        formatPercent(row.weightedReal, 2),
        row.statusLabel,
        row.customText
      ].join(" "));
    }

    function collectRows() {
      if (!project || typeof project !== "object") {
        return [];
      }

      ensurePackageControls(project);
      ensureReviewControls(project);
      ensureReviewMilestones(project);

      const fields = Array.isArray(project.fields) ? project.fields : [];
      const deliverables = Array.isArray(project.deliverables) ? project.deliverables : [];
      const packageControls = Array.isArray(project.packageControls) ? project.packageControls : [];
      const reviewControls = Array.isArray(project.reviewControls) ? project.reviewControls : [];
      const reviewMilestones = Array.isArray(project.reviewMilestones) ? project.reviewMilestones : [];
      const packageFieldIds = getPackageControlFieldIds(fields);
      const metricsByPackageKey = collectDeliverableMetricsByPackageKey(project, packageFieldIds);
      const milestoneMap = new Map(reviewMilestones.map((item) => [item.id, item]));
      const rows = [];

      const creatorFieldId = findFieldIdByAlias(fields, ["creador"]);
      const phaseFieldId = findFieldIdByAlias(fields, ["fase"]);
      const sectorFieldId = findFieldIdByAlias(fields, ["sector"]);
      const levelFieldId = findFieldIdByAlias(fields, ["nivel"]);
      const typeFieldId = findFieldIdByAlias(fields, ["tipo"]);

      deliverables.forEach((item) => {
        const rowRefs = normalizePlainObject(item?.rowRefs);
        const projectData = getFieldRowData(fields, packageFieldIds.projectFieldId, rowRefs[packageFieldIds.projectFieldId]);
        const disciplineData = getFieldRowData(fields, packageFieldIds.disciplineFieldId, rowRefs[packageFieldIds.disciplineFieldId]);
        const systemData = getFieldRowData(fields, packageFieldIds.systemFieldId, rowRefs[packageFieldIds.systemFieldId]);
        const packageData = getFieldRowData(fields, packageFieldIds.packageFieldId, rowRefs[packageFieldIds.packageFieldId]);
        const creatorData = getFieldRowData(fields, creatorFieldId, rowRefs[creatorFieldId]);
        const phaseData = getFieldRowData(fields, phaseFieldId, rowRefs[phaseFieldId]);
        const sectorData = getFieldRowData(fields, sectorFieldId, rowRefs[sectorFieldId]);
        const levelData = getFieldRowData(fields, levelFieldId, rowRefs[levelFieldId]);
        const typeData = getFieldRowData(fields, typeFieldId, rowRefs[typeFieldId]);
        const customValues = getCustomValuesFromRefs(project, rowRefs);
        const customText = Object.values(customValues).filter((value) => !!value).join(" ");
        const startDate = sanitizeDateInput(item?.startDate || "");
        const endDate = sanitizeDateInput(item?.endDate || "");
        const invalidDates = isDateRangeInvalid(startDate, endDate);
        const progress = invalidDates
          ? { label: "Fechas invertidas", percent: null }
          : buildProgressSnapshot(startDate, endDate);
        const baseUnitsRaw = sanitizeBaseUnits(item?.baseUnits);
        const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
        const realProgressRaw = sanitizeRealProgress(item?.realProgress);
        const realProgress = realProgressRaw === "" ? null : realProgressRaw;
        const row = {
          source: "deliverable",
          sourceLabel: getSourceLabel("deliverable"),
          itemId: item?.id,
          itemLabel: trimOrFallback(item?.code, "") || trimOrFallback(item?.ref, ""),
          ref: trimOrFallback(item?.ref, ""),
          code: trimOrFallback(item?.code, ""),
          startDate,
          endDate,
          monthStart: startDate ? startDate.slice(0, 7) : "",
          monthEnd: endDate ? endDate.slice(0, 7) : "",
          baseUnits,
          programmedPercent: progress.percent,
          realProgress,
          weightedProgrammed: null,
          weightedReal: null,
          invalidDates,
          statusLabel: progress.label,
          createdAt: trimOrFallback(item?.createdAt, ""),
          projectLabel: projectData.label || project.name,
          disciplineLabel: disciplineData.label,
          systemLabel: systemData.label,
          packageLabel: packageData.label,
          milestoneLabel: "",
          creatorLabel: creatorData.label,
          phaseLabel: phaseData.label,
          sectorLabel: sectorData.label,
          levelLabel: levelData.label,
          typeLabel: typeData.label,
          customValues,
          customText
        };
        row.searchBlob = buildSearchBlob(row);
        rows.push(row);
      });

      packageControls.forEach((item) => {
        const projectData = getFieldRowData(fields, packageFieldIds.projectFieldId, item?.projectRowId);
        const disciplineData = getFieldRowData(fields, packageFieldIds.disciplineFieldId, item?.disciplineRowId);
        const systemData = getFieldRowData(fields, packageFieldIds.systemFieldId, item?.systemRowId);
        const packageData = getFieldRowData(fields, packageFieldIds.packageFieldId, item?.packageRowId);
        const startDate = sanitizeDateInput(item?.startDate || "");
        const endDate = sanitizeDateInput(item?.endDate || "");
        const invalidDates = isDateRangeInvalid(startDate, endDate);
        const progress = invalidDates
          ? { label: "Fechas invertidas", percent: null }
          : buildProgressSnapshot(startDate, endDate);
        const metrics = metricsByPackageKey.get(item?.key) || null;
        const baseUnitsRaw = sanitizeBaseUnits(metrics?.baseUnitsTotal);
        const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
        const realProgressRaw = sanitizeRealProgress(item?.realProgress);
        const realProgress = realProgressRaw === "" ? null : realProgressRaw;
        const keyLabel = buildPackageControlCode(project, [
          { code: projectData.code, name: projectData.name },
          { code: disciplineData.code, name: disciplineData.name },
          { code: systemData.code, name: systemData.name },
          { code: packageData.code, name: packageData.name }
        ]) || trimOrFallback(item?.key, "");
        const row = {
          source: "package",
          sourceLabel: getSourceLabel("package"),
          itemId: item?.id,
          itemLabel: keyLabel,
          ref: "",
          code: keyLabel,
          startDate,
          endDate,
          monthStart: startDate ? startDate.slice(0, 7) : "",
          monthEnd: endDate ? endDate.slice(0, 7) : "",
          baseUnits,
          programmedPercent: progress.percent,
          realProgress,
          weightedProgrammed: null,
          weightedReal: null,
          invalidDates,
          statusLabel: progress.label,
          createdAt: trimOrFallback(item?.createdAt, ""),
          projectLabel: projectData.label || project.name,
          disciplineLabel: disciplineData.label,
          systemLabel: systemData.label,
          packageLabel: packageData.label,
          milestoneLabel: "",
          creatorLabel: "",
          phaseLabel: "",
          sectorLabel: "",
          levelLabel: "",
          typeLabel: "",
          customValues: {},
          customText: ""
        };
        row.searchBlob = buildSearchBlob(row);
        rows.push(row);
      });

      reviewControls.forEach((item) => {
        const projectData = getFieldRowData(fields, packageFieldIds.projectFieldId, item?.projectRowId);
        const disciplineData = getFieldRowData(fields, packageFieldIds.disciplineFieldId, item?.disciplineRowId);
        const systemData = getFieldRowData(fields, packageFieldIds.systemFieldId, item?.systemRowId);
        const milestone = milestoneMap.get(item?.milestoneId) || null;
        const milestoneLabel = trimOrFallback(milestone?.name, "");
        const startDate = sanitizeDateInput(item?.startDate || "");
        const endDate = sanitizeDateInput(item?.endDate || "");
        const invalidDates = isDateRangeInvalid(startDate, endDate);
        const progress = invalidDates
          ? { label: "Fechas invertidas", percent: null }
          : buildProgressSnapshot(startDate, endDate);
        const baseUnitsRaw = sanitizeBaseUnits(milestone?.baseUnits ?? item?.baseUnits);
        const baseUnits = baseUnitsRaw === "" ? 0 : baseUnitsRaw;
        const realProgressRaw = sanitizeRealProgress(item?.realProgress);
        const realProgress = realProgressRaw === "" ? null : realProgressRaw;
        const combination = buildPackageControlCode(project, [
          { code: projectData.code, name: projectData.name },
          { code: disciplineData.code, name: disciplineData.name },
          { code: systemData.code, name: systemData.name }
        ]);
        const row = {
          source: "review-control",
          sourceLabel: getSourceLabel("review-control"),
          itemId: item?.id,
          itemLabel: [combination, milestoneLabel].filter((token) => !!token).join(" | "),
          ref: "",
          code: combination,
          startDate,
          endDate,
          monthStart: startDate ? startDate.slice(0, 7) : "",
          monthEnd: endDate ? endDate.slice(0, 7) : "",
          baseUnits,
          programmedPercent: progress.percent,
          realProgress,
          weightedProgrammed: null,
          weightedReal: null,
          invalidDates,
          statusLabel: progress.label,
          createdAt: trimOrFallback(item?.createdAt, ""),
          projectLabel: projectData.label || project.name,
          disciplineLabel: disciplineData.label,
          systemLabel: systemData.label,
          packageLabel: "",
          milestoneLabel,
          creatorLabel: "",
          phaseLabel: "",
          sectorLabel: "",
          levelLabel: "",
          typeLabel: "",
          customValues: {},
          customText: ""
        };
        row.searchBlob = buildSearchBlob(row);
        rows.push(row);
      });

      const totalBySource = new Map();
      rows.forEach((row) => {
        const current = totalBySource.get(row.source) || 0;
        totalBySource.set(row.source, current + (row.baseUnits || 0));
      });

      rows.forEach((row) => {
        const sourceTotal = totalBySource.get(row.source) || 0;
        const incidence = computeProjectIncidenceRatio(row.baseUnits, sourceTotal);
        row.weightedProgrammed = computeWeightedProjectProgress(incidence, row.programmedPercent);
        row.weightedReal = computeWeightedProjectProgress(incidence, row.realProgress);
        row.searchBlob = buildSearchBlob(row);
      });

      return rows;
    }

    function filterRows(rows, config) {
      if (!Array.isArray(rows) || !config) {
        return [];
      }

      const localQuery = normalizeLookup(config.textFilter || "");
      const globalQuery = normalizeLookup(getGlobalSearchQuery());
      const queries = [localQuery, globalQuery].filter((item) => !!item);
      const source = normalizeSource(config.source || "all");
      const startDate = sanitizeDateInput(config.startDate || "");
      const endDate = sanitizeDateInput(config.endDate || "");
      const crossFilterScope = normalizeCrossFilterScope(config.crossFilterScope || "all");
      const crossFilters = normalizeCrossFilters(config.crossFilters);

      return rows.filter((row) => {
        if (source !== "all" && row.source !== source) {
          return false;
        }
        if (config.invalidOnly && !row.invalidDates) {
          return false;
        }
        if (startDate && (!row.startDate || row.startDate < startDate)) {
          return false;
        }
        if (endDate && (!row.endDate || row.endDate > endDate)) {
          return false;
        }
        if (crossFilters.length > 0) {
          const matchesCross = crossFilters.every((filter) => rowMatchesCrossFilter(row, filter, crossFilterScope));
          if (!matchesCross) {
            return false;
          }
        }
        if (queries.length === 0) {
          return true;
        }
        const blob = trimString(row?.searchBlob, "");
        return queries.every((query) => blob.includes(query));
      });
    }

    return Object.freeze({
      buildSearchBlob,
      collectRows,
      filterRows
    });
  }

  globalScope.MIDPBILocalData = Object.freeze({
    createLocalProjectDataService
  });
})(typeof window !== "undefined" ? window : globalThis);
