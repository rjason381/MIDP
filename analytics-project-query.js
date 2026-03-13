"use strict";

(function attachMidpBiProjectQuery(globalScope) {
  const FIELDS_BY_SOURCE = Object.freeze({
    deliverable: Object.freeze([
      Object.freeze({ name: "Ref", type: "text" }),
      Object.freeze({ name: "Nomenclatura", type: "text" }),
      Object.freeze({ name: "Proyecto", type: "text", groupBy: "proyecto" }),
      Object.freeze({ name: "Disciplina", type: "text", groupBy: "disciplina" }),
      Object.freeze({ name: "Sistema", type: "text", groupBy: "sistema" }),
      Object.freeze({ name: "Paquete", type: "text", groupBy: "paquete" }),
      Object.freeze({ name: "Creador", type: "text", groupBy: "creador" }),
      Object.freeze({ name: "Fase", type: "text", groupBy: "fase" }),
      Object.freeze({ name: "Sector", type: "text", groupBy: "sector" }),
      Object.freeze({ name: "Nivel", type: "text", groupBy: "nivel" }),
      Object.freeze({ name: "Tipo", type: "text", groupBy: "tipo" }),
      Object.freeze({ name: "Fecha inicio", type: "date" }),
      Object.freeze({ name: "Fecha fin", type: "date" }),
      Object.freeze({ name: "UP Base", type: "number", metric: "baseunits" }),
      Object.freeze({ name: "Programado", type: "percent", metric: "programmedavg" }),
      Object.freeze({ name: "Avance Real", type: "percent", metric: "realavg" }),
      Object.freeze({ name: "Avance Programado Proyecto", type: "calc", metric: "weightedprogrammed" }),
      Object.freeze({ name: "Avance Real Proyecto", type: "calc", metric: "weightedreal" }),
      Object.freeze({ name: "Estado", type: "text" }),
      Object.freeze({ name: "Creado", type: "date" })
    ]),
    package: Object.freeze([
      Object.freeze({ name: "Combinacion", type: "text" }),
      Object.freeze({ name: "Proyecto", type: "text", groupBy: "proyecto" }),
      Object.freeze({ name: "Disciplina", type: "text", groupBy: "disciplina" }),
      Object.freeze({ name: "Sistema", type: "text", groupBy: "sistema" }),
      Object.freeze({ name: "Paquete", type: "text", groupBy: "paquete" }),
      Object.freeze({ name: "Fecha inicio", type: "date" }),
      Object.freeze({ name: "Fecha fin", type: "date" }),
      Object.freeze({ name: "UP Base", type: "number", metric: "baseunits" }),
      Object.freeze({ name: "Programado", type: "percent", metric: "programmedavg" }),
      Object.freeze({ name: "Avance Real", type: "percent", metric: "realavg" }),
      Object.freeze({ name: "Avance Programado Proyecto", type: "calc", metric: "weightedprogrammed" }),
      Object.freeze({ name: "Avance Real Proyecto", type: "calc", metric: "weightedreal" }),
      Object.freeze({ name: "Estado", type: "text" }),
      Object.freeze({ name: "Creado", type: "date" })
    ]),
    "review-control": Object.freeze([
      Object.freeze({ name: "Combinacion", type: "text" }),
      Object.freeze({ name: "Hito", type: "text", groupBy: "hito" }),
      Object.freeze({ name: "Proyecto", type: "text", groupBy: "proyecto" }),
      Object.freeze({ name: "Disciplina", type: "text", groupBy: "disciplina" }),
      Object.freeze({ name: "Sistema", type: "text", groupBy: "sistema" }),
      Object.freeze({ name: "Fecha inicio", type: "date" }),
      Object.freeze({ name: "Fecha fin", type: "date" }),
      Object.freeze({ name: "UP Base", type: "number", metric: "baseunits" }),
      Object.freeze({ name: "Programado", type: "percent", metric: "programmedavg" }),
      Object.freeze({ name: "Avance Real", type: "percent", metric: "realavg" }),
      Object.freeze({ name: "Avance Programado Proyecto", type: "calc", metric: "weightedprogrammed" }),
      Object.freeze({ name: "Avance Real Proyecto", type: "calc", metric: "weightedreal" }),
      Object.freeze({ name: "Estado", type: "text" }),
      Object.freeze({ name: "Creado", type: "date" })
    ])
  });

  const METRIC_CATALOG = Object.freeze([
    Object.freeze({ name: "Record Count", type: "number", metric: "count" }),
    Object.freeze({ name: "UP Base total", type: "number", metric: "baseunits" }),
    Object.freeze({ name: "UP Base promedio", type: "calc", metric: "baseavg" }),
    Object.freeze({ name: "UP Base max", type: "calc", metric: "basemax" }),
    Object.freeze({ name: "UP Base min", type: "calc", metric: "basemin" }),
    Object.freeze({ name: "Avance real promedio", type: "percent", metric: "realavg" }),
    Object.freeze({ name: "Avance real max", type: "percent", metric: "realmax" }),
    Object.freeze({ name: "Avance real min", type: "percent", metric: "realmin" }),
    Object.freeze({ name: "Avance programado promedio", type: "percent", metric: "programmedavg" }),
    Object.freeze({ name: "Avance programado max", type: "percent", metric: "programmedmax" }),
    Object.freeze({ name: "Avance programado min", type: "percent", metric: "programmedmin" }),
    Object.freeze({ name: "Avance real proyecto", type: "percent", metric: "weightedreal" }),
    Object.freeze({ name: "Avance programado proyecto", type: "percent", metric: "weightedprogrammed" }),
    Object.freeze({ name: "Brecha real - programado", type: "calc", metric: "weightedgap" }),
    Object.freeze({ name: "Fechas invertidas", type: "calc", metric: "invaliddates" })
  ]);

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

  function getCoreApi() {
    if (!globalScope.MIDPBICore || typeof globalScope.MIDPBICore !== "object") {
      return null;
    }
    return globalScope.MIDPBICore;
  }

  function createProjectQueryService(options) {
    const safeOptions = options && typeof options === "object" ? options : {};
    const project = safeOptions.project || null;
    const getSourceLabel = typeof safeOptions.getSourceLabel === "function"
      ? safeOptions.getSourceLabel
      : (value) => trimString(value, "");
    const getCustomFieldDefs = typeof safeOptions.getCustomFieldDefs === "function"
      ? safeOptions.getCustomFieldDefs
      : () => [];
    const normalizeSource = typeof safeOptions.normalizeSource === "function"
      ? safeOptions.normalizeSource
      : (value) => trimString(value, "all");
    const collectRows = typeof safeOptions.collectRows === "function"
      ? safeOptions.collectRows
      : () => [];
    const filterRows = typeof safeOptions.filterRows === "function"
      ? safeOptions.filterRows
      : (rows) => (Array.isArray(rows) ? rows.slice() : []);
    const getExternalSchemaFieldDefs = typeof safeOptions.getExternalSchemaFieldDefs === "function"
      ? safeOptions.getExternalSchemaFieldDefs
      : () => [];
    const trimOrFallback = typeof safeOptions.trimOrFallback === "function"
      ? safeOptions.trimOrFallback
      : trimString;

    function getCatalogFields(source = "all") {
      const selectedSource = normalizeSource(source || "all");
      const selectedSources = selectedSource === "all"
        ? ["deliverable", "package", "review-control"]
        : [selectedSource];
      const catalog = [];

      selectedSources.forEach((sourceKey) => {
        const sourceLabel = getSourceLabel(sourceKey);
        const sourceFields = FIELDS_BY_SOURCE[sourceKey] || [];
        sourceFields.forEach((field) => {
          catalog.push({
            source: sourceLabel,
            name: field.name,
            type: field.type,
            groupBy: trimOrFallback(field.groupBy, ""),
            metric: trimOrFallback(field.metric, "")
          });
        });
      });

      if (selectedSource !== "all") {
        getExternalSchemaFieldDefs(project, selectedSource).forEach((field) => {
          catalog.push({
            source: getSourceLabel(selectedSource),
            name: trimOrFallback(field.label || field.name, "Campo"),
            type: trimOrFallback(field.type, "text"),
            groupBy: `field:${trimOrFallback(field.id, "")}`,
            metric: ""
          });
        });
      }

      if (selectedSources.includes("deliverable")) {
        getCustomFieldDefs(project).forEach((field) => {
          catalog.push({
            source: getSourceLabel("deliverable"),
            name: field.label,
            type: "text",
            groupBy: `field:${field.id}`,
            metric: ""
          });
        });
      }

      METRIC_CATALOG.forEach((item) => {
        catalog.push({
          source: "Metricas",
          name: item.name,
          type: item.type,
          groupBy: "",
          metric: item.metric
        });
      });

      const seen = new Set();
      return catalog.filter((item) => {
        const key = `${item.source}|${item.name}`;
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      });
    }

    function getSchemaFields(source = "all") {
      return getCatalogFields(source).map((field, index) => {
        const id = trimOrFallback(
          field.metric || field.groupBy || `${field.source}_${field.name}_${index + 1}`,
          `field_${index + 1}`
        );
        const kind = trimOrFallback(field.metric, "") ? "measure" : "dimension";
        return {
          id,
          name: trimOrFallback(field.name, `Field ${index + 1}`),
          label: trimOrFallback(field.name, `Field ${index + 1}`),
          role: kind,
          dataType: trimOrFallback(field.type, "string"),
          source: trimOrFallback(field.source, "")
        };
      });
    }

    function createAdapter() {
      const core = getCoreApi();
      if (!project || !core || typeof core.createLocalQueryAdapter !== "function") {
        return null;
      }
      return core.createLocalQueryAdapter({
        getFields(request) {
          return getSchemaFields(request?.source || "all");
        },
        getRows(request) {
          const baseConfig = normalizePlainObject(project?.biConfig);
          const filterConfig = normalizePlainObject(request?.filtersConfig);
          const mergedConfig = { ...baseConfig, ...filterConfig };
          return filterRows(collectRows(project), mergedConfig);
        }
      });
    }

    function queryRows(rawConfig = null) {
      if (!project) {
        return [];
      }
      const projectConfig = normalizePlainObject(project?.biConfig);
      const requestConfig = normalizePlainObject(rawConfig);
      const mergedConfig = { ...projectConfig, ...requestConfig };
      const adapter = createAdapter();
      if (!adapter || typeof adapter.runQuery !== "function") {
        return filterRows(collectRows(project), mergedConfig);
      }
      const result = adapter.runQuery({
        datasetId: project.id,
        source: normalizeSource(mergedConfig.source || projectConfig.source || "all"),
        fields: ["*"],
        filtersConfig: mergedConfig
      });
      return Array.isArray(result?.rows) ? result.rows : [];
    }

    function queryQuickSightRows() {
      return queryRows({
        source: "all",
        startDate: "",
        endDate: "",
        textFilter: "",
        invalidOnly: false,
        crossFilters: [],
        crossFilterScope: "all"
      });
    }

    return Object.freeze({
      getCatalogFields,
      getSchemaFields,
      createAdapter,
      queryRows,
      queryQuickSightRows
    });
  }

  globalScope.MIDPBIProjectQuery = Object.freeze({
    createProjectQueryService
  });
})(typeof window !== "undefined" ? window : globalThis);
