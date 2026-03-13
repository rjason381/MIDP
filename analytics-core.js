"use strict";

(function attachMidpBiCore(globalScope) {
  const QUERY_VERSION = 1;

  function trimString(value, fallback = "") {
    if (typeof value !== "string") {
      return fallback;
    }
    const clean = value.trim();
    return clean || fallback;
  }

  function clampInteger(value, fallback, min, max) {
    const parsed = Number.parseInt(String(value ?? "").trim(), 10);
    if (Number.isNaN(parsed)) {
      return Math.max(min, Math.min(max, fallback));
    }
    return Math.max(min, Math.min(max, parsed));
  }

  function normalizeStringArray(value) {
    if (!Array.isArray(value)) {
      return [];
    }
    return value
      .map((item) => trimString(item, ""))
      .filter((item) => !!item);
  }

  function normalizePlainObject(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return {};
    }
    return { ...value };
  }

  function normalizeQueryRequest(rawRequest) {
    const source = rawRequest && typeof rawRequest === "object" ? rawRequest : {};
    const filtersConfig = normalizePlainObject(source.filtersConfig);
    const limitFallback = source.limit ?? 0;

    return {
      version: QUERY_VERSION,
      datasetId: trimString(source.datasetId, "default"),
      source: trimString(source.source || filtersConfig.source || "all", "all"),
      fields: normalizeStringArray(source.fields),
      groupBy: normalizeStringArray(source.groupBy),
      filtersConfig,
      limit: clampInteger(limitFallback, 0, 0, 100000)
    };
  }

  function normalizeSchemaField(rawField, index = 0) {
    const source = rawField && typeof rawField === "object" ? rawField : {};
    const fallbackId = `field_${index + 1}`;
    const fallbackLabel = trimString(source.name, fallbackId);
    const roleToken = trimString(source.role, "").toLowerCase();
    const role = roleToken === "measure" ? "measure" : "dimension";
    return {
      id: trimString(source.id, fallbackId),
      name: trimString(source.name, fallbackLabel),
      label: trimString(source.label, fallbackLabel),
      role,
      dataType: trimString(source.dataType || source.type, "string"),
      source: trimString(source.source, "")
    };
  }

  function createQueryResult(rawRows, meta = {}) {
    const rows = Array.isArray(rawRows) ? rawRows.slice() : [];
    const fields = Array.isArray(meta.fields)
      ? meta.fields.map((field, index) => normalizeSchemaField(field, index))
      : [];
    const request = normalizeQueryRequest(meta.request);
    const executionMs = Number.isFinite(meta.executionMs)
      ? Math.max(0, Math.round(meta.executionMs * 100) / 100)
      : 0;

    return {
      version: QUERY_VERSION,
      request,
      source: trimString(meta.source || request.source, "all"),
      fields,
      rows,
      rowCount: rows.length,
      executionMs
    };
  }

  function createLocalQueryAdapter(handlers) {
    const safeHandlers = handlers && typeof handlers === "object" ? handlers : {};
    const getRows = typeof safeHandlers.getRows === "function"
      ? safeHandlers.getRows
      : () => [];
    const getFields = typeof safeHandlers.getFields === "function"
      ? safeHandlers.getFields
      : () => [];

    return {
      kind: "local",
      getSchema(rawRequest) {
        const request = normalizeQueryRequest(rawRequest);
        return getFields(request).map((field, index) => normalizeSchemaField(field, index));
      },
      runQuery(rawRequest) {
        const request = normalizeQueryRequest(rawRequest);
        const startedAt = Date.now();
        const rows = getRows(request);
        const executionMs = Date.now() - startedAt;
        return createQueryResult(rows, {
          request,
          source: request.source,
          fields: getFields(request),
          executionMs
        });
      }
    };
  }

  globalScope.MIDPBICore = Object.freeze({
    QUERY_VERSION,
    normalizeQueryRequest,
    normalizeSchemaField,
    createQueryResult,
    createLocalQueryAdapter
  });
})(typeof window !== "undefined" ? window : globalThis);
