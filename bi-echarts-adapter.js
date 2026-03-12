"use strict";

(function attachMidpEChartsAdapter(globalScope) {
  const SUPPORTED_TYPES = new Set([
    "bar",
    "bar_horizontal",
    "bar_stacked",
    "line",
    "area",
    "combo",
    "pie",
    "donut",
    "treemap",
    "funnel",
    "scatter",
    "bubble",
    "waterfall",
    "pareto"
  ]);
  const BAR_FAMILY_TYPES = new Set(["bar", "bar_horizontal", "bar_stacked"]);
  const HORIZONTAL_BAR_TYPES = new Set(["bar_horizontal"]);
  const STACKED_BAR_TYPES = new Set(["bar_stacked"]);
  const PIE_LIKE_TYPES = new Set(["pie", "donut", "funnel"]);
  const NO_GRID_TYPES = new Set(["pie", "donut", "treemap", "funnel"]);

  function trimString(value, fallback = "") {
    if (typeof value !== "string") {
      return fallback;
    }
    const clean = value.trim();
    return clean || fallback;
  }

  function normalizeChartType(value) {
    const token = trimString(value, "").toLowerCase();
    return SUPPORTED_TYPES.has(token) ? token : "";
  }

  function normalizeLegendPosition(value) {
    const token = trimString(value, "").toLowerCase();
    return new Set(["top", "right", "bottom", "left"]).has(token) ? token : "right";
  }

  function normalizeNumber(value, fallback = 0) {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : fallback;
  }

  function clampNumber(value, min, max, fallback) {
    const numeric = normalizeNumber(value, fallback);
    return Math.max(min, Math.min(max, numeric));
  }

  function normalizeAxisLimit(value) {
    if (value === null || value === undefined || String(value).trim() === "") {
      return null;
    }
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : null;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function truncateLabel(value, maxChars) {
    const safe = String(value ?? "");
    const safeMax = Math.max(4, Math.round(normalizeNumber(maxChars, 14)));
    if (safe.length <= safeMax) {
      return safe;
    }
    return `${safe.slice(0, Math.max(1, safeMax - 1))}\u2026`;
  }

  function normalizeColorArray(colors) {
    if (!Array.isArray(colors)) {
      return [];
    }
    return colors
      .map((item) => trimString(item, ""))
      .filter((item) => !!item);
  }

  function getApi() {
    const api = globalScope && globalScope.echarts;
    if (!api || typeof api.init !== "function" || typeof api.getInstanceByDom !== "function") {
      return null;
    }
    return api;
  }

  function canRenderChartType(chartType) {
    return !!normalizeChartType(chartType);
  }

  function getParamSource(params) {
    if (Array.isArray(params)) {
      return params[0] || null;
    }
    return params || null;
  }

  function getRowIndexFromParams(params) {
    const source = getParamSource(params);
    if (!source || typeof source !== "object") {
      return -1;
    }
    const rawIndex = Number(source?.data?.rawIndex);
    if (Number.isInteger(rawIndex) && rawIndex >= 0) {
      return rawIndex;
    }
    const dataIndex = Number(source?.dataIndex);
    if (Number.isInteger(dataIndex) && dataIndex >= 0) {
      return dataIndex;
    }
    return -1;
  }

  function getRowFromParams(payload, params) {
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];
    const rowIndex = getRowIndexFromParams(params);
    return {
      rowIndex,
      row: rowIndex >= 0 && rowIndex < rows.length ? rows[rowIndex] : null
    };
  }

  function disposeQuickSightChart(surface) {
    const api = getApi();
    if (!api || !(surface instanceof HTMLElement)) {
      return false;
    }
    const instance = api.getInstanceByDom(surface);
    if (!instance) {
      return false;
    }
    instance.dispose();
    return true;
  }

  function buildAxisTextStyle(payload) {
    const visual = payload.visualSettings || {};
    return {
      color: "#5d7085",
      fontFamily: trimString(visual.fontFamilyAxis, "Segoe UI"),
      fontSize: clampNumber(visual.fontSizeAxis, 8, 20, 11)
    };
  }

  function buildSeriesLabelStyle(payload) {
    const visual = payload.visualSettings || {};
    return {
      color: "#20354a",
      fontFamily: trimString(visual.fontFamilyLabel, "Segoe UI"),
      fontSize: clampNumber(visual.fontSizeLabels, 8, 20, 11)
    };
  }

  function getLegendEntries(payload, chartType) {
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];
    if (!rows.length || chartType === "treemap") {
      return [];
    }
    if (PIE_LIKE_TYPES.has(chartType)) {
      return Array.from(new Set(rows
        .map((row, index) => trimString(row?.groupLabel || row?.label, `Serie ${index + 1}`))
        .filter((item) => !!item)));
    }
    const breakdownEntries = Array.from(new Set(rows
      .map((row) => trimString(row?.breakdownLabel, ""))
      .filter((item) => !!item)));
    if (breakdownEntries.length > 0) {
      return breakdownEntries;
    }
    const metricLabel = trimString(payload?.metricLabel, "Serie");
    return metricLabel ? [metricLabel] : [];
  }

  function resolveLegend(payload, chartType) {
    const visual = payload.visualSettings || {};
    const legendEntries = getLegendEntries(payload, chartType);
    if (!visual.showLegend || chartType === "treemap" || legendEntries.length === 0) {
      return null;
    }
    const position = normalizeLegendPosition(visual.legendPosition);
    const vertical = position === "left" || position === "right";
    const singleMetricLegend = !PIE_LIKE_TYPES.has(chartType) && legendEntries.length <= 1;
    const orient = singleMetricLegend ? "horizontal" : (vertical ? "vertical" : "horizontal");
    if (singleMetricLegend) {
      return {
        show: true,
        type: "plain",
        orient,
        top: 8,
        right: 12,
        data: legendEntries,
        textStyle: {
          fontFamily: trimString(visual.fontFamilyLabel, "Segoe UI"),
          fontSize: clampNumber(visual.fontSizeLabels, 8, 20, 11),
          color: "#516579"
        }
      };
    }
    return {
      show: true,
      type: legendEntries.length > 8 ? "scroll" : "plain",
      orient,
      top: position === "top" ? 8 : (position === "bottom" ? null : 12),
      bottom: position === "bottom" ? 8 : null,
      left: position === "left" ? 8 : (position === "right" ? null : 18),
      right: position === "right" ? 8 : null,
      data: legendEntries,
      textStyle: {
        fontFamily: trimString(visual.fontFamilyLabel, "Segoe UI"),
        fontSize: clampNumber(visual.fontSizeLabels, 8, 20, 11),
        color: "#516579"
      }
    };
  }

  function buildTooltipFormatter(payload) {
    const visual = payload.visualSettings || {};
    const chartType = normalizeChartType(payload?.chartType);
    const compactMode = trimString(visual.tooltipMode, "").toLowerCase() === "compact";
    return function formatter(params) {
      if (Array.isArray(params) && params.length > 0 && chartType !== "pareto") {
        const axisLabel = trimString(params[0]?.axisValueLabel || params[0]?.name, "");
        const parts = axisLabel ? [`<strong>${escapeHtml(axisLabel)}</strong>`] : [];
        params.forEach((item) => {
          const resolvedItem = getRowFromParams(payload, item);
          const seriesName = trimString(item?.seriesName, trimString(payload.metricLabel, "Serie"));
          const rawValue = Array.isArray(item?.data?.value) ? item.data.value[1] : item?.data?.value;
          const valueText = resolveFormattedValueText(payload, resolvedItem.row || null, item, { rawValue });
          if (!valueText && !Number.isFinite(normalizeNumber(rawValue, Number.NaN))) {
            return;
          }
          parts.push(`${escapeHtml(seriesName)}: ${escapeHtml(valueText || String(rawValue ?? ""))}`);
        });
        return parts.join("<br>");
      }
      const source = getParamSource(params);
      const resolved = getRowFromParams(payload, params);
      const row = resolved.row || {};
      const label = row.groupLabel || row.label || source?.name || "";
      const rawValue = Array.isArray(source?.value) ? source.value[1] : source?.value;
      const valueText = resolveFormattedValueText(payload, row, params, { rawValue });
      if (chartType === "pareto" && Array.isArray(params)) {
        const cumulativeItem = params.find((item) => trimString(item?.seriesName, "").toLowerCase() === "acumulado") || null;
        const cumulativeValue = Number(cumulativeItem?.data?.paretoPercent);
        const parts = [
          `<strong>${escapeHtml(label)}</strong>`,
          `Valor: ${escapeHtml(valueText)}`
        ];
        if (Number.isFinite(cumulativeValue)) {
          parts.push(`Acumulado: ${escapeHtml(formatPercentValue(cumulativeValue))}`);
        }
        if (row.breakdownLabel) {
          parts.push(`Grupo: ${escapeHtml(row.breakdownLabel)}`);
        }
        return parts.join("<br>");
      }
      if (compactMode) {
        return `${escapeHtml(label)}: <strong>${escapeHtml(valueText)}</strong>`;
      }
      const parts = [
        `<strong>${escapeHtml(label)}</strong>`,
        `Valor: ${escapeHtml(valueText)}`
      ];
      if ((chartType === "scatter" || chartType === "bubble") && Number.isFinite(row.xValue)) {
        parts.push(`Eje X: ${escapeHtml(String(row.xValue))}`);
      }
      if (chartType === "bubble" && Number.isFinite(row.bubbleSize)) {
        parts.push(`Tam. burbuja: ${escapeHtml(String(row.bubbleSize))}`);
      }
      if (row.breakdownLabel) {
        parts.push(`Grupo: ${escapeHtml(row.breakdownLabel)}`);
      }
      if (Number.isFinite(row.count)) {
        parts.push(`Filas: ${escapeHtml(String(row.count))}`);
      }
      if (Number.isFinite(row.baseUnits)) {
        parts.push(`UP base: ${escapeHtml(String(row.baseUnits))}`);
      }
      return parts.join("<br>");
    };
  }

  function isPercentMetric(metric) {
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
    ]).has(trimString(metric, "").toLowerCase());
  }

  function formatNumberValue(value, decimals) {
    const safeValue = normalizeNumber(value, 0);
    const digits = Math.max(0, Math.min(4, Math.round(normalizeNumber(decimals, 1))));
    try {
      return new Intl.NumberFormat("es-PE", {
        minimumFractionDigits: 0,
        maximumFractionDigits: digits
      }).format(safeValue);
    } catch (_error) {
      return String(Math.round(safeValue * (10 ** digits)) / (10 ** digits));
    }
  }

  function formatPercentValue(value, decimals = 1) {
    const safeValue = normalizeNumber(value, 0);
    const digits = Math.max(0, Math.min(4, Math.round(normalizeNumber(decimals, 1))));
    return `${safeValue.toFixed(digits).replace(/\.?0+$/, "")}%`;
  }

  function formatVisualValue(value, payload, options = {}) {
    const visual = payload?.visualSettings || {};
    const decimals = clampNumber(visual.valueDecimals, 0, 4, 1);
    const numberFormat = trimString(visual.numberFormat, "auto").toLowerCase();
    const percentHint = options.percentHint === true;
    if (!Number.isFinite(normalizeNumber(value, Number.NaN))) {
      return "";
    }
    if (numberFormat === "percent" || (numberFormat === "auto" && percentHint)) {
      return formatPercentValue(value, decimals);
    }
    if (numberFormat === "currency_pen") {
      return `S/ ${formatNumberValue(value, decimals)}`;
    }
    if (numberFormat === "hours") {
      return `${formatNumberValue(value, decimals)} h`;
    }
    return formatNumberValue(value, decimals);
  }

  function resolveFormattedValueText(payload, row, params, options = {}) {
    const chartType = normalizeChartType(payload?.chartType);
    const source = getParamSource(params);
    const visual = payload?.visualSettings || {};
    const numberFormat = trimString(visual.numberFormat, "auto").toLowerCase();
    const rawValue = Number.isFinite(normalizeNumber(options.rawValue, Number.NaN))
      ? normalizeNumber(options.rawValue, 0)
      : normalizeNumber(Array.isArray(source?.value) ? source.value[1] : source?.value, Number.NaN);
    const piePercent = Number.isFinite(normalizeNumber(options.piePercent, Number.NaN))
      ? normalizeNumber(options.piePercent, 0)
      : normalizeNumber(source?.percent, Number.NaN);
    if ((chartType === "pie" || chartType === "donut") && numberFormat === "percent" && Number.isFinite(piePercent)) {
      return formatPercentValue(piePercent, clampNumber(visual.valueDecimals, 0, 4, 1));
    }
    const formatted = Number.isFinite(rawValue)
      ? formatVisualValue(rawValue, payload, { percentHint: isPercentMetric(payload?.metric) })
      : "";
    return formatted || trimString(row?.valueText, String(rawValue ?? ""));
  }

  function formatCompactNumber(value) {
    const safeValue = normalizeNumber(value, 0);
    try {
      return new Intl.NumberFormat("es-PE", {
        maximumFractionDigits: Math.abs(safeValue) >= 100 ? 0 : 2
      }).format(safeValue);
    } catch (_error) {
      return String(Math.round(safeValue * 100) / 100);
    }
  }

  function buildAxisLabelOptions(payload) {
    const visual = payload.visualSettings || {};
    const labelMaxChars = clampNumber(visual.labelMaxChars, 4, 40, 16);
    return {
      ...buildAxisTextStyle(payload),
      formatter(value) {
        return truncateLabel(value, labelMaxChars);
      }
    };
  }

  function buildDataLabelFormatter(payload) {
    return function formatter(params) {
      const resolved = getRowFromParams(payload, params);
      const row = resolved.row || null;
      if (!row) {
        return resolveFormattedValueText(payload, null, params);
      }
      return resolveFormattedValueText(payload, row, params, { rawValue: row.value });
    };
  }

  function buildSeriesValueLabelFormatter(payload) {
    return function formatter(params) {
      const resolved = getRowFromParams(payload, params);
      const rawValue = Number.isFinite(normalizeNumber(params?.data?.value, Number.NaN))
        ? normalizeNumber(params.data.value, 0)
        : normalizeNumber(Array.isArray(params?.value) ? params.value[1] : params?.value, Number.NaN);
      if (!Number.isFinite(normalizeNumber(rawValue, Number.NaN)) || normalizeNumber(rawValue, 0) === 0) {
        return "";
      }
      return resolveFormattedValueText(payload, resolved.row || null, params, { rawValue });
    };
  }

  function buildValueLabelFormatter(payload) {
    return function formatter(params) {
      const resolved = getRowFromParams(payload, params);
      if (!resolved.row) {
        return resolveFormattedValueText(payload, null, params);
      }
      return resolveFormattedValueText(payload, resolved.row, params, { rawValue: resolved.row.value });
    };
  }

  function buildNameValueLabelFormatter(payload, includeValue) {
    const visual = payload.visualSettings || {};
    const labelMaxChars = clampNumber(visual.labelMaxChars, 4, 40, 16);
    const valueFormatter = buildValueLabelFormatter(payload);
    return function formatter(params) {
      const source = getParamSource(params);
      const label = truncateLabel(source?.name || "", labelMaxChars);
      if (!includeValue) {
        return label;
      }
      return `${label}\n${valueFormatter(params)}`;
    };
  }

  function buildTreemapLabelFormatter(payload, showValueText) {
    const visual = payload.visualSettings || {};
    const labelMaxChars = clampNumber(visual.labelMaxChars, 4, 40, 14);
    const valueFormatter = buildValueLabelFormatter(payload);
    return function formatter(params) {
      const source = getParamSource(params);
      const label = truncateLabel(source?.name || "", labelMaxChars);
      if (!visual.showDataLabels || !showValueText) {
        return label;
      }
      return `${label}\n${valueFormatter(params)}`;
    };
  }

  function buildGrid(payload, chartType) {
    if (NO_GRID_TYPES.has(chartType)) {
      return undefined;
    }
    const visual = payload.visualSettings || {};
    const legendEntries = getLegendEntries(payload, chartType);
    const position = normalizeLegendPosition(visual.legendPosition);
    const hasLegend = !!visual.showLegend && legendEntries.length > 0;
    const grid = {
      top: 20,
      right: 14,
      bottom: 42,
      left: 54,
      containLabel: true
    };
    if (!hasLegend) {
      return grid;
    }
    if (position === "top") {
      grid.top += legendEntries.length > 6 ? 42 : 28;
    } else if (position === "bottom") {
      grid.bottom += legendEntries.length > 6 ? 50 : 30;
    } else if (position === "left" && legendEntries.length > 1) {
      grid.left += legendEntries.length > 10 ? 152 : 128;
    } else if (position === "right" && legendEntries.length > 1) {
      grid.right += legendEntries.length > 10 ? 152 : 128;
    }
    return grid;
  }

  function buildTargetLine(payload) {
    const visual = payload.visualSettings || {};
    const chartType = normalizeChartType(payload?.chartType);
    const value = Number(visual.targetLineValue);
    if (!visual.showTargetLine || !Number.isFinite(value)) {
      return undefined;
    }
    const axisKey = HORIZONTAL_BAR_TYPES.has(chartType) ? "xAxis" : "yAxis";
    return {
      symbol: "none",
      silent: true,
      label: {
        show: true,
        formatter: trimString(visual.targetLineLabel, "Objetivo"),
        color: trimString(visual.targetLineColor, "#d45555"),
        fontSize: clampNumber(visual.fontSizeLabels, 8, 20, 11)
      },
      lineStyle: {
        color: trimString(visual.targetLineColor, "#d45555"),
        width: 1.4,
        type: "dashed"
      },
      data: [{ [axisKey]: value }]
    };
  }

  function buildPieSeries(payload, chartType) {
    const visual = payload.visualSettings || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const typeConfig = payload.typeConfig || {};
    const radius = chartType === "donut"
      ? [`${Math.round(clampNumber(typeConfig.innerRadiusRatio, 0.35, 0.8, 0.58) * 100)}%`, "78%"]
      : ["0%", "78%"];
    return {
      name: trimString(payload.metricLabel, "Serie"),
      type: "pie",
      radius,
      center: ["50%", "54%"],
      avoidLabelOverlap: true,
      label: {
        show: !!visual.showDataLabels,
        ...buildSeriesLabelStyle(payload),
        formatter: buildDataLabelFormatter(payload)
      },
      labelLine: {
        show: !!visual.showDataLabels
      },
      data: rows.map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        name: row?.groupLabel || row?.label || `Serie ${index + 1}`,
        rawIndex: index,
        itemStyle: {
          color: payload.colors[index] || undefined
        }
      }))
    };
  }

  function buildTreemapSeries(payload) {
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const typeConfig = payload.typeConfig || {};
    const minTilePercent = clampNumber(typeConfig.minTilePercent, 0, 25, 1) / 100;
    const showValueText = typeConfig.showValueText !== false;
    const positiveTotal = rows.reduce((sum, row) => {
      const value = normalizeNumber(row?.value, 0);
      return value > 0 ? sum + value : sum;
    }, 0);
    let data = rows
      .map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        name: row?.groupLabel || row?.label || `Item ${index + 1}`,
        rawIndex: index,
        itemStyle: {
          color: payload.colors[index] || undefined
        }
      }))
      .filter((item) => item.value > 0);
    if (positiveTotal > 0 && minTilePercent > 0) {
      const filtered = data.filter((item) => (item.value / positiveTotal) >= minTilePercent);
      if (filtered.length > 0) {
        data = filtered;
      }
    }
    return {
      name: trimString(payload.metricLabel, "Serie"),
      type: "treemap",
      roam: false,
      nodeClick: false,
      breadcrumb: { show: false },
      label: {
        show: true,
        overflow: "truncate",
        ...buildSeriesLabelStyle(payload),
        formatter: buildTreemapLabelFormatter(payload, showValueText)
      },
      upperLabel: { show: false },
      itemStyle: {
        borderColor: "rgba(255,255,255,0.9)",
        borderWidth: 2,
        gapWidth: 2
      },
      emphasis: {
        itemStyle: {
          borderColor: "#ffffff",
          borderWidth: 3
        }
      },
      data
    };
  }

  function buildFunnelSeries(payload) {
    const visual = payload.visualSettings || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const typeConfig = payload.typeConfig || {};
    const minSegmentPercent = clampNumber(typeConfig.minSegmentPercent, 0, 60, 2) / 100;
    const sortMode = trimString(typeConfig.sortMode, "desc").toLowerCase() === "asc" ? "asc" : "desc";
    const sorted = rows
      .map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        name: row?.groupLabel || row?.label || `Paso ${index + 1}`,
        rawIndex: index,
        itemStyle: {
          color: payload.colors[index] || undefined
        }
      }))
      .sort((left, right) => sortMode === "asc" ? left.value - right.value : right.value - left.value);
    const maxValue = Math.max(1, ...sorted.map((item) => item.value));
    let data = sorted.filter((item) => (item.value / maxValue) >= minSegmentPercent);
    if (data.length === 0) {
      data = sorted.slice(0, 1);
    }
    return {
      name: trimString(payload.metricLabel, "Serie"),
      type: "funnel",
      sort: sortMode === "asc" ? "ascending" : "descending",
      top: 18,
      bottom: 18,
      left: 16,
      width: "72%",
      minSize: "10%",
      maxSize: "100%",
      gap: 3,
      label: {
        show: true,
        position: visual.showDataLabels ? "inside" : "right",
        ...buildSeriesLabelStyle(payload),
        formatter: buildNameValueLabelFormatter(payload, !!visual.showDataLabels)
      },
      labelLine: {
        show: !visual.showDataLabels,
        length: 8,
        lineStyle: {
          color: "rgba(93, 112, 133, 0.6)"
        }
      },
      itemStyle: {
        borderColor: "#ffffff",
        borderWidth: 2
      },
      emphasis: {
        focus: "series"
      },
      data
    };
  }

  function buildCartesianSeries(payload, chartType) {
    const visual = payload.visualSettings || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const firstColor = payload.colors[0] || "#2f7ed8";
    const lineStyleType = trimString(visual.seriesLineStyle, "solid");
    const markerStyle = trimString(visual.seriesMarkerStyle, "circle").toLowerCase();
    const isBarFamily = BAR_FAMILY_TYPES.has(chartType);
    const showSymbol = !isBarFamily && markerStyle !== "none";
    const canUseBreakdownSeries = new Set(["bar", "bar_horizontal", "bar_stacked", "line", "area"]).has(chartType)
      && rows.some((row) => !!trimString(row?.breakdownLabel, ""));
    const targetLine = buildTargetLine(payload);
    if (canUseBreakdownSeries) {
      const categories = [];
      const categoryIndex = new Map();
      rows.forEach((row, index) => {
        const category = trimString(row?.groupLabel || row?.label, `Item ${index + 1}`);
        if (!categoryIndex.has(category)) {
          categoryIndex.set(category, categories.length);
          categories.push(category);
        }
      });
      const seriesOrder = [];
      const seriesBuckets = new Map();
      rows.forEach((row, index) => {
        const seriesName = trimString(row?.breakdownLabel, "Serie");
        const category = trimString(row?.groupLabel || row?.label, `Item ${index + 1}`);
        const categorySlot = categoryIndex.get(category);
        if (!Number.isInteger(categorySlot)) {
          return;
        }
        if (!seriesBuckets.has(seriesName)) {
          seriesOrder.push(seriesName);
          seriesBuckets.set(seriesName, {
            color: payload.colors[index] || payload.colors[seriesOrder.length - 1] || firstColor,
            data: categories.map(() => ({ value: 0, rawIndex: -1, valueText: "" }))
          });
        }
        const bucket = seriesBuckets.get(seriesName);
        bucket.data[categorySlot] = {
          value: normalizeNumber(row?.value, 0),
          rawIndex: index,
          valueText: trimString(row?.valueText, ""),
          itemStyle: {
            color: payload.colors[index] || bucket.color
          }
        };
      });
      return seriesOrder.map((seriesName, seriesIndex) => {
        const bucket = seriesBuckets.get(seriesName);
        const seriesColor = bucket?.color || payload.colors[seriesIndex] || firstColor;
        const series = {
          name: seriesName,
          type: isBarFamily ? "bar" : "line",
          smooth: !isBarFamily && !!visual.smoothLines,
          showSymbol,
          symbol: showSymbol ? markerStyle : "circle",
          symbolSize: clampNumber(visual.seriesMarkerSize, 2, 18, 6),
          lineStyle: {
            color: seriesColor,
            width: clampNumber(visual.lineWidth, 1, 6, 2),
            type: lineStyleType
          },
          itemStyle: {
            color: seriesColor
          },
          label: {
            show: !!visual.showDataLabels,
            ...buildSeriesLabelStyle(payload),
            formatter: buildSeriesValueLabelFormatter(payload)
          },
          emphasis: {
            focus: "series"
          },
          data: bucket?.data || []
        };
        if (chartType === "area") {
          series.areaStyle = {
            color: seriesColor,
            opacity: clampNumber(visual.areaOpacity, 0.05, 0.8, 0.2)
          };
        }
        if (isBarFamily) {
          series.barWidth = `${Math.round(clampNumber(visual.barWidthRatio, 0.2, 0.9, 0.64) * 100)}%`;
          series.lineStyle = undefined;
          series.showSymbol = false;
        }
        if (trimString(visual.stackMode, "none") !== "none" || STACKED_BAR_TYPES.has(chartType)) {
          series.stack = "qs-stack";
        }
        if (targetLine && seriesIndex === 0) {
          series.markLine = targetLine;
        }
        return series;
      });
    }
    const common = {
      name: trimString(payload.metricLabel, "Serie"),
      type: isBarFamily ? "bar" : "line",
      smooth: !isBarFamily && !!visual.smoothLines,
      showSymbol,
      symbol: showSymbol ? markerStyle : "circle",
      symbolSize: clampNumber(visual.seriesMarkerSize, 2, 18, 6),
      lineStyle: {
        color: firstColor,
        width: clampNumber(visual.lineWidth, 1, 6, 2),
        type: lineStyleType
      },
      itemStyle: {
        color: firstColor
      },
      label: {
        show: !!visual.showDataLabels,
        ...buildSeriesLabelStyle(payload),
        formatter: buildDataLabelFormatter(payload)
      },
      emphasis: {
        focus: "series"
      },
      data: rows.map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        rawIndex: index,
        valueText: trimString(row?.valueText, ""),
        itemStyle: {
          color: payload.colors[index] || firstColor
        }
      }))
    };
    if (chartType === "area") {
      common.areaStyle = {
        color: firstColor,
        opacity: clampNumber(visual.areaOpacity, 0.05, 0.8, 0.2)
      };
    }
    if (isBarFamily) {
      common.barWidth = `${Math.round(clampNumber(visual.barWidthRatio, 0.2, 0.9, 0.64) * 100)}%`;
      common.lineStyle = undefined;
      common.showSymbol = false;
    }
    if (trimString(visual.stackMode, "none") !== "none" || STACKED_BAR_TYPES.has(chartType)) {
      common.stack = "qs-stack";
    }
    if (targetLine) {
      common.markLine = targetLine;
    }
    return common;
  }

  function buildScatterSeries(payload, bubbleMode) {
    const visual = payload.visualSettings || {};
    const typeConfig = payload.typeConfig || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const markerStyle = trimString(visual.seriesMarkerStyle, "circle").toLowerCase();
    const pointAlpha = clampNumber(typeConfig.pointAlpha, 0.15, 1, bubbleMode ? 0.95 : 1);
    const minRadius = bubbleMode
      ? clampNumber(typeConfig.minBubbleRadius, 3, 60, 6)
      : clampNumber(typeConfig.minPointRadius, 3, 40, 7);
    const maxRadius = bubbleMode
      ? clampNumber(typeConfig.maxBubbleRadius, minRadius, 80, 40)
      : minRadius;
    const series = {
      name: trimString(payload.metricLabel, "Serie"),
      type: "scatter",
      symbol: markerStyle === "none" ? "circle" : markerStyle,
      label: {
        show: !!visual.showDataLabels,
        position: "top",
        ...buildSeriesLabelStyle(payload),
        formatter: buildValueLabelFormatter(payload)
      },
      emphasis: {
        focus: "series"
      },
      data: rows.map((row, index) => {
        const xValue = normalizeNumber(row?.xValue, index + 1);
        const yValue = normalizeNumber(row?.value, 0);
        const bubbleSize = bubbleMode
          ? clampNumber(row?.bubbleSize, minRadius, maxRadius, minRadius)
          : minRadius;
        return {
          value: bubbleMode ? [xValue, yValue, bubbleSize] : [xValue, yValue],
          rawIndex: index,
          symbolSize: Math.max(6, bubbleSize * 2),
          itemStyle: {
            color: payload.colors[index] || undefined,
            opacity: pointAlpha
          }
        };
      })
    };
    const targetLine = buildTargetLine(payload);
    if (targetLine) {
      series.markLine = targetLine;
    }
    return series;
  }

  function buildComboSeries(payload) {
    const visual = payload.visualSettings || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const lineColor = trimString(payload.colors?.[0], "#1f5ea8");
    const markerStyle = trimString(visual.seriesMarkerStyle, "circle").toLowerCase();
    const showSymbol = markerStyle !== "none";
    const targetLine = buildTargetLine(payload);
    const barSeries = {
      name: `${trimString(payload.metricLabel, "Serie")} barras`,
      type: "bar",
      barWidth: `${Math.round(clampNumber(visual.barWidthRatio, 0.2, 0.9, 0.64) * 100)}%`,
      label: {
        show: !!visual.showDataLabels,
        position: "top",
        ...buildSeriesLabelStyle(payload),
        formatter: buildValueLabelFormatter(payload)
      },
      data: rows.map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        rawIndex: index,
        itemStyle: {
          color: payload.colors[index] || undefined,
          opacity: 0.45
        }
      }))
    };
    const lineSeries = {
      name: `${trimString(payload.metricLabel, "Serie")} linea`,
      type: "line",
      smooth: !!visual.smoothLines,
      showSymbol,
      symbol: showSymbol ? markerStyle : "circle",
      symbolSize: clampNumber(visual.seriesMarkerSize, 2, 18, 6),
      lineStyle: {
        color: lineColor,
        width: clampNumber(visual.lineWidth, 1, 6, 2),
        type: trimString(visual.seriesLineStyle, "solid")
      },
      itemStyle: {
        color: lineColor
      },
      emphasis: {
        focus: "series"
      },
      data: rows.map((row, index) => ({
        value: normalizeNumber(row?.value, 0),
        rawIndex: index
      }))
    };
    if (targetLine) {
      lineSeries.markLine = targetLine;
    }
    return [barSeries, lineSeries];
  }

  function buildWaterfallChartMeta(payload) {
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const items = [];
    let running = 0;
    let minValue = 0;
    let maxValue = 0;
    rows.forEach((row, index) => {
      const delta = normalizeNumber(row?.value, 0);
      const start = running;
      const end = running + delta;
      items.push({
        index,
        rawIndex: index,
        start,
        end,
        delta,
        label: row?.groupLabel || row?.label || `Item ${index + 1}`,
        valueText: row?.valueText || formatCompactNumber(delta)
      });
      running = end;
      minValue = Math.min(minValue, start, end);
      maxValue = Math.max(maxValue, start, end);
    });
    const spread = Math.max(1, maxValue - minValue);
    return {
      items,
      total: running,
      minValue: minValue - (spread * 0.08),
      maxValue: maxValue + (spread * 0.08)
    };
  }

  function buildWaterfallSeries(payload) {
    const visual = payload.visualSettings || {};
    return {
      name: trimString(payload.metricLabel, "Serie"),
      type: "custom",
      renderItem(params, api) {
        const categoryIndex = api.value(0);
        const startCoord = api.coord([categoryIndex, api.value(1)]);
        const endCoord = api.coord([categoryIndex, api.value(2)]);
        const delta = normalizeNumber(api.value(3), 0);
        const rawIndex = Math.max(0, Math.round(normalizeNumber(api.value(4), categoryIndex)));
        const valueText = trimString(api.value(5), formatCompactNumber(delta));
        const barWidth = Math.max(6, api.size([1, 0])[0] * clampNumber(visual.barWidthRatio, 0.2, 0.9, 0.58));
        const x = startCoord[0] - (barWidth / 2);
        const y = Math.min(startCoord[1], endCoord[1]);
        const height = Math.max(2, Math.abs(startCoord[1] - endCoord[1]));
        const color = trimString(payload.colors?.[rawIndex], delta >= 0 ? "#2f7ed8" : "#e06b6b");
        const children = [{
          type: "rect",
          transition: ["shape"],
          shape: {
            x,
            y,
            width: barWidth,
            height
          },
          style: api.style({
            fill: color,
            stroke: color,
            lineWidth: 1
          })
        }];
        if (visual.showDataLabels) {
          children.push({
            type: "text",
            silent: true,
            style: {
              x: startCoord[0],
              y: Math.max(10, y - 6),
              text: valueText,
              textAlign: "center",
              textVerticalAlign: "bottom",
              fill: "#20354a",
              font: `${clampNumber(visual.fontSizeLabels, 8, 20, 11)}px ${trimString(visual.fontFamilyLabel, "Segoe UI")}`
            }
          });
        }
        return {
          type: "group",
          children
        };
      },
      encode: {
        x: 0,
        y: [1, 2],
        tooltip: 3
      },
      data: buildWaterfallChartMeta(payload).items.map((item) => ([
        item.index,
        item.start,
        item.end,
        item.delta,
        item.rawIndex,
        item.valueText
      ])),
      z: 3
    };
  }

  function buildWaterfallTotalGraphic(payload, total) {
    const typeConfig = payload.typeConfig || {};
    if (typeConfig.showTotalBar === false) {
      return undefined;
    }
    return [{
      type: "text",
      top: 8,
      right: 10,
      style: {
        text: `${trimString(typeConfig.totalLabel, "Total")}: ${formatCompactNumber(total)}`,
        textAlign: "right",
        fill: "#1f4d78",
        font: `600 11px ${trimString(payload?.visualSettings?.fontFamilyLabel, "Segoe UI")}`
      },
      silent: true
    }];
  }

  function buildParetoSeries(payload) {
    const visual = payload.visualSettings || {};
    const typeConfig = payload.typeConfig || {};
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const safeValues = rows.map((row) => Math.max(0, normalizeNumber(row?.value, 0)));
    const total = safeValues.reduce((sum, value) => sum + value, 0);
    const cumulativePercentages = [];
    let running = 0;
    safeValues.forEach((value) => {
      running += value;
      cumulativePercentages.push(total <= 0 ? 0 : ((running / total) * 100));
    });
    const lineColor = trimString(payload.colors?.[0], "#235fa3");
    const showCumulativeLine = typeConfig.showCumulativeLine !== false;
    const targetPercent = clampNumber(typeConfig.targetPercent, 1, 100, 80);
    return [
      {
        name: trimString(payload.metricLabel, "Valor"),
        type: "bar",
        barWidth: `${Math.round(clampNumber(visual.barWidthRatio, 0.2, 0.9, 0.62) * 100)}%`,
        itemStyle: {
          borderRadius: [4, 4, 0, 0]
        },
        label: {
          show: !!visual.showDataLabels,
          position: "top",
          ...buildSeriesLabelStyle(payload),
          formatter: buildValueLabelFormatter(payload)
        },
        data: rows.map((row, index) => ({
          value: safeValues[index],
          rawIndex: index,
          itemStyle: {
            color: payload.colors[index] || undefined
          }
        }))
      },
      {
        name: "Acumulado",
        type: "line",
        yAxisIndex: 1,
        smooth: !!visual.smoothLines,
        showSymbol: showCumulativeLine && trimString(visual.seriesMarkerStyle, "circle").toLowerCase() !== "none",
        symbol: trimString(visual.seriesMarkerStyle, "circle").toLowerCase() === "none"
          ? "circle"
          : trimString(visual.seriesMarkerStyle, "circle").toLowerCase(),
        symbolSize: clampNumber(visual.seriesMarkerSize, 2, 18, 6),
        lineStyle: {
          color: lineColor,
          width: showCumulativeLine ? clampNumber(visual.lineWidth, 1, 6, 2) : 0,
          type: trimString(visual.seriesLineStyle, "solid")
        },
        itemStyle: {
          color: lineColor,
          opacity: showCumulativeLine ? 1 : 0
        },
        label: {
          show: !!visual.showDataLabels && showCumulativeLine,
          position: "top",
          ...buildSeriesLabelStyle(payload),
          formatter(params) {
            return formatPercentValue(Number(params?.data?.paretoPercent));
          }
        },
        markLine: {
          symbol: "none",
          silent: true,
          label: {
            show: true,
            formatter: `Meta ${formatPercentValue(targetPercent)}`,
            color: trimString(visual.targetLineColor, "#d45555"),
            fontSize: clampNumber(visual.fontSizeLabels, 8, 20, 11)
          },
          lineStyle: {
            color: trimString(visual.targetLineColor, "#d45555"),
            width: 1.4,
            type: "dashed"
          },
          data: [{ yAxis: targetPercent }]
        },
        data: cumulativePercentages.map((pct, index) => ({
          value: pct,
          rawIndex: index,
          paretoPercent: pct
        }))
      }
    ];
  }

  function buildCenterGraphic(payload) {
    const chartType = normalizeChartType(payload.chartType);
    if (chartType !== "donut") {
      return undefined;
    }
    const typeConfig = payload.typeConfig || {};
    const title = trimString(typeConfig.centerTitle, "");
    if (!title) {
      return undefined;
    }
    return [{
      type: "text",
      left: "center",
      top: "46%",
      style: {
        text: title,
        textAlign: "center",
        fill: "#607488",
        font: "600 12px 'Segoe UI'"
      },
      silent: true
    }];
  }

  function buildGridLineStyle(visual) {
    return {
      color: `rgba(80, 101, 122, ${clampNumber(visual.gridOpacity, 0.1, 1, 0.35)})`,
      type: clampNumber(visual.gridDash, 0, 12, 4) > 0 ? "dashed" : "solid"
    };
  }

  function buildQuickSightOption(payload) {
    const chartType = normalizeChartType(payload.chartType);
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const visual = payload.visualSettings || {};
    const categories = new Set(["bar", "bar_horizontal", "bar_stacked", "line", "area"]).has(chartType) && rows.some((row) => !!trimString(row?.breakdownLabel, ""))
      ? Array.from(new Set(rows.map((row, index) => trimString(row?.groupLabel || row?.label, `Item ${index + 1}`))))
      : rows.map((row, index) => row?.groupLabel || row?.label || `Item ${index + 1}`);
    const axisLabel = buildAxisLabelOptions(payload);
    const axisTextStyle = buildAxisTextStyle(payload);
    const tooltip = {
      trigger: new Set(["bar", "bar_horizontal", "bar_stacked", "line", "area", "combo"]).has(chartType) ? "axis" : "item",
      confine: true,
      backgroundColor: "rgba(18, 31, 46, 0.92)",
      borderWidth: 0,
      textStyle: {
        color: "#f5f8fb",
        fontFamily: trimString(visual.fontFamilyTooltip, "Segoe UI"),
        fontSize: clampNumber(visual.fontSizeTooltip, 8, 20, 11)
      },
      formatter: buildTooltipFormatter(payload)
    };
    const option = {
      animation: visual.loadAnimation !== false,
      color: normalizeColorArray(payload.colors),
      tooltip,
      legend: resolveLegend(payload, chartType),
      graphic: buildCenterGraphic(payload),
      backgroundColor: "transparent",
      grid: buildGrid(payload, chartType),
      series: []
    };

    if (chartType === "pie" || chartType === "donut") {
      option.series = [buildPieSeries(payload, chartType)];
      return option;
    }

    if (chartType === "treemap") {
      option.series = [buildTreemapSeries(payload)];
      return option;
    }

    if (chartType === "funnel") {
      option.series = [buildFunnelSeries(payload)];
      return option;
    }

    if (chartType === "scatter" || chartType === "bubble") {
      const canUseLogScale = trimString(visual.axisScale, "linear") === "log"
        && rows.every((row) => normalizeNumber(row?.value, 0) > 0);
      option.xAxis = {
        type: "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisXLabel, ""),
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel: axisTextStyle,
        nameTextStyle: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
      option.yAxis = {
        type: canUseLogScale ? "log" : "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisYLabel, ""),
        min: normalizeAxisLimit(visual.axisMin),
        max: normalizeAxisLimit(visual.axisMax),
        nameTextStyle: axisTextStyle,
        axisLabel: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
      option.series = [buildScatterSeries(payload, chartType === "bubble")];
      return option;
    }

    if (chartType === "combo") {
      const canUseLogScale = trimString(visual.axisScale, "linear") === "log"
        && rows.every((row) => normalizeNumber(row?.value, 0) > 0);
      option.xAxis = {
        type: "category",
        show: visual.showAxisLabels !== false,
        data: categories,
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel
      };
      option.yAxis = {
        type: canUseLogScale ? "log" : "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisYLabel, ""),
        min: normalizeAxisLimit(visual.axisMin),
        max: normalizeAxisLimit(visual.axisMax),
        nameTextStyle: axisTextStyle,
        axisLabel: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
      option.series = buildComboSeries(payload);
      return option;
    }

    if (chartType === "waterfall") {
      const meta = buildWaterfallChartMeta(payload);
      option.graphic = (option.graphic || []).concat(buildWaterfallTotalGraphic(payload, meta.total) || []);
      option.xAxis = {
        type: "category",
        show: visual.showAxisLabels !== false,
        data: categories,
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel
      };
      option.yAxis = {
        type: "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisYLabel, ""),
        min: Number.isFinite(Number(visual.axisMin)) ? Number(visual.axisMin) : meta.minValue,
        max: Number.isFinite(Number(visual.axisMax)) ? Number(visual.axisMax) : meta.maxValue,
        nameTextStyle: axisTextStyle,
        axisLabel: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
      option.series = [buildWaterfallSeries(payload)];
      return option;
    }

    if (chartType === "pareto") {
      option.tooltip.trigger = "axis";
      option.xAxis = {
        type: "category",
        show: visual.showAxisLabels !== false,
        data: categories,
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel
      };
      option.yAxis = [
        {
          type: "value",
          show: visual.showAxisLabels !== false,
          name: trimString(visual.axisYLabel, ""),
          min: normalizeAxisLimit(visual.axisMin),
          max: normalizeAxisLimit(visual.axisMax),
          nameTextStyle: axisTextStyle,
          axisLabel: axisTextStyle,
          splitLine: {
            show: !!visual.showGrid,
            lineStyle: buildGridLineStyle(visual)
          }
        },
        {
          type: "value",
          min: 0,
          max: 100,
          position: "right",
          axisLabel: {
            ...axisTextStyle,
            formatter(value) {
              return formatPercentValue(value);
            }
          },
          splitLine: { show: false }
        }
      ];
      option.series = buildParetoSeries(payload);
      return option;
    }

    const canUseLogScale = trimString(visual.axisScale, "linear") === "log"
      && rows.every((row) => normalizeNumber(row?.value, 0) > 0);
    if (HORIZONTAL_BAR_TYPES.has(chartType)) {
      option.xAxis = {
        type: canUseLogScale ? "log" : "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisXLabel, ""),
        min: normalizeAxisLimit(visual.axisMin),
        max: normalizeAxisLimit(visual.axisMax),
        nameTextStyle: axisTextStyle,
        axisLabel: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
      option.yAxis = {
        type: "category",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisYLabel, ""),
        data: categories,
        nameTextStyle: axisTextStyle,
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel
      };
    } else {
      option.xAxis = {
        type: "category",
        show: visual.showAxisLabels !== false,
        data: categories,
        axisLine: {
          show: true,
          lineStyle: { color: "rgba(80, 101, 122, 0.34)" }
        },
        axisTick: { show: false },
        axisLabel
      };
      option.yAxis = {
        type: canUseLogScale ? "log" : "value",
        show: visual.showAxisLabels !== false,
        name: trimString(visual.axisYLabel, ""),
        min: normalizeAxisLimit(visual.axisMin),
        max: normalizeAxisLimit(visual.axisMax),
        nameTextStyle: axisTextStyle,
        axisLabel: axisTextStyle,
        splitLine: {
          show: !!visual.showGrid,
          lineStyle: buildGridLineStyle(visual)
        }
      };
    }
    const cartesianSeries = buildCartesianSeries(payload, chartType);
    option.series = Array.isArray(cartesianSeries) ? cartesianSeries : [cartesianSeries];
    return option;
  }

  function renderQuickSightChart(surface, payload) {
    const api = getApi();
    if (!api || !(surface instanceof HTMLElement)) {
      return { rendered: false, reason: "missing_api" };
    }
    const chartType = normalizeChartType(payload?.chartType);
    if (!chartType) {
      return { rendered: false, reason: "unsupported" };
    }
    const chart = api.getInstanceByDom(surface) || api.init(surface, null, { renderer: "canvas" });
    chart.setOption(buildQuickSightOption(payload), true);
    chart.resize();
    return { rendered: true, chart };
  }

  globalScope.MIDPEChartsAdapter = Object.freeze({
    canRenderChartType,
    disposeQuickSightChart,
    renderQuickSightChart
  });
})(typeof window !== "undefined" ? window : globalThis);
