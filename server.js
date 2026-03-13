const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const os = require("os");

loadLocalEnv();

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const BASIC_AUTH_USER = String(process.env.BASIC_AUTH_USER || "").trim();
const BASIC_AUTH_PASS = String(process.env.BASIC_AUTH_PASS || "").trim();
const BASIC_AUTH_ENABLED = BASIC_AUTH_USER.length > 0 && BASIC_AUTH_PASS.length > 0;
const APS_CLIENT_ID = String(process.env.APS_CLIENT_ID || "").trim();
const APS_CLIENT_SECRET = String(process.env.APS_CLIENT_SECRET || "").trim();
const APS_CALLBACK_URL = String(process.env.APS_CALLBACK_URL || `http://127.0.0.1:${PORT}/api/aps/auth/callback`).trim();
const APS_BASE_URL = "https://developer.api.autodesk.com";
const APS_USERINFO_URL = "https://api.userprofile.autodesk.com/userinfo";
const APS_SCOPE_LIST = String(process.env.APS_SCOPES || "data:read viewables:read account:read").trim();
const APS_ENABLED = true;
const APS_CONFIGURED = APS_CLIENT_ID.length > 0 && APS_CALLBACK_URL.length > 0;
const APS_SESSION_COOKIE = "midp_aps_session";
const APPDATA_DIR = String(process.env.APPDATA || path.join(os.homedir(), "AppData", "Roaming")).trim();
const ACC_LOCAL_EXTRACT_DIR = String(
  process.env.ACC_LOCAL_EXTRACT_DIR || path.join(os.homedir(), "Downloads", "autodesk_data_extract")
).trim();
const ACC_LOCAL_PROJECT_JSON_PATHS = [
  path.join(APPDATA_DIR, "ACC Issues Explorer", ".aps-project.json"),
  path.join(APPDATA_DIR, "acc", ".aps-project.json"),
  path.join(os.homedir(), "OneDrive", "Documentos", "Programacion", "BIM", "Jason", "ACC", ".aps-project.json")
];
const ACC_LOCAL_ISSUES_EXPORT_DIRS = [
  path.join(APPDATA_DIR, "ACC Issues Explorer", "out"),
  path.join(os.homedir(), "OneDrive", "Documentos", "Programacion", "BIM", "Jason", "ACC", "out")
];
const apsSessions = new Map();

function loadLocalEnv() {
  const envFiles = [".env.local", ".env"];
  envFiles.forEach((filename) => {
    const filePath = path.join(__dirname, filename);
    if (!fs.existsSync(filePath)) {
      return;
    }
    const contents = fs.readFileSync(filePath, "utf8");
    contents.split(/\r?\n/).forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) {
        return;
      }
      const separatorIndex = line.indexOf("=");
      if (separatorIndex <= 0) {
        return;
      }
      const key = line.slice(0, separatorIndex).trim();
      if (!key || Object.prototype.hasOwnProperty.call(process.env, key)) {
        return;
      }
      let value = line.slice(separatorIndex + 1).trim();
      if (
        (value.startsWith("\"") && value.endsWith("\"")) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      process.env[key] = value;
    });
  });
}

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".ico": "image/x-icon",
  ".webp": "image/webp"
};

function writeSecurityHeaders(res, contentType = "text/plain; charset=utf-8") {
  res.setHeader("Content-Type", contentType);
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader("Cross-Origin-Resource-Policy", "same-origin");
  res.setHeader(
    "Content-Security-Policy",
    "default-src 'self'; script-src 'self' 'unsafe-eval' https://developer.api.autodesk.com https://*.autodesk.com; style-src 'self' 'unsafe-inline' https://developer.api.autodesk.com https://*.autodesk.com; img-src 'self' data: blob: https://developer.api.autodesk.com https://*.autodesk.com; font-src 'self' data: https://developer.api.autodesk.com https://*.autodesk.com; connect-src 'self' https://developer.api.autodesk.com https://api.userprofile.autodesk.com https://*.autodesk.com; worker-src 'self' blob:; object-src 'none'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
  );
  res.setHeader("Cache-Control", "no-store, max-age=0");
}

function writeJson(res, statusCode, payload) {
  writeSecurityHeaders(res, "application/json; charset=utf-8");
  res.writeHead(statusCode);
  res.end(JSON.stringify(payload));
}

function safeEqual(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  if (left.length !== right.length) {
    return false;
  }
  return crypto.timingSafeEqual(left, right);
}

function requestIsAuthorized(req) {
  if (!BASIC_AUTH_ENABLED) {
    return true;
  }
  const raw = String(req.headers.authorization || "");
  if (!raw.startsWith("Basic ")) {
    return false;
  }
  let decoded = "";
  try {
    decoded = Buffer.from(raw.slice(6), "base64").toString("utf8");
  } catch {
    return false;
  }
  const separatorIndex = decoded.indexOf(":");
  if (separatorIndex <= 0) {
    return false;
  }
  const user = decoded.slice(0, separatorIndex);
  const pass = decoded.slice(separatorIndex + 1);
  return safeEqual(user, BASIC_AUTH_USER) && safeEqual(pass, BASIC_AUTH_PASS);
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      writeSecurityHeaders(res, "text/plain; charset=utf-8");
      res.writeHead(error.code === "ENOENT" ? 404 : 500);
      res.end(error.code === "ENOENT" ? "404 Not Found" : "500 Internal Server Error");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const mimeType = MIME_TYPES[ext] || "application/octet-stream";
    writeSecurityHeaders(res, mimeType);
    res.writeHead(200);
    res.end(content);
  });
}

function parseCsvRows(raw) {
  const rows = [];
  let currentRow = [];
  let currentValue = "";
  let inQuotes = false;
  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    const nextChar = raw[index + 1];
    if (char === "\"") {
      if (inQuotes && nextChar === "\"") {
        currentValue += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (char === "," && !inQuotes) {
      currentRow.push(currentValue);
      currentValue = "";
      continue;
    }
    if ((char === "\n" || char === "\r") && !inQuotes) {
      currentRow.push(currentValue);
      const hasContent = currentRow.some((value) => String(value || "").length > 0);
      if (hasContent) {
        rows.push(currentRow);
      }
      currentRow = [];
      currentValue = "";
      if (char === "\r" && nextChar === "\n") {
        index += 1;
      }
      continue;
    }
    currentValue += char;
  }
  if (currentValue.length > 0 || currentRow.length > 0) {
    currentRow.push(currentValue);
    const hasContent = currentRow.some((value) => String(value || "").length > 0);
    if (hasContent) {
      rows.push(currentRow);
    }
  }
  return rows;
}

function readCsvRecords(filePath) {
  if (!fs.existsSync(filePath)) {
    return [];
  }
  const raw = fs.readFileSync(filePath, "utf8");
  const rows = parseCsvRows(raw);
  if (rows.length === 0) {
    return [];
  }
  const headers = rows[0].map((header) => String(header || "").trim());
  return rows.slice(1).map((cells) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = String(cells[index] || "").trim();
    });
    return record;
  });
}

function readJsonFile(filePath) {
  if (!fs.existsSync(filePath)) {
    return null;
  }
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return null;
  }
}

function normalizeLocalAccProjectId(rawValue) {
  return String(rawValue || "").trim().replace(/^b\./i, "");
}

function normalizeLocalAccAccountId(rawValue) {
  return String(rawValue || "").trim().replace(/^b\./i, "");
}

function looksLikeAccUuid(rawValue) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(rawValue || "").trim());
}

function looksLikeMeaningfulLocalModelName(value) {
  const name = String(value || "").trim();
  if (!name) {
    return false;
  }
  return !/^(\{?3d\}?|2d view|\[\d+\]\s*model|\(\d+\)|model|3d[_\s-]*exportacion[_\s-]*navis|vista\s*3d)$/i.test(name);
}

function inferLocalAccModelType(fileName) {
  const safeName = String(fileName || "").trim().toLowerCase();
  if (safeName.endsWith(".ifc")) {
    return "acc-ifc";
  }
  return "acc-rvt";
}

function chooseLocalAccModelName(nameSet, fallbackIndex) {
  const names = Array.from(nameSet || []).map((item) => String(item || "").trim()).filter(Boolean);
  const preferredFileName = names.find((name) => /\.[a-z0-9]{2,6}$/i.test(name) && looksLikeMeaningfulLocalModelName(name));
  if (preferredFileName) {
    return preferredFileName;
  }
  const preferredLabel = names.find((name) => looksLikeMeaningfulLocalModelName(name));
  if (preferredLabel) {
    return preferredLabel;
  }
  return `Modelo detectado ${fallbackIndex}`;
}

function buildLocalAccIssueReferenceRow(rawIssue, projectName, modelLike, fallbackIndex) {
  const issue = rawIssue && typeof rawIssue === "object" ? rawIssue : {};
  const model = modelLike && typeof modelLike === "object" ? modelLike : {};
  const viewable = model.viewable && typeof model.viewable === "object" ? model.viewable : {};
  return {
    issueId: trimOrFallback(issue.id, `issue-${fallbackIndex + 1}`),
    issueDisplayId: trimOrFallback(issue.displayId, ""),
    issueTitle: trimOrFallback(issue.title, ""),
    issueStatus: trimOrFallback(issue.status, ""),
    issueTypeId: trimOrFallback(issue.issueTypeId, ""),
    issueSubtypeId: trimOrFallback(issue.issueSubtypeId, ""),
    issueCreatedAt: trimOrFallback(issue.createdAt, ""),
    issueUpdatedAt: trimOrFallback(issue.updatedAt, ""),
    issueDueDate: trimOrFallback(issue.dueDate, ""),
    issueAssignedTo: trimOrFallback(issue.assignedTo, ""),
    issueAssignedToType: trimOrFallback(issue.assignedToType, ""),
    issueLocation: trimOrFallback(issue.locationDetails, ""),
    issueDescription: trimOrFallback(issue.description, ""),
    projectName: trimOrFallback(projectName, ""),
    modelLineageUrn: trimOrFallback(model.lineageUrn || model.urn, ""),
    modelViewerUrn: trimOrFallback(model.viewerUrn, ""),
    modelGuid: trimOrFallback(model.modelGuid || viewable.guid, ""),
    modelName: trimOrFallback(model.displayName || model.name || viewable.name, ""),
    modelViewableName: trimOrFallback(viewable.name, ""),
    modelViewableId: trimOrFallback(viewable.viewableId, ""),
    modelIs3D: model.is3D ? "Si" : "No",
    sourceType: trimOrFallback(model.sourceType, "")
  };
}

function buildLocalAccIssueModelEntries(projectId, projectName, modelBuckets) {
  const entries = Array.from(modelBuckets.values()).map((bucket, index) => {
    let modelName = chooseLocalAccModelName(bucket.names, index + 1);
    if (/^Modelo detectado \d+$/i.test(modelName) && bucket.is3D) {
      modelName = `Modelo ACC 3D ${index + 1}`;
    }
    const sourceUrn = String(bucket.versionUrn || "").trim();
    return {
      id: `local-model:${projectId}:${sourceUrn || index + 1}`,
      kind: "model",
      itemId: sourceUrn || `local-item-${index + 1}`,
      versionUrn: sourceUrn,
      viewerUrn: trimOrFallback(bucket.viewerUrn, ""),
      modelGuid: trimOrFallback(bucket.modelGuid, ""),
      accScopes: "",
      type: inferLocalAccModelType(modelName),
      name: modelName,
      metaLabel: `${bucket.referenceCount} referencia(s) local(es)${bucket.viewerUrn ? " | Viewer listo" : ""}`,
      projectId: projectId ? `b.${projectId}` : "",
      projectName: String(projectName || "").trim(),
      hubId: "",
      hubName: "",
      propertyRows: Array.isArray(bucket.rows) ? bucket.rows.slice() : []
    };
  }).filter((entry) => {
    const hasFileName = /\.[a-z0-9]{2,6}$/i.test(entry.name);
    const hasViewer = !!trimOrFallback(entry.viewerUrn, "");
    const hasRows = Array.isArray(entry.propertyRows) && entry.propertyRows.length > 0;
    const looksGeneric = /^Modelo detectado \d+$/i.test(entry.name);
    return hasFileName || hasViewer || (hasRows && !looksGeneric);
  });
  return entries.sort((left, right) => left.name.localeCompare(right.name, "es"));
}

function collectLocalAccSources() {
  const extractAccountsPath = path.join(ACC_LOCAL_EXTRACT_DIR, "admin_accounts.csv");
  const extractProjectsPath = path.join(ACC_LOCAL_EXTRACT_DIR, "admin_projects.csv");
  const availableSources = [];
  const accounts = [];
  const accountsById = new Map();
  const projectMap = new Map();
  const projectEntries = new Map();

  const upsertProject = (projectLike) => {
    const normalizedProjectId = normalizeLocalAccProjectId(projectLike.id || projectLike.projectId);
    const normalizedAccountId = normalizeLocalAccAccountId(projectLike.accountId);
    const safeName = String(projectLike.name || "").trim();
    const key = normalizedProjectId || safeName.toLowerCase();
    if (!key || !safeName) {
      return null;
    }
    const existing = projectMap.get(key) || {
      id: normalizedProjectId,
      dataProjectId: normalizedProjectId ? `b.${normalizedProjectId}` : "",
      accountId: normalizedAccountId,
      hubId: normalizedAccountId ? `b.${normalizedAccountId}` : "",
      accountName: "",
      hubName: "",
      name: safeName,
      classification: "",
      status: "",
      source: "",
      sourceEndpoint: "",
      sources: []
    };
    existing.id = existing.id || normalizedProjectId;
    existing.dataProjectId = existing.dataProjectId || (normalizedProjectId ? `b.${normalizedProjectId}` : "");
    existing.accountId = existing.accountId || normalizedAccountId;
    existing.hubId = existing.hubId || (existing.accountId ? `b.${existing.accountId}` : "");
    existing.accountName = existing.accountName || String(projectLike.accountName || "").trim();
    existing.hubName = existing.hubName || String(projectLike.hubName || projectLike.accountName || "").trim();
    existing.name = existing.name || safeName;
    existing.classification = existing.classification || String(projectLike.classification || "").trim();
    existing.status = existing.status || String(projectLike.status || "").trim();
    existing.source = existing.source || String(projectLike.source || "").trim();
    existing.sourceEndpoint = existing.sourceEndpoint || String(projectLike.sourceEndpoint || "").trim();
    const sourceLabel = String(projectLike.sourceLabel || projectLike.source || "").trim();
    if (sourceLabel && !existing.sources.includes(sourceLabel)) {
      existing.sources.push(sourceLabel);
    }
    projectMap.set(key, existing);
    return existing;
  };

  if (fs.existsSync(extractAccountsPath) && fs.existsSync(extractProjectsPath)) {
    availableSources.push(ACC_LOCAL_EXTRACT_DIR);
    readCsvRecords(extractAccountsPath).forEach((record) => {
      const accountId = normalizeLocalAccAccountId(record.bim360_account_id);
      if (!accountId) {
        return;
      }
      const account = {
        id: accountId,
        name: String(record.display_name || "").trim()
      };
      accounts.push(account);
      accountsById.set(accountId, account);
    });
    readCsvRecords(extractProjectsPath).forEach((record) => {
      const accountId = normalizeLocalAccAccountId(record.bim360_account_id);
      const account = accountsById.get(accountId) || null;
      upsertProject({
        id: record.id,
        accountId,
        accountName: account ? account.name : "",
        hubName: account ? account.name : "",
        name: record.name,
        classification: record.classification,
        status: record.status,
        source: "local-acc-extract",
        sourceEndpoint: "admin_projects.csv",
        sourceLabel: "Extract local ACC"
      });
    });
  }

  ACC_LOCAL_PROJECT_JSON_PATHS.forEach((filePath) => {
    const payload = readJsonFile(filePath);
    if (!payload || typeof payload !== "object") {
      return;
    }
    if (!looksLikeAccUuid(normalizeLocalAccProjectId(payload.projectId))) {
      return;
    }
    availableSources.push(filePath);
    upsertProject({
      id: payload.projectId,
      name: payload.projectName,
      source: "local-project-cache",
      sourceEndpoint: filePath,
      sourceLabel: path.basename(path.dirname(filePath)) || path.basename(filePath)
    });
  });

  ACC_LOCAL_ISSUES_EXPORT_DIRS.forEach((dirPath) => {
    if (!fs.existsSync(dirPath)) {
      return;
    }
    availableSources.push(dirPath);
    const csvFiles = fs.readdirSync(dirPath)
      .filter((name) => /^issues-.*\.csv$/i.test(name))
      .map((name) => path.join(dirPath, name))
      .sort((left, right) => {
        try {
          return fs.statSync(right).mtimeMs - fs.statSync(left).mtimeMs;
        } catch {
          return 0;
        }
      });
    csvFiles.forEach((filePath) => {
      readCsvRecords(filePath).forEach((record) => {
        const normalizedProjectId = normalizeLocalAccProjectId(record.projectId);
        if (!looksLikeAccUuid(normalizedProjectId)) {
          return;
        }
        const project = upsertProject({
          id: normalizedProjectId,
          name: record.projectName,
          source: "local-issues-export",
          sourceEndpoint: filePath,
          sourceLabel: path.basename(path.dirname(filePath)) || "issues export"
        });
        if (!project) {
          return;
        }
        const projectKey = project.id || normalizedProjectId || project.name.toLowerCase();
        if (!projectKey) {
          return;
        }
        let rawJson = null;
        try {
          rawJson = record.raw_json ? JSON.parse(record.raw_json) : null;
        } catch {
          rawJson = null;
        }
        if (!rawJson || typeof rawJson !== "object") {
          return;
        }
        const modelBuckets = projectEntries.get(projectKey) || new Map();
        const registerModel = (modelLike, referenceRow) => {
          const safeModel = modelLike && typeof modelLike === "object" ? modelLike : {};
          const safeUrn = String(safeModel.lineageUrn || safeModel.urn || safeModel.versionUrn || "").trim();
          const safeName = String(safeModel.displayName || safeModel.name || "").trim();
          const safeViewerUrn = String(safeModel.viewerUrn || "").trim();
          const safeGuid = String(safeModel.modelGuid || safeModel.guid || "").trim();
          const is3D = !!safeModel.is3D;
          if (!safeUrn && !safeName) {
            return;
          }
          const modelKey = safeUrn || safeName.toLowerCase();
          const bucket = modelBuckets.get(modelKey) || {
            versionUrn: safeUrn,
            viewerUrn: safeViewerUrn,
            modelGuid: safeGuid,
            is3D,
            names: new Set(),
            referenceCount: 0,
            rows: [],
            rowKeys: new Set()
          };
          if (safeUrn && !bucket.versionUrn) {
            bucket.versionUrn = safeUrn;
          }
          if (safeViewerUrn && !bucket.viewerUrn) {
            bucket.viewerUrn = safeViewerUrn;
          }
          if (safeGuid && !bucket.modelGuid) {
            bucket.modelGuid = safeGuid;
          }
          bucket.is3D = bucket.is3D || is3D;
          if (safeName) {
            bucket.names.add(safeName);
          }
          bucket.referenceCount += 1;
          if (referenceRow && typeof referenceRow === "object") {
            const rowKey = `${trimOrFallback(referenceRow.issueId, "")}|${trimOrFallback(referenceRow.modelLineageUrn, "")}|${trimOrFallback(referenceRow.modelGuid, "")}`;
            if (rowKey && !bucket.rowKeys.has(rowKey)) {
              bucket.rowKeys.add(rowKey);
              bucket.rows.push(referenceRow);
            }
          }
          modelBuckets.set(modelKey, bucket);
        };
        const linkedDocuments = Array.isArray(rawJson.linkedDocuments) ? rawJson.linkedDocuments : [];
        linkedDocuments.forEach((item, itemIndex) => {
          const viewable = item?.details?.viewable && typeof item.details.viewable === "object"
            ? item.details.viewable
            : {};
          const viewerUrn = trimOrFallback(
            item?.details?.viewerState?.seedURN
            || item?.details?.viewerState?.seedUrn,
            ""
          );
          const modelLike = {
            sourceType: trimOrFallback(item?.type, "linkedDocument"),
            urn: item?.urn,
            lineageUrn: item?.urn,
            viewerUrn,
            modelGuid: viewable?.guid,
            displayName: viewable?.name || item?.name || item?.details?.name || "",
            viewable,
            is3D: !!viewable?.is3D
          };
          registerModel(
            modelLike,
            buildLocalAccIssueReferenceRow(rawJson, project.name, modelLike, itemIndex)
          );
        });
        const placements = Array.isArray(rawJson.placements) ? rawJson.placements : [];
        placements.forEach((item, itemIndex) => {
          const viewable = item?.viewable && typeof item.viewable === "object"
            ? item.viewable
            : {};
          const modelLike = {
            sourceType: trimOrFallback(item?.type, "placement"),
            lineageUrn: item?.lineageUrn,
            urn: item?.lineageUrn,
            displayName: viewable?.name || item?.name || "",
            modelGuid: viewable?.guid,
            viewable,
            is3D: !!viewable?.is3D
          };
          registerModel(
            modelLike,
            buildLocalAccIssueReferenceRow(rawJson, project.name, modelLike, itemIndex)
          );
        });
        if (modelBuckets.size > 0) {
          projectEntries.set(projectKey, modelBuckets);
        }
      });
    });
  });

  const projects = Array.from(projectMap.values())
    .map((project) => ({
      id: project.id,
      dataProjectId: project.dataProjectId,
      accountId: project.accountId,
      hubId: project.hubId,
      accountName: project.accountName,
      hubName: project.hubName || project.accountName,
      name: project.name,
      classification: project.classification,
      status: project.status,
      source: project.source,
      sourceEndpoint: project.sourceEndpoint,
      sources: project.sources,
      entriesCount: 0
    }))
    .sort((left, right) => left.name.localeCompare(right.name, "es"));

  const projectEntriesPayload = Array.from(projectEntries.entries()).reduce((acc, [projectKey, modelBuckets]) => {
    const project = projects.find((item) => item.id === projectKey || item.name.toLowerCase() === projectKey);
    if (!project) {
      return acc;
    }
    acc[projectKey] = buildLocalAccIssueModelEntries(project.id, project.name, modelBuckets);
    project.entriesCount = acc[projectKey].length;
    return acc;
  }, {});

  return {
    available: projects.length > 0 || accounts.length > 0,
    dir: ACC_LOCAL_EXTRACT_DIR,
    accounts,
    projects,
    sources: Array.from(new Set(availableSources.filter(Boolean))),
    projectEntries: projectEntriesPayload
  };
}

function getLocalAccExtractSummary() {
  const localSources = collectLocalAccSources();
  return {
    available: localSources.available,
    dir: localSources.dir,
    accounts: localSources.accounts,
    projects: localSources.projects,
    sources: localSources.sources
  };
}

function getLocalAccProjectEntries(projectId) {
  const normalizedProjectId = normalizeLocalAccProjectId(projectId);
  const localSources = collectLocalAccSources();
  const projectKey = normalizedProjectId || String(projectId || "").trim().toLowerCase();
  const entries = localSources.projectEntries[normalizedProjectId] || localSources.projectEntries[projectKey] || [];
  return {
    available: entries.length > 0,
    projectId: normalizedProjectId ? `b.${normalizedProjectId}` : String(projectId || "").trim(),
    entries
  };
}

function parseCookies(req) {
  const raw = String(req.headers.cookie || "");
  return raw.split(";").reduce((acc, segment) => {
    const [name, ...valueParts] = segment.trim().split("=");
    if (!name) {
      return acc;
    }
    acc[name] = decodeURIComponent(valueParts.join("=") || "");
    return acc;
  }, {});
}

function getSession(req) {
  const cookies = parseCookies(req);
  const sessionId = String(cookies[APS_SESSION_COOKIE] || "").trim();
  if (!sessionId) {
    return { sessionId: "", session: null };
  }
  return {
    sessionId,
    session: apsSessions.get(sessionId) || null
  };
}

function createSession(res) {
  const sessionId = crypto.randomUUID();
  const session = {
    id: sessionId,
    createdAt: new Date().toISOString(),
    aps: null
  };
  apsSessions.set(sessionId, session);
  res.setHeader("Set-Cookie", `${APS_SESSION_COOKIE}=${encodeURIComponent(sessionId)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=28800`);
  return { sessionId, session };
}

function ensureSession(req, res) {
  const existing = getSession(req);
  if (existing.session) {
    return existing;
  }
  return createSession(res);
}

function clearSession(res, sessionId) {
  if (sessionId) {
    apsSessions.delete(sessionId);
  }
  res.setHeader("Set-Cookie", `${APS_SESSION_COOKIE}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`);
}

function base64UrlEncode(value) {
  return Buffer.from(value)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function generateApsCodeVerifier() {
  return base64UrlEncode(crypto.randomBytes(64));
}

function createApsCodeChallenge(codeVerifier) {
  return crypto.createHash("sha256").update(String(codeVerifier || "")).digest("base64url");
}

function decodeJwtPayload(token) {
  const raw = String(token || "").trim();
  const parts = raw.split(".");
  if (parts.length < 2) {
    return null;
  }
  try {
    const payload = parts[1]
      .replace(/-/g, "+")
      .replace(/_/g, "/");
    const normalized = `${payload}${"=".repeat((4 - (payload.length % 4 || 4)) % 4)}`;
    return JSON.parse(Buffer.from(normalized, "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function getApsStatusPayload(session) {
  const apsState = session?.aps || null;
  const tokenClaims = decodeJwtPayload(apsState?.accessToken);
  return {
    enabled: APS_ENABLED,
    configured: APS_CONFIGURED,
    authenticated: !!apsState?.accessToken,
    clientId: APS_CLIENT_ID,
    callbackUrl: APS_CALLBACK_URL,
    userName: String(apsState?.userName || "").trim(),
    userEmail: String(apsState?.userEmail || "").trim(),
    userId: String(apsState?.userId || "").trim(),
    expiresAt: String(apsState?.expiresAt || "").trim(),
    scopes: Array.isArray(apsState?.scopes) ? apsState.scopes.slice() : APS_SCOPE_LIST.split(/\s+/).filter(Boolean),
    tokenClaims: tokenClaims
      ? {
          client_id: String(tokenClaims.client_id || "").trim(),
          userid: String(tokenClaims.userid || tokenClaims.userId || "").trim(),
          scope: Array.isArray(tokenClaims.scope) ? tokenClaims.scope.slice() : tokenClaims.scope || "",
          aud: tokenClaims.aud || null,
          iss: tokenClaims.iss || null,
          exp: tokenClaims.exp || null
        }
      : null
  };
}

async function readJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8").trim();
  if (!raw) {
    return {};
  }
  return JSON.parse(raw);
}

async function requestApsToken(params) {
  const body = new URLSearchParams(params);
  const credentials = Buffer.from(`${APS_CLIENT_ID}:${APS_CLIENT_SECRET}`).toString("base64");
  const response = await fetch(`${APS_BASE_URL}/authentication/v2/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Authorization: `Basic ${credentials}`
    },
    body
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`APS token ${response.status}: ${errorText}`);
  }
  return response.json();
}

async function requestApsTwoLeggedToken(scopes) {
  return requestApsToken({
    grant_type: "client_credentials",
    scope: scopes
  });
}

async function requestApsThreeLeggedToken(code, codeVerifier) {
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    client_id: APS_CLIENT_ID,
    code,
    redirect_uri: APS_CALLBACK_URL,
    code_verifier: String(codeVerifier || "").trim()
  });
  if (APS_CLIENT_SECRET) {
    body.set("client_secret", APS_CLIENT_SECRET);
  }
  const response = await fetch(`${APS_BASE_URL}/authentication/v2/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`APS token ${response.status}: ${errorText}`);
  }
  return response.json();
}

async function fetchApsUserProfile(accessToken) {
  if (!accessToken) {
    return {
      userName: "",
      userEmail: "",
      userId: ""
    };
  }
  try {
    const response = await fetch(APS_USERINFO_URL, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json"
      }
    });
    if (!response.ok) {
      return {
        userName: "",
        userEmail: "",
        userId: ""
      };
    }
    const payload = await response.json();
    return {
      userName: String(payload.name || payload.preferred_username || payload.email || "").trim(),
      userEmail: String(payload.email || payload.preferred_username || "").trim(),
      userId: String(payload.sub || payload.userId || payload.uid || "").trim()
    };
  } catch {
    return {
      userName: "",
      userEmail: "",
      userId: ""
    };
  }
}

function buildApsAuthorizeUrl(session) {
  const scopes = APS_SCOPE_LIST.split(/\s+/).filter(Boolean).join(" ");
  const state = crypto.randomUUID();
  const codeVerifier = generateApsCodeVerifier();
  const codeChallenge = createApsCodeChallenge(codeVerifier);
  session.apsAuth = {
    state,
    codeVerifier,
    createdAt: new Date().toISOString()
  };
  const params = new URLSearchParams({
    response_type: "code",
    client_id: APS_CLIENT_ID,
    redirect_uri: APS_CALLBACK_URL,
    scope: scopes,
    state,
    code_challenge: codeChallenge,
    code_challenge_method: "S256"
  });
  return `${APS_BASE_URL}/authentication/v2/authorize?${params.toString()}`;
}

function sanitizeApsResourcePath(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  if (/^https?:\/\//i.test(raw)) {
    try {
      const parsed = new URL(raw);
      if (parsed.origin !== APS_BASE_URL) {
        return "";
      }
      return `${parsed.pathname}${parsed.search}`;
    } catch {
      return "";
    }
  }
  return raw.startsWith("/") ? raw : `/${raw}`;
}

async function handleApsApi(req, res, pathname, searchParams) {
  const { sessionId, session } = ensureSession(req, res);

  if (pathname === "/api/acc/local-extract/summary" && req.method === "GET") {
    writeJson(res, 200, getLocalAccExtractSummary());
    return true;
  }

  const localProjectEntriesMatch = pathname.match(/^\/api\/acc\/local-extract\/projects\/([^/]+)\/entries$/);
  if (localProjectEntriesMatch && req.method === "GET") {
    writeJson(res, 200, getLocalAccProjectEntries(decodeURIComponent(localProjectEntriesMatch[1] || "")));
    return true;
  }

  if (pathname === "/api/aps/config" && req.method === "GET") {
    writeJson(res, 200, {
      enabled: APS_ENABLED,
      configured: APS_CONFIGURED,
      clientId: APS_CLIENT_ID,
      callbackUrl: APS_CALLBACK_URL,
      scopes: APS_SCOPE_LIST.split(/\s+/).filter(Boolean)
    });
    return true;
  }

  if (pathname === "/api/aps/auth/status" && req.method === "GET") {
    writeJson(res, 200, getApsStatusPayload(session));
    return true;
  }

  if (pathname === "/api/aps/auth/login" && req.method === "GET") {
    if (!APS_CONFIGURED) {
      writeJson(res, 503, {
        enabled: APS_ENABLED,
        configured: false,
        message: "Configura APS_CLIENT_ID y APS_CALLBACK_URL."
      });
      return true;
    }
    session.apsAuth = null;
    writeSecurityHeaders(res, "text/plain; charset=utf-8");
    res.writeHead(302, { Location: buildApsAuthorizeUrl(session) });
    res.end();
    return true;
  }

  if (pathname === "/api/aps/auth/callback" && req.method === "GET") {
    if (!APS_CONFIGURED) {
      writeSecurityHeaders(res, "text/html; charset=utf-8");
      res.writeHead(503);
      res.end("<h1>APS no configurado</h1>");
      return true;
    }
    const code = String(searchParams.get("code") || "").trim();
    const state = String(searchParams.get("state") || "").trim();
    if (!code) {
      writeSecurityHeaders(res, "text/html; charset=utf-8");
      res.writeHead(400);
      res.end("<h1>APS callback invalido</h1>");
      return true;
    }
    const expectedState = String(session?.apsAuth?.state || "").trim();
    const codeVerifier = String(session?.apsAuth?.codeVerifier || "").trim();
    if (!expectedState || !codeVerifier || !state || state !== expectedState) {
      writeSecurityHeaders(res, "text/html; charset=utf-8");
      res.writeHead(400);
      res.end("<h1>APS callback invalido</h1><p>OAuth state o PKCE invalido.</p>");
      return true;
    }
    try {
      const tokenData = await requestApsThreeLeggedToken(code, codeVerifier);
      const userProfile = await fetchApsUserProfile(tokenData.access_token);
      session.aps = {
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || "",
        scopes: APS_SCOPE_LIST.split(/\s+/).filter(Boolean),
        expiresAt: new Date(Date.now() + (Number(tokenData.expires_in || 0) * 1000)).toISOString(),
        userName: userProfile.userName,
        userEmail: userProfile.userEmail,
        userId: userProfile.userId
      };
      session.apsAuth = null;
      writeSecurityHeaders(res, "text/html; charset=utf-8");
      res.writeHead(200);
      res.end("<h1>APS conectado</h1><p>Ya puedes volver a MIDP y sincronizar modelos ACC / RVT.</p>");
    } catch (error) {
      session.apsAuth = null;
      writeSecurityHeaders(res, "text/html; charset=utf-8");
      res.writeHead(500);
      res.end(`<h1>Error APS</h1><pre>${String(error.message || error)}</pre>`);
    }
    return true;
  }

  if (pathname === "/api/aps/auth/logout" && req.method === "GET") {
    if (session) {
      session.aps = null;
      session.apsAuth = null;
    }
    clearSession(res, sessionId);
    writeJson(res, 200, { ok: true });
    return true;
  }

  if (pathname === "/api/aps/auth/token" && req.method === "GET") {
    if (!APS_CONFIGURED) {
      writeJson(res, 503, {
        enabled: APS_ENABLED,
        configured: false,
        message: "APS no configurado."
      });
      return true;
    }
    try {
      if (session?.aps?.accessToken) {
        const expiresAt = Date.parse(String(session.aps.expiresAt || ""));
        const expiresIn = Number.isFinite(expiresAt)
          ? Math.max(60, Math.floor((expiresAt - Date.now()) / 1000))
          : 3600;
        writeJson(res, 200, {
          access_token: session.aps.accessToken,
          token_type: "Bearer",
          expires_in: expiresIn,
          expires_at: session.aps.expiresAt
        });
        return true;
      }
      const tokenData = await requestApsTwoLeggedToken("viewables:read data:read");
      writeJson(res, 200, tokenData);
    } catch (error) {
      writeJson(res, 500, { message: String(error.message || error) });
    }
    return true;
  }

  if (pathname === "/api/aps/model-properties/proxy" && req.method === "POST") {
    if (!APS_CONFIGURED) {
      writeJson(res, 503, { message: "APS no configurado." });
      return true;
    }
    const accessToken = session?.aps?.accessToken || "";
    if (!accessToken) {
      writeJson(res, 401, { message: "Debes autenticarte con APS primero." });
      return true;
    }
    try {
      const payload = await readJsonBody(req);
      const resourcePath = sanitizeApsResourcePath(payload.resourcePath || payload.path || "");
      if (!resourcePath) {
        writeJson(res, 400, { message: "resourcePath invalido." });
        return true;
      }
      const method = String(payload.method || "POST").trim().toUpperCase();
      const forwardBody = payload.body && typeof payload.body === "object" ? payload.body : null;
      const allowedHeaders = {};
      const requestedHeaders = payload.headers && typeof payload.headers === "object" ? payload.headers : {};
      const headerAllowList = {
        "x-ads-acm-scopes": "x-ads-acm-scopes",
        "x-ads-acm-namespace": "x-ads-acm-namespace",
        "x-ads-acm-check-groups": "x-ads-acm-check-groups",
        "region": "Region",
        "user-id": "User-Id"
      };
      Object.entries(requestedHeaders).forEach(([headerName, headerValue]) => {
        const normalizedName = String(headerName || "").trim().toLowerCase();
        const allowedName = headerAllowList[normalizedName];
        if (allowedName && typeof headerValue === "string" && headerValue.trim()) {
          allowedHeaders[allowedName] = headerValue.trim();
        }
      });
      const response = await fetch(`${APS_BASE_URL}${resourcePath}`, {
        method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
          ...(forwardBody ? { "Content-Type": "application/json" } : {}),
          ...allowedHeaders
        },
        body: forwardBody ? JSON.stringify(forwardBody) : undefined
      });
      const contentType = String(response.headers.get("content-type") || "");
      const responseText = await response.text();
      if (!response.ok) {
        writeJson(res, response.status, {
          ok: false,
          resourcePath,
          message: responseText
        });
        return true;
      }
      if (contentType.includes("application/json")) {
        writeSecurityHeaders(res, "application/json; charset=utf-8");
        res.writeHead(200);
        res.end(responseText);
        return true;
      }
      writeJson(res, 200, {
        ok: true,
        resourcePath,
        responseText
      });
    } catch (error) {
      writeJson(res, 500, { message: String(error.message || error) });
    }
    return true;
  }

  return false;
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || `127.0.0.1:${PORT}`}`);
    const pathname = decodeURIComponent(url.pathname);

    if (!requestIsAuthorized(req)) {
      writeSecurityHeaders(res, "text/plain; charset=utf-8");
      res.setHeader("WWW-Authenticate", "Basic realm=\"MIDP\", charset=\"UTF-8\"");
      res.writeHead(401);
      res.end("401 Unauthorized");
      return;
    }

    if (pathname.startsWith("/api/aps/") || pathname.startsWith("/api/acc/")) {
      const handled = await handleApsApi(req, res, pathname, url.searchParams);
      if (handled) {
        return;
      }
    }

    if (!new Set(["GET", "HEAD"]).has(req.method || "")) {
      writeSecurityHeaders(res, "text/plain; charset=utf-8");
      res.writeHead(405);
      res.end("405 Method Not Allowed");
      return;
    }

    const requested = pathname === "/" ? "/index.html" : pathname;
    const safePath = path.normalize(requested).replace(/^(\.\.[/\\])+/, "");
    const absolutePath = path.join(ROOT, safePath);

    if (!absolutePath.startsWith(ROOT)) {
      writeSecurityHeaders(res, "text/plain; charset=utf-8");
      res.writeHead(403);
      res.end("403 Forbidden");
      return;
    }

    sendFile(res, absolutePath);
  } catch (error) {
    writeSecurityHeaders(res, "application/json; charset=utf-8");
    res.writeHead(500);
    res.end(JSON.stringify({ message: String(error.message || error) }));
  }
});

server.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`MIDP web running on port ${PORT}`);
});
