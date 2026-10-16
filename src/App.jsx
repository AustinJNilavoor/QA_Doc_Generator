import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import ExcelJS from "exceljs";
import "./App.css";

/* ==========================================================================
   CONSTANTS
   ========================================================================== */

// Column layout of the "Test Cases" worksheet.
// NOTE: verify against the real template and edit here only (single source of truth).
const COLUMNS = [
  { key: "step", header: "STEP NO", width: 14 },
  { key: "id", header: "TEST CASE ID", width: 22 },
  { key: "input", header: "INPUT", width: 34 },
  { key: "expectedOutput", header: "EXPECTED OUTPUT", width: 44 },
  { key: "actualOutput", header: "ACTUAL OUTPUT", width: 60 },
  { key: "status", header: "STATUS", width: 16 },
  { key: "apiUsed", header: "API USED", width: 44 },
  { key: "remarks", header: "REMARKS", width: 30 },
  { key: "testType", header: "TEST TYPE", width: 16 },
];

const FLOW_OPTIONS = ["API to API", "API to File", "File to API", "File to File", "Other"];
const STATUS_OPTIONS = ["Passed", "Failed", "Partially passed", "Blocked", "Not Tested"];
const TEST_TYPES = ["Positive", "Negative", "Validation"];

// ARGB colours (estimates from the screenshot - replace with values from the real template)
const COLORS = {
  label: "FFB6D7A8",
  value: "FFFFF2CC",
  title: "FF6D9EEB",
  white: "FFFFFFFF",
  evidenceHeading: "FFDDEBF7",
  black: "FF000000",
  status: {
    Passed: { fill: "FFD9EAD3", font: "FF274E13" },
    Failed: { fill: "FFF4CCCC", font: "FF990000" },
    "Partially passed": { fill: "FFFCE5CD", font: "FF7F6000" },
    Blocked: { fill: "FFD9D9D9", font: "FF434343" },
    "Not Tested": { fill: "FFEFEFEF", font: "FF666666" },
  },
};

const IMAGE_SIZES = {
  original: { label: "Original size", maxWidth: 1400 },
  standard: { label: "Standard (max 900 px wide)", maxWidth: 900 },
  compact: { label: "Compact (max 600 px wide)", maxWidth: 600 },
};

const EVIDENCE_COLS = 8;
const EVIDENCE_COL_WIDTH = 16; // ~117 px each -> ~940 px usable
const IMAGE_ROW_PT = 15; // 20 px

const DB_NAME = "qa-doc-generator";
const DB_STORE = "drafts";
const DRAFT_KEY = "current";

/* ==========================================================================
   SMALL UTILITIES
   ========================================================================== */

const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;

const pad = (n, len = 3) => String(n).padStart(len, "0");

const sanitizeFilename = (s) =>
  String(s || "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, "")
    .trim();

const isValidUrl = (s) => {
  try {
    const u = new URL(String(s).trim());
    return u.protocol === "http:" || u.protocol === "https:";
  } catch {
    return false;
  }
};

const SENSITIVE_RE =
  /("?(?:access_token|refresh_token|token|password|passwd|client_secret|secret|api[_-]?key|authorization)"?\s*[:=]\s*)("[^"]*"|[^\s,}]+)/gi;

const looksSensitive = (text) => {
  SENSITIVE_RE.lastIndex = 0;
  return SENSITIVE_RE.test(text || "");
};

const redact = (text) =>
  (text || "").replace(SENSITIVE_RE, '$1"***REDACTED***"');

const buildFileName = (meta) => {
  const parts = [meta.client, meta.interfaceName, meta.ticketNo]
    .map(sanitizeFilename)
    .filter(Boolean);
  return `${parts.length ? parts.join("_") : "QA_Document"}_QA.xlsx`;
};

const newTestCase = (existing, prefix) => {
  const used = new Set(existing.map((t) => t.id));
  let n = existing.length + 1;
  while (used.has(`${prefix}${pad(n)}`)) n++;
  return {
    key: uid(),
    id: `${prefix}${pad(n)}`,
    stepNo: "",
    stepName: "",
    input: "",
    expectedOutput: "",
    actualOutput: "",
    status: "Not Tested",
    apiUsed: "",
    remarks: "",
    testType: "Positive",
    evidence: [],
  };
};

const initialDoc = () => ({
  metadata: { client: "", interfaceName: "", flow: "API to API", ticketNo: "" },
  settings: { idPrefix: "TC_API_", imageSize: "standard", fileName: "" },
  testCases: [],
});

const moveItem = (arr, from, to) => {
  if (to < 0 || to >= arr.length) return arr;
  const copy = [...arr];
  const [item] = copy.splice(from, 1);
  copy.splice(to, 0, item);
  return copy;
};

const blobToDataUrl = (blob) =>
  new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });

const dataUrlToBlob = async (url) => (await fetch(url)).blob();

const getImageSize = (blob) =>
  new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.naturalWidth, height: img.naturalHeight });
      URL.revokeObjectURL(url);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image could not be read"));
    };
    img.src = url;
  });

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
};

/* ==========================================================================
   DRAFT SERVICE (IndexedDB)
   ========================================================================== */

const openDB = () =>
  new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(DB_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });

const dbPut = async (key, value) => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
};

const dbGet = async (key) => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readonly");
    const req = tx.objectStore(DB_STORE).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
};

const dbDelete = async (key) => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
};

const exportDraftJson = async (doc) => {
  const testCases = [];
  for (const tc of doc.testCases) {
    const evidence = [];
    for (const ev of tc.evidence) {
      evidence.push({
        ...ev,
        blob: undefined,
        imageData: ev.blob ? await blobToDataUrl(ev.blob) : null,
      });
    }
    testCases.push({ ...tc, evidence });
  }
  return JSON.stringify({ version: 1, ...doc, testCases }, null, 2);
};

const importDraftJson = async (text) => {
  const data = JSON.parse(text);
  if (!data || !Array.isArray(data.testCases) || !data.metadata) {
    throw new Error("This file is not a valid QA draft.");
  }
  const testCases = [];
  for (const tc of data.testCases) {
    const evidence = [];
    for (const ev of tc.evidence || []) {
      const { imageData, ...rest } = ev;
      evidence.push({ ...rest, blob: imageData ? await dataUrlToBlob(imageData) : null });
    }
    testCases.push({ ...tc, key: tc.key || uid(), evidence });
  }
  const base = initialDoc();
  return {
    metadata: { ...base.metadata, ...data.metadata },
    settings: { ...base.settings, ...(data.settings || {}) },
    testCases,
  };
};

/* ==========================================================================
   VALIDATION SERVICE
   ========================================================================== */

const validateDocument = (doc) => {
  const errors = [];
  const warnings = [];
  const { metadata, testCases } = doc;

  if (!metadata.client.trim()) errors.push("Client is required.");
  if (!metadata.interfaceName.trim()) errors.push("Interface is required.");
  if (!metadata.flow.trim()) errors.push("Flow is required.");
  if (!String(metadata.ticketNo).trim()) errors.push("Ticket No is required.");
  if (testCases.length === 0) errors.push("Add at least one test case.");

  const seen = new Map();
  testCases.forEach((tc, i) => {
    const id = tc.id.trim();
    if (!id) {
      errors.push(`Test case #${i + 1} has no ID.`);
      return;
    }
    if (seen.has(id)) errors.push(`Duplicate test case ID: ${id}. Please assign a unique ID.`);
    seen.set(id, true);

    if (looksSensitive(tc.actualOutput))
      warnings.push(`${id}: Actual output may contain credentials or tokens.`);
    tc.evidence.forEach((ev, j) => {
      if (!ev.blob) errors.push(`${id}: screenshot ${j + 1} has no image.`);
      else if (ev.blob.size > 5 * 1024 * 1024)
        warnings.push(
          `${id}: screenshot ${j + 1} is ${(ev.blob.size / 1048576).toFixed(1)} MB - consider a smaller image size.`
        );
    });
  });

  return { errors, warnings };
};

/* ==========================================================================
   EXCEL SERVICE
   ========================================================================== */

const thin = { style: "thin", color: { argb: COLORS.black } };
const allThin = { top: thin, left: thin, bottom: thin, right: thin };

const estimateLines = (text, width) => {
  if (text === undefined || text === null || text === "") return 1;
  const perLine = Math.max(1, Math.floor(width * 1.05));
  return String(text)
    .split(/\r?\n/)
    .reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / perLine)), 0);
};

const buildWorkbook = async (doc, onProgress = () => {}) => {
  const { metadata, testCases, settings } = doc;
  const wb = new ExcelJS.Workbook();
  wb.creator = "QA Documentation Generator";
  wb.created = new Date();

  const main = wb.addWorksheet("Test Cases", {
    views: [{ showGridLines: false }],
  });
  const lastCol = COLUMNS.length;
  const lastColLetter = main.getColumn(lastCol).letter;

  COLUMNS.forEach((c, i) => {
    main.getColumn(i + 1).width = c.width;
  });

  // Evidence sheet names (deterministic, Excel-safe)
  const evidenceSheetName = new Map();
  let evCounter = 0;
  testCases.forEach((tc) => {
    if (tc.evidence.length > 0) {
      evCounter++;
      evidenceSheetName.set(tc.key, `Evidence_${pad(evCounter)}`);
    }
  });

  /* ---- Metadata (rows 1-4) ---- */
  const metaRows = [
    ["Client", metadata.client],
    ["Interface", metadata.interfaceName],
    ["Flow", metadata.flow],
    ["Ticket No", metadata.ticketNo],
  ];
  metaRows.forEach(([label, value], i) => {
    const r = i + 1;
    const row = main.getRow(r);
    row.height = 21;
    main.mergeCells(r, 2, r, lastCol);
    const a = main.getCell(r, 1);
    a.value = label;
    a.font = { name: "Arial", size: 12, bold: true, color: { argb: COLORS.black } };
    a.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.label } };
    a.alignment = { horizontal: "left", vertical: "middle" };
    a.border = allThin;

    const b = main.getCell(r, 2);
    b.value = value;
    b.font = { name: "Arial", size: 12, bold: true, color: { argb: COLORS.black } };
    b.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.value } };
    b.alignment = { horizontal: "left", vertical: "middle" };
    for (let c = 2; c <= lastCol; c++) {
      const cell = main.getCell(r, c);
      cell.border = allThin;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.value } };
    }
  });

  /* ---- Title (row 5) ---- */
  const titleRow = 5;
  main.mergeCells(`A${titleRow}:${lastColLetter}${titleRow}`);
  main.getRow(titleRow).height = 28;
  const title = main.getCell(`A${titleRow}`);
  title.value = "TEST CASES";
  title.font = { name: "Arial", size: 16, bold: true, color: { argb: COLORS.black } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  for (let c = 1; c <= lastCol; c++) {
    const cell = main.getCell(titleRow, c);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.title } };
    cell.border = { ...allThin, bottom: { style: "medium", color: { argb: COLORS.black } } };
  }

  /* ---- Header (row 6) ---- */
  const headerRow = 6;
  main.getRow(headerRow).height = 22;
  COLUMNS.forEach((c, i) => {
    const cell = main.getCell(headerRow, i + 1);
    cell.value = c.header;
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.black } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.white } };
    cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    cell.border = allThin;
  });
  main.views = [{ showGridLines: false, state: "frozen", ySplit: headerRow }];

  /* ---- Data rows ---- */
  let lastStepLabel = null;
  testCases.forEach((tc, idx) => {
    const r = headerRow + 1 + idx;
    const row = main.getRow(r);

    const stepLabel =
      tc.stepNo || tc.stepName
        ? [tc.stepNo, tc.stepName].filter(Boolean).join(" — ")
        : "";
    const showStep = stepLabel && stepLabel !== lastStepLabel;
    if (stepLabel) lastStepLabel = stepLabel;

    const values = {
      step: showStep ? stepLabel : "",
      id: tc.id,
      input: tc.input,
      expectedOutput: tc.expectedOutput,
      actualOutput: tc.actualOutput,
      status: tc.status,
      apiUsed: tc.apiUsed,
      remarks: tc.remarks,
      testType: tc.testType,
    };

    let maxLines = 1;
    COLUMNS.forEach((col, i) => {
      const cell = row.getCell(i + 1);
      const val = values[col.key];
      cell.value = val;
      cell.font = { name: "Arial", size: 10, color: { argb: COLORS.black } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.white } };
      cell.alignment = { horizontal: "left", vertical: "top", wrapText: true };
      cell.border = allThin;
      maxLines = Math.max(maxLines, estimateLines(val, col.width));

      if (col.key === "id") {
        const sheetName = evidenceSheetName.get(tc.key);
        if (sheetName) {
          cell.value = { text: tc.id, hyperlink: `#'${sheetName}'!A1` };
          cell.font = {
            name: "Arial",
            size: 10,
            underline: true,
            color: { argb: "FF0563C1" },
          };
        }
      }
      if (col.key === "apiUsed" && val && isValidUrl(val) && String(val).length < 2000) {
        cell.value = { text: String(val).trim(), hyperlink: String(val).trim() };
        cell.font = { name: "Arial", size: 10, underline: true, color: { argb: "FF0563C1" } };
      }
      if (col.key === "status" && COLORS.status[val]) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: COLORS.status[val].fill },
        };
        cell.font = {
          name: "Arial",
          size: 10,
          bold: true,
          color: { argb: COLORS.status[val].font },
        };
      }
    });
    row.height = Math.min(409, Math.max(20, maxLines * 13 + 6));
    onProgress(Math.round(((idx + 1) / Math.max(1, testCases.length)) * 40));
  });

  /* ---- Evidence worksheets ---- */
  const maxImgWidth = IMAGE_SIZES[settings.imageSize]?.maxWidth ?? 900;
  const withEvidence = testCases.filter((t) => t.evidence.length > 0);

  for (let t = 0; t < withEvidence.length; t++) {
    const tc = withEvidence[t];
    const sheetName = evidenceSheetName.get(tc.key);
    const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
    for (let c = 1; c <= EVIDENCE_COLS; c++) ws.getColumn(c).width = EVIDENCE_COL_WIDTH;
    const lastLetter = ws.getColumn(EVIDENCE_COLS).letter;
    const totalWidthPx = EVIDENCE_COLS * EVIDENCE_COL_WIDTH * 7.2;

    let r = 1;
    ws.mergeCells(`A${r}:${lastLetter}${r}`);
    const head = ws.getCell(`A${r}`);
    head.value = `${tc.id}${tc.stepName ? ` — ${tc.stepName}` : ""}`;
    head.font = { name: "Arial", size: 14, bold: true };
    head.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.title } };
    head.alignment = { vertical: "middle", horizontal: "left" };
    ws.getRow(r).height = 26;
    r++;

    ws.mergeCells(`A${r}:${lastLetter}${r}`);
    const back = ws.getCell(`A${r}`);
    back.value = { text: "← Back to Test Cases", hyperlink: "#'Test Cases'!A1" };
    back.font = { name: "Arial", size: 10, underline: true, color: { argb: "FF0563C1" } };
    r += 2;

    for (let e = 0; e < tc.evidence.length; e++) {
      const ev = tc.evidence[e];

      // Heading
      ws.mergeCells(`A${r}:${lastLetter}${r}`);
      const h = ws.getCell(`A${r}`);
      h.value = `Screenshot ${e + 1}: ${ev.heading || "Untitled"}`;
      h.font = { name: "Arial", size: 12, bold: true };
      h.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.evidenceHeading } };
      h.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      for (let c = 1; c <= EVIDENCE_COLS; c++) ws.getCell(r, c).border = allThin;
      ws.getRow(r).height = 22;
      r++;

      // Image
      let dims;
      try {
        dims = await getImageSize(ev.blob);
      } catch {
        throw new Error(`Screenshot "${ev.heading || e + 1}" in ${tc.id} could not be read.`);
      }
      const targetW = Math.min(dims.width, maxImgWidth, totalWidthPx);
      const scale = targetW / dims.width;
      const targetH = Math.round(dims.height * scale);
      const rowsNeeded = Math.ceil(targetH / (IMAGE_ROW_PT / 0.75)) + 1;

      const type = ev.blob.type || "";
      const extension = type.includes("png") ? "png" : "jpeg";
      const imageId = wb.addImage({ buffer: await ev.blob.arrayBuffer(), extension });
      ws.addImage(imageId, {
        tl: { col: 0.05, row: r - 1 + 0.1 },
        ext: { width: Math.round(targetW), height: targetH },
      });
      for (let i = 0; i < rowsNeeded; i++) ws.getRow(r + i).height = IMAGE_ROW_PT;
      r += rowsNeeded;

      // Description
      ws.mergeCells(`A${r}:${lastLetter}${r}`);
      const d = ws.getCell(`A${r}`);
      d.value = ev.description ? `Description: ${ev.description}` : "";
      d.font = { name: "Arial", size: 10 };
      d.alignment = { vertical: "top", horizontal: "left", wrapText: true };
      for (let c = 1; c <= EVIDENCE_COLS; c++) ws.getCell(r, c).border = allThin;
      ws.getRow(r).height = Math.max(
        20,
        estimateLines(d.value, EVIDENCE_COLS * EVIDENCE_COL_WIDTH) * 13 + 6
      );
      r += 2;
    }
    onProgress(40 + Math.round(((t + 1) / withEvidence.length) * 50));
  }

  onProgress(95);
  const buffer = await wb.xlsx.writeBuffer();
  onProgress(100);
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};

/* ==========================================================================
   SMALL COMPONENTS
   ========================================================================== */

function EvidenceThumb({ blob, alt }) {
  const [url, setUrl] = useState(null);
  useEffect(() => {
    if (!blob) return undefined;
    const u = URL.createObjectURL(blob);
    setUrl(u);
    return () => URL.revokeObjectURL(u);
  }, [blob]);
  if (!url) return <div className="thumb-empty">No image</div>;
  return <img className="thumb" src={url} alt={alt} />;
}

function Field({ label, htmlFor, hint, children }) {
  return (
    <div className="field">
      <label htmlFor={htmlFor}>{label}</label>
      {children}
      {hint && <span className="hint">{hint}</span>}
    </div>
  );
}

function StatusBadge({ status }) {
  const cls = status.toLowerCase().replace(/\s+/g, "-");
  return <span className={`badge badge-${cls}`}>{status}</span>;
}

function ConfirmDialog({ title, message, confirmLabel = "Confirm", onConfirm, onCancel }) {
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label={title}>
      <div className="modal">
        <h3>{title}</h3>
        <p>{message}</p>
        <div className="modal-actions">
          <button className="btn" onClick={onCancel}>Cancel</button>
          <button className="btn btn-danger" onClick={onConfirm}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

/* ==========================================================================
   DOCUMENT DETAILS
   ========================================================================== */

function DocumentDetails({ metadata, onChange }) {
  const set = (k) => (e) => onChange({ ...metadata, [k]: e.target.value });
  return (
    <section className="card">
      <h2>Document Information</h2>
      <div className="grid-2">
        <Field label="Client *" htmlFor="client">
          <input id="client" value={metadata.client} onChange={set("client")} placeholder="e.g. Mankind Pharma" />
        </Field>
        <Field label="Interface *" htmlFor="interface">
          <input id="interface" value={metadata.interfaceName} onChange={set("interfaceName")} placeholder="e.g. OutletMaster_Bizom_API_To_SAP" />
        </Field>
        <Field label="Flow *" htmlFor="flow">
          <select id="flow" value={metadata.flow} onChange={set("flow")}>
            {FLOW_OPTIONS.map((f) => <option key={f}>{f}</option>)}
          </select>
        </Field>
        <Field label="Ticket No *" htmlFor="ticket">
          <input id="ticket" value={metadata.ticketNo} onChange={set("ticketNo")} placeholder="e.g. 8388033" />
        </Field>
      </div>
    </section>
  );
}

/* ==========================================================================
   EVIDENCE EDITOR
   ========================================================================== */

function EvidenceEditor({ evidence, onChange, notify }) {
  const fileInput = useRef(null);
  const replaceTarget = useRef(null);
  const replaceInput = useRef(null);

  const accept = ["image/png", "image/jpeg"];

  const addFiles = (files) => {
    const added = [];
    Array.from(files).forEach((f) => {
      if (!accept.includes(f.type)) {
        notify(`"${f.name}" skipped: only PNG and JPEG are supported.`, "error");
        return;
      }
      added.push({
        id: uid(),
        heading: f.name.replace(/\.[^.]+$/, ""),
        description: "",
        blob: f,
        fileName: f.name,
      });
    });
    if (added.length) onChange([...evidence, ...added]);
  };

  // Paste screenshots directly from the clipboard
  const onPaste = (e) => {
    const files = Array.from(e.clipboardData?.files || []).filter((f) => accept.includes(f.type));
    if (files.length) {
      e.preventDefault();
      addFiles(files);
    }
  };

  const update = (id, patch) => onChange(evidence.map((ev) => (ev.id === id ? { ...ev, ...patch } : ev)));

  return (
    <div className="evidence" onPaste={onPaste} tabIndex={0}>
      <div className="evidence-head">
        <h4>Evidence ({evidence.length})</h4>
        <button className="btn" onClick={() => fileInput.current.click()}>+ Add Screenshot</button>
        <input
          ref={fileInput}
          type="file"
          accept="image/png,image/jpeg"
          multiple
          hidden
          onChange={(e) => {
            addFiles(e.target.files);
            e.target.value = "";
          }}
        />
        <input
          ref={replaceInput}
          type="file"
          accept="image/png,image/jpeg"
          hidden
          onChange={(e) => {
            const f = e.target.files[0];
            if (f && accept.includes(f.type) && replaceTarget.current) {
              update(replaceTarget.current, { blob: f, fileName: f.name });
            }
            e.target.value = "";
          }}
        />
      </div>
      <p className="hint">Click here and press Ctrl+V to paste a screenshot from the clipboard.</p>

      {evidence.length === 0 && <div className="empty-small">No screenshots yet. Without screenshots, no evidence sheet or link is created.</div>}

      {evidence.map((ev, i) => (
        <div className="evidence-card" key={ev.id}>
          <EvidenceThumb blob={ev.blob} alt={ev.heading} />
          <div className="evidence-fields">
            <Field label={`Screenshot ${i + 1} heading`}>
              <input value={ev.heading} onChange={(e) => update(ev.id, { heading: e.target.value })} />
            </Field>
            <Field label="Description">
              <textarea rows={2} value={ev.description} onChange={(e) => update(ev.id, { description: e.target.value })} />
            </Field>
            <div className="btn-row">
              <button className="btn btn-sm" onClick={() => { replaceTarget.current = ev.id; replaceInput.current.click(); }}>Replace</button>
              <button className="btn btn-sm" disabled={i === 0} onClick={() => onChange(moveItem(evidence, i, i - 1))}>Move Up</button>
              <button className="btn btn-sm" disabled={i === evidence.length - 1} onClick={() => onChange(moveItem(evidence, i, i + 1))}>Move Down</button>
              <button className="btn btn-sm btn-danger" onClick={() => onChange(evidence.filter((x) => x.id !== ev.id))}>Remove</button>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ==========================================================================
   TEST CASE EDITOR
   ========================================================================== */

function TestCaseEditor({ tc, allCases, prefix, onChange, notify }) {
  const set = (k) => (e) => onChange({ ...tc, [k]: e.target.value });
  const isDuplicate = allCases.filter((t) => t.id.trim() === tc.id.trim()).length > 1;
  const sensitive = looksSensitive(tc.actualOutput);

  const autoId = () => {
    const used = new Set(allCases.filter((t) => t.key !== tc.key).map((t) => t.id));
    let n = allCases.findIndex((t) => t.key === tc.key) + 1;
    while (used.has(`${prefix}${pad(n)}`)) n++;
    onChange({ ...tc, id: `${prefix}${pad(n)}` });
  };

  const doRedact = () => {
    if (window.confirm("Replace detected token/password/secret values in Actual Output with ***REDACTED***? You can undo by retyping the value.")) {
      onChange({ ...tc, actualOutput: redact(tc.actualOutput) });
      notify("Sensitive values were redacted in the Actual Output field.", "info");
    }
  };

  return (
    <section className="card editor">
      <h2>Edit Test Case</h2>

      <div className="grid-3">
        <Field label="Step No" htmlFor="stepNo">
          <input id="stepNo" value={tc.stepNo} onChange={set("stepNo")} placeholder="1" />
        </Field>
        <Field label="Step Name" htmlFor="stepName" hint="Same step no + name on consecutive cases shows the label once.">
          <input id="stepName" value={tc.stepName} onChange={set("stepName")} placeholder="Token" />
        </Field>
        <Field label="Test Case ID *" htmlFor="tcid">
          <div className="input-with-btn">
            <input id="tcid" className={isDuplicate || !tc.id.trim() ? "invalid" : ""} value={tc.id} onChange={set("id")} />
            <button className="btn btn-sm" type="button" onClick={autoId}>Auto</button>
          </div>
          {isDuplicate && <span className="error-text">Duplicate ID</span>}
        </Field>
      </div>

      <Field label="Input" htmlFor="input">
        <textarea id="input" rows={3} value={tc.input} onChange={set("input")} />
      </Field>
      <Field label="Expected Output" htmlFor="expected">
        <textarea id="expected" rows={4} value={tc.expectedOutput} onChange={set("expectedOutput")} />
      </Field>
      <Field label="Actual Output" htmlFor="actual" hint="Stored exactly as typed (JSON/XML is not parsed or modified).">
        <textarea id="actual" className="mono" rows={8} value={tc.actualOutput} onChange={set("actualOutput")} spellCheck={false} />
      </Field>
      {sensitive && (
        <div className="alert alert-warn">
          <span>Actual Output appears to contain a token, password or secret.</span>
          <button className="btn btn-sm" onClick={doRedact}>Redact before export</button>
        </div>
      )}

      <div className="grid-3">
        <Field label="Status" htmlFor="status">
          <select id="status" value={tc.status} onChange={set("status")}>
            {STATUS_OPTIONS.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Test Type" htmlFor="type">
          <select id="type" value={tc.testType} onChange={set("testType")}>
            {TEST_TYPES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="API Used" htmlFor="api">
          <input id="api" value={tc.apiUsed} onChange={set("apiUsed")} placeholder="https://..." />
        </Field>
      </div>
      <Field label="Remarks" htmlFor="remarks">
        <textarea id="remarks" rows={2} value={tc.remarks} onChange={set("remarks")} />
      </Field>

      <EvidenceEditor
        evidence={tc.evidence}
        onChange={(evidence) => onChange({ ...tc, evidence })}
        notify={notify}
      />
    </section>
  );
}

/* ==========================================================================
   TEST CASES PAGE
   ========================================================================== */

function TestCasesPage({ doc, setDoc, notify }) {
  const { testCases, settings } = doc;
  const [selectedKey, setSelectedKey] = useState(testCases[0]?.key ?? null);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [typeFilter, setTypeFilter] = useState("All");
  const [deleteKey, setDeleteKey] = useState(null);

  useEffect(() => {
    if (!testCases.find((t) => t.key === selectedKey)) setSelectedKey(testCases[0]?.key ?? null);
  }, [testCases, selectedKey]);

  const selected = testCases.find((t) => t.key === selectedKey);

  const setCases = (fn) => setDoc((d) => ({ ...d, testCases: fn(d.testCases) }));

  const add = () => {
    const tc = newTestCase(testCases, settings.idPrefix);
    const last = testCases[testCases.length - 1];
    if (last) {
      tc.stepNo = last.stepNo;
      tc.stepName = last.stepName;
    }
    setCases((list) => [...list, tc]);
    setSelectedKey(tc.key);
  };

  const duplicate = (tc) => {
    const copy = {
      ...tc,
      key: uid(),
      id: newTestCase(testCases, settings.idPrefix).id,
      evidence: tc.evidence.map((e) => ({ ...e, id: uid() })),
    };
    setCases((list) => {
      const idx = list.findIndex((t) => t.key === tc.key);
      const next = [...list];
      next.splice(idx + 1, 0, copy);
      return next;
    });
    setSelectedKey(copy.key);
  };

  const filtered = testCases.filter((t) => {
    const q = search.trim().toLowerCase();
    return (
      (!q || t.id.toLowerCase().includes(q) || t.stepName.toLowerCase().includes(q)) &&
      (statusFilter === "All" || t.status === statusFilter) &&
      (typeFilter === "All" || t.testType === typeFilter)
    );
  });

  const filtersActive = search || statusFilter !== "All" || typeFilter !== "All";

  return (
    <div className="split">
      <aside className="card list-pane">
        <div className="list-head">
          <h2>Test Cases ({testCases.length})</h2>
          <button className="btn btn-primary" onClick={add}>+ Add</button>
        </div>

        <input className="search" placeholder="Search by ID or step…" value={search} onChange={(e) => setSearch(e.target.value)} aria-label="Search test cases" />
        <div className="filter-row">
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} aria-label="Filter by status">
            <option>All</option>
            {STATUS_OPTIONS.map((s) => <option key={s}>{s}</option>)}
          </select>
          <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} aria-label="Filter by type">
            <option>All</option>
            {TEST_TYPES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </div>

        {testCases.length === 0 && <div className="empty-small">No test cases yet. Click “+ Add”.</div>}

        <ul className="tc-list">
          {filtered.map((tc) => {
            const idx = testCases.findIndex((t) => t.key === tc.key);
            return (
              <li key={tc.key} className={`tc-item ${tc.key === selectedKey ? "active" : ""}`}>
                <button className="tc-main" onClick={() => setSelectedKey(tc.key)}>
                  <strong>{tc.id || "(no ID)"}</strong>
                  <span className="tc-sub">
                    {tc.stepName ? `${tc.stepNo ? tc.stepNo + " — " : ""}${tc.stepName}` : "No step"} · {tc.testType}
                  </span>
                  <span className="tc-meta">
                    <StatusBadge status={tc.status} />
                    <span className="muted">{tc.evidence.length} screenshot{tc.evidence.length === 1 ? "" : "s"}</span>
                  </span>
                </button>
                <div className="tc-actions">
                  <button className="icon-btn" title="Move up" disabled={!!filtersActive || idx === 0} onClick={() => setCases((l) => moveItem(l, idx, idx - 1))}>↑</button>
                  <button className="icon-btn" title="Move down" disabled={!!filtersActive || idx === testCases.length - 1} onClick={() => setCases((l) => moveItem(l, idx, idx + 1))}>↓</button>
                  <button className="icon-btn" title="Duplicate" onClick={() => duplicate(tc)}>⧉</button>
                  <button className="icon-btn danger" title="Delete" onClick={() => setDeleteKey(tc.key)}>✕</button>
                </div>
              </li>
            );
          })}
          {testCases.length > 0 && filtered.length === 0 && <li className="empty-small">No matches.</li>}
        </ul>
      </aside>

      <div className="editor-pane">
        {selected ? (
          <TestCaseEditor
            key={selected.key}
            tc={selected}
            allCases={testCases}
            prefix={settings.idPrefix}
            notify={notify}
            onChange={(next) => setCases((list) => list.map((t) => (t.key === next.key ? next : t)))}
          />
        ) : (
          <div className="card empty">Select or add a test case to start editing.</div>
        )}
      </div>

      {deleteKey && (
        <ConfirmDialog
          title="Delete test case?"
          message="This also removes its screenshots from the draft. This cannot be undone."
          confirmLabel="Delete"
          onCancel={() => setDeleteKey(null)}
          onConfirm={() => {
            setCases((l) => l.filter((t) => t.key !== deleteKey));
            setDeleteKey(null);
          }}
        />
      )}
    </div>
  );
}

/* ==========================================================================
   PREVIEW PAGE
   ========================================================================== */

function PreviewPage({ doc, validation }) {
  const { metadata, testCases } = doc;
  const sheetNames = [];
  let n = 0;
  testCases.forEach((tc) => {
    if (tc.evidence.length) sheetNames.push({ tc, name: `Evidence_${pad(++n)}` });
  });
  const linkFor = (tc) => sheetNames.find((s) => s.tc.key === tc.key)?.name;

  let lastStep = null;

  return (
    <>
      <section className="card">
        <h2>Validation</h2>
        {validation.errors.length === 0 && validation.warnings.length === 0 && (
          <div className="alert alert-ok">No issues found. Ready to export.</div>
        )}
        {validation.errors.map((e, i) => <div key={`e${i}`} className="alert alert-error">{e}</div>)}
        {validation.warnings.map((w, i) => <div key={`w${i}`} className="alert alert-warn">{w}</div>)}
      </section>

      <section className="card">
        <h2>Sheet 1 — Test Cases</h2>
        <p className="hint">Approximate preview. Merged cells, row heights and image placement may differ slightly in Excel — always open the exported file to confirm.</p>
        <div className="sheet-scroll">
          <table className="sheet">
            <tbody>
              {[["Client", metadata.client], ["Interface", metadata.interfaceName], ["Flow", metadata.flow], ["Ticket No", metadata.ticketNo]].map(([l, v]) => (
                <tr key={l}>
                  <td className="meta-label">{l}</td>
                  <td className="meta-value" colSpan={COLUMNS.length - 1}>{v}</td>
                </tr>
              ))}
              <tr><td className="sheet-title" colSpan={COLUMNS.length}>TEST CASES</td></tr>
              <tr>{COLUMNS.map((c) => <th key={c.key}>{c.header}</th>)}</tr>
              {testCases.map((tc) => {
                const label = [tc.stepNo, tc.stepName].filter(Boolean).join(" — ");
                const show = label && label !== lastStep;
                if (label) lastStep = label;
                const link = linkFor(tc);
                return (
                  <tr key={tc.key}>
                    <td>{show ? label : ""}</td>
                    <td>{link ? <span className="link">{tc.id} → {link}</span> : tc.id}</td>
                    <td className="pre">{tc.input}</td>
                    <td className="pre">{tc.expectedOutput}</td>
                    <td className="pre mono">{tc.actualOutput}</td>
                    <td><StatusBadge status={tc.status} /></td>
                    <td className="pre">{tc.apiUsed}</td>
                    <td className="pre">{tc.remarks}</td>
                    <td>{tc.testType}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>

      <section className="card">
        <h2>Evidence Worksheets ({sheetNames.length})</h2>
        {sheetNames.length === 0 && <div className="empty-small">No test case has screenshots.</div>}
        {sheetNames.map(({ tc, name }) => (
          <div key={name} className="evidence-preview">
            <h4>{name} <span className="muted">· {tc.id}</span></h4>
            <ol>
              {tc.evidence.map((ev) => (
                <li key={ev.id}>
                  <strong>{ev.heading || "Untitled"}</strong>
                  {ev.description && <span className="muted"> — {ev.description}</span>}
                </li>
              ))}
            </ol>
          </div>
        ))}
      </section>
    </>
  );
}

/* ==========================================================================
   SETTINGS PAGE
   ========================================================================== */

function SettingsPage({ doc, setDoc, onClearDraft }) {
  const { settings, metadata } = doc;
  const set = (k) => (e) => setDoc((d) => ({ ...d, settings: { ...d.settings, [k]: e.target.value } }));
  return (
    <section className="card">
      <h2>Settings</h2>
      <div className="grid-2">
        <Field label="Test case ID prefix" htmlFor="prefix" hint="Used by “+ Add”, “Auto” and “Duplicate”. Example: TC_BZ_API_ → TC_BZ_API_001">
          <input id="prefix" value={settings.idPrefix} onChange={set("idPrefix")} />
        </Field>
        <Field label="Screenshot size in Excel" htmlFor="imgsize" hint="Aspect ratio is always preserved. Images are not re-compressed.">
          <select id="imgsize" value={settings.imageSize} onChange={set("imageSize")}>
            {Object.entries(IMAGE_SIZES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </Field>
        <Field label="Output file name" htmlFor="fname" hint={`Leave blank to use: ${buildFileName(metadata)}`}>
          <input id="fname" value={settings.fileName} onChange={set("fileName")} placeholder={buildFileName(metadata)} />
        </Field>
      </div>
      <hr />
      <button className="btn btn-danger" onClick={onClearDraft}>Clear everything and start new document</button>
    </section>
  );
}

/* ==========================================================================
   APP
   ========================================================================== */

const PAGES = [
  { id: "document", label: "Document Details" },
  { id: "testcases", label: "Test Cases" },
  { id: "preview", label: "Preview" },
  { id: "settings", label: "Settings" },
];

export default function App() {
  const [doc, setDoc] = useState(initialDoc);
  const [page, setPage] = useState("document");
  const [toast, setToast] = useState(null);
  const [progress, setProgress] = useState(null);
  const [savedDraftExists, setSavedDraftExists] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);
  const importInput = useRef(null);
  const firstRender = useRef(true);

  const notify = useCallback((message, type = "info") => {
    setToast({ message, type, id: Date.now() });
  }, []);

  useEffect(() => {
    if (!toast) return undefined;
    const t = setTimeout(() => setToast(null), 5000);
    return () => clearTimeout(t);
  }, [toast]);

  // Detect an existing draft on startup
  useEffect(() => {
    dbGet(DRAFT_KEY)
      .then((d) => setSavedDraftExists(!!d))
      .catch(() => {});
  }, []);

  // Track unsaved changes
  useEffect(() => {
    if (firstRender.current) {
      firstRender.current = false;
      return;
    }
    setDirty(true);
  }, [doc]);

  useEffect(() => {
    const handler = (e) => {
      if (dirty) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [dirty]);

  const validation = useMemo(() => validateDocument(doc), [doc]);

  const setMetadata = (metadata) => setDoc((d) => ({ ...d, metadata }));

  /* ---- Drafts ---- */
  const saveDraft = async () => {
    try {
      await dbPut(DRAFT_KEY, { savedAt: Date.now(), doc });
      setDirty(false);
      setSavedDraftExists(true);
      notify("Draft saved in this browser. Use “Export Draft” for a portable backup.", "success");
    } catch (err) {
      notify(`Could not save draft: ${err.message}`, "error");
    }
  };

  const restoreDraft = async () => {
    try {
      const saved = await dbGet(DRAFT_KEY);
      if (!saved) return notify("No saved draft found.", "error");
      setDoc(saved.doc);
      setDirty(false);
      setSavedDraftExists(false);
      notify("Draft restored.", "success");
    } catch (err) {
      notify(`Draft could not be loaded: ${err.message}`, "error");
    }
  };

  const discardDraft = async () => {
    await dbDelete(DRAFT_KEY).catch(() => {});
    setSavedDraftExists(false);
  };

  const exportDraft = async () => {
    try {
      const json = await exportDraftJson(doc);
      downloadBlob(new Blob([json], { type: "application/json" }), `${sanitizeFilename(doc.metadata.ticketNo) || "QA"}_draft.json`);
      notify("Draft exported.", "success");
    } catch (err) {
      notify(`Draft export failed: ${err.message}`, "error");
    }
  };

  const importDraft = async (file) => {
    try {
      const imported = await importDraftJson(await file.text());
      setDoc(imported);
      setPage("testcases");
      notify("Draft imported.", "success");
    } catch (err) {
      notify(`Import failed: ${err.message}`, "error");
    }
  };

  /* ---- Export ---- */
  const exportExcel = async () => {
    const v = validateDocument(doc);
    if (v.errors.length) {
      setPage("preview");
      notify(`Fix ${v.errors.length} issue(s) before exporting.`, "error");
      return;
    }
    if (v.warnings.length && !window.confirm(`${v.warnings.length} warning(s):\n\n${v.warnings.join("\n")}\n\nExport anyway?`)) {
      return;
    }
    try {
      setProgress(0);
      const blob = await buildWorkbook(doc, setProgress);
      const name = doc.settings.fileName.trim()
        ? doc.settings.fileName.trim().replace(/\.xlsx$/i, "") + ".xlsx"
        : buildFileName(doc.metadata);
      downloadBlob(blob, sanitizeFilename(name.replace(/\.xlsx$/i, "")) + ".xlsx");
      notify("Workbook generated and downloaded.", "success");
    } catch (err) {
      notify(`Export failed: ${err.message}`, "error");
    } finally {
      setProgress(null);
    }
  };

  const clearAll = async () => {
    setDoc(initialDoc());
    await dbDelete(DRAFT_KEY).catch(() => {});
    setConfirmClear(false);
    setDirty(false);
    setPage("document");
    notify("Started a new document.", "info");
  };

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <span className="brand-mark">QA</span>
          <h1>QA Documentation Generator</h1>
          {dirty && <span className="unsaved">● Unsaved changes</span>}
        </div>
        <div className="header-actions">
          <button className="btn" onClick={() => importInput.current.click()}>Import Draft</button>
          <input ref={importInput} type="file" accept="application/json,.json" hidden onChange={(e) => { if (e.target.files[0]) importDraft(e.target.files[0]); e.target.value = ""; }} />
          <button className="btn" onClick={exportDraft}>Export Draft</button>
          <button className="btn" onClick={saveDraft}>Save Draft</button>
          <button className="btn btn-primary" onClick={exportExcel} disabled={progress !== null}>
            {progress !== null ? `Generating… ${progress}%` : "Export Excel"}
          </button>
        </div>
        {progress !== null && <div className="progress"><div style={{ width: `${progress}%` }} /></div>}
      </header>

      <div className="layout">
        <nav className="sidebar" aria-label="Sections">
          {PAGES.map((p) => (
            <button key={p.id} className={`nav-item ${page === p.id ? "active" : ""}`} onClick={() => setPage(p.id)}>
              {p.label}
              {p.id === "testcases" && <span className="count">{doc.testCases.length}</span>}
              {p.id === "preview" && validation.errors.length > 0 && <span className="count count-error">{validation.errors.length}</span>}
            </button>
          ))}
          <div className="sidebar-note">
            Everything runs in your browser. No data is uploaded.
          </div>
        </nav>

        <main className="content">
          {savedDraftExists && (
            <div className="alert alert-info">
              <span>A saved draft was found in this browser.</span>
              <span className="btn-row">
                <button className="btn btn-sm btn-primary" onClick={restoreDraft}>Restore</button>
                <button className="btn btn-sm" onClick={discardDraft}>Discard</button>
              </span>
            </div>
          )}

          {page === "document" && <DocumentDetails metadata={doc.metadata} onChange={setMetadata} />}
          {page === "testcases" && <TestCasesPage doc={doc} setDoc={setDoc} notify={notify} />}
          {page === "preview" && <PreviewPage doc={doc} validation={validation} />}
          {page === "settings" && <SettingsPage doc={doc} setDoc={setDoc} onClearDraft={() => setConfirmClear(true)} />}
        </main>
      </div>

      {toast && (
        <div className={`toast toast-${toast.type}`} role="status" onClick={() => setToast(null)}>
          {toast.message}
        </div>
      )}

      {confirmClear && (
        <ConfirmDialog
          title="Start a new document?"
          message="All current data and the saved browser draft will be deleted. Export a draft first if you want to keep it."
          confirmLabel="Clear everything"
          onCancel={() => setConfirmClear(false)}
          onConfirm={clearAll}
        />
      )}
    </div>
  );
}