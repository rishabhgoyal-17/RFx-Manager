import React, { useState } from "react";
import { Layers, Settings, Users, FileText, Sparkles, Inbox, UploadCloud, CheckCircle2, AlertTriangle, XCircle, Paperclip } from "lucide-react";import * as XLSX from "xlsx";
import mammoth from "mammoth";

// ---------- Design tokens ----------
// Clean B2B SaaS palette — indigo primary on neutral slate surfaces.
const C = {
  ink: "#0F172A",        // slate-900, primary text
  paper: "#F8FAFC",      // slate-50, app background
  paperRaised: "#FFFFFF",
  line: "#E2E8F0",       // slate-200, hairline borders
  lineStrong: "#CBD5E1", // slate-300
  amber: "#4F46E5",      // indigo-600 — primary accent, action, AI-touched
  amberSoft: "#EEF2FF",  // indigo-50
  slate: "#64748B",      // slate-500, secondary text
  green: "#16A34A",
  greenSoft: "#F0FDF4",
  red: "#DC2626",
  redSoft: "#FEF2F2",
};

const SHADOW_SM = "0 1px 2px rgba(15,23,42,0.06)";
const SHADOW_MD = "0 4px 16px rgba(15,23,42,0.08)";
const SHADOW_LG = "0 12px 32px rgba(15,23,42,0.14)";

const FONT_HEAD = "'Inter', -apple-system, sans-serif";
const FONT_UI = "'Inter', -apple-system, sans-serif";
const FONT_MONO = "'IBM Plex Mono', ui-monospace, monospace";

// ---------- Column type options ----------
const COLUMN_TYPES = ["text", "number", "select", "boolean"];

// ---------- Seed data ----------
const seedCategories = [
  {
    categoryId: "cat_pkg",
    name: "Corrugated Packaging",
    description:
      "Cardboard shipping boxes and packaging materials, typically specified by ply count, size, and GSM (paper weight). Quoted per piece or per volume tier.",
  },
];

const seedConfigs = {
  cat_pkg: {
    categoryId: "cat_pkg",
    defaultColumns: [
      { key: "description", label: "Description", type: "text", required: true, aiSuggested: false },
      { key: "qty", label: "Quantity", type: "number", required: true, aiSuggested: false },
      { key: "unitOfQuote", label: "Unit of Quote", type: "select", options: ["per piece", "per 100 pieces", "per kg"], required: true, aiSuggested: false },
      { key: "ply", label: "Ply", type: "select", options: ["3-ply", "5-ply", "7-ply"], required: false, aiSuggested: false },
      { key: "gsm", label: "GSM", type: "number", required: false, aiSuggested: false },
    ],
  },
};

// ---------- API call: suggest columns from category description ----------
async function suggestColumns(categoryName, categoryDescription) {
  const systemPrompt = `You are helping configure a procurement RFx (request for quotation) system.
Given a purchasing category name and description, propose a sensible set of line-item columns a buyer would need to specify each item precisely enough for vendors to quote accurately.

Always include these base columns first: description (text), qty (number), unitOfQuote (select, with 3-5 sensible unit options for this category).
Then add 2-5 category-specific specification columns that matter for THIS category (e.g. dimensions, material, tolerance, grade, size, weight, voltage — whatever is actually relevant).

Respond with ONLY a JSON array, no prose, no markdown fences. Each item:
{ "key": "camelCaseKey", "label": "Human Label", "type": "text"|"number"|"select"|"boolean", "options": [...] (only if type is select), "required": true|false }`;

  const userPrompt = `Category name: ${categoryName}\nCategory description: ${categoryDescription}`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: systemPrompt,
      messages: [{ role: "user", content: userPrompt }],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return parsed.map((c) => ({ ...c, aiSuggested: true }));
}

// ---------- API call: draft line items from a buyer's free-text description ----------
async function draftLineItems(columnSchema, userMessage, existingLineItems) {
  const schemaDesc = columnSchema
    .map((c) => `- ${c.key} (${c.label}, type: ${c.type}${c.options ? ", options: " + c.options.join("/") : ""}${c.required ? ", required" : ""})`)
    .join("\n");

  const systemPrompt = `You are an RFx (request for quotation) drafting assistant for a procurement buyer.
The line items table has EXACTLY these columns — you must only populate these keys, never invent new ones:
${schemaDesc}

The buyer will describe what they need in free text. Produce a list of line items as a JSON array.
Each line item is an object: { "values": { <columnKey>: <value>, ... } }
Where the buyer was vague (e.g. didn't specify every size), use reasonable, clearly-labeled assumptions — the buyer will review and edit every row, so it's fine to propose a reasonable draft rather than asking for every detail up front.
If the buyer's message is a small edit request against existing line items (e.g. "add 2 more sizes", "remove line 3", "change all ply to 5-ply") rather than a fresh draft, apply that edit to the existing items and return the FULL updated list, not just the diff.

Existing line items (may be empty): ${JSON.stringify(existingLineItems || [])}

Respond with ONLY a JSON array of line items, no prose, no markdown fences.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      system: systemPrompt,
      messages: [{ role: "user", content: userMessage }],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return parsed.map((li) => ({ lineId: "li_" + Math.random().toString(36).slice(2, 8), lineQuestions: [], ...li }));
}

// ---------- API call: draft line items from an attached spec image ----------
async function draftLineItemsFromImage(columnSchema, imageFile, existingLineItems) {
  const schemaDesc = columnSchema
    .map((c) => `- ${c.key} (${c.label}, type: ${c.type}${c.options ? ", options: " + c.options.join("/") : ""}${c.required ? ", required" : ""})`)
    .join("\n");

  const systemPrompt = `You are an RFx drafting assistant. The buyer has attached an image (e.g. a spec sheet, BOM, or photo of a rate card) instead of typing a description.
Read the image directly and extract line items matching EXACTLY these columns — never invent new keys:
${schemaDesc}

Existing line items (may be empty, merge/extend rather than duplicate): ${JSON.stringify(existingLineItems || [])}

Respond with ONLY a JSON array of line items, no prose, no markdown fences.
Each item: { "values": { <columnKey>: <value>, ... } }`;

  const b64 = await readFileAsBase64(imageFile);
  const mediaType = imageFile.type || "image/png";

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      system: systemPrompt,
      messages: [
        {
          role: "user",
          content: [
            { type: "text", text: "Extract the line items from this attached image." },
            { type: "image", source: { type: "base64", media_type: mediaType, data: b64 } },
          ],
        },
      ],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return parsed.map((li) => ({ lineId: "li_" + Math.random().toString(36).slice(2, 8), lineQuestions: [], ...li }));
}

// ---------- API call: suggest RFx-level questionnaire ----------
async function suggestQuestionnaire(categoryName, categoryDescription) {
  const systemPrompt = `You are configuring an RFx questionnaire for a procurement category.
Given the category, propose 3-6 RFx-level questions asked ONCE per vendor (not per line item) — covering things like certifications, capacity, lead time, defect rate, years in business, compliance.

For each question, decide:
- "kind": "gate" if it's a pass/fail eligibility check (e.g. "ISO 9001 certified?") that should exclude a vendor outright if failed
- "kind": "ranking" if it's a scalar/numeric factor that should influence ranking alongside price (e.g. defect rate, lead time) rather than being a strict pass/fail

For "ranking" questions, also include "betterDirection": "lower" or "higher" (e.g. defect rate: lower is better; capacity: higher is better).

Respond with ONLY a JSON array, no prose, no markdown fences. Each item:
{ "question": "...", "type": "boolean"|"number"|"text", "kind": "gate"|"ranking", "betterDirection": "lower"|"higher"|null }`;

  const userPrompt = `Category name: ${categoryName}\nCategory description: ${categoryDescription}`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1200,
      system: systemPrompt,
      messages: [{ role: "user", content: userPrompt }],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return parsed.map((q) => ({ qId: "q_" + Math.random().toString(36).slice(2, 8), level: "rfx", aiSuggested: true, ...q }));
}

// ---------- API call: flag unusual line items and propose line-level questions ----------
async function flagUnusualLines(columnSchema, lineItems) {
  const schemaDesc = columnSchema.map((c) => c.key).join(", ");
  const systemPrompt = `You review a procurement RFx's line items (columns: ${schemaDesc}) and flag ONLY the ones that are genuinely unusual enough to warrant a specific clarifying question to vendors — e.g. an unusually tight tolerance, an unusual material, a hazardous or safety-relevant spec, an oddly large or small quantity relative to the rest, or a spec that's ambiguous as written.
Most line items are ordinary and should NOT be flagged — only flag a line if a specific, non-generic question genuinely adds value for that line.

Respond with ONLY a JSON array, no prose, no markdown fences. Each item:
{ "lineId": "<the lineId of the unusual line item>", "question": "<the specific clarifying question>" }
If nothing is unusual, respond with an empty array [].`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1500,
      system: systemPrompt,
      messages: [{ role: "user", content: JSON.stringify(lineItems) }],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  return JSON.parse(cleaned);
}

// ==================== MODULE 3: RFx EVALUATOR ====================

// ---------- File reading helpers ----------
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("Could not read file"));
    r.readAsArrayBuffer(file);
  });
}
function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = () => reject(new Error("Could not read file"));
    r.readAsDataURL(file);
  });
}
function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("Could not read file"));
    r.readAsText(file);
  });
}

// ---------- Convert an arbitrary vendor file into API content block(s) ----------
// This is real per-format parsing — the model never sees pre-digested "clean" data.
async function fileToContentBlocks(file) {
  const name = file.name.toLowerCase();

  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buf = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(buf, { type: "array" });
    let raw = "";
    wb.SheetNames.forEach((sheetName) => {
      const sheet = wb.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
      raw += `--- Sheet: ${sheetName} ---\n${csv}\n\n`;
    });
    return [{ type: "text", text: `[Vendor file: ${file.name}, parsed from Excel — raw grid below, exactly as the vendor structured it]\n\n${raw}` }];
  }

  if (name.endsWith(".docx")) {
    const buf = await readFileAsArrayBuffer(file);
    const result = await mammoth.extractRawText({ arrayBuffer: buf });
    return [{ type: "text", text: `[Vendor file: ${file.name}, parsed from Word — raw text below, paragraph structure as written]\n\n${result.value}` }];
  }

  if (name.endsWith(".pdf")) {
    const b64 = await readFileAsBase64(file);
    return [
      { type: "text", text: `[Vendor file: ${file.name} — original PDF attached below, read directly]` },
      { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
    ];
  }

  if (name.match(/\.(png|jpg|jpeg|webp|gif)$/)) {
    const b64 = await readFileAsBase64(file);
    const mediaType = name.endsWith(".png") ? "image/png" : name.match(/\.(jpg|jpeg)$/) ? "image/jpeg" : name.endsWith(".webp") ? "image/webp" : "image/gif";
    return [
      { type: "text", text: `[Vendor file: ${file.name} — image attached below, e.g. a photo of a rate card. Read it directly, including anything at an angle, blurred, or partially illegible.]` },
      { type: "image", source: { type: "base64", media_type: mediaType, data: b64 } },
    ];
  }

  // plain text / email / anything else
  const text = await readFileAsText(file);
  return [{ type: "text", text: `[Vendor file: ${file.name}, plain text]\n\n${text}` }];
}

// ---------- Core extraction call: one vendor document → one StructuredQuote ----------
async function extractVendorQuote(rfx, vendorName, file) {
  const lineItemsRef = rfx.lineItems.map((li) => ({ lineId: li.lineId, ...li.values }));
  const columnKeys = rfx.columnSchema.map((c) => c.key);
  const unitCol = rfx.columnSchema.find((c) => c.key === "unitOfQuote" || c.key.toLowerCase().includes("unit"));

  const systemPrompt = `You are a procurement RFx evaluator. A vendor has replied to an RFx with a document in whatever format they chose — you must extract their quote and align it against the EXACT line items that were requested, without inventing or guessing anything you are not given.

REQUESTED LINE ITEMS (the RFx — this is the ground truth to match against):
${JSON.stringify(lineItemsRef, null, 1)}

Column keys on each requested line item: ${columnKeys.join(", ")}
${unitCol ? `The unit of quote the buyer specified is the "${unitCol.key}" field per line — if the vendor quoted in a DIFFERENT unit (e.g. "per box" when the buyer asked "per 100 pieces") and did not state a conversion (e.g. pieces per box), you must NOT guess or silently convert. Set unitMatchesRFx to false and quotedValue to what the vendor literally stated, with a resolutionNote explaining the mismatch.` : ""}

RFx-level questionnaire (answer once per vendor, if answered in the document):
${JSON.stringify(rfx.questionnaire.map((q) => ({ qId: q.qId, question: q.question, type: q.type })), null, 1)}

RULES — these are strict:
1. Match the vendor's quoted items to the closest requested lineId by description/spec. If a requested line item was NOT quoted at all, still include it in lineItems with quotedValue: null and resolutionNote: "not quoted by vendor".
2. Never invent a price, a unit conversion, or resolve an ambiguous reference (e.g. "same as last year") unless the actual reference value is present somewhere in what you were given. If you cannot resolve something, say so explicitly in resolutionNote and leave quotedValue as literally stated (or null).
3. confidence should be "low" wherever the source text was ambiguous, hard to read (e.g. blurred/rotated image, garbled OCR-like text), a range instead of a single number, or required interpretation to match a line item.
4. sourceSnippet must be the actual excerpt/region of the document you read the value from — never fabricate a snippet.
5. Currency: capture exactly what the vendor stated (e.g. "USD" vs "INR") — do not silently convert.

Respond with ONLY a JSON object, no prose, no markdown fences, matching exactly:
{
  "lineItems": [
    { "lineId": "...", "quotedValue": number|null, "unit": "...", "unitMatchesRFx": true|false, "currency": "...", "confidence": "high"|"low", "sourceSnippet": "...", "resolutionNote": "..."|null }
  ],
  "coverage": { "requested": <int>, "quoted": <int> },
  "questionnaireAnswers": [ { "qId": "...", "answer": "...", "sourceSnippet": "..." } ],
  "unresolvedFlags": [ "..." ]
}`;

  const contentBlocks = await fileToContentBlocks(file);

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 6000,
      system: systemPrompt,
      messages: [{ role: "user", content: contentBlocks }],
    }),
  });

  const data = await response.json();
  const textBlock = data.content?.find((b) => b.type === "text");
  if (!textBlock) throw new Error("No text response from model");
  const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return { vendorName, fileName: file.name, extractedAt: new Date().toISOString(), ...parsed };
}

// ---------- Small UI atoms ----------
function TabButton({ active, onClick, children, icon }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 10,
        textAlign: "left",
        padding: "9px 12px",
        margin: "2px 10px",
        width: "calc(100% - 20px)",
        background: active ? C.amberSoft : "transparent",
        border: "none",
        borderRadius: 8,
        fontFamily: FONT_UI,
        fontSize: 13.5,
        fontWeight: active ? 600 : 500,
        color: active ? C.amber : C.slate,
        cursor: "pointer",
        letterSpacing: 0.1,
        transition: "background 0.12s ease, color 0.12s ease",
      }}
      onMouseEnter={(e) => { if (!active) e.currentTarget.style.background = "#F1F5F9"; }}
      onMouseLeave={(e) => { if (!active) e.currentTarget.style.background = "transparent"; }}
    >
      {icon && <span style={{ width: 18, display: "flex", alignItems: "center", justifyContent: "center", opacity: active ? 1 : 0.65 }}>{icon}</span>}
      {children}
    </button>
  );
}

function Pill({ children, tone = "neutral" }) {
  const tones = {
    neutral: { bg: "#F1F5F9", fg: C.slate },
    amber: { bg: C.amberSoft, fg: C.amber },
    green: { bg: C.greenSoft, fg: C.green },
    red: { bg: C.redSoft, fg: C.red },
  };
  const t = tones[tone];
  return (
    <span
      style={{
        display: "inline-block",
        padding: "3px 9px",
        borderRadius: 999,
        background: t.bg,
        color: t.fg,
        fontFamily: FONT_UI,
        fontSize: 11,
        fontWeight: 700,
        letterSpacing: 0.2,
      }}
    >
      {children}
    </span>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <label style={{ display: "block", fontFamily: FONT_UI, fontSize: 12.5, fontWeight: 600, color: C.ink, marginBottom: 7, textTransform: "none" }}>
        {label}
      </label>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%",
  padding: "10px 12px",
  border: `1px solid ${C.line}`,
  borderRadius: 8,
  fontFamily: FONT_UI,
  fontSize: 14,
  color: C.ink,
  background: C.paperRaised,
  boxSizing: "border-box",
  outline: "none",
  transition: "border-color 0.12s ease, box-shadow 0.12s ease",
};

// ---------- Category Tab ----------
function CategoryTab({ categories, setCategories, onOpenConfig }) {
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [error, setError] = useState("");

  function addCategory() {
    if (!name.trim() || !description.trim()) {
      setError("Both name and description are required — the description is what the AI uses to suggest columns.");
      return;
    }
    const categoryId = "cat_" + name.toLowerCase().replace(/[^a-z0-9]+/g, "_").slice(0, 24) + "_" + Math.random().toString(36).slice(2, 6);
    setCategories([...categories, { categoryId, name: name.trim(), description: description.trim() }]);
    setName("");
    setDescription("");
    setError("");
  }

  return (
    <div style={{ maxWidth: 1280 }}>
      <h1 style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, letterSpacing: -0.3, color: C.ink, marginBottom: 4 }}>Categories</h1>
      <p style={{ fontFamily: FONT_UI, fontSize: 14, color: C.slate, marginBottom: 28, maxWidth: 640, lineHeight: 1.5 }}>
        A category defines a purchasing domain — corrugated packaging, IT hardware, MRO spares. The description
        matters: it's what the system reads to propose relevant line-item columns in Configuration.
      </p>

      <div style={{ display: "flex", gap: 28, alignItems: "flex-start" }}>
        <div style={{ background: C.paperRaised, border: `1px solid ${C.line}`, borderRadius: 12, padding: 22, boxShadow: SHADOW_SM, width: 340, flexShrink: 0 }}>
          <div style={{ fontFamily: FONT_HEAD, fontSize: 14, fontWeight: 700, color: C.ink, marginBottom: 14 }}>New category</div>
          <Field label="Category name">
            <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Corrugated Packaging" />
          </Field>
          <Field label="Description">
            <textarea
              style={{ ...inputStyle, minHeight: 96, resize: "vertical", fontFamily: FONT_UI }}
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="What is this category, how is it typically specified and quoted?"
            />
          </Field>
          {error && <div style={{ color: C.red, fontFamily: FONT_UI, fontSize: 13, marginBottom: 12 }}>{error}</div>}
          <button
            onClick={addCategory}
            style={{
              width: "100%",
              background: C.amber,
              color: "#fff",
              border: "none",
              padding: "10px 18px",
              borderRadius: 8,
              fontFamily: FONT_UI,
              fontSize: 14,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Add category
          </button>
        </div>

        <div style={{ flex: 1, minWidth: 0 }}>
          {categories.length === 0 ? (
            <EmptyState
              icon={<Layers size={20} />}
              title="No categories yet"
              description="Add your first purchasing category on the left — packaging, IT hardware, MRO spares — anything you buy repeatedly."
            />
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
              {categories.map((c) => (
                <div
                  key={c.categoryId}
                  style={{
                    border: `1px solid ${C.line}`,
                    borderRadius: 12,
                    padding: "18px 20px",
                    background: C.paperRaised,
                    boxShadow: SHADOW_SM,
                    display: "flex",
                    flexDirection: "column",
                    gap: 10,
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
                    <div style={{ width: 30, height: 30, borderRadius: 8, background: C.amberSoft, color: C.amber, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                      <Layers size={15} />
                    </div>
                    <div style={{ fontFamily: FONT_UI, fontWeight: 700, fontSize: 15, color: C.ink }}>{c.name}</div>
                  </div>
                  <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.slate, lineHeight: 1.5, flex: 1 }}>{c.description}</div>
                  <button
                    onClick={() => onOpenConfig(c.categoryId)}
                    style={{
                      alignSelf: "flex-start",
                      background: "transparent",
                      border: `1px solid ${C.lineStrong}`,
                      color: C.ink,
                      padding: "7px 12px",
                      borderRadius: 8,
                      fontFamily: FONT_UI,
                      fontSize: 12.5,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    Configure →
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------- Configuration Tab ----------
function ConfigurationTab({ categories, configs, setConfigs, selectedCategoryId, setSelectedCategoryId }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const category = categories.find((c) => c.categoryId === selectedCategoryId);
  const config = selectedCategoryId ? configs[selectedCategoryId] : null;
  const columns = config?.defaultColumns || [];

  async function handleSuggest() {
    if (!category) return;
    setLoading(true);
    setError("");
    try {
      const suggested = await suggestColumns(category.name, category.description);
      setConfigs({
        ...configs,
        [selectedCategoryId]: { categoryId: selectedCategoryId, defaultColumns: suggested },
      });
    } catch (e) {
      setError("Couldn't get suggestions from the model: " + e.message);
    } finally {
      setLoading(false);
    }
  }

  function updateColumn(idx, patch) {
    const next = columns.map((col, i) => (i === idx ? { ...col, ...col_patch_fix(col, patch) } : col));
    setConfigs({ ...configs, [selectedCategoryId]: { categoryId: selectedCategoryId, defaultColumns: next } });
  }
  function col_patch_fix(col, patch) {
    return patch;
  }

  function removeColumn(idx) {
    const next = columns.filter((_, i) => i !== idx);
    setConfigs({ ...configs, [selectedCategoryId]: { categoryId: selectedCategoryId, defaultColumns: next } });
  }

  function addBlankColumn() {
    const next = [...columns, { key: "field_" + Math.random().toString(36).slice(2, 6), label: "New column", type: "text", required: false, aiSuggested: false }];
    setConfigs({ ...configs, [selectedCategoryId]: { categoryId: selectedCategoryId, defaultColumns: next } });
  }

  return (
    <div style={{ maxWidth: 1200 }}>
      <h1 style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, letterSpacing: -0.3, color: C.ink, marginBottom: 4 }}>Configuration</h1>
      <p style={{ fontFamily: FONT_UI, fontSize: 14, color: C.slate, marginBottom: 24, maxWidth: 620, lineHeight: 1.5 }}>
        Default line-item columns per category. These seed every new RFx in this category — edited per-RFx without
        changing the default here.
      </p>

      <div style={{ marginBottom: 24 }}>
        <label style={{ display: "block", fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 6 }}>Category</label>
        <select
          style={{ ...inputStyle, maxWidth: 320 }}
          value={selectedCategoryId || ""}
          onChange={(e) => setSelectedCategoryId(e.target.value)}
        >
          <option value="" disabled>
            Select a category…
          </option>
          {categories.map((c) => (
            <option key={c.categoryId} value={c.categoryId}>
              {c.name}
            </option>
          ))}
        </select>
      </div>

      {!category && (
        <EmptyState
          icon={<Settings size={20} />}
          title={categories.length === 0 ? "No categories yet" : "Select a category"}
          description={
            categories.length === 0
              ? "Add a category in the Category tab first, then come back here to configure its default columns."
              : "Pick a category above to view or edit its default line-item columns."
          }
        />
      )}

      {category && (
        <>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 14,
              paddingBottom: 14,
              borderBottom: `1px solid ${C.line}`,
            }}
          >
            <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.slate }}>
              {columns.length} column{columns.length !== 1 ? "s" : ""} configured for <strong style={{ color: C.ink }}>{category.name}</strong>
            </div>
            <button
              onClick={handleSuggest}
              disabled={loading}
              style={{
                background: loading ? C.line : C.amber,
                color: loading ? C.slate : "#fff",
                border: "none",
                padding: "9px 16px",
                borderRadius: 8,
                fontFamily: FONT_UI,
                fontSize: 13,
                fontWeight: 600,
                cursor: loading ? "default" : "pointer",
              }}
            >
              {loading ? "Asking the model…" : columns.length > 0 ? "Re-suggest from description" : "Suggest columns from description"}
            </button>
          </div>

          {error && (
            <div style={{ color: C.red, fontFamily: FONT_UI, fontSize: 13, marginBottom: 16, padding: "10px 12px", background: C.redSoft, borderRadius: 8 }}>
              {error}
            </div>
          )}

          <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", background: C.paperRaised, boxShadow: SHADOW_SM }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_UI, fontSize: 13 }}>
              <thead>
                <tr style={{ background: C.paper, borderBottom: `1px solid ${C.line}` }}>
                  <th style={thStyle}>Key</th>
                  <th style={thStyle}>Label</th>
                  <th style={thStyle}>Type</th>
                  <th style={thStyle}>Options (if select)</th>
                  <th style={thStyle}>Required</th>
                  <th style={thStyle}>Source</th>
                  <th style={thStyle}></th>
                </tr>
              </thead>
              <tbody>
                {columns.map((col, idx) => (
                  <tr key={idx} style={{ borderBottom: `1px solid ${C.line}` }}>
                    <td style={tdStyle}>
                      <span style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.slate }}>{col.key}</span>
                    </td>
                    <td style={tdStyle}>
                      <input
                        style={{ ...cellInputStyle }}
                        value={col.label}
                        onChange={(e) => updateColumn(idx, { label: e.target.value })}
                      />
                    </td>
                    <td style={tdStyle}>
                      <select style={cellInputStyle} value={col.type} onChange={(e) => updateColumn(idx, { type: e.target.value })}>
                        {COLUMN_TYPES.map((t) => (
                          <option key={t} value={t}>
                            {t}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td style={tdStyle}>
                      {col.type === "select" ? (
                        <input
                          style={cellInputStyle}
                          value={(col.options || []).join(", ")}
                          onChange={(e) => updateColumn(idx, { options: e.target.value.split(",").map((s) => s.trim()).filter(Boolean) })}
                          placeholder="option1, option2"
                        />
                      ) : (
                        <span style={{ color: C.lineStrong }}>—</span>
                      )}
                    </td>
                    <td style={tdStyle}>
                      <input type="checkbox" checked={!!col.required} onChange={(e) => updateColumn(idx, { required: e.target.checked })} />
                    </td>
                    <td style={tdStyle}>{col.aiSuggested ? <Pill tone="amber">AI</Pill> : <Pill tone="neutral">manual</Pill>}</td>
                    <td style={tdStyle}>
                      <button
                        onClick={() => removeColumn(idx)}
                        style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontFamily: FONT_UI, fontSize: 12 }}
                      >
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
                {columns.length === 0 && (
                  <tr>
                    <td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: C.slate, fontStyle: "italic", padding: 24 }}>
                      No columns yet — suggest from description, or add one manually below.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <button
            onClick={addBlankColumn}
            style={{
              marginTop: 12,
              background: "transparent",
              border: `1px dashed ${C.lineStrong}`,
              color: C.slate,
              padding: "8px 14px",
              borderRadius: 8,
              fontFamily: FONT_UI,
              fontSize: 13,
              cursor: "pointer",
            }}
          >
            + Add column manually
          </button>
        </>
      )}
    </div>
  );
}

const thStyle = {
  textAlign: "left",
  padding: "10px 14px",
  fontFamily: FONT_UI,
  fontSize: 11,
  fontWeight: 700,
  color: C.slate,
  letterSpacing: 0.4,
  textTransform: "uppercase",
  borderBottom: `1px solid ${C.line}`,
};
const tdStyle = { padding: "10px 14px", verticalAlign: "middle" };
const cellInputStyle = {
  width: "100%",
  padding: "7px 9px",
  border: `1px solid #EDF1F5`,
  borderRadius: 6,
  fontFamily: FONT_UI,
  fontSize: 13,
  background: "#FBFCFE",
  boxSizing: "border-box",
  transition: "border-color 0.12s ease, background 0.12s ease",
};

function EmptyState({ icon, title, description, ctaLabel, onCta }) {
  return (
    <div
      style={{
        border: `1px dashed ${C.lineStrong}`,
        borderRadius: 12,
        background: "#FBFCFE",
        padding: "48px 32px",
        textAlign: "center",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: 6,
      }}
    >
      <div style={{ width: 44, height: 44, borderRadius: 10, background: C.amberSoft, color: C.amber, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 6 }}>
        {icon}
      </div>
      <div style={{ fontFamily: FONT_HEAD, fontSize: 15, fontWeight: 700, color: C.ink }}>{title}</div>
      <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.slate, maxWidth: 380, lineHeight: 1.5, marginBottom: ctaLabel ? 10 : 0 }}>{description}</div>
      {ctaLabel && (
        <button
          onClick={onCta}
          style={{ background: C.amber, color: "#fff", border: "none", padding: "9px 16px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
        >
          {ctaLabel}
        </button>
      )}
    </div>
  );
}
function ThinkingDots() {
  const [tick, setTick] = useState(0);
  React.useEffect(() => {
    const id = setInterval(() => setTick((t) => t + 1), 400);
    return () => clearInterval(id);
  }, []);
  const dots = ".".repeat((tick % 3) + 1);
  return <span>thinking{dots}</span>;
}

// ---------- Modal shell ----------
function Modal({ title, onClose, children, width = 480 }) {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(15,23,42,0.5)",
        backdropFilter: "blur(2px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 50,
      }}
      onClick={onClose}
    >
      <div
        style={{ background: C.paperRaised, borderRadius: 14, width, maxWidth: "90vw", maxHeight: "85vh", overflow: "auto", boxShadow: SHADOW_LG }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: `1px solid ${C.line}` }}>
          <div style={{ fontFamily: FONT_HEAD, fontSize: 15.5, fontWeight: 700, color: C.ink }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: C.slate, cursor: "pointer", lineHeight: 1 }}>×</button>
        </div>
        <div style={{ padding: 20 }}>{children}</div>
      </div>
    </div>
  );
}

// ---------- Vendor Tab ----------
function emptyVendorForm() {
  return {
    name: "",
    contactPerson: "",
    mobile: "",
    email: "",
    address: "",
    gstin: "",
    categoryIds: [],
  };
}

function VendorTab({ vendors, setVendors, categories }) {
  const [search, setSearch] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  const [form, setForm] = useState(emptyVendorForm());
  const [editingId, setEditingId] = useState(null);
  const [expandedId, setExpandedId] = useState(null);

  const filtered =
    search.trim().length >= 2
      ? vendors.filter(
        (v) =>
          v.name.toLowerCase().includes(search.toLowerCase()) ||
          v.contactPerson.toLowerCase().includes(search.toLowerCase()) ||
          v.gstin.toLowerCase().includes(search.toLowerCase())
      )
      : vendors;

  const noMatchButSearching = search.trim().length >= 2 && filtered.length === 0;

  function openAdd() {
    setForm({ ...emptyVendorForm(), name: search.trim().length >= 2 ? search : "" });
    setEditingId(null);
    setModalOpen(true);
  }

  function openEdit(v) {
    setForm({ ...v });
    setEditingId(v.vendorId);
    setModalOpen(true);
  }

  function toggleCategory(catId) {
    setForm((f) => {
      const has = f.categoryIds.includes(catId);
      return { ...f, categoryIds: has ? f.categoryIds.filter((id) => id !== catId) : [...f.categoryIds, catId] };
    });
  }

  function save() {
    if (!form.name.trim()) return;
    if (editingId) {
      setVendors(vendors.map((v) => (v.vendorId === editingId ? { ...v, ...form } : v)));
    } else {
      const vendorId = "ven_" + Math.random().toString(36).slice(2, 8);
      setVendors([...vendors, { vendorId, ...form, quoteHistory: [] }]);
    }
    setModalOpen(false);
  }

  function removeVendor(vendorId) {
    setVendors(vendors.filter((v) => v.vendorId !== vendorId));
  }

  return (
    <div style={{ maxWidth: 1280 }}>
      <h1 style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, letterSpacing: -0.3, color: C.ink, marginBottom: 4 }}>Vendors</h1>
      <p style={{ fontFamily: FONT_UI, fontSize: 14, color: C.slate, marginBottom: 24, maxWidth: 620, lineHeight: 1.5 }}>
        Onboard vendors once, reuse across every RFx. Quote history accumulates here over time — this is what
        resolves "same as last year" style responses on future RFx's instead of leaving them unresolved.
      </p>

      <div style={{ display: "flex", gap: 10, marginBottom: 18 }}>
        <input
          style={{ ...inputStyle, maxWidth: 340 }}
          placeholder="Search vendors (name, contact, GSTIN)…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <button
          onClick={openAdd}
          style={{
            background: C.amber,
            color: C.paper,
            border: "none",
            padding: "9px 16px",
            borderRadius: 8,
            fontFamily: FONT_UI,
            fontSize: 13,
            fontWeight: 600,
            cursor: "pointer",
            whiteSpace: "nowrap",
          }}
        >
          + Add vendor
        </button>
      </div>

      {noMatchButSearching && (
        <div
          style={{
            background: C.amberSoft,
            border: `1px solid ${C.amber}`,
            borderRadius: 8,
            padding: "10px 14px",
            marginBottom: 16,
            fontFamily: FONT_UI,
            fontSize: 13,
            color: C.ink,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <span>No vendor matches "{search}".</span>
          <button
            onClick={openAdd}
            style={{ background: C.amber, color: "#fff", border: "none", padding: "6px 12px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
          >
            Add "{search}" as new vendor
          </button>
        </div>
      )}

      {vendors.length === 0 ? (
        <EmptyState
          icon={<Users size={20} />}
          title="No vendors onboarded"
          description="Add vendors here once — they'll be reusable across every RFx, and quote history builds up automatically over time."
          ctaLabel="+ Add your first vendor"
          onCta={openAdd}
        />
      ) : (
        <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", background: C.paperRaised, boxShadow: SHADOW_SM }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_UI, fontSize: 13 }}>
            <thead>
              <tr style={{ background: C.paper, borderBottom: `1px solid ${C.line}` }}>
                <th style={thStyle}>Vendor</th>
                <th style={thStyle}>Contact</th>
                <th style={thStyle}>GSTIN</th>
                <th style={thStyle}>Categories</th>
                <th style={thStyle}>Quote history</th>
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((v) => (
                <React.Fragment key={v.vendorId}>
                  <tr style={{ borderBottom: `1px solid ${C.line}` }}>
                    <td style={tdStyle}>
                      <div style={{ fontWeight: 600, color: C.ink }}>{v.name}</div>
                      <div style={{ fontSize: 11, color: C.slate, fontFamily: FONT_MONO }}>{v.vendorId}</div>
                    </td>
                    <td style={tdStyle}>
                      <div>{v.contactPerson || "—"}</div>
                      <div style={{ fontSize: 11, color: C.slate }}>{v.mobile}</div>
                    </td>
                    <td style={tdStyle}>
                      <span style={{ fontFamily: FONT_MONO, fontSize: 12 }}>{v.gstin || "—"}</span>
                    </td>
                    <td style={tdStyle}>
                      {v.categoryIds.length === 0 ? (
                        <span style={{ color: C.lineStrong }}>—</span>
                      ) : (
                        v.categoryIds.map((cid) => {
                          const cat = categories.find((c) => c.categoryId === cid);
                          return (
                            <span key={cid} style={{ marginRight: 4 }}>
                              <Pill tone="neutral">{cat ? cat.name : cid}</Pill>
                            </span>
                          );
                        })
                      )}
                    </td>
                    <td style={tdStyle}>
                      {v.quoteHistory && v.quoteHistory.length > 0 ? (
                        <button
                          onClick={() => setExpandedId(expandedId === v.vendorId ? null : v.vendorId)}
                          style={{ background: "none", border: "none", color: C.amber, fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
                        >
                          {v.quoteHistory.length} past quote{v.quoteHistory.length !== 1 ? "s" : ""} {expandedId === v.vendorId ? "▲" : "▼"}
                        </button>
                      ) : (
                        <span style={{ color: C.lineStrong, fontStyle: "italic", fontSize: 12 }}>none yet</span>
                      )}
                    </td>
                    <td style={tdStyle}>
                      <button onClick={() => openEdit(v)} style={{ background: "none", border: "none", color: C.slate, cursor: "pointer", fontFamily: FONT_UI, fontSize: 12, marginRight: 10 }}>
                        Edit
                      </button>
                      <button onClick={() => removeVendor(v.vendorId)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontFamily: FONT_UI, fontSize: 12 }}>
                        Remove
                      </button>
                    </td>
                  </tr>
                  {expandedId === v.vendorId && v.quoteHistory && v.quoteHistory.length > 0 && (
                    <tr>
                      <td colSpan={6} style={{ background: C.paper, padding: "10px 20px" }}>
                        <div style={{ fontFamily: FONT_UI, fontSize: 12, color: C.slate }}>
                          Quote history view — populated once Module 3 (Evaluator) runs and stores quotes against this vendor.
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
              {filtered.length === 0 && !noMatchButSearching && vendors.length > 0 && (
                <tr>
                  <td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: C.slate, fontStyle: "italic", padding: 24 }}>
                    No matches.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {modalOpen && (
        <Modal title={editingId ? "Edit vendor" : "Add vendor"} onClose={() => setModalOpen(false)}>
          <Field label="Vendor name">
            <input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. Shree Packaging Co." />
          </Field>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Contact person">
                <input style={inputStyle} value={form.contactPerson} onChange={(e) => setForm({ ...form, contactPerson: e.target.value })} />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Mobile">
                <input style={inputStyle} value={form.mobile} onChange={(e) => setForm({ ...form, mobile: e.target.value })} />
              </Field>
            </div>
          </div>
          <Field label="Email">
            <input style={inputStyle} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} />
          </Field>
          <Field label="Address">
            <textarea style={{ ...inputStyle, minHeight: 60 }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
          </Field>
          <Field label="GSTIN">
            <input style={inputStyle} value={form.gstin} onChange={(e) => setForm({ ...form, gstin: e.target.value })} placeholder="22AAAAA0000A1Z5" />
          </Field>
          <Field label="Categories served">
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
              {categories.length === 0 && <span style={{ color: C.slate, fontSize: 13, fontStyle: "italic" }}>No categories yet — add one in the Category tab.</span>}
              {categories.map((c) => {
                const active = form.categoryIds.includes(c.categoryId);
                return (
                  <button
                    key={c.categoryId}
                    onClick={() => toggleCategory(c.categoryId)}
                    style={{
                      padding: "6px 12px",
                      borderRadius: 8,
                      border: `1px solid ${active ? C.amber : C.line}`,
                      background: active ? C.amberSoft : "#fff",
                      color: active ? C.amber : C.slate,
                      fontFamily: FONT_UI,
                      fontSize: 12,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    {c.name}
                  </button>
                );
              })}
            </div>
          </Field>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
            <button onClick={() => setModalOpen(false)} style={{ background: "none", border: `1px solid ${C.line}`, color: C.slate, padding: "9px 16px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, cursor: "pointer" }}>
              Cancel
            </button>
            <button onClick={save} style={{ background: C.amber, color: C.paper, border: "none", padding: "9px 18px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              {editingId ? "Save changes" : "Add vendor"}
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ---------- RFx Tab ----------
function emptyRFx() {
  return {
    rfxId: "rfx_" + Math.random().toString(36).slice(2, 8),
    title: "",
    categoryId: null,
    columnSchema: [],
    lineItems: [],
    questionnaire: [],
    vendorIds: [],
    status: "draft",
  };
}

function RFxTab({ categories, configs, vendors, rfxList, setRfxList }) {
  const [current, setCurrent] = useState(emptyRFx());
  const [chatInput, setChatInput] = useState("");
  const [chatLog, setChatLog] = useState([]);
  const [drafting, setDrafting] = useState(false);
  const [suggestingQ, setSuggestingQ] = useState(false);
  const [flaggingLines, setFlaggingLines] = useState(false);
  const [error, setError] = useState("");
  const [sentBanner, setSentBanner] = useState("");
  const [attachedImage, setAttachedImage] = useState(null);

  const category = categories.find((c) => c.categoryId === current.categoryId);
  const eligibleVendors = category ? vendors.filter((v) => v.categoryIds.includes(category.categoryId)) : [];

  function pickCategory(categoryId) {
    const conf = configs[categoryId];
    setCurrent({
      ...emptyRFx(),
      categoryId,
      columnSchema: conf ? conf.defaultColumns.map((c) => ({ ...c })) : [],
      title: categories.find((c) => c.categoryId === categoryId)?.name + " RFx",
    });
    setChatLog([]);
  }

  async function handleChatSubmit() {
    if (!chatInput.trim() || current.columnSchema.length === 0) return;
    const msg = chatInput.trim();
    setChatLog((l) => [...l, { role: "user", text: msg }, { role: "assistant", text: "…", thinking: true }]);
    setChatInput("");
    setDrafting(true);
    setError("");
    try {
      const drafted = await draftLineItems(current.columnSchema, msg, current.lineItems);
      setCurrent((c) => ({ ...c, lineItems: drafted }));
      setChatLog((l) => [...l.filter((m) => !m.thinking), { role: "assistant", text: `Drafted ${drafted.length} line item(s). Review and edit the table below — nothing is final.` }]);
    } catch (e) {
      setError("Draft failed: " + e.message);
      setChatLog((l) => [...l.filter((m) => !m.thinking), { role: "assistant", text: "Couldn't draft that — see error below." }]);
    } finally {
      setDrafting(false);
    }
  }

  async function handleImageDraft() {
    if (!attachedImage || current.columnSchema.length === 0) return;
    setChatLog((l) => [...l, { role: "user", text: `📎 ${attachedImage.name}` }, { role: "assistant", text: "…", thinking: true }]);
    setDrafting(true);
    setError("");
    try {
      const drafted = await draftLineItemsFromImage(current.columnSchema, attachedImage, current.lineItems);
      setCurrent((c) => ({ ...c, lineItems: drafted }));
      setChatLog((l) => [...l.filter((m) => !m.thinking), { role: "assistant", text: `Read ${attachedImage.name} and drafted ${drafted.length} line item(s). Review and edit below.` }]);
    } catch (e) {
      setError("Image draft failed: " + e.message);
      setChatLog((l) => [...l.filter((m) => !m.thinking), { role: "assistant", text: "Couldn't read that image — see error below." }]);
    } finally {
      setDrafting(false);
      setAttachedImage(null);
    }
  }

  async function handleSuggestQuestionnaire() {
    if (!category) return;
    setSuggestingQ(true);
    setError("");
    try {
      const qs = await suggestQuestionnaire(category.name, category.description);
      setCurrent((c) => ({ ...c, questionnaire: [...c.questionnaire, ...qs] }));
    } catch (e) {
      setError("Questionnaire suggestion failed: " + e.message);
    } finally {
      setSuggestingQ(false);
    }
  }

  async function handleFlagUnusual() {
    if (current.lineItems.length === 0) return;
    setFlaggingLines(true);
    setError("");
    try {
      const flags = await flagUnusualLines(current.columnSchema, current.lineItems);
      setCurrent((c) => ({
        ...c,
        lineItems: c.lineItems.map((li) => {
          const match = flags.find((f) => f.lineId === li.lineId);
          if (!match) return li;
          return { ...li, lineQuestions: [...li.lineQuestions, { qId: "lq_" + Math.random().toString(36).slice(2, 6), question: match.question, aiSuggested: true }] };
        }),
      }));
    } catch (e) {
      setError("Flagging failed: " + e.message);
    } finally {
      setFlaggingLines(false);
    }
  }

  function updateCell(lineId, key, value) {
    setCurrent((c) => ({
      ...c,
      lineItems: c.lineItems.map((li) => (li.lineId === lineId ? { ...li, values: { ...li.values, [key]: value } } : li)),
    }));
  }

  function removeLine(lineId) {
    setCurrent((c) => ({ ...c, lineItems: c.lineItems.filter((li) => li.lineId !== lineId) }));
  }

  function addBlankLine() {
    const values = {};
    current.columnSchema.forEach((col) => (values[col.key] = ""));
    setCurrent((c) => ({ ...c, lineItems: [...c.lineItems, { lineId: "li_" + Math.random().toString(36).slice(2, 8), values, lineQuestions: [] }] }));
  }

  function removeLineQuestion(lineId, qId) {
    setCurrent((c) => ({
      ...c,
      lineItems: c.lineItems.map((li) => (li.lineId === lineId ? { ...li, lineQuestions: li.lineQuestions.filter((q) => q.qId !== qId) } : li)),
    }));
  }

  function updateQuestion(qId, patch) {
    setCurrent((c) => ({ ...c, questionnaire: c.questionnaire.map((q) => (q.qId === qId ? { ...q, ...patch } : q)) }));
  }
  function removeQuestion(qId) {
    setCurrent((c) => ({ ...c, questionnaire: c.questionnaire.filter((q) => q.qId !== qId) }));
  }
  function addBlankQuestion() {
    setCurrent((c) => ({
      ...c,
      questionnaire: [...c.questionnaire, { qId: "q_" + Math.random().toString(36).slice(2, 8), level: "rfx", question: "", type: "text", kind: "gate", betterDirection: null, aiSuggested: false }],
    }));
  }

  function toggleVendor(vendorId) {
    setCurrent((c) => ({
      ...c,
      vendorIds: c.vendorIds.includes(vendorId) ? c.vendorIds.filter((id) => id !== vendorId) : [...c.vendorIds, vendorId],
    }));
  }

  function sendRfx() {
    const finalized = { ...current, status: "sent", sentAt: new Date().toISOString() };
    setRfxList([...rfxList, finalized]);
    setSentBanner(`RFx "${finalized.title}" sent to ${finalized.vendorIds.length} vendor(s) — simulated (no real email dispatched).`);
    setCurrent(emptyRFx());
    setChatLog([]);
  }

  return (
    <div style={{ maxWidth: 1320 }}>
      <h1 style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, letterSpacing: -0.3, color: C.ink, marginBottom: 4 }}>RFx</h1>
      <p style={{ fontFamily: FONT_UI, fontSize: 14, color: C.slate, marginBottom: 20, maxWidth: 640, lineHeight: 1.5 }}>
        Draft by describing what you need, then edit directly — the table is the source of truth, chat is for
        drafting and bulk edits only.
      </p>

      {sentBanner && (
        <div style={{ background: C.greenSoft, border: `1px solid ${C.green}`, borderRadius: 8, padding: "10px 14px", marginBottom: 18, fontFamily: FONT_UI, fontSize: 13, color: C.ink }}>
          {sentBanner}
        </div>
      )}

      <div style={{ marginBottom: 20, display: "flex", gap: 16, alignItems: "flex-end" }}>
        <div>
          <label style={{ display: "block", fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 6 }}>Category</label>
          <select style={{ ...inputStyle, width: 260 }} value={current.categoryId || ""} onChange={(e) => pickCategory(e.target.value)}>
            <option value="" disabled>
              Select a category…
            </option>
            {categories.map((c) => (
              <option key={c.categoryId} value={c.categoryId}>
                {c.name}
              </option>
            ))}
          </select>
        </div>
        <div style={{ flex: 1 }}>
          <label style={{ display: "block", fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 6 }}>RFx title</label>
          <input style={inputStyle} value={current.title} onChange={(e) => setCurrent({ ...current, title: e.target.value })} disabled={!current.categoryId} />
        </div>
      </div>

      {!current.categoryId && (
        <EmptyState
          icon={<FileText size={20} />}
          title={categories.length === 0 ? "No categories to draft against" : "Pick a category to start"}
          description={
            categories.length === 0
              ? "Create a category first — its default columns seed every RFx drafted in it."
              : "Choose a category above, then describe what you need in the chat box and the draft will appear as an editable table."
          }
        />
      )}

      {current.categoryId && (
        <>
          {/* Chat drafting */}
          <div style={{ border: `1px solid ${C.line}`, borderRadius: 12, background: C.paperRaised, marginBottom: 24, overflow: "hidden", boxShadow: SHADOW_SM }}>
            <div style={{ display: "flex", alignItems: "center", gap: 7, padding: "11px 16px", borderBottom: `1px solid ${C.line}`, background: "linear-gradient(180deg, #FAFBFF, #F8FAFC)", fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.amber, letterSpacing: 0.3 }}>
              <Sparkles size={13} />
              DRAFT / BULK-EDIT VIA CHAT
            </div>
            {chatLog.length > 0 && (
              <div style={{ padding: "12px 16px", maxHeight: 160, overflow: "auto", borderBottom: `1px solid ${C.line}` }}>
                {chatLog.map((m, i) => (
                  <div key={i} style={{ marginBottom: 8, fontFamily: FONT_UI, fontSize: 13 }}>
                    <span style={{ fontWeight: 700, color: m.role === "user" ? C.ink : C.amber }}>{m.role === "user" ? "You: " : "Assistant: "}</span>
                    {m.thinking ? (
                      <span style={{ color: C.slate, display: "inline-flex", gap: 3, verticalAlign: "middle" }}>
                        <ThinkingDots />
                      </span>
                    ) : (
                      <span style={{ color: C.slate }}>{m.text}</span>
                    )}
                  </div>
                ))}
              </div>
            )}



            {attachedImage && (
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "0 16px 10px", fontFamily: FONT_UI, fontSize: 12, color: C.slate }}>
                <FileText size={13} />
                <span>{attachedImage.name}</span>
                <button onClick={() => setAttachedImage(null)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>×</button>
                <button
                  onClick={handleImageDraft}
                  disabled={drafting}
                  style={{ marginLeft: "auto", background: C.amber, color: "#fff", border: "none", padding: "5px 12px", borderRadius: 6, fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
                >
                  Scan & draft from image
                </button>
              </div>
            )}
            <div style={{ padding: 12, display: "flex", gap: 8 }}>
              <label style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 40, border: `1px solid ${C.line}`, borderRadius: 8, cursor: "pointer", color: C.slate, flexShrink: 0 }}>
                <Paperclip size={16} />
                <input
                  type="file"
                  accept="image/*"
                  style={{ display: "none" }}
                  onChange={(e) => e.target.files[0] && setAttachedImage(e.target.files[0])}
                />
              </label>
              <input
                style={{ ...inputStyle, flex: 1 }}
                placeholder='e.g. "30 line items, 3-ply and 5-ply boxes in 5 sizes, quote per 100 pieces" or "add 2 more sizes to the 5-ply lines"'
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleChatSubmit()}
                disabled={drafting}
              />
              <button
                onClick={handleChatSubmit}
                disabled={drafting}
                style={{ background: drafting ? C.line : C.amber, color: drafting ? C.slate : "#fff", border: "none", padding: "0 18px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, fontWeight: 600, cursor: drafting ? "default" : "pointer" }}
              >
                {drafting ? "Thinking…" : "Send"}
              </button>
            </div>
            <div style={{ padding: 12, display: "flex", gap: 8 }}>
              <input
                style={{ ...inputStyle, flex: 1 }}
                placeholder='e.g. "30 line items, 3-ply and 5-ply boxes in 5 sizes, quote per 100 pieces" or "add 2 more sizes to the 5-ply lines"'
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleChatSubmit()}
                disabled={drafting}
              />
              <button
                onClick={handleChatSubmit}
                disabled={drafting}
                style={{ background: drafting ? C.line : C.amber, color: drafting ? C.slate : "#fff", border: "none", padding: "0 18px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, fontWeight: 600, cursor: drafting ? "default" : "pointer" }}
              >
                {drafting ? "Drafting…" : "Send"}
              </button>
            </div>
          </div>

          {error && <div style={{ color: C.red, fontFamily: FONT_UI, fontSize: 13, marginBottom: 16, padding: "10px 12px", background: C.redSoft, borderRadius: 8 }}>{error}</div>}

          {/* Line items table */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.slate }}>
              {current.lineItems.length} line item{current.lineItems.length !== 1 ? "s" : ""}
            </div>
            <button
              onClick={handleFlagUnusual}
              disabled={flaggingLines || current.lineItems.length === 0}
              style={{
                background: "transparent",
                border: `1px solid ${C.amber}`,
                color: C.amber,
                padding: "7px 12px",
                borderRadius: 8,
                fontFamily: FONT_UI,
                fontSize: 12,
                fontWeight: 600,
                cursor: flaggingLines ? "default" : "pointer",
                opacity: current.lineItems.length === 0 ? 0.4 : 1,
              }}
            >
              {flaggingLines ? "Reviewing lines…" : "AI: flag unusual lines"}
            </button>
          </div>

          <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "auto", background: C.paperRaised, marginBottom: 12, boxShadow: SHADOW_SM }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_UI, fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.paper, borderBottom: `1px solid ${C.line}` }}>
                  <th style={{ ...thStyle, width: 28 }}>#</th>
                  {current.columnSchema.map((col) => (
                    <th key={col.key} style={thStyle}>
                      {col.label}
                      {col.required && <span style={{ color: C.red }}> *</span>}
                    </th>
                  ))}
                  <th style={thStyle}>Line questions</th>
                  <th style={thStyle}></th>
                </tr>
              </thead>
              <tbody>
                {current.lineItems.map((li, idx) => (
                  <tr key={li.lineId} style={{ borderBottom: `1px solid ${C.line}` }}>
                    <td style={{ ...tdStyle, color: C.slate }}>{idx + 1}</td>
                    {current.columnSchema.map((col) => (
                      <td key={col.key} style={tdStyle}>
                        {col.type === "select" ? (
                          <select style={cellInputStyle} value={li.values[col.key] || ""} onChange={(e) => updateCell(li.lineId, col.key, e.target.value)}>
                            <option value="">—</option>
                            {(col.options || []).map((o) => (
                              <option key={o} value={o}>
                                {o}
                              </option>
                            ))}
                          </select>
                        ) : col.type === "boolean" ? (
                          <input type="checkbox" checked={!!li.values[col.key]} onChange={(e) => updateCell(li.lineId, col.key, e.target.checked)} />
                        ) : (
                          <input
                            style={cellInputStyle}
                            type={col.type === "number" ? "number" : "text"}
                            value={li.values[col.key] ?? ""}
                            onChange={(e) => updateCell(li.lineId, col.key, e.target.value)}
                          />
                        )}
                      </td>
                    ))}
                    <td style={{ ...tdStyle, minWidth: 180 }}>
                      {li.lineQuestions.length === 0 ? (
                        <span style={{ color: C.lineStrong, fontStyle: "italic", fontSize: 11 }}>—</span>
                      ) : (
                        li.lineQuestions.map((q) => (
                          <div key={q.qId} style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 2 }}>
                            <Pill tone="amber">Q</Pill>
                            <span style={{ fontSize: 11, color: C.ink }}>{q.question}</span>
                            <button onClick={() => removeLineQuestion(li.lineId, q.qId)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 11 }}>
                              ×
                            </button>
                          </div>
                        ))
                      )}
                    </td>
                    <td style={tdStyle}>
                      <button onClick={() => removeLine(li.lineId)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
                {current.lineItems.length === 0 && (
                  <tr>
                    <td colSpan={current.columnSchema.length + 3} style={{ ...tdStyle, textAlign: "center", color: C.slate, fontStyle: "italic", padding: 24 }}>
                      No line items yet — draft via chat above, or add manually.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          <button onClick={addBlankLine} style={{ marginBottom: 32, background: "transparent", border: `1px dashed ${C.lineStrong}`, color: C.slate, padding: "8px 14px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, cursor: "pointer" }}>
            + Add line item manually
          </button>

          {/* Questionnaire */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontFamily: FONT_HEAD, fontSize: 15.5, fontWeight: 700, color: C.ink }}>RFx-level questionnaire</div>
            <button
              onClick={handleSuggestQuestionnaire}
              disabled={suggestingQ}
              style={{ background: suggestingQ ? C.line : C.amber, color: suggestingQ ? C.slate : "#fff", border: "none", padding: "8px 14px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, cursor: suggestingQ ? "default" : "pointer" }}
            >
              {suggestingQ ? "Suggesting…" : "Suggest questionnaire"}
            </button>
          </div>
          <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", background: C.paperRaised, marginBottom: 12, boxShadow: SHADOW_SM }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_UI, fontSize: 13 }}>
              <thead>
                <tr style={{ background: C.paper, borderBottom: `1px solid ${C.line}` }}>
                  <th style={thStyle}>Question</th>
                  <th style={thStyle}>Type</th>
                  <th style={thStyle}>Kind</th>
                  <th style={thStyle}>Better direction</th>
                  <th style={thStyle}>Source</th>
                  <th style={thStyle}></th>
                </tr>
              </thead>
              <tbody>
                {current.questionnaire.map((q) => (
                  <tr key={q.qId} style={{ borderBottom: `1px solid ${C.line}` }}>
                    <td style={tdStyle}>
                      <input style={cellInputStyle} value={q.question} onChange={(e) => updateQuestion(q.qId, { question: e.target.value })} />
                    </td>
                    <td style={tdStyle}>
                      <select style={cellInputStyle} value={q.type} onChange={(e) => updateQuestion(q.qId, { type: e.target.value })}>
                        <option value="boolean">boolean</option>
                        <option value="number">number</option>
                        <option value="text">text</option>
                      </select>
                    </td>
                    <td style={tdStyle}>
                      <select style={cellInputStyle} value={q.kind} onChange={(e) => updateQuestion(q.qId, { kind: e.target.value })}>
                        <option value="gate">gate (pass/fail)</option>
                        <option value="ranking">ranking factor</option>
                      </select>
                    </td>
                    <td style={tdStyle}>
                      {q.kind === "ranking" ? (
                        <select style={cellInputStyle} value={q.betterDirection || "lower"} onChange={(e) => updateQuestion(q.qId, { betterDirection: e.target.value })}>
                          <option value="lower">lower is better</option>
                          <option value="higher">higher is better</option>
                        </select>
                      ) : (
                        <span style={{ color: C.lineStrong }}>—</span>
                      )}
                    </td>
                    <td style={tdStyle}>{q.aiSuggested ? <Pill tone="amber">AI</Pill> : <Pill tone="neutral">manual</Pill>}</td>
                    <td style={tdStyle}>
                      <button onClick={() => removeQuestion(q.qId)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
                {current.questionnaire.length === 0 && (
                  <tr>
                    <td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: C.slate, fontStyle: "italic", padding: 20 }}>
                      No questions yet.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          <button onClick={addBlankQuestion} style={{ marginBottom: 32, background: "transparent", border: `1px dashed ${C.lineStrong}`, color: C.slate, padding: "8px 14px", borderRadius: 8, fontFamily: FONT_UI, fontSize: 13, cursor: "pointer" }}>
            + Add question manually
          </button>

          {/* Vendors */}
          <div style={{ fontFamily: FONT_HEAD, fontSize: 15.5, fontWeight: 700, color: C.ink, marginBottom: 10 }}>Send to vendors</div>
          {eligibleVendors.length === 0 ? (
            <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.slate, fontStyle: "italic", marginBottom: 20 }}>
              No vendors are onboarded for this category yet — add some in the Vendor tab.
            </div>
          ) : (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 20 }}>
              {eligibleVendors.map((v) => {
                const active = current.vendorIds.includes(v.vendorId);
                return (
                  <button
                    key={v.vendorId}
                    onClick={() => toggleVendor(v.vendorId)}
                    style={{
                      padding: "8px 14px",
                      borderRadius: 8,
                      border: `1px solid ${active ? C.amber : C.line}`,
                      background: active ? C.amberSoft : "#fff",
                      color: active ? C.amber : C.slate,
                      fontFamily: FONT_UI,
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    {active ? "✓ " : ""}
                    {v.name}
                  </button>
                );
              })}
            </div>
          )}

          <button
            onClick={sendRfx}
            disabled={current.lineItems.length === 0 || current.vendorIds.length === 0}
            style={{
              background: current.lineItems.length === 0 || current.vendorIds.length === 0 ? C.line : C.amber,
              color: current.lineItems.length === 0 || current.vendorIds.length === 0 ? C.slate : C.paper,
              border: "none",
              padding: "11px 22px",
              borderRadius: 8,
              fontFamily: FONT_UI,
              fontSize: 14,
              fontWeight: 600,
              cursor: current.lineItems.length === 0 || current.vendorIds.length === 0 ? "default" : "pointer",
            }}
          >
            Send RFx (simulated) →
          </button>

          {rfxList.length > 0 && (
            <div style={{ marginTop: 40, paddingTop: 20, borderTop: `1px solid ${C.line}` }}>
              <div style={{ fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.slate, marginBottom: 10 }}>SENT RFx's (outbox log)</div>
              {rfxList.map((r) => (
                <div key={r.rfxId} style={{ fontFamily: FONT_UI, fontSize: 13, color: C.ink, marginBottom: 6 }}>
                  <Pill tone="green">sent</Pill> <strong>{r.title}</strong> — {r.lineItems.length} lines, {r.vendorIds.length} vendors, {new Date(r.sentAt).toLocaleString()}
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ---------- Evaluator Tab ----------
function confidencePill(li) {
  if (li.quotedValue === null || li.quotedValue === undefined) return <Pill tone="red">not quoted</Pill>;
  if (li.unitMatchesRFx === false) return <Pill tone="red">unit mismatch</Pill>;
  if (li.confidence === "low") return <Pill tone="amber">low confidence</Pill>;
  return <Pill tone="green">ok</Pill>;
}

function SourceModal({ cell, lineDesc, onClose }) {
  if (!cell) return null;
  return (
    <Modal title="Source" onClose={onClose} width={520}>
      <div style={{ fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.slate, marginBottom: 4 }}>LINE ITEM</div>
      <div style={{ fontFamily: FONT_UI, fontSize: 14, color: C.ink, marginBottom: 16 }}>{lineDesc}</div>

      <div style={{ fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.slate, marginBottom: 4 }}>EXTRACTED VALUE</div>
      <div style={{ fontFamily: FONT_UI, fontSize: 15, color: C.ink, marginBottom: 16 }}>
        {cell.quotedValue === null || cell.quotedValue === undefined ? "— not quoted —" : `${cell.quotedValue} ${cell.currency || ""} ${cell.unit ? "/ " + cell.unit : ""}`}
        <span style={{ marginLeft: 10 }}>{confidencePill(cell)}</span>
      </div>

      {cell.resolutionNote && (
        <>
          <div style={{ fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.slate, marginBottom: 4 }}>RESOLUTION NOTE</div>
          <div style={{ fontFamily: FONT_UI, fontSize: 13, color: C.red, marginBottom: 16, background: C.redSoft, padding: "8px 12px", borderRadius: 8 }}>{cell.resolutionNote}</div>
        </>
      )}

      <div style={{ fontFamily: FONT_UI, fontSize: 12, fontWeight: 700, color: C.slate, marginBottom: 4 }}>SOURCE EXCERPT (as read from the document)</div>
      <div style={{ fontFamily: FONT_MONO, fontSize: 12.5, color: C.ink, background: "#F8FAFC", padding: "12px 14px", borderRadius: 8, border: `1px solid ${C.line}`, whiteSpace: "pre-wrap", lineHeight: 1.5 }}>
        {cell.sourceSnippet || "— no snippet captured —"}
      </div>
    </Modal>
  );
}

function EvaluatorTab({ rfxList, vendors, setVendors }) {
  const [selectedRfxId, setSelectedRfxId] = useState(rfxList[0]?.rfxId || null);
  const [uploads, setUploads] = useState({}); // vendorId -> File
  const [statuses, setStatuses] = useState({}); // vendorId -> "idle"|"extracting"|"done"|"error"
  const [quotes, setQuotes] = useState({}); // vendorId -> StructuredQuote
  const [errors, setErrors] = useState({});
  const [sourceCell, setSourceCell] = useState(null); // { cell, lineDesc }

  const rfx = rfxList.find((r) => r.rfxId === selectedRfxId);
  const rfxVendors = rfx ? vendors.filter((v) => rfx.vendorIds.includes(v.vendorId)) : [];
  const descKey = rfx?.columnSchema.find((c) => c.key === "description")?.key || "description";

  function pickFile(vendorId, file) {
    setUploads((u) => ({ ...u, [vendorId]: file }));
    setStatuses((s) => ({ ...s, [vendorId]: "idle" }));
  }

  async function runExtraction(vendorId, vendorName) {
    const file = uploads[vendorId];
    if (!file || !rfx) return;
    setStatuses((s) => ({ ...s, [vendorId]: "extracting" }));
    setErrors((e) => ({ ...e, [vendorId]: null }));
    try {
      const quote = await extractVendorQuote(rfx, vendorName, file);
      setQuotes((q) => ({ ...q, [vendorId]: quote }));
      setStatuses((s) => ({ ...s, [vendorId]: "done" }));
      // Persist to vendor's quote history — this is what resolves "same as last year" on future RFx's.
      setVendors((allVendors) =>
        allVendors.map((v) =>
          v.vendorId === vendorId
            ? { ...v, quoteHistory: [...(v.quoteHistory || []), { rfxId: rfx.rfxId, submittedAt: quote.extractedAt, structuredQuote: quote }] }
            : v
        )
      );
    } catch (e) {
      setStatuses((s) => ({ ...s, [vendorId]: "error" }));
      setErrors((er) => ({ ...er, [vendorId]: e.message }));
    }
  }

  function statusPill(vendorId) {
    const st = statuses[vendorId] || "idle";
    if (st === "extracting") return <Pill tone="amber">extracting…</Pill>;
    if (st === "done") return <Pill tone="green">extracted</Pill>;
    if (st === "error") return <Pill tone="red">failed</Pill>;
    return <Pill tone="neutral">not run</Pill>;
  }

  const extractedVendorIds = rfxVendors.filter((v) => quotes[v.vendorId]).map((v) => v.vendorId);

  return (
    <div style={{ maxWidth: 1400 }}>
      <h1 style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, letterSpacing: -0.3, color: C.ink, marginBottom: 4 }}>Evaluator</h1>
      <p style={{ fontFamily: FONT_UI, fontSize: 14, color: C.slate, marginBottom: 24, maxWidth: 700, lineHeight: 1.5 }}>
        Upload each vendor's response in whatever format it arrived — .xlsx, .docx, .pdf, or an image. Each one is
        read directly and aligned against this RFx's exact line items and units. Nothing is cleaned up beforehand.
      </p>

      {rfxList.length === 0 ? (
        <EmptyState icon={<Inbox size={20} />} title="No RFx's sent yet" description="Send an RFx from the RFx tab first — vendor responses get evaluated against it here." />
      ) : (
        <>
          <div style={{ marginBottom: 24 }}>
            <label style={{ display: "block", fontFamily: FONT_UI, fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 6 }}>RFx</label>
            <select style={{ ...inputStyle, maxWidth: 380 }} value={selectedRfxId || ""} onChange={(e) => { setSelectedRfxId(e.target.value); setUploads({}); setStatuses({}); setQuotes({}); setErrors({}); }}>
              {rfxList.map((r) => (
                <option key={r.rfxId} value={r.rfxId}>
                  {r.title} — {r.lineItems.length} lines, {r.vendorIds.length} vendors
                </option>
              ))}
            </select>
          </div>

          {/* Upload slots */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 14, marginBottom: 32 }}>
            {rfxVendors.map((v) => (
              <div key={v.vendorId} style={{ border: `1px solid ${C.line}`, borderRadius: 12, padding: 16, background: C.paperRaised, boxShadow: SHADOW_SM }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <div style={{ fontFamily: FONT_UI, fontWeight: 700, fontSize: 14, color: C.ink }}>{v.name}</div>
                  {statusPill(v.vendorId)}
                </div>
                <label
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    border: `1px dashed ${C.lineStrong}`,
                    borderRadius: 8,
                    padding: "10px 12px",
                    cursor: "pointer",
                    fontFamily: FONT_UI,
                    fontSize: 12.5,
                    color: C.slate,
                    marginBottom: 10,
                  }}
                >
                  <UploadCloud size={15} />
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{uploads[v.vendorId]?.name || "Choose file (.xlsx, .docx, .pdf, image, .txt)"}</span>
                  <input type="file" style={{ display: "none" }} onChange={(e) => e.target.files[0] && pickFile(v.vendorId, e.target.files[0])} accept=".xlsx,.xls,.docx,.pdf,.png,.jpg,.jpeg,.webp,.txt" />
                </label>
                <button
                  onClick={() => runExtraction(v.vendorId, v.name)}
                  disabled={!uploads[v.vendorId] || statuses[v.vendorId] === "extracting"}
                  style={{
                    width: "100%",
                    background: !uploads[v.vendorId] || statuses[v.vendorId] === "extracting" ? C.line : C.amber,
                    color: !uploads[v.vendorId] || statuses[v.vendorId] === "extracting" ? C.slate : "#fff",
                    border: "none",
                    padding: "8px 12px",
                    borderRadius: 8,
                    fontFamily: FONT_UI,
                    fontSize: 12.5,
                    fontWeight: 600,
                    cursor: !uploads[v.vendorId] || statuses[v.vendorId] === "extracting" ? "default" : "pointer",
                  }}
                >
                  {statuses[v.vendorId] === "extracting" ? "Reading document…" : "Extract"}
                </button>
                {quotes[v.vendorId] && (
                  <div style={{ marginTop: 8, fontFamily: FONT_UI, fontSize: 11.5, color: C.slate }}>
                    Coverage: {quotes[v.vendorId].coverage?.quoted ?? "?"} / {quotes[v.vendorId].coverage?.requested ?? rfx.lineItems.length} lines quoted
                  </div>
                )}
                {errors[v.vendorId] && <div style={{ marginTop: 8, fontFamily: FONT_UI, fontSize: 11.5, color: C.red }}>{errors[v.vendorId]}</div>}
              </div>
            ))}
          </div>

          {/* Comparison grid — deterministic merge of already-extracted StructuredQuotes, no AI call */}
          <div style={{ fontFamily: FONT_HEAD, fontSize: 15.5, fontWeight: 700, color: C.ink, marginBottom: 10 }}>Comparison</div>
          {extractedVendorIds.length === 0 ? (
            <EmptyState icon={<FileText size={20} />} title="No quotes extracted yet" description="Upload and extract at least one vendor's response above to see the comparison grid." />
          ) : (
            <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "auto", background: C.paperRaised, boxShadow: SHADOW_SM }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_UI, fontSize: 12.5 }}>
                <thead>
                  <tr style={{ background: C.paper, borderBottom: `1px solid ${C.line}` }}>
                    <th style={{ ...thStyle, position: "sticky", left: 0, background: C.paper, zIndex: 1 }}>Line item</th>
                    {rfxVendors
                      .filter((v) => extractedVendorIds.includes(v.vendorId))
                      .map((v) => (
                        <th key={v.vendorId} style={thStyle}>
                          {v.name}
                          <div style={{ fontWeight: 500, color: C.slate, textTransform: "none", fontSize: 11, marginTop: 2 }}>
                            {quotes[v.vendorId].coverage?.quoted ?? "?"}/{quotes[v.vendorId].coverage?.requested ?? rfx.lineItems.length} quoted
                          </div>
                        </th>
                      ))}
                  </tr>
                </thead>
                <tbody>
                  {rfx.lineItems.map((li) => (
                    <tr key={li.lineId} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ ...tdStyle, position: "sticky", left: 0, background: C.paperRaised, fontWeight: 600, minWidth: 180 }}>{li.values[descKey] || li.lineId}</td>
                      {rfxVendors
                        .filter((v) => extractedVendorIds.includes(v.vendorId))
                        .map((v) => {
                          const q = quotes[v.vendorId];
                          const cell = q.lineItems.find((x) => x.lineId === li.lineId);
                          if (!cell) return <td key={v.vendorId} style={tdStyle}><span style={{ color: C.lineStrong }}>—</span></td>;
                          return (
                            <td
                              key={v.vendorId}
                              style={{ ...tdStyle, cursor: "pointer", minWidth: 150 }}
                              onClick={() => setSourceCell({ cell, lineDesc: li.values[descKey] || li.lineId })}
                            >
                              <div style={{ fontWeight: 600, color: C.ink }}>
                                {cell.quotedValue === null || cell.quotedValue === undefined ? "—" : `${cell.quotedValue} ${cell.currency || ""}`}
                              </div>
                              <div style={{ marginTop: 3 }}>{confidencePill(cell)}</div>
                            </td>
                          );
                        })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <p style={{ fontFamily: FONT_UI, fontSize: 12, color: C.slate, marginTop: 10, fontStyle: "italic" }}>
            Click any cell to see the exact source excerpt it was read from.
          </p>
        </>
      )}

      {sourceCell && <SourceModal cell={sourceCell.cell} lineDesc={sourceCell.lineDesc} onClose={() => setSourceCell(null)} />}
    </div>
  );
}

// ---------- Root ----------
export default function RFxSystem() {
  const [tab, setTab] = useState("category");
  const [categories, setCategories] = useState(seedCategories);
  const [configs, setConfigs] = useState(seedConfigs);
  const [selectedCategoryId, setSelectedCategoryId] = useState(seedCategories[0]?.categoryId || null);
  const [vendors, setVendors] = useState([]);
  const [rfxList, setRfxList] = useState([]);

  return (
    <div style={{ display: "flex", height: "100vh", background: C.paper, fontFamily: FONT_UI }}>
      <div style={{ width: 216, borderRight: `1px solid ${C.line}`, paddingTop: 18, background: C.paperRaised, flexShrink: 0, height: "100%", overflow: "auto" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 18px 18px", borderBottom: `1px solid ${C.line}`, marginBottom: 10, paddingBottom: 16 }}>
          <div style={{ width: 26, height: 26, borderRadius: 7, background: C.amber, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontFamily: FONT_HEAD, fontSize: 13, fontWeight: 800 }}>
            R
          </div>
          <div style={{ fontFamily: FONT_HEAD, fontSize: 15.5, color: C.ink, fontWeight: 700, letterSpacing: -0.2 }}>RFx System</div>
        </div>
        <TabButton icon={<Layers size={15} />} active={tab === "category"} onClick={() => setTab("category")}>Category</TabButton>
        <TabButton icon={<Settings size={15} />} active={tab === "configuration"} onClick={() => setTab("configuration")}>Configuration</TabButton>
        <TabButton icon={<Users size={15} />} active={tab === "vendor"} onClick={() => setTab("vendor")}>Vendor</TabButton>
        <TabButton icon={<FileText size={15} />} active={tab === "rfx"} onClick={() => setTab("rfx")}>RFx</TabButton>
        <TabButton icon={<CheckCircle2 size={15} />} active={tab === "evaluator"} onClick={() => setTab("evaluator")}>Evaluator</TabButton>
      </div>

      <div style={{ flex: 1, padding: "32px 40px", overflow: "auto", height: "100%", boxSizing: "border-box" }}>
        {tab === "category" && (
          <CategoryTab
            categories={categories}
            setCategories={setCategories}
            onOpenConfig={(id) => {
              setSelectedCategoryId(id);
              setTab("configuration");
            }}
          />
        )}
        {tab === "configuration" && (
          <ConfigurationTab
            categories={categories}
            configs={configs}
            setConfigs={setConfigs}
            selectedCategoryId={selectedCategoryId}
            setSelectedCategoryId={setSelectedCategoryId}
          />
        )}
        {tab === "vendor" && <VendorTab vendors={vendors} setVendors={setVendors} categories={categories} />}
        {tab === "rfx" && <RFxTab categories={categories} configs={configs} vendors={vendors} rfxList={rfxList} setRfxList={setRfxList} />}
        {tab === "evaluator" && <EvaluatorTab rfxList={rfxList} vendors={vendors} setVendors={setVendors} />}
      </div>
    </div>
  );
}
