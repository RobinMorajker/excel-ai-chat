const PRESETS = {
  "OpenRouter":"https://openrouter.ai/api/v1","OpenAI":"https://api.openai.com/v1",
  "Anthropic":"https://api.anthropic.com/v1","Google Gemini":"https://generativelanguage.googleapis.com/v1beta/openai",
  "Groq":"https://api.groq.com/openai/v1","Mistral":"https://api.mistral.ai/v1","DeepSeek":"https://api.deepseek.com/v1",
  "xAI":"https://api.x.ai/v1","Together":"https://api.together.xyz/v1","Fireworks":"https://api.fireworks.ai/inference/v1",
  "Cerebras":"https://api.cerebras.ai/v1","Ollama (local)":"http://localhost:11434/v1",
  "LM Studio (local)":"http://localhost:1234/v1","Custom":""
};
const $ = (id) => document.getElementById(id);
const SYSTEM = "You are an expert Excel assistant working inside the user's open workbook. Always read before writing. Prefer formulas over hardcoded numbers. Report briefly what you changed. " +
  "Each user message ends with [Selection: …]: the cells the user had selected in Excel when sending it. Apply the request to that selection unless the message names its own cells, ranges or sheets. " +
  "Named ones take precedence; anything left unnamed (it, this, these, here) still means the selection, and cells named without a sheet are on the selection's sheet.";
// The chat is a tree of turns. A turn is one prompt and everything the AI did for it: msgs go to the model, shown is what
// the log displays. kids are the forks that follow a turn and pick is the one on screen. The root holds the system prompt.
const turn = (parent, msgs) => ({ parent, msgs, shown: [], kids: [], pick: 0 });
const fresh = () => turn(null, [{ role: "system", content: SYSTEM }]);
// forkAt: the turn picked with "Fork from here", whose next prompt starts a new fork. busy: the AI is answering.
let root = fresh(), forkAt = null, busy = false, modelCache = [];

const fn = (name, description, properties = {}, required = []) =>
  ({ type: "function", function: { name, description, parameters: { type: "object", properties, required } } });
const str = { type: "string" };
const TOOLS = [
  fn("list_sheets", "List worksheets and their used ranges"),
  fn("get_selection", "Get the user's selected range address and values"),
  fn("read_range", "Read values and formulas", { sheet: str, address: str }, ["sheet", "address"]),
  fn("write_range", "Write a 2D grid starting at address. Strings beginning with = are formulas.",
     { sheet: str, address: str, data: { type: "array", items: { type: "array", items: str } } }, ["sheet", "address", "data"]),
  fn("add_sheet", "Create a worksheet", { name: str }, ["name"]),
  fn("format_range", "Format a range", { sheet: str, address: str, bold: { type: "boolean" }, fill: str,
     fontColor: str, numberFormat: str, autofit: { type: "boolean" } }, ["sheet", "address"]),
  fn("clear_range", "Clear contents of a range", { sheet: str, address: str }, ["sheet", "address"])
];

async function runTool(name, a) {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    if (name === "list_sheets") {
      const ws = wb.worksheets; ws.load("items/name"); await ctx.sync();
      const u = ws.items.map(s => { const r = s.getUsedRangeOrNullObject(true); r.load("address"); return [s.name, r]; });
      await ctx.sync();
      return u.map(([n, r]) => ({ name: n, used: r.isNullObject ? null : r.address }));
    }
    if (name === "get_selection") { const r = wb.getSelectedRange(); r.load("address,values"); await ctx.sync(); return { address: r.address, values: r.values }; }
    if (name === "add_sheet") { wb.worksheets.add(a.name).activate(); await ctx.sync(); return { ok: true }; }
    const sh = wb.worksheets.getItem(a.sheet);
    if (name === "read_range") { const r = sh.getRange(a.address); r.load("address,values,formulas"); await ctx.sync(); return { address: r.address, values: r.values, formulas: r.formulas }; }
    if (name === "write_range") {
      const cols = Math.max(...a.data.map(r => r.length));
      const grid = a.data.map(r => { const x = r.map(v => v ?? ""); while (x.length < cols) x.push(""); return x; });
      const r = sh.getRange(a.address).getCell(0, 0).getResizedRange(grid.length - 1, cols - 1);
      r.formulas = grid; r.load("address,values"); await ctx.sync();
      return { written: r.address, values: r.values };
    }
    if (name === "format_range") {
      const r = sh.getRange(a.address);
      if (a.bold != null) r.format.font.bold = a.bold;
      if (a.fill) r.format.fill.color = a.fill;
      if (a.fontColor) r.format.font.color = a.fontColor;
      if (a.numberFormat) { r.load("rowCount,columnCount"); await ctx.sync();
        r.numberFormat = Array.from({ length: r.rowCount }, () => Array(r.columnCount).fill(a.numberFormat)); }
      if (a.autofit) r.format.autofitColumns();
      await ctx.sync(); return { ok: true };
    }
    if (name === "clear_range") { sh.getRange(a.address).clear("Contents"); await ctx.sync(); return { ok: true }; }
    return { error: "Unknown tool " + name };
  });
}

const base = () => $("base").value.replace(/\/+$/, "");
function headers() {
  const h = { "Content-Type": "application/json" }, k = $("key").value.trim();
  if (k) h.Authorization = "Bearer " + k;
  if (base().includes("openrouter")) { h["HTTP-Referer"] = "https://excel.local"; h["X-Title"] = "Excel Custom AI"; }
  if (base().includes("anthropic.com")) { h["x-api-key"] = k; h["anthropic-dangerous-direct-browser-access"] = "true"; }
  if (base().includes("azure.com")) h["api-key"] = k;
  return h;
}
// Turns on screen: from the root, follow each turn's picked fork, stopping at a pending "Fork from here".
function path() { const p = []; for (let n = root; n !== forkAt && n.kids[n.pick]; ) p.push(n = n.kids[n.pick]); return p; }
// What the model gets for turn n: the system prompt and the turns above n on n's own branch, whatever is on screen.
const thread = (n) => n ? [...thread(n.parent), ...n.msgs] : [];
// Record a log entry for turn n; redraw if that turn is on screen.
function show(n, role, text) { n.shown.push([role, text]); if (path().includes(n)) render(true); }

const el = (tag, cls, text) => { const e = document.createElement(tag); e.className = cls; e.textContent = text; return e; };
// Edit and Fork are off while the AI answers, so two answers never edit the workbook at once; ‹ › stay on.
const btn = (text, title, onclick, off = busy) => Object.assign(el("button", "", text), { title, onclick, disabled: off });
const bar = (...items) => { const d = el("div", "acts", ""); d.append(...items); return d; };

// myopt: redraws the whole branch on every log entry; draw only the new entry if long chats get slow.
function render(toEnd) {
  const log = $("log"), top = log.scrollTop, p = path();
  log.replaceChildren(...p.flatMap((n, i) => [promptBubble(n), ...n.shown.map(([role, text]) => el("div", "msg " + role, text)),
    ...(i < p.length - 1 ? [bar(btn("Fork from here", "Start a new fork after this answer; this branch stays",
      () => { forkAt = n; render(true); $("input").focus(); }))] : [])]),
    ...(forkAt ? [pending()] : []));
  log.scrollTop = toEnd ? 1e9 : top;
  $("send").disabled = busy;
}

// A prompt bubble: its text, ‹ i/n › when other forks start at the same point, and Edit.
function promptBubble(n) {
  const d = el("div", "msg user", n.msgs[0].content), all = n.parent.kids, i = all.indexOf(n);
  const go = (j, text, label) => { const b = btn(text, label, () => { n.parent.pick = j; forkAt = null; render(); }, j < 0 || j === all.length);
    b.setAttribute("aria-label", label); return b; };
  d.append(bar(...(all.length > 1 ? [go(i - 1, "‹", "Previous fork"), `${i + 1}/${all.length}`, go(i + 1, "›", "Next fork")] : []),
    btn("Edit", "Edit this prompt: the edit becomes a new fork and this one stays", () => edit(n, d))));
  return d;
}

// Swap prompt n's bubble d for a box holding exactly what the model got, [Selection: …] tag included.
// Send asks it as a new fork beside n; Cancel or Esc puts the bubble back.
function edit(n, d) {
  const box = el("textarea", "", ""), w = el("div", "msg user", ""), go = () => box.value.trim() && ask(n.parent, box.value.trim());
  box.value = n.msgs[0].content;
  box.onkeydown = (e) => { if (e.key === "Enter" && e.ctrlKey) go(); if (e.key === "Escape") render(); };
  w.append(box, bar(btn("Send", "Ask this as a new fork (Ctrl+Enter)", go), btn("Cancel", "Keep the prompt as it is (Esc)", () => render(), false)));
  d.replaceWith(w); box.focus();
}

// Shown after "Fork from here" until the next prompt is sent.
function pending() {
  const d = el("div", "msg fork", "New fork: your next prompt continues from the answer above.");
  d.append(bar(btn("Cancel", "Back to the branch you were on", () => { forkAt = null; render(); }, false)));
  return d;
}
function persist() { ["base", "key", "model", "preset"].forEach(k => localStorage.setItem("cai_" + k, $(k).value)); }

// The cells selected in Excel, e.g. `sheet "Ward 1", B2:D10, F1`; "" when no cells are selected (chart, shape) or Excel can't be read.
async function selection() {
  try {
    return await Excel.run(async (ctx) => {
      const s = ctx.workbook.getSelectedRanges(); s.areas.load("items/address"); s.worksheet.load("name"); await ctx.sync();
      // Area addresses carry the sheet ('My Sheet'!B2:D10): keep what follows the last "!", which a cell reference never contains.
      return `sheet ${JSON.stringify(s.worksheet.name)}, ` + s.areas.items.map(a => a.address.slice(a.address.lastIndexOf("!") + 1)).join(", ");
    });
  } catch { return ""; }
}

// Send the input box. It continues the branch on screen, which ends at the turn picked with "Fork from here" if there is one.
async function send() {
  const text = $("input").value.trim(); if (!text || busy) return;
  $("input").value = "";
  await ask(path().at(-1) || root, text);
}

// Add prompt text as a new turn under parent (a new fork when parent already has a follow-up) and let the AI answer it.
async function ask(parent, text) {
  if (busy) return; busy = true; persist();
  // Text already ending in [Selection: …] (an edited prompt keeps its tag) goes as is; otherwise tag what is selected now.
  const sel = /\[Selection: [^\]]*\]$/.test(text) ? "" : await selection(), content = sel ? `${text}\n\n[Selection: ${sel}]` : text;
  const n = turn(parent, [{ role: "user", content }]);
  parent.pick = parent.kids.push(n) - 1; forkAt = null; render(true);
  try {
    for (let i = 0; i < 20; i++) {
      const res = await fetch(base() + "/chat/completions", { method: "POST", headers: headers(),
        body: JSON.stringify({ model: $("model").value.trim(), messages: thread(n), tools: TOOLS }) });
      if (!res.ok) { show(n, "error", res.status + ": " + (await res.text()).slice(0, 500)); break; }
      const msg = (await res.json()).choices[0].message;
      n.msgs.push({ role: "assistant", content: msg.content ?? "", tool_calls: msg.tool_calls });
      if (msg.content) show(n, "ai", msg.content);
      if (!msg.tool_calls?.length) break;
      for (const tc of msg.tool_calls) {
        show(n, "tool", "⚙ " + tc.function.name + " " + tc.function.arguments);
        let out; try { out = await runTool(tc.function.name, JSON.parse(tc.function.arguments || "{}")); }
        catch (e) { out = { error: e.message }; }
        n.msgs.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(out).slice(0, 20000) });
      }
    }
  } catch (e) { show(n, "error", e.message + " (check URL, key, or CORS)"); }
  busy = false; render();
}

async function loadModels() {
  persist(); $("status").textContent = "Loading…";
  try {
    const j = await (await fetch(base() + "/models", { headers: headers() })).json();
    modelCache = (j.data || j.models || []).map(m => ({
      id: m.id || m.name, name: m.name || m.id, ctx: m.context_length || "",
      inPrice: m.pricing ? +m.pricing.prompt * 1e6 : "", outPrice: m.pricing ? +m.pricing.completion * 1e6 : "",
      tools: m.supported_parameters ? m.supported_parameters.includes("tools") : null
    })).filter(m => !$("toolsOnly").checked || m.tools !== false)
      .sort((a, b) => a.id.localeCompare(b.id));
    // DOM nodes, not innerHTML: model names come from the provider and must never run as HTML on the page that holds your key.
    $("models").replaceChildren(...modelCache.map(m => new Option(
      `${m.name}${m.ctx ? " · " + Math.round(m.ctx / 1000) + "k" : ""}${m.inPrice !== "" ? ` · $${m.inPrice.toFixed(2)}/$${m.outPrice.toFixed(2)} per M` : ""}`, m.id)));
    $("status").textContent = modelCache.length + " models loaded";
  } catch (e) { $("status").textContent = "Failed: " + e.message; }
}

async function modelsToSheet() {
  if (!modelCache.length) await loadModels();
  await Excel.run(async (ctx) => {
    let ws = ctx.workbook.worksheets.getItemOrNullObject("Models"); await ctx.sync();
    if (ws.isNullObject) ws = ctx.workbook.worksheets.add("Models"); else ws.getRange().clear();
    const rows = [["Model ID", "Name", "Context", "$ / M input", "$ / M output", "Tools"],
      ...modelCache.map(m => [m.id, m.name, m.ctx, m.inPrice, m.outPrice, m.tools == null ? "" : String(m.tools)])];
    const r = ws.getRange("A1").getResizedRange(rows.length - 1, 5); r.values = rows;
    ws.getRange("A1:F1").format.font.bold = true; ws.getRange("A1:F1").format.fill.color = "#1F4E78";
    ws.getRange("A1:F1").format.font.color = "#FFFFFF"; r.format.autofitColumns(); ws.freezePanes.freezeRows(1);
    ws.activate(); await ctx.sync();
  });
}

Office.onReady(() => {
  $("preset").innerHTML = Object.keys(PRESETS).map(p => `<option>${p}</option>`).join("");
  ["base", "key", "model", "preset"].forEach(k => { const v = localStorage.getItem("cai_" + k); if (v) $(k).value = v; });
  if (!$("base").value) $("base").value = PRESETS[$("preset").value];
  $("preset").onchange = () => { $("base").value = PRESETS[$("preset").value]; persist(); };
  $("load").onclick = loadModels; $("toSheet").onclick = modelsToSheet; $("send").onclick = send;
  $("reset").onclick = () => { root = fresh(); forkAt = null; render(); };
  $("input").onkeydown = (e) => { if (e.key === "Enter" && e.ctrlKey) send(); };
});
