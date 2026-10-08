const PRESETS = {
  "OpenRouter":"https://openrouter.ai/api/v1","OpenAI":"https://api.openai.com/v1",
  "Anthropic":"https://api.anthropic.com/v1","Google Gemini":"https://generativelanguage.googleapis.com/v1beta/openai",
  "Groq":"https://api.groq.com/openai/v1","Mistral":"https://api.mistral.ai/v1","DeepSeek":"https://api.deepseek.com/v1",
  "xAI":"https://api.x.ai/v1","Together":"https://api.together.xyz/v1","Fireworks":"https://api.fireworks.ai/inference/v1",
  "Cerebras":"https://api.cerebras.ai/v1","Ollama (local)":"http://localhost:11434/v1",
  "LM Studio (local)":"http://localhost:1234/v1","Custom":""
};
const $ = (id) => document.getElementById(id);
const SYSTEM = "You are an expert Excel assistant working inside the user's open workbook. Always read before writing. Prefer formulas over hardcoded numbers. Report briefly what you changed.";
let messages = [], modelCache = [];

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
function show(role, text) {
  const d = document.createElement("div"); d.className = "msg " + role; d.textContent = text;
  $("log").appendChild(d); $("log").scrollTop = 1e9;
}
function persist() { ["base", "key", "model", "preset"].forEach(k => localStorage.setItem("cai_" + k, $(k).value)); }

async function send() {
  const text = $("input").value.trim(); if (!text) return;
  $("input").value = ""; persist();
  if (!messages.length) messages.push({ role: "system", content: SYSTEM });
  messages.push({ role: "user", content: text }); show("user", text);
  $("send").disabled = true;
  try {
    for (let i = 0; i < 20; i++) {
      const res = await fetch(base() + "/chat/completions", { method: "POST", headers: headers(),
        body: JSON.stringify({ model: $("model").value.trim(), messages, tools: TOOLS }) });
      if (!res.ok) { show("error", res.status + ": " + (await res.text()).slice(0, 500)); break; }
      const msg = (await res.json()).choices[0].message;
      messages.push({ role: "assistant", content: msg.content ?? "", tool_calls: msg.tool_calls });
      if (msg.content) show("ai", msg.content);
      if (!msg.tool_calls?.length) break;
      for (const tc of msg.tool_calls) {
        show("tool", "⚙ " + tc.function.name + " " + tc.function.arguments);
        let out; try { out = await runTool(tc.function.name, JSON.parse(tc.function.arguments || "{}")); }
        catch (e) { out = { error: e.message }; }
        messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(out).slice(0, 20000) });
      }
    }
  } catch (e) { show("error", e.message + " (check URL, key, or CORS)"); }
  $("send").disabled = false;
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
  $("reset").onclick = () => { messages = []; $("log").innerHTML = ""; };
  $("input").onkeydown = (e) => { if (e.key === "Enter" && e.ctrlKey) send(); };
});
