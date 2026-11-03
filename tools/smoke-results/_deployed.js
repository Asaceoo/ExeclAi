
const bridge = window.chrome.webview.hostObjects.sync.excelAiBridge;
const asyncBridge = window.chrome.webview.hostObjects.excelAiBridge;
const ids = ["messages","workbookChip","sheetChip","selectionChip","insertSelectionBtn","refreshContext","prompt","sendButton","connectionState","sessionBadge","summaryBadge","compressButton","sessionSelect","createSessionButton","deleteSessionButton","refbar","refAddress","addRefButton","addSelectionRef","refList","historyList","rollbackAllButton","toolsList","profileSelect","newProfileButton","deleteProfileButton","profileName","presetSelect","baseUrl","apiKey","modelName","temperature","protocolMode","testButton","saveProfileButton","skillList","activeSkillName","confirmOverlay","riskBadge","confirmHeading","confirmTool","confirmTarget","confirmReason","confirmCodeWrap","codeLang","confirmCode","copyCodeButton","confirmHint","approveButton","rejectButton","toast","errorBar","controlledEnabled","controlledTimeout","controlledOutputKb","controlledPython","cliAllow","cliDeny","exeAllow","exeDeny","scriptAllow","scriptDeny","controlledPath","reloadControlledButton","saveControlledButton","atList","stopButton","writeModeSelect","topWriteMode","maxToolCalls","aboutInfo"];
const elements = Object.fromEntries(ids.map(id => [id, document.getElementById(id)]));
let activeConfirmation = null;
let contextTimer = null;
let currentReferences = [];
let currentSessions = [];
let currentToolCatalog = [];
let atState = { open:false, selected:0, items:[] };
let currentHistory = [];
let lastSession = null;

function parseResult(result){ return JSON.parse(result || "{}"); }
function showToast(message){ elements.toast.textContent = message; elements.toast.style.display = "block"; setTimeout(() => elements.toast.style.display = "none", 1800); }
function showError(message){ elements.errorBar.textContent = message || "发生未知错误。"; elements.errorBar.style.display = "block"; setTimeout(() => elements.errorBar.style.display = "none", 6500); }
function append(content, kind){ const item = document.createElement("div"); item.className = kind + " message"; item.textContent = content || "(空回复)"; elements.messages.appendChild(item); elements.messages.scrollTop = elements.messages.scrollHeight; return item; }

function buildToolCard(ev){
  const card = document.createElement("div"); card.className = "toolcard";
  const head = document.createElement("div"); head.className = "head";
  head.setAttribute("title", "点击折叠 / 展开");
  const caret = document.createElement("span"); caret.className = "caret"; caret.textContent = "▼";
  const status = document.createElement("span"); status.textContent = ev.ok ? "✓" : "✗";
  const name = document.createElement("b"); name.setAttribute("title", ev.tool);
  name.style.cssText = "flex:1 1 auto;min-width:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;display:block"; name.textContent = ev.tool;
  const body = document.createElement("div"); body.className = "body";
  head.append(caret, status, name);
  if(ev.ok && ev.summary){ const s = document.createElement("span"); s.className = "muted"; s.setAttribute("title", ev.summary);
    s.style.cssText = "flex:0 1 auto;max-width:50%;min-width:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;margin-left:4px"; s.textContent = " — " + ev.summary; head.appendChild(s); }
  if(!ev.ok && ev.error_message){ const line = document.createElement("div"); line.style.color = "var(--danger)"; line.textContent = ev.error_message; body.appendChild(line); }
  if(!ev.ok && ev.confirmed === false){ const d = document.createElement("div"); d.className = "muted"; d.textContent = "（用户取消，未执行）"; body.appendChild(d); }
  if(!ev.ok && ev.retryable){
    const btn = document.createElement("button"); btn.className = "retry-btn"; btn.textContent = "重试";
    btn.addEventListener("click", async (e) => {
      e.stopPropagation();
      btn.disabled = true; btn.textContent = "重试中…";
      try{
        const r = parseResult(await asyncBridge.RetryTool(ev.tool));
        if(!r.ok) throw new Error(r.error || "重试失败。");
        body.innerHTML = "";
        const okLine = document.createElement("div"); okLine.style.color = "#047857"; okLine.textContent = "✓ 已重试成功";
        body.appendChild(okLine);
        if(r.summary){ const sl = document.createElement("div"); sl.className = "muted"; sl.textContent = r.summary; body.appendChild(sl); }
        status.textContent = "✓"; card.classList.add("collapsed"); showToast("重试成功");
      }catch(err){
        btn.disabled = false; btn.textContent = "重试";
        const e2 = document.createElement("div"); e2.style.color = "var(--danger)"; e2.textContent = err.message || String(err);
        body.appendChild(e2);
      }
    });
    body.appendChild(btn);
  }
  if(ev.ok) card.classList.add("success"); else card.classList.remove("success");
  card.appendChild(head); card.appendChild(body);
  head.addEventListener("click", () => { card.classList.toggle("collapsed"); });
  return card;
}
function setBusy(busy){ elements.sendButton.disabled = busy; elements.saveProfileButton.disabled = busy; elements.testButton.disabled = busy; elements.connectionState.textContent = busy ? "处理中..." : "就绪"; let s = document.getElementById("busyStatus"); if(busy && !s){ s = document.createElement("div"); s.id="busyStatus"; s.className="status"; s.textContent="正在读取 Excel 上下文并调用模型..."; elements.messages.appendChild(s); elements.messages.scrollTop=elements.messages.scrollHeight; } else if(!busy && s){ s.remove(); } }

function renderContext(context){
  if(!context || !context.Workbook){ elements.workbookChip.textContent = "工作簿：未打开"; elements.sheetChip.textContent = "工作表：—"; elements.selectionChip.textContent = "选区：—"; return; }
  const selection = context.Selection;
  elements.workbookChip.textContent = "工作簿：" + (context.Workbook.Name || "—") + (context.Workbook.IsReadOnly ? "（只读）" : "");
  elements.sheetChip.textContent = "工作表：" + (context.Worksheet ? context.Worksheet.Name : "—");
  elements.selectionChip.textContent = "选区：" + (selection ? `${selection.Address} · ${selection.RowCount}×${selection.ColumnCount}` : "—");
}
async function refreshContext(){ try{ const result = parseResult(await asyncBridge.GetContext()); if(result.ok) renderContext(result.context); }catch(error){} }
async function loadSession(){ try{ const result = parseResult(await asyncBridge.GetSession()); if(result.ok){ renderSessionSelect(result.sessions, result.active_session_id); renderSession(result.session); } }catch(error){} }
function formatTime(value){ const time = new Date(value + "Z"); const diff = Math.max(0, Date.now() - time.getTime()) / 60000; if(diff < 1) return "刚刚"; if(diff < 60) return `${Math.floor(diff)} 分钟前`; if(diff < 1440) return `${Math.floor(diff / 60)} 小时前`; return time.toLocaleString("zh-CN"); }

const historyToolNames = { "write_excel_cell":"写入单元格","write_excel_range":"批量写入","write_cells_batch":"批量写入","format_excel_range":"格式设置","clean_excel_data":"数据清洗","filter_excel_data":"数据筛选","merge_or_split_cells":"合并/拆分单元格","insert_recommended_chart":"插入图表","create_excel_chart":"创建图表","create_excel_pivot_summary":"透视汇总","add_excel_worksheet":"新增工作表","rename_excel_worksheet":"重命名工作表","delete_excel_worksheet":"删除工作表","create_excel_sparkline":"插入迷你图","import_delimited_file":"导入文本","import_excel_file_sheet":"导入工作表","merge_delimited_files":"合并文件","generate_analysis_report":"生成报告","apply_page_setup":"页面设置","export_sheets_to_pdf":"导出PDF" };
function historyTarget(operation){
  const affected = operation.affected || {};
  if(operation.undo_kind === "chart-insert") return (affected.sheet ? affected.sheet + "!" : "") + (affected.chart_name ? "图表 " + affected.chart_name : "图表");
  if(operation.undo_kind === "sparkline-insert") return affected.dest_range || "迷你图";
  const address = affected.address || "";
  const sheet = affected.sheet || "";
  if(address) return sheet ? sheet + "!" + address : address;
  return "目标未知";
}
function renderHistory(operations){
  currentHistory = operations || [];
  elements.historyList.innerHTML = "";
  const undoables = (operations || []).filter(o => o && o.can_undo);
  elements.rollbackAllButton.style.display = operations && operations.length > 0 ? "inline-block" : "none";
  elements.rollbackAllButton.disabled = undoables.length === 0;
  if(!operations || operations.length === 0){ elements.historyList.innerHTML = '<div class="empty">暂无可展示的修改记录。</div>'; return; }
  operations.forEach(operation => {
    const row = document.createElement("div"); row.className = "record";
    const info = document.createElement("div");
    const title = document.createElement("div");
    const name = document.createElement("span"); name.textContent = historyToolNames[operation.tool] || operation.tool;
    title.append(name);
    if(operation.turn_label){ const badge = document.createElement("span"); badge.className = "turn-chip"; badge.textContent = operation.turn_label; title.append(badge); }
    title.className = "record-title";
    const detail = document.createElement("small"); detail.textContent = `${historyTarget(operation)} · ${operation.undo_kind || "cells"} · ${formatTime(operation.created_at_utc)} · ${operation.can_undo ? "可回滚" : "不可回滚"}`;
    info.append(title, detail);
    const action = document.createElement("button"); action.className = "btn"; action.textContent = operation.can_undo ? "回滚" : "—"; action.disabled = !operation.can_undo;
    action.addEventListener("click", async () => { action.disabled = true; try{ const result = parseResult(await asyncBridge.RollbackOperation(operation.id)); if(!result.ok) throw new Error(result.error || "回滚失败。"); showToast("已回滚"); renderHistory(result.history); }catch(error){ action.disabled = false; showError(error.message || String(error)); } });
    row.append(info, action); elements.historyList.appendChild(row);
  });
}
async function loadHistory(){
  try{
    const result = parseResult(await asyncBridge.GetHistory());
    if(result.ok){
      renderHistory(result.history);
      // 历史刷新后重建消息，让每条 AI 回复下方按最新 turn_label 显示回滚按钮
      if(lastSession){ renderMessagesFromPayload(lastSession); elements.messages.scrollTop = elements.messages.scrollHeight; }
    }
  }catch(error){ /* 历史刷新失败时保持原有内容 */ }
}

function renderMessagesFromPayload(session){
  elements.messages.innerHTML = "";
  const messages = (session && session.messages) || [];
  if(messages.length === 0){ elements.messages.innerHTML = '<div class="assistant message">新对话。选择范围后描述需求，普通修改会先确认。</div>'; return; }
  messages.forEach(item => {
    const kind = item.role === "user" ? "user" : (item.role === "system" ? "system" : "assistant");
    if(kind === "assistant"){ appendAssistant(item); }
    else { append(item.content || "(空回复)", kind); }
  });
  elements.messages.scrollTop = elements.messages.scrollHeight;
}

// 工具名 → 人性化动作文案（思考过程中的「可读步骤」）
function readableStep(ev){
  const map = { "get_excel_range_values_and_formulas":"读取单元格","get_current_selection_values_and_formulas":"读取当前选区","preview_excel_sheet":"预览工作表结构","write_excel_range":"写入数据","write_excel_cells_batch":"批量写入","format_excel_range":"设置格式","clean_excel_data":"清洗数据","filter_excel_data":"数据筛选","recommend_excel_chart":"推荐图表类型","create_excel_chart":"创建图表","insert_recommended_chart":"插入图表","generate_analysis_report":"生成分析报告","export_selection_to_markdown":"导出 Markdown","import_markdown_to_excel":"导入 Markdown","analyze_range_statistics":"统计分析","analyze_column_correlation":"相关性分析","analyze_value_distribution":"分布统计","create_excel_sparkline":"插入迷你图","create_excel_pivot_summary":"透视汇总","merge_or_split_cells":"合并/拆分单元格","apply_page_setup":"页面设置","export_sheets_to_pdf":"导出 PDF" };
  const label = map[ev.tool] || ev.tool;
  let extra = "";
  if(ev.summary){ try{ const j = JSON.parse(ev.summary); extra = j.address ? (j.address + " · " + (j.rows ? j.rows + "行" : "") + (j.columns ? j.columns + "列" : "")) : ev.summary; }catch(_){ extra = ev.summary; } }
  return (ev.ok ? "→ " : "✗ ") + label + (extra ? " ｜ " + extra : "");
}

// 思考过程面板：模型推理文字（如有）+ 工具步骤可读文案
function buildThinking(msg){
  const block = document.createElement("div");
  block.style.cssText = "margin:3px 0 5px;background:#f0f7ff;border:1px solid #bfdbfe;border-radius:8px;padding:7px 9px;font-size:11px;color:#1e3a8a;cursor:pointer";
  const head = document.createElement("div"); head.style.cssText = "display:flex;align-items:center;gap:6px;font-weight:600";
  const care = document.createElement("span"); care.style.cssText = "width:12px;color:var(--muted);flex:none;transition:transform .12s"; care.textContent = "▼";
  const lab = document.createElement("span"); lab.textContent = "思考过程";
  head.append(care, lab);
  const body = document.createElement("div"); body.style.cssText = "padding:5px 2px 1px";
  if(msg.reasoning && String(msg.reasoning).trim()){
    const r = document.createElement("div"); r.style.cssText = "white-space:pre-wrap;word-break:break-word;margin-bottom:6px;max-height:220px;overflow:auto;padding-bottom:4px;border-bottom:1px dashed #bfdbfe"; r.textContent = String(msg.reasoning).trim(); body.appendChild(r);
  }
  const steps = (msg.tool_events || []).map(readableStep).filter(Boolean);
  if(steps.length){ const s = document.createElement("div"); s.style.cssText = "display:grid;gap:3px"; steps.forEach(txt => { const d = document.createElement("div"); d.textContent = txt; s.appendChild(d); }); body.appendChild(s); }
  block.append(head, body);
  head.addEventListener("click", () => { const isCollapsed = block.classList.toggle("collapsed"); body.style.display = isCollapsed ? "none" : ""; care.textContent = isCollapsed ? "▶" : "▼"; });
  return block;
}

// 每条 AI 回复下方的「逐个操作回滚」小按钮
function buildRollbackStrip(msg){
  if(!msg.turn_label || !currentHistory.length) return null;
  const ops = currentHistory.filter(o => o && o.turn_label === msg.turn_label && o.can_undo);
  if(ops.length === 0) return null;
  const wrap = document.createElement("div"); wrap.style.cssText = "display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:4px 0 6px";
  const cap = document.createElement("span"); cap.className = "muted"; cap.style.cssText = "font-size:11px"; cap.textContent = "↩ 本回复可回滚：";
  wrap.appendChild(cap);
  ops.forEach(op => {
    const b = document.createElement("button"); b.className = "mini-rollback"; b.setAttribute("title", (historyToolNames[op.tool] || op.tool) + "·" + historyTarget(op));
    b.style.cssText = "font-size:11px;padding:2px 8px;border:1px solid #fca5a5;color:#b91c1c;background:#fef2f2;border-radius:7px;cursor:pointer";
    b.textContent = historyTarget(op) || op.tool;
    b.addEventListener("click", async () => { b.disabled = true; try{ const r = parseResult(await asyncBridge.RollbackOperation(op.id)); if(!r.ok) throw new Error(r.error || "回滚失败。"); showToast("已回滚"); renderHistory(r.history || []); renderSession(lastSession); }catch(e){ b.disabled = false; showError(e.message || String(e)); } });
    wrap.appendChild(b);
  });
  return wrap;
}

function appendAssistant(item){
  const hasTools = item.tool_events && item.tool_events.length;
  const hasReason = item.reasoning && String(item.reasoning).trim();
  // 1) 思考过程
  if(hasTools || hasReason){ elements.messages.appendChild(buildThinking(item)); }
  // 2) 工具卡片
  if(hasTools){
    const te = document.createElement("div"); te.style.cssText = "margin:0 0 2px";
    item.tool_events.forEach(ev => te.appendChild(buildToolCard(ev)));
    elements.messages.appendChild(te);
  }
  // 3) AI 返回结果
  if(item.content){ append(item.content, "assistant"); }
  // 4) 逐个操作回滚小按钮（紧贴 AI 回复下方）
  const rb = buildRollbackStrip(item); if(rb){ elements.messages.appendChild(rb); }
}

function renderSession(session){
  lastSession = session || null;
  const badge = elements.sessionBadge;
  renderMessagesFromPayload(session);
  const sumBadge = elements.summaryBadge;
  if(session && session.has_summary){ sumBadge.style.display = "inline-flex"; } else { sumBadge.style.display = "none"; }
  elements.compressButton.disabled = !session || (session.turn_count === 0 && !session.has_summary);
  if(!session || !session.turn_count){ badge.style.display = "none"; badge.className = "session-badge"; return; }
  badge.style.display = "inline-flex";
  badge.className = "session-badge" + (session.restored ? " restored" : "");
  badge.textContent = session.restored ? `📂 已恢复 ${session.turn_count} 轮` : `💬 ${session.turn_count} 轮`;
}

function renderSessionSelect(sessions, activeId){
  currentSessions = sessions || [];
  elements.sessionSelect.innerHTML = "";
  currentSessions.forEach(s => {
    const opt = document.createElement("option");
    opt.value = s.Id;
    opt.textContent = (s.Id === activeId ? "● " : "") + s.Title + " · " + s.TurnCount + "轮";
    opt.selected = (s.Id === activeId);
    if (s.Active) opt.classList.add("active-opt");
    elements.sessionSelect.appendChild(opt);
  });
  if(currentSessions.length === 0){ elements.deleteSessionButton.disabled = true; }
  else { elements.deleteSessionButton.disabled = false; }
}

function renderReferences(references){
  currentReferences = references || [];
  elements.refList.innerHTML = "";
  const bar = elements.refbar;
  if(currentReferences.length === 0){ bar.style.display = "none"; bar.innerHTML = ""; }
  else {
    bar.style.display = "flex";
    bar.innerHTML = '<span class="muted">引用：</span>';
    currentReferences.forEach(ref => {
      const chip = document.createElement("span"); chip.className = "refchip";
      const label = document.createElement("span"); label.textContent = ref.Address + " · " + ref.estimated_tokens + " tok";
      const close = document.createElement("span"); close.className = "x"; close.textContent = "✕";
      close.addEventListener("click", () => removeReference(ref.Id));
      chip.append(label, close); bar.appendChild(chip);
    });
  }
  if(currentReferences.length === 0){ elements.refList.innerHTML = '<div class="empty">尚未添加引用。可输入地址或引用当前选区。</div>'; return; }
  currentReferences.forEach(ref => {
    const row = document.createElement("div"); row.className = "record";
    const info = document.createElement("div");
    const title = document.createElement("div"); title.textContent = ref.Address;
    const detail = document.createElement("small"); detail.textContent = `${ref.RowCount} 行 × ${ref.ColumnCount} 列 · 约 ${ref.estimated_tokens} tokens`;
    info.append(title, detail);
    const action = document.createElement("button"); action.className = "btn danger"; action.textContent = "移除";
    action.addEventListener("click", () => removeReference(ref.Id));
    row.append(info, action); elements.refList.appendChild(row);
  });
}
async function removeReference(id){ try{ const result = parseResult(await asyncBridge.RemoveReference(id)); if(!result.ok) throw new Error(result.error || "移除失败。"); renderReferences(result.references); }catch(error){ showError(error.message || String(error)); } }

function renderProfiles(profiles, activeId, selectFirst){
  profiles = profiles || [];
  elements.profileSelect.innerHTML = "";
  profiles.forEach(p => { const opt = document.createElement("option"); opt.value = p.id; opt.textContent = p.name; if(p.id === activeId) opt.selected = true; elements.profileSelect.appendChild(opt); });
  if(selectFirst && profiles.length > 0){ const target = profiles.find(p => p.id === activeId) || profiles[0]; loadProfileForm(target); }
}

function loadProfileForm(p){
  if(!p) return;
  elements.profileName.value = p.name || "";
  elements.baseUrl.value = p.BaseUrl || "";
  elements.modelName.value = p.ModelName || "";
  elements.temperature.value = p.temperature || 0.1;
  elements.protocolMode.value = (p.protocol || "Auto");
  elements.apiKey.placeholder = p.api_key_set ? "已保存，留空则不修改" : "sk-...";
  elements.apiKey.value = "";
}

function renderSkills(skills){
  elements.skillList.innerHTML = "";
  if(!skills || skills.length === 0){ elements.skillList.innerHTML = '<div class="empty">暂无技能。将 SKILL.md 放入 skills 目录后重启。</div>'; return; }
  skills.forEach(s => {
    const row = document.createElement("div"); row.className = "record";
    const info = document.createElement("div");
    const title = document.createElement("div"); title.textContent = "/" + s.Name + " " + (s.Version ? "v"+s.Version : "") + (s.Valid ? "" : "（无效）");
    const detail = document.createElement("small"); detail.textContent = s.Description || s.InvalidReason || "";
    info.append(title, detail);
    row.append(info); elements.skillList.appendChild(row);
  });
}

const THIRD_PHASE_TOOLS = new Set([
  "preview_external_file","import_delimited_file","import_excel_file_sheet","merge_delimited_files",
  "analyze_range_statistics","analyze_column_correlation","analyze_value_distribution","recommend_excel_chart",
  "create_excel_sparkline","export_excel_chart_image"
]);
const FOURTH_PHASE_TOOLS = new Set([
  "insert_recommended_chart","generate_analysis_report"
]);
const JRTOOL_PHASE_TOOLS = new Set([
  "extract_text_characters","insert_text_into_cell","delete_text_matches","convert_text_to_pinyin","convert_chinese_variant","amount_to_uppercase",
  "bankers_round_values","pick_numbers_to_sum","fill_random_numbers","fill_number_sequence","generate_permutation_list",
  "insert_blank_rows_every_n","delete_blank_rows_range","make_payslip_rows","fill_blank_cells_down","merge_duplicate_cells","fuzzy_match_columns","wrap_formulas_with_iferror","convert_2d_to_1d","convert_1d_to_2d","create_pivot_table",
  "split_sheet_by_field","create_sheet_by_value","compare_two_sheets","protect_sheet_watermark",
  "copy_range_as_image","copy_visible_range_only","smart_paste_delimiter"
]);
const TOOL_RISK_LABELS = { ReadOnly:"只读", NormalWrite:"普通写入", StructuralChange:"结构变更", Dangerous:"高危" };
const TOOL_RISK_ORDER = ["ReadOnly","NormalWrite","StructuralChange","Dangerous"];

function renderTools(tools){
  currentToolCatalog = (tools || []).slice();
  elements.toolsList.innerHTML = "";
  if(!tools || tools.length === 0){ elements.toolsList.innerHTML = '<div class="empty">工具目录为空或尚未加载。</div>'; return; }
  const groups = {};
  tools.forEach(tool => { (groups[tool.risk_level || "ReadOnly"] = groups[tool.risk_level || "ReadOnly"] || []).push(tool); });
  TOOL_RISK_ORDER.forEach(risk => {
    const list = groups[risk] || [];
    list.sort((a, b) => a.name.localeCompare(b.name));
    const head = document.createElement("div");
    head.className = "muted";
    head.style.cssText = "margin:6px 0 4px;font-weight:700;color:#334155";
    head.textContent = (TOOL_RISK_LABELS[risk] || risk) + "（" + list.length + " 个）";
    elements.toolsList.appendChild(head);
    list.forEach(tool => {
      const row = document.createElement("div"); row.className = "record";
      const info = document.createElement("div");
      const title = document.createElement("div");
      title.textContent = tool.name;
      if(THIRD_PHASE_TOOLS.has(tool.name)){
        const badge = document.createElement("span");
        badge.textContent = "三期";
        badge.style.cssText = "margin-left:6px;font-size:10px;color:var(--brand);background:#eaf1ff;border:1px solid #bcd3f7;border-radius:99px;padding:1px 6px";
        title.appendChild(badge);
      }
      if(FOURTH_PHASE_TOOLS.has(tool.name)){
        const badge = document.createElement("span");
        badge.textContent = "四期";
        badge.style.cssText = "margin-left:6px;font-size:10px;color:#7c3aed;background:#f3e8ff;border:1px solid #d8b4fe;border-radius:99px;padding:1px 6px";
        title.appendChild(badge);
      }
      if(JRTOOL_PHASE_TOOLS.has(tool.name)){
        const badge = document.createElement("span");
        badge.textContent = "JR";
        badge.style.cssText = "margin-left:6px;font-size:10px;color:#0f766e;background:#ccfbf1;border:1px solid #5eead4;border-radius:99px;padding:1px 6px";
        title.appendChild(badge);
      }
      const detail = document.createElement("small"); detail.textContent = tool.description || "";
      info.append(title, detail);
      row.append(info); elements.toolsList.appendChild(row);
    });
  });
}

function closeConfirm(){
  if(activeConfirmation){ asyncBridge.RespondConfirmation(activeConfirmation.call_id, false); activeConfirmation = null; }
  elements.confirmOverlay.classList.remove("open");
}

window.chrome.webview.addEventListener("message", event => {
  const data = event.data || {};
  if(data.type === "busy"){ setBusy(data.payload === true); }
  else if(data.type === "tool-error"){ showError(String(data.payload || "工具执行失败。")); }
  else if(data.type === "confirmation-request"){
    activeConfirmation = data.payload;
    if(data.payload.code && String(data.payload.code).trim()){
      elements.riskBadge.textContent = "受控执行";
      elements.riskBadge.style.color = "#b45309"; elements.riskBadge.style.background = "#fff7ed"; elements.riskBadge.style.borderColor = "#fed7aa";
      elements.confirmHeading.textContent = "脚本 / 代码预览";
      elements.confirmCode.textContent = String(data.payload.code);
      elements.codeLang.textContent = (data.payload.code_language ? data.payload.code_language + " · " : "") + "等待你审阅";
      elements.confirmCodeWrap.style.display = "block";
      elements.confirmHint.textContent = "执行前已创建快照；不满意可在历史中回滚。仅在你确认后才执行。";
      elements.approveButton.textContent = "执行代码";
      elements.copyCodeButton.textContent = "复制";
    }else{
      const level = /danger/i.test(data.payload.risk_level) ? "L3 \u9ad8\u5371" : /structural/i.test(data.payload.risk_level) ? "L2 \u7ed3\u6784\u53d8\u66f4" : "L1 \u666e\u901a\u5199\u5165";
      elements.riskBadge.textContent = level; elements.riskBadge.style.color = ""; elements.riskBadge.style.background = ""; elements.riskBadge.style.borderColor = "";
      elements.confirmHeading.textContent = "确认修改 Excel";
      elements.confirmCodeWrap.style.display = "none";
      elements.confirmHint.textContent = "执行前已创建快照；完成后可在历史中回滚。";
      elements.approveButton.textContent = "允许执行";
    }
    elements.confirmTool.textContent = data.payload.tool; elements.confirmTarget.textContent = data.payload.target; elements.confirmReason.textContent = data.payload.reason;
    elements.confirmOverlay.classList.add("open");
  }
});

async function send(){
  const prompt = elements.prompt.value.trim(); if(!prompt || elements.sendButton.disabled) return;
  append(prompt, "user"); elements.prompt.value = ""; setBusy(true);
  elements.sendButton.style.display = "none"; elements.stopButton.style.display = "";
  try{
    const result = parseResult(await asyncBridge.SendPrompt(prompt));
    if(!result.ok){
      if(result.stopped){ append("已停止生成。", "system"); }
      else { throw new Error(result.error || "处理失败。"); }
    }else{
      renderSessionSelect(result.sessions, result.session?.id);
      renderSession(result.ok ? (result.session || result) : null);
      if(result.used_skill){ elements.activeSkillName.style.display = "inline"; elements.activeSkillName.className = "skill-tag"; elements.activeSkillName.textContent = "技能已加载"; }
    }
  }
  catch(error){ append(error.message || String(error), "error"); showError(error.message || String(error)); }
  finally{ setBusy(false); loadHistory(); elements.sendButton.style.display = ""; elements.stopButton.style.display = "none"; }
}

elements.stopButton.addEventListener("click", () => { try{ asyncBridge.StopGeneration(); }catch(_){} });

elements.sendButton.addEventListener("click", send);

// ---- 预制宏按钮：内建 + 用户自定义，存 localStorage ----
const QUICK_BUILTIN = [
  { name: "结构摘要", prompt: "分析当前选区的结构和数据质量，给出摘要和清洗建议。" },
  { name: "清洗数据", prompt: "清洗当前选区：去空行、修复日期、填充空白。" },
  { name: "生成公式", prompt: "在当前选区右侧生成销售汇总公式。" },
  { name: "透视图表", prompt: "根据当前数据创建透视表或图表。" }
];
const QUICK_KEY = "xai_macro_presets_v1";
function loadCustomMacros(){ try{ const arr = JSON.parse(localStorage.getItem(QUICK_KEY) || "[]"); return Array.isArray(arr) ? arr : []; }catch(_){ return []; } }
function saveCustomMacros(list){ try{ localStorage.setItem(QUICK_KEY, JSON.stringify(list)); }catch(_){} }
function renderQuickRow(){
  const row = document.getElementById("quickRow"); if(!row) return;
  row.innerHTML = "";
  const custom = loadCustomMacros();
  QUICK_BUILTIN.concat(custom).forEach(p => {
    const chip = document.createElement("span");
    chip.className = "quick-chip"; chip.textContent = p.name;
    chip.title = p.prompt;
    chip.addEventListener("click", () => { elements.prompt.value = p.prompt; elements.prompt.focus(); });
    const isBuiltin = QUICK_BUILTIN.some(b => b.name === p.name);
    if(!isBuiltin){
      chip.classList.add("custom");
      const del = document.createElement("i"); del.className = "macro-del"; del.textContent = "×";
      del.setAttribute("role", "button"); del.title = "删除此宏";
      del.addEventListener("click", (ev) => { ev.stopPropagation(); if(!confirm(`删除宏「${p.name}」？`)) return; saveCustomMacros(loadCustomMacros().filter(x => x.name !== p.name)); renderQuickRow(); });
      chip.appendChild(del);
    }
    row.appendChild(chip);
  });
  const add = document.createElement("span"); add.className = "quick-chip add"; add.textContent = "+ 新增宏";
  add.addEventListener("click", () => {
    const name = prompt("宏名称："); if(!name || !name.trim()) return;
    const content = prompt(`宏「${name.trim()}」的指令内容：`); if(content == null || !String(content).trim()) return;
    const list = loadCustomMacros();
    if(list.some(x => x.name === name.trim())){ showError("该宏名称已存在。"); return; }
    list.push({ name: name.trim(), prompt: String(content).trim() });
    saveCustomMacros(list); renderQuickRow(); showToast("宏已保存");
  });
  row.appendChild(add);
}
renderQuickRow();

// ---- D1：输入框 @ 工具补全 ----
function currentAtToken(textarea){
  const value = textarea.value, pos = textarea.selectionStart;
  const before = value.slice(0, pos);
  const lastSpace = Math.max(before.lastIndexOf(" "), before.lastIndexOf("\n"));
  const token = before.slice(lastSpace + 1);
  return token.startsWith("@") ? token.slice(1) : null;
}
function refreshAtSelected(){ Array.from(elements.atList.children).forEach((el, i) => { el.classList.toggle("selected", i === atState.selected); }); }
function renderAtMenu(items){
  atState.items = items; atState.selected = 0;
  const box = elements.atList;
  box.innerHTML = "";
  if(items.length === 0){ const empty = document.createElement("div"); empty.style.cssText = "padding:7px 10px;color:var(--muted)"; empty.textContent = "无匹配工具"; box.appendChild(empty); box.classList.add("open"); return; }
  items.forEach((tool, i) => {
    const div = document.createElement("div"); div.className = "at-item" + (i === 0 ? " selected" : "");
    const b = document.createElement("b"); b.textContent = "@" + tool.name;
    const risk = document.createElement("span"); risk.className = "at-risk"; risk.textContent = TOOL_RISK_LABELS[tool.risk_level] || tool.risk_level;
    const desc = document.createElement("span"); desc.className = "at-desc"; desc.textContent = tool.description || "";
    div.append(b, desc, risk);
    div.addEventListener("mousedown", (e) => { e.preventDefault(); insertAtTool(tool.name); });
    div.addEventListener("mousemove", () => { atState.selected = i; refreshAtSelected(); });
    box.appendChild(div);
  });
  box.classList.add("open");
}
function openAtMenu(token){
  atState.open = true;
  const items = currentToolCatalog.filter(t => t.name.toLowerCase().includes(token.toLowerCase())).slice(0, 12);
  renderAtMenu(items);
}
function closeAtMenu(){ atState.open = false; elements.atList.classList.remove("open"); }
function insertAtTool(toolName){
  const ta = elements.prompt, value = ta.value, pos = ta.selectionStart;
  const before = value.slice(0, pos), after = value.slice(ta.selectionEnd);
  const lastSpace = Math.max(before.lastIndexOf(" "), before.lastIndexOf("\n"));
  const tokenStart = lastSpace + 1;
  const newBefore = before.slice(0, tokenStart) + "@" + toolName;
  ta.value = newBefore + " " + after;
  const newPos = newBefore.length + 1;
  ta.focus(); ta.setSelectionRange(newPos, newPos);
  closeAtMenu();
}
function handlePromptInput(){
  const token = currentAtToken(elements.prompt);
  if(token !== null){ openAtMenu(token); } else { closeAtMenu(); }
}
elements.prompt.addEventListener("input", handlePromptInput);
elements.prompt.addEventListener("keydown", event => {
  if(atState.open){
    if(event.key === "ArrowDown" && atState.items.length){ event.preventDefault(); atState.selected = (atState.selected + 1) % atState.items.length; refreshAtSelected(); return; }
    if(event.key === "ArrowUp" && atState.items.length){ event.preventDefault(); atState.selected = (atState.selected - 1 + atState.items.length) % atState.items.length; refreshAtSelected(); return; }
    if(event.key === "Enter" && atState.items.length){ event.preventDefault(); insertAtTool(atState.items[atState.selected].name); return; }
    if(event.key === "Escape"){ closeAtMenu(); return; }
  }
  if(event.key === "Enter" && !event.shiftKey){ event.preventDefault(); send(); }
});
elements.prompt.addEventListener("blur", () => { setTimeout(closeAtMenu, 120); });
elements.refreshContext.addEventListener("click", async () => { await refreshContext(); showToast("选区已刷新"); });
elements.rollbackAllButton.addEventListener("click", async () => {
  if(!confirm("将按最新→最早顺序回滚当前工作簿的全部可撤销操作，整个过程不可中断。确定继续？")) return;
  setBusy(true);
  try{
    const result = parseResult(await asyncBridge.RollbackAllOperations());
    if(!result.ok && !(result.failed && result.failed > 0)) throw new Error(result.error || "回滚失败。");
    renderHistory(result.history);
    showToast(`已回滚 ${result.rolled_back ?? 0} 条${result.skipped ? `，跳过 ${result.skipped} 条` : ""}${result.failed ? `，失败 ${result.failed} 条` : ""}`);
  }catch(error){ showError(error.message || String(error)); }
  finally{ setBusy(false); }
});

function applySessionUpdate(result, toastText){
  renderSessionSelect(result.sessions, result.session?.id);
  renderSession(result.session);
  // 会话(含恢复)渲染后同步刷新当前工作簿历史：让每条 AI 回复下方能按最新 turn_label 显示回滚按钮
  loadHistory();
  if(toastText) showToast(toastText);
}
async function createSession(){ try{ const result = parseResult(await asyncBridge.CreateSession()); if(!result.ok) throw new Error(result.error || "操作失败。"); elements.activeSkillName.style.display = "none"; applySessionUpdate(result, "已开启新对话"); }catch(error){ showError(error.message || String(error)); } }
async function deleteSession(){ const id = elements.sessionSelect.value; if(!id) return; if(!confirm("删除当前对话？该对话的多轮记录将永久清除。")) return; try{ const result = parseResult(await asyncBridge.DeleteSession(id)); if(!result.ok) throw new Error(result.error || "删除失败。"); elements.activeSkillName.style.display = "none"; applySessionUpdate(result, "已删除对话"); }catch(error){ showError(error.message || String(error)); } }
async function switchSession(){ const id = elements.sessionSelect.value; if(!id) return; try{ const result = parseResult(await asyncBridge.SwitchSession(id)); if(!result.ok) throw new Error(result.error || "切换失败。"); elements.activeSkillName.style.display = "none"; applySessionUpdate(result); }catch(error){ showError(error.message || String(error)); } }
async function compressContext(){
  if(!confirm("把当前多轮对话压缩成一段摘要后继续？原始对话会被清空，摘要将作为后续上下文保留。")) return;
  setBusy(true);
  try{
    const result = parseResult(await asyncBridge.CompressContext());
    if(!result.ok) throw new Error(result.error || "压缩失败。");
    applySessionUpdate(result, "已压缩为摘要，可继续对话");
    elements.messages.innerHTML = "";
    const note = document.createElement("div");
    note.className = "assistant message";
    const title = document.createElement("b"); title.textContent = "已压缩上下文。后续对话将以这段摘要继续：";
    const pre = document.createElement("div");
    pre.className = "summary-block"; pre.style.cssText = "margin-top:6px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:8px;padding:7px;color:#065f46;white-space:pre-wrap;word-break:break-word";
    pre.textContent = result.summary || "";
    note.append(title, pre);
    elements.messages.appendChild(note);
  }catch(error){ showError(error.message || String(error)); }
  finally{ setBusy(false); }
}
elements.createSessionButton.addEventListener("click", createSession);
elements.deleteSessionButton.addEventListener("click", deleteSession);
elements.sessionSelect.addEventListener("change", switchSession);
elements.compressButton.addEventListener("click", compressContext);
async function setWriteMode(mode){
  elements.writeModeSelect.value = mode;
  elements.topWriteMode.value = mode;
  try{ const res = parseResult(await asyncBridge.SetWritePreference(mode)); if(res.ok) showToast(mode === "formula" ? "已切换为：公式优先" : mode === "value" ? "已切换为：直接写值" : "已切换为：自动（按操作默认）"); else showError(res.error || "保存失败。"); }catch(error){ showError(error.message || String(error)); }
}
elements.writeModeSelect.addEventListener("change", () => { setWriteMode(elements.writeModeSelect.value); });
elements.topWriteMode.addEventListener("change", () => { setWriteMode(elements.topWriteMode.value); });

elements.maxToolCalls.addEventListener("change", async () => {
  const value = clampInt(elements.maxToolCalls.value, 1, 50, 15);
  elements.maxToolCalls.value = value;
  try{ const res = parseResult(await asyncBridge.SetMaxToolCalls(value)); if(res.ok) showToast("单轮工具调用上限已设为 " + value + " 次"); else showError(res.error || "保存失败。"); }catch(error){ showError(error.message || String(error)); }
});

elements.presetSelect.addEventListener("change", () => {
  const presets = {
    deepseek: { base:"https://api.deepseek.com/v1", model:"deepseek-chat" },
    qwen: { base:"https://dashscope.aliyuncs.com/compatible-mode/v1", model:"qwen-plus" },
    openai: { base:"https://api.openai.com/v1", model:"gpt-4o-mini" },
    kimi: { base:"https://api.moonshot.cn/v1", model:"moonshot-v1-8k" }
  };
  const p = presets[elements.presetSelect.value];
  if(p){ elements.baseUrl.value = p.base; elements.modelName.value = p.model; if(elements.presetSelect.value !== "custom"){ const m = elements.modelName.value; elements.presetSelect.value = ""; elements.modelName.value = m; } }
});

elements.addRefButton.addEventListener("click", () => { const address = elements.refAddress.value.trim(); if(!address){ showError("请输入引用地址。"); return; } addReference(address); });
elements.refAddress.addEventListener("keydown", e => { if(e.key === "Enter"){ e.preventDefault(); elements.addRefButton.click(); } });
elements.addSelectionRef.addEventListener("click", async () => { try{ const c = parseResult(await asyncBridge.GetContext()); if(!c.ok || !c.context || !c.context.Selection) throw new Error("当前没有选中区域。"); const sel = c.context.Selection; const address = (c.context.Worksheet ? c.context.Worksheet.Name + "!" : "") + sel.Address; await addReference(address); }catch(error){ showError(error.message || String(error)); } });
async function insertSelectionToPrompt(){ try{ const c = parseResult(await asyncBridge.GetContext()); if(!c.ok || !c.context || !c.context.Selection) throw new Error("当前没有选中区域。"); const sel = c.context.Selection; const sheet = c.context.Worksheet ? c.context.Worksheet.Name + "!" : ""; const areas = String(sel.Address || "").split(",").map(a => a.trim()).filter(Boolean).map(a => "[" + sheet + a + "]"); if(areas.length === 0) throw new Error("当前没有选中区域。"); const fragment = areas.join(" "); const ta = elements.prompt; ta.value = (ta.value.trim() ? ta.value.replace(/\s*$/, "") + " " : "") + fragment; const pos = ta.value.length; ta.focus(); ta.setSelectionRange(pos, pos); showToast("已插入选区：" + areas.join("，") + "（可继续选择其他区域再插入）"); }catch(error){ showError(error.message || String(error)); } }
elements.insertSelectionBtn.addEventListener("click", insertSelectionToPrompt);
async function addReference(address){ try{ const result = parseResult(await asyncBridge.AddReference(address)); if(!result.ok) throw new Error(result.error || "添加失败。"); renderReferences(result.references); elements.refAddress.value = ""; showToast("已添加引用"); }catch(error){ showError(error.message || String(error)); } }

elements.profileSelect.addEventListener("change", () => { creatingNewProfile = false; const id = elements.profileSelect.value; if(id){ const current = activeProfiles.find(p => p.id === id); if(current) loadProfileForm(current); } });

let activeProfiles = [];
let creatingNewProfile = false;
elements.saveProfileButton.addEventListener("click", async () => { setBusy(true); try{
  const result = parseResult(await asyncBridge.SaveProfile(elements.profileName.value.trim(), elements.baseUrl.value.trim(), elements.apiKey.value, elements.modelName.value.trim(), elements.protocolMode.value, elements.temperature.value, creatingNewProfile));
  if(!result.ok) throw new Error(result.error || "保存失败。"); creatingNewProfile = false; activeProfiles = result.profiles || []; renderProfiles(activeProfiles, result.active_profile_id, true); showToast("方案已保存"); }catch(error){ showError(error.message || String(error)); }finally{ setBusy(false); } });

elements.testButton.addEventListener("click", async () => { setBusy(true); elements.connectionState.textContent = "测试连接..."; try{
  const saved = parseResult(bridge.SaveProfile(elements.profileName.value.trim(), elements.baseUrl.value.trim(), elements.apiKey.value, elements.modelName.value.trim(), elements.protocolMode.value, elements.temperature.value, creatingNewProfile));
  if(!saved.ok) throw new Error(saved.error || "保存配置失败。"); creatingNewProfile = false; activeProfiles = saved.profiles || []; renderProfiles(activeProfiles, saved.active_profile_id, false);
  const tested = parseResult(await asyncBridge.TestConnection()); if(!tested.ok) throw new Error(tested.error || "连接失败。"); elements.connectionState.textContent = "已连接"; showToast(tested.message || "连接成功"); }catch(error){ elements.connectionState.textContent = "连接失败"; showError(error.message || String(error)); }finally{ setBusy(false); } });

elements.newProfileButton.addEventListener("click", () => { creatingNewProfile = true; elements.profileName.placeholder = "在此输入新方案名称并保存"; elements.profileName.value = ""; elements.baseUrl.value = ""; elements.apiKey.value = ""; elements.apiKey.placeholder = "sk-..."; elements.modelName.value = ""; elements.protocolMode.value = "Auto"; showToast("已创建空白方案，填写后保存"); });

elements.deleteProfileButton.addEventListener("click", async () => { if(!activeProfiles || activeProfiles.length <= 1){ showError("至少保留一个方案。"); return; } const id = elements.profileSelect.value; if(!id) return; try{ const result = parseResult(await asyncBridge.DeleteProfile(id)); if(!result.ok) throw new Error(result.error || "删除失败。"); activeProfiles = result.profiles || []; renderProfiles(activeProfiles, result.active_profile_id, true); showToast("已删除方案"); }catch(error){ showError(error.message || String(error)); } });

// nav
document.querySelectorAll(".nav button").forEach(btn => btn.addEventListener("click", () => {
  document.querySelectorAll(".nav button").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  btn.classList.add("active"); document.getElementById(btn.dataset.view).classList.add("active");
  if(btn.dataset.view === "view-history") loadHistory();
}));

elements.approveButton.addEventListener("click", () => { if(activeConfirmation){ const c = activeConfirmation; activeConfirmation = null; asyncBridge.RespondConfirmation(c.call_id, true); } elements.confirmOverlay.classList.remove("open"); });
elements.rejectButton.addEventListener("click", closeConfirm);
elements.copyCodeButton.addEventListener("click", () => { const code = elements.confirmCode.textContent || ""; if(!code) return; const ta = document.createElement("textarea"); ta.value = code; ta.style.position = "absolute"; ta.style.left = "-9999px"; document.body.appendChild(ta); ta.select(); try{ document.execCommand("copy"); elements.copyCodeButton.textContent = "已复制"; }catch(_){ } document.body.removeChild(ta); setTimeout(() => elements.copyCodeButton.textContent = "复制", 1600); });

function splitLines(value){ return (value||"").split(/\r?\n/).map(s=>s.trim()).filter(Boolean); }
function joinLines(arr){ return (arr||[]).join("\n"); }
function clampInt(v, min, max, d){ const n = parseInt(v,10); if(isNaN(n)) return d; return Math.max(min, Math.min(max, n)); }

async function refreshControlledExecution(){
  try{
    const r = parseResult(await asyncBridge.GetControlledExecution());
    if(!r.ok){ showError(r.error||"读取受控执行配置失败。"); return; }
    const c = r.config || {};
    elements.controlledEnabled.checked = !!c.enabled;
    elements.controlledTimeout.value = c.timeout_seconds ?? 30;
    elements.controlledOutputKb.value = c.output_limit_kb ?? 64;
    elements.controlledPython.value = c.python_path || "python";
    elements.cliAllow.value = joinLines(c.cli_allowlist); elements.cliDeny.value = joinLines(c.cli_denylist);
    elements.exeAllow.value = joinLines(c.exe_allowlist); elements.exeDeny.value = joinLines(c.exe_denylist);
    elements.scriptAllow.value = joinLines(c.script_allowlist); elements.scriptDeny.value = joinLines(c.script_denylist);
    elements.controlledPath.textContent = "配置文件：" + (r.file_path || "");
  }catch(error){ showError(error.message || String(error)); }
}

async function saveControlledExecution(){
  const payload = {
    enabled: elements.controlledEnabled.checked,
    timeout_seconds: clampInt(elements.controlledTimeout.value, 1, 3600, 30),
    output_limit_kb: clampInt(elements.controlledOutputKb.value, 1, 10240, 64),
    python_path: elements.controlledPython.value.trim() || "python",
    cli_allowlist: splitLines(elements.cliAllow.value),
    cli_denylist: splitLines(elements.cliDeny.value),
    exe_allowlist: splitLines(elements.exeAllow.value),
    exe_denylist: splitLines(elements.exeDeny.value),
    script_allowlist: splitLines(elements.scriptAllow.value),
    script_denylist: splitLines(elements.scriptDeny.value)
  };
  try{
    const r = parseResult(await asyncBridge.SetControlledExecution(JSON.stringify(payload)));
    if(!r.ok) throw new Error(r.error || "保存失败。");
    showToast("受控执行配置已保存并生效。");
    refreshControlledExecution();
  }catch(error){ showError(error.message || String(error)); }
}

elements.saveControlledButton.addEventListener("click", saveControlledExecution);
elements.reloadControlledButton.addEventListener("click", refreshControlledExecution);

async function refreshAbout(){
  try{
    const r = parseResult(await asyncBridge.GetVersion());
    if(!r.ok) throw new Error(r.error || "读取版本失败。");
    const sc = r.self_check || {};
    const kv = elements.aboutInfo.querySelectorAll("b");
    kv[0].textContent = r.informational + "（" + r.version + "）";
    kv[1].textContent = sc.excel_connected ? "已连接" : "未连接";
    kv[2].textContent = sc.webview2_loader ? "存在" : "缺失";
    kv[3].textContent = sc.skill_files ? "已加载" : "无";
    kv[1].style.color = sc.excel_connected ? "#047857" : "#b42318";
    kv[2].style.color = sc.webview2_loader ? "#047857" : "#b42318";
    kv[3].style.color = sc.skill_files ? "#047857" : "#b45309";
  }catch(error){ }
}

try{
  const initialized = parseResult(bridge.Initialize());
  if(!initialized.ok) throw new Error(initialized.error || "初始化失败。");
  renderContext(initialized.context); renderHistory(initialized.history);
  activeProfiles = initialized.profiles || []; renderProfiles(activeProfiles, initialized.active_profile_id, true);
  renderReferences(initialized.references || []);
  renderTools(initialized.tools || []);
  renderSkills(initialized.skills || []);
  renderSessionSelect(initialized.sessions || [], initialized.session?.id || initialized.active_session_id);
  renderSession(initialized.session);
  loadHistory(); // 覆盖 Excel 启动初期 ActiveWorkbook 尚未就绪导致历史为空的情况，重渲染出回滚按钮
  const writeMode = initialized.write_preference || "auto";
  elements.writeModeSelect.value = writeMode;
  elements.topWriteMode.value = writeMode;
  elements.maxToolCalls.value = clampInt(initialized.max_tool_calls_per_turn, 1, 50, 15);
  refreshControlledExecution();
  refreshAbout();
}catch(error){ showError(error.message || String(error)); }
contextTimer = setInterval(refreshContext, 2500);
