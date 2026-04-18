// v3.1.5 版本说明
// 已包含此前升级到 v3.1.0 时要求保留的内容：
// 1. 保留 Outlook 自动签名，正文从邮件正文第一行插入，并与签名之间固定保留 1 个空行。
// 2. 任务日期输入同时支持 YYYY-MM-DD 与 YYYY/MM/DD，写入 Outlook 时统一输出为 YYYY/MM/DD。
// 3. 批量任务输入框示例、日期校验、报错提示已同步支持上述两种日期格式。
//
// 已包含此前升级到 v3.1.1 时要求保留的内容：
// 1. 红 / 黄 / 绿状态提示与固定 RPA 状态节点。
// 2. 黄色提示改为轻量复用 toast，避免长时间运行时频繁堆叠弹层。
//
// 已包含此前升级到 v3.1.2 时要求保留的内容：
// 1. 最小化按钮改为明确的“最小化 / 展开”文字切换，并新增面板展开态 / 最小化态标识，方便人工与 RPA 识别。
// 2. 模板设置新增“单模板模式 / 随机抽取模式”，并按“总体变化尽量多、与上一封差异尽量大”的原则预排固定序列。
// 3. 当前正在运行的批次会锁定启动时模板配置；中途修改模板仅从下一批任务开始生效。
// 4. 日志新增当前邮件所使用的模板组合编号，例如 S2-O4-B1-C3。
//
// 本次相对 v3.1.2 新增：
// 1. 原“变体库模式”正式改名为“随机抽取模式”，主页尽量不变，相关编辑能力全部收纳到设置界面。
// 2. 设置界面中的随机抽取模式改为摘要视图；详细编辑改为大尺寸居中编辑层，避免 Subject / Opening / Body / CTA 挤在一起看不清。
// 3. 随机抽取库编辑器改为逐条独立行编辑，固定显示 S1/S2...、O1/O2...、B1/B2...、C1/C2...，并支持增删行与保存/取消。
// 4. 随机抽取项底层改为数组持久化，支持多行正文类条目，避免旧版用单个 textarea 时分隔不清与内容混行的问题。
//
// 本次相对 v3.1.3 新增：
// 1. 正文插入前会自动将首段正文处理为正式信件格式：称呼行后保留 1 个空行，并只对正文首段首行添加缩进。
// 2. 保持 Ox + 两个换行 + Bx + 两个换行 + Cx 的主体结构不变；仅修正首段缩进与称呼后空行，不影响后续 B / C 段落。
//
// 本次相对 v3.1.4 新增：
// 1. 两个联系人之间的等待提示改为显示 3 秒后再淡出，便于人工与 RPA 观察。
// 2. 两个联系人之间的处理间隔新增下限保护，最小不得低于 5 秒；即使设置里填得更小，运行时也会自动按 5 秒执行。

(function () {
  'use strict';

  const PANEL_ID = 'outlook-rpa-schedule-panel-v3-1-5';
  const STYLE_ID = 'outlook-rpa-schedule-panel-style-v3-1-5';
  const STORAGE_KEY = 'outlook_rpa_schedule_panel_v3_1_5_state';
  const MAX_LOG_LINES = 500;
  const META = { version: '3.1.5', author: '陈柏安' };
  const WAIT_TOAST_MS = 3000;
  const DEFAULT_PLACEHOLDER = '[Name/Team]';
  const DEFAULT_SUBJECT = 'Paper Packaging Supplier Introduction – Zaya Packaging';
  const DEFAULT_BODY = 'Hello [Name/Team],\n\nWe’re Zaya Packaging, a direct paper packaging manufacturer with production in China and Vietnam.\n\nWe supply folding cartons, rigid gift boxes, collapsible rigid boxes, paper bags, hang tags, and labels for beauty, consumer goods, and retail brands.\n\nMay I ask what the best route would be for us to be considered as a packaging supplier to your company? If there is a supplier registration portal or a packaging-related procurement/sourcing contact we should approach, we’d greatly appreciate your direction.';
  const DEFAULT_TEMPLATE_MODE = 'variants';
  const DEFAULT_VARIANT_SUBJECTS = [
    'Paper Packaging Supplier Introduction – Zaya Packaging',
    'Secondary Paper Packaging Supplier Inquiry',
    'Packaging Supplier Registration Inquiry',
    'China + Vietnam Paper Packaging Support',
  ];
  const DEFAULT_VARIANT_OPENINGS = [
    'Hello [Name/Team],\n\nWe’re Zaya Packaging, a direct paper packaging manufacturer with production in China and Vietnam.',
    'Hello [Name/Team],\n\nWe’re Zaya Packaging, a direct manufacturer of premium paper packaging with production in China and Vietnam.',
    'Hello [Name/Team],\n\nWe’re Zaya Packaging, a paper packaging manufacturer supporting beauty, consumer goods, and retail brands from our China and Vietnam production bases.',
    'Hello [Name/Team],\n\nWe’re Zaya Packaging, a direct paper packaging supplier with production support in both China and Vietnam.',
  ];
  const DEFAULT_VARIANT_BODIES = [
    'We supply folding cartons, rigid gift boxes, collapsible rigid boxes, paper bags, hang tags, and labels for beauty, consumer goods, and retail brands.',
    'Our core paper packaging range includes folding cartons, rigid boxes, collapsible rigid boxes, paper bags, hang tags, and labels for secondary packaging programs.',
    'We support packaging programs with folding cartons, gift boxes, paper bags, hang tags, and labels, with flexible production options depending on project needs.',
    'With production support from both China and Vietnam, we can offer flexible options around lead time, sourcing priorities, and program requirements across a range of paper packaging formats.',
  ];
  const DEFAULT_VARIANT_CTAS = [
    'May I ask what the best route would be for us to be considered as a packaging supplier to your company? If there is a supplier registration portal or a packaging-related procurement/sourcing contact we should approach, we’d greatly appreciate your direction.',
    'Could you please advise whether your company has a preferred process for reviewing new packaging suppliers? If there is a registration portal or the right contact person, we’d appreciate your guidance.',
    'If your team is open to reviewing new paper packaging suppliers, we’d be grateful if you could point us to the appropriate procurement/sourcing contact or vendor onboarding channel.',
    'We’re reaching out to ask about the best way to introduce our company to the right team. If there is a supplier registration path or a packaging-related contact we should start with, we’d appreciate your advice.',
  ];


  const VARIANT_GROUPS = [
    { key: 'subjects', label: 'Subject', prefix: 'S', defaults: DEFAULT_VARIANT_SUBJECTS },
    { key: 'openings', label: 'Opening', prefix: 'O', defaults: DEFAULT_VARIANT_OPENINGS },
    { key: 'bodies', label: 'Body', prefix: 'B', defaults: DEFAULT_VARIANT_BODIES },
    { key: 'ctas', label: 'CTA', prefix: 'C', defaults: DEFAULT_VARIANT_CTAS },
  ];

  const LANG = {
    newMail: ['新邮件', 'New mail'],
    send: ['发送', 'Send'],
    scheduleSend: ['计划发送', 'Schedule send'],
    customTime: ['自定义时间', 'Custom time'],
    cancel: ['取消', 'Cancel'],
    subjectLabel: ['主题', 'Subject'],
    subjectPlaceholder: ['添加主题', 'Add a subject'],
    recipientLabel: ['收件人', 'To'],
    bodyLabel: ['邮件正文', 'Message body'],
    moreSendOptions: ['更多发送选项', 'More send options', '展开以查看更多发送选项'],
    loading: ['加载', 'Loading'],
  };

  const STATE = {
    running: false,
    paused: false,
    stopRequested: false,
    currentIndex: 0,
    tasks: [],
    els: {},
    demoBox: null,
    modal: null,
    toast: null,
    toastTimer: 0,
    rpaStatus: null,
    variantCursor: 0,
    variantSignature: '',
    successCount: 0,
    variantDraft: null,
  };

  if (document.getElementById(PANEL_ID)) return;

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  const low = (s) => norm(s).toLowerCase();
  const esc = (s) => String(s || '').replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));

  function now() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  function visible(el) {
    if (!el || !document.contains(el)) return false;
    const st = getComputedStyle(el);
    const r = el.getBoundingClientRect();
    return st.display !== 'none' && st.visibility !== 'hidden' && Number(st.opacity || '1') !== 0 && r.width > 0 && r.height > 0;
  }

  function text(el) {
    return norm(el?.innerText || el?.textContent || el?.value || el?.getAttribute?.('aria-label') || el?.getAttribute?.('title') || '');
  }

  function eqAny(v, arr) {
    const t = low(v);
    return arr.some((x) => t === low(x));
  }

  function includesAny(v, arr) {
    const t = low(v);
    return arr.some((x) => t.includes(low(x)));
  }

  function qsaVisible(sel, root = document) {
    return Array.from(root.querySelectorAll(sel)).filter(visible);
  }

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function fireInput(el) {
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function fireKey(el, key, code) {
    ['keydown', 'keyup'].forEach((type) => {
      el.dispatchEvent(new KeyboardEvent(type, { key, code: code || key, bubbles: true, cancelable: true }));
    });
  }

  function setInputValue(el, value) {
    const proto = Object.getPrototypeOf(el);
    const desc = Object.getOwnPropertyDescriptor(proto, 'value');
    if (desc && desc.set) desc.set.call(el, value);
    else el.value = value;
    fireInput(el);
  }

  function prependComposeBody(el, content) {
    const text = String(content || '');
    if ('value' in el) {
      const existing = String(el.value || '');
      const merged = existing ? `${text}\n\n${existing}` : text;
      setInputValue(el, merged);
      return;
    }
    if (!el || !el.isContentEditable) {
      el.textContent = text;
      fireInput(el);
      return;
    }

    const lines = text.split(/\r?\n/);
    const frag = document.createDocumentFragment();

    lines.forEach((line) => {
      const div = document.createElement('div');
      if (line) div.textContent = line;
      else div.appendChild(document.createElement('br'));
      frag.appendChild(div);
    });

    const blank = document.createElement('div');
    blank.appendChild(document.createElement('br'));
    frag.appendChild(blank);

    const range = document.createRange();
    range.setStart(el, 0);
    range.collapse(true);
    range.insertNode(frag);
    fireInput(el);
  }

  function typeSemicolon(el) {
    el.focus();
    if ('value' in el) {
      const current = String(el.value || '');
      setInputValue(el, current + ';');
    } else if (el.isContentEditable) {
      document.execCommand('insertText', false, ';');
      fireInput(el);
    }
    fireKey(el, ';', 'Semicolon');
    fireKey(el, 'Enter', 'Enter');
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('blur', { bubbles: true }));
  }

  function click(el) {
    if (!el) throw new Error('目标元素不存在，无法点击。');
    el.scrollIntoView({ block: 'center', inline: 'center' });
    el.focus && el.focus();
    if (typeof el.click === 'function') el.click();
    else el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
  }

  function setRpaStatus(status, detail = '', errorCode = '') {
    if (!STATE.rpaStatus) return;
    [STATE.rpaStatus, STATE.els.miniRpaStatus].filter(Boolean).forEach((node) => {
      node.textContent = status;
      node.dataset.rpaStatus = status;
      node.dataset.detail = detail || '';
      node.dataset.errorCode = errorCode || '';
      node.setAttribute('aria-label', `RPA status: ${status}`);
    });
  }

  function showPageModal(type, title, message, options = {}) {
    if (!STATE.modal) return;
    hideStatusToast();
    const symbolMap = { success: '✅', error: '❌', warn: '⚠' };
    STATE.modal.dataset.modalType = type || 'warn';
    STATE.modal.dataset.rpaStatus = options.rpaStatus || type || 'warn';
    STATE.modal.dataset.errorCode = options.errorCode || '';
    STATE.modal.querySelector('.orp-modal-symbol').textContent = symbolMap[type] || '⚠';
    STATE.modal.querySelector('.orp-modal-title').textContent = title;
    STATE.modal.querySelector('.orp-modal-msg').textContent = message;
    STATE.modal.classList.add('open');
  }

  function hidePageModal() {
    if (!STATE.modal) return;
    STATE.modal.classList.remove('open');
    STATE.modal.dataset.modalType = '';
    STATE.modal.dataset.rpaStatus = '';
    STATE.modal.dataset.errorCode = '';
  }

  function showStatusToast(type, message, durationMs = WAIT_TOAST_MS) {
    if (!STATE.toast) return;
    window.clearTimeout(STATE.toastTimer);
    STATE.toast.dataset.toastType = type || 'warn';
    STATE.toast.dataset.rpaStatus = type || 'warn';
    STATE.toast.querySelector('.orp-toast-text').textContent = message;
    STATE.toast.classList.add('open');
    if (durationMs > 0) {
      STATE.toastTimer = window.setTimeout(() => {
        hideStatusToast();
      }, durationMs);
    }
  }

  function hideStatusToast() {
    if (!STATE.toast) return;
    window.clearTimeout(STATE.toastTimer);
    STATE.toast.classList.remove('open');
    STATE.toast.dataset.toastType = '';
    STATE.toast.dataset.rpaStatus = '';
  }

  function errorCodeFromMessage(msg) {
    const m = String(msg || '');
    if (/发送下拉按钮/.test(m)) return 'ERR_SEND_MENU_NOT_FOUND';
    if (/计划发送.*菜单项/.test(m)) return 'ERR_SCHEDULE_MENU_NOT_FOUND';
    if (/自定义时间.*按钮/.test(m)) return 'ERR_CUSTOM_TIME_BUTTON_NOT_FOUND';
    if (/写信界面就绪/.test(m)) return 'ERR_COMPOSE_READY_TIMEOUT';
    if (/自定义时间弹窗就绪/.test(m)) return 'ERR_CUSTOM_DIALOG_READY_TIMEOUT';
    if (/日期格式错误/.test(m)) return 'ERR_TASK_DATE_FORMAT';
    if (/时间格式错误/.test(m)) return 'ERR_TASK_TIME_FORMAT';
    if (/邮箱格式错误|没有有效邮箱/.test(m)) return 'ERR_TASK_EMAIL_INVALID';
    if (/主题模板不能为空/.test(m)) return 'ERR_SUBJECT_TEMPLATE_EMPTY';
    if (/正文模板不能为空/.test(m)) return 'ERR_BODY_TEMPLATE_EMPTY';
    if (/姓名占位符不能为空/.test(m)) return 'ERR_NAME_PLACEHOLDER_EMPTY';
    if (/随机抽取库 Subject/.test(m)) return 'ERR_VARIANT_SUBJECT_EMPTY';
    if (/随机抽取库 Opening/.test(m)) return 'ERR_VARIANT_OPENING_EMPTY';
    if (/随机抽取库 Body/.test(m)) return 'ERR_VARIANT_BODY_EMPTY';
    if (/随机抽取库 CTA/.test(m)) return 'ERR_VARIANT_CTA_EMPTY';
    if (/草稿详情编辑态|禁止启动/.test(m)) return 'ERR_UNSAFE_START_STATE';
    if (/当前页面不是可安全启动/.test(m)) return 'ERR_UNSAFE_PAGE';
    if (/未找到收件人输入框/.test(m)) return 'ERR_RECIPIENT_INPUT_NOT_FOUND';
    if (/已停止/.test(m)) return 'ERR_STOPPED_BY_USER';
    return 'ERR_UNKNOWN';
  }

  function normalizeTaskDate(date) {
    const raw = String(date || '').trim();
    const m = raw.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})$/);
    if (!m) return '';
    return `${m[1]}/${m[2]}/${m[3]}`;
  }

  function logHtml(html, type) {
    const box = STATE.els.log;
    if (!box) return;
    const line = document.createElement('div');
    line.className = 'orp-log-line' + (type ? ` ${type}` : '');
    line.innerHTML = `<span class="orp-log-time">[${now()}]</span> ${html}`;
    box.appendChild(line);
    while (box.childNodes.length > MAX_LOG_LINES) box.removeChild(box.firstChild);
    box.scrollTop = box.scrollHeight;
  }

  function logInfo(msg) {
    logHtml(esc(msg));
  }

  function logStep(no, msg, extraHtml = '') {
    logHtml(`<span class="orp-log-step">Step.${no}:</span> ${esc(msg)}${extraHtml ? ` ${extraHtml}` : ''}`);
  }

  function logTaskStart(task, templateCode = '') {
    const html = [
      `<span class="orp-log-step">Task:</span>`,
      `<span class="orp-log-email">${esc(task.emails.join('; '))}</span>`,
      `| <span class="orp-log-name">${esc(task.name)}</span>`,
      `| <span class="orp-log-time-strong">${esc(task.date)} ${esc(task.time)}</span>`,
      templateCode ? `| <span class="orp-log-step">${esc(templateCode)}</span>` : '',
    ].join(' ');
    logHtml(html);
  }

  function logWarn(msg) {
    logHtml(esc(msg), 'warn');
  }

  function logError(msg) {
    logHtml(esc(msg), 'error');
  }

  function logSuccess(msg) {
    logHtml(esc(msg), 'success');
  }

  function setStatus(s) {
    STATE.els.status.textContent = s;
    if (STATE.els.miniStatus) STATE.els.miniStatus.textContent = s;
  }

  function setProgress(cur, total) {
    STATE.els.progress.textContent = `${cur} / ${total}`;
    if (STATE.els.miniProgress) STATE.els.miniProgress.textContent = `${cur} / ${total}`;
  }

  function setPanelMinimized(minimized) {
    const root = STATE.els.root;
    if (!root) return;
    root.classList.toggle('minimized', !!minimized);
    root.dataset.panelState = minimized ? 'minimized' : 'expanded';
    if (STATE.els.minToggle) {
      STATE.els.minToggle.textContent = minimized ? '展开' : '最小化';
      STATE.els.minToggle.title = minimized ? '展开' : '最小化';
      STATE.els.minToggle.setAttribute('aria-label', minimized ? '展开' : '最小化');
    }
  }

  function nonEmptyLines(raw) {
    return String(raw || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  }

  function cleanVariantItems(items) {
    return (Array.isArray(items) ? items : []).map((x) => String(x == null ? '' : x).replace(/\r\n/g, '\n').trim()).filter(Boolean);
  }

  function parseVariantItems(raw, fallback = []) {
    const source = raw == null ? '' : String(raw).trim();
    if (!source) return cleanVariantItems(fallback);
    if (source.startsWith('[')) {
      try {
        const parsed = JSON.parse(source);
        if (Array.isArray(parsed)) return cleanVariantItems(parsed);
      } catch (e) {
        // 兼容旧版逐行文本格式。
      }
    }
    return cleanVariantItems(String(raw || '').split(/\r?\n/));
  }

  function serializeVariantItems(items) {
    return JSON.stringify(cleanVariantItems(items));
  }

  function setStoredVariantItems(el, items) {
    if (el) el.value = serializeVariantItems(items);
  }

  function getStoredVariantItems(el, fallback = []) {
    return parseVariantItems(el ? el.value : '', fallback);
  }

  function joinVariantDefaults(lines) {
    return serializeVariantItems(lines || []);
  }

  function ensureLeadingIndent(line) {
    const raw = String(line || '');
    if (!raw.trim()) return raw;
    if (/^(?:\u3000\u3000|\t| {2,})/.test(raw)) return raw;
    return `\u3000\u3000${raw.replace(/^\s+/, '')}`;
  }

  function normalizeFormalBody(body) {
    const lines = String(body || '').replace(/\r\n/g, '\n').split('\n');
    let greetingIdx = -1;
    for (let i = 0; i < lines.length; i += 1) {
      if (lines[i].trim()) {
        greetingIdx = i;
        break;
      }
    }
    if (greetingIdx < 0) return '';
    let firstParagraphIdx = -1;
    for (let i = greetingIdx + 1; i < lines.length; i += 1) {
      if (lines[i].trim()) {
        firstParagraphIdx = i;
        break;
      }
    }
    if (firstParagraphIdx < 0) return lines.join('\n').trim();
    const head = lines.slice(0, greetingIdx + 1).map((x) => x.replace(/\s+$/, ''));
    const tail = lines.slice(firstParagraphIdx);
    tail[0] = ensureLeadingIndent(tail[0]);
    return head.concat([''], tail).join('\n');
  }

  function getVariantPartsFromFields() {
    return {
      namePlaceholder: STATE.els.namePlaceholder?.value || DEFAULT_PLACEHOLDER,
      subjects: getStoredVariantItems(STATE.els.variantSubjects, DEFAULT_VARIANT_SUBJECTS),
      openings: getStoredVariantItems(STATE.els.variantOpenings, DEFAULT_VARIANT_OPENINGS),
      bodies: getStoredVariantItems(STATE.els.variantBodies, DEFAULT_VARIANT_BODIES),
      ctas: getStoredVariantItems(STATE.els.variantCtas, DEFAULT_VARIANT_CTAS),
    };
  }

  function countVariantCombos(parts) {
    return (parts.subjects.length || 0) * (parts.openings.length || 0) * (parts.bodies.length || 0) * (parts.ctas.length || 0);
  }

  function buildVariantSummaryState() {
    const parts = getVariantPartsFromFields();
    const totalCombos = countVariantCombos(parts);
    const sequence = totalCombos ? buildVariantSequence(parts) : [];
    const signature = buildVariantSignature(parts);
    const cursor = STATE.variantSignature === signature ? STATE.variantCursor : 0;
    const nextCode = sequence.length ? sequence[cursor % sequence.length].code : '--';
    return { parts, totalCombos, sequence, signature, cursor, nextCode };
  }

  function refreshVariantSummary() {
    if (!STATE.els.variantCountSubjects) return;
    const summary = buildVariantSummaryState();
    STATE.els.variantCountSubjects.textContent = String(summary.parts.subjects.length || 0);
    STATE.els.variantCountOpenings.textContent = String(summary.parts.openings.length || 0);
    STATE.els.variantCountBodies.textContent = String(summary.parts.bodies.length || 0);
    STATE.els.variantCountCtas.textContent = String(summary.parts.ctas.length || 0);
    STATE.els.variantNextCode.textContent = summary.nextCode;
    STATE.els.variantTotalCombos.textContent = String(summary.totalCombos || 0);
  }

  function buildVariantBody(opening, body, cta) {
    return `${opening}\n\n${body}\n\n${cta}`;
  }

  function buildVariantSignature(parts) {
    return JSON.stringify({
      mode: 'variants',
      namePlaceholder: parts.namePlaceholder,
      subjects: parts.subjects,
      openings: parts.openings,
      bodies: parts.bodies,
      ctas: parts.ctas,
    });
  }

  function lexComboKey(combo) {
    return [combo.s, combo.o, combo.b, combo.c].map((n) => String(n).padStart(4, '0')).join('-');
  }

  function comboImmediateDiff(a, b) {
    if (!a || !b) return 0;
    let score = 0;
    if (a.s !== b.s) score += 12;
    if (a.o !== b.o) score += 9;
    if (a.b !== b.b) score += 9;
    if (a.c !== b.c) score += 8;
    return score;
  }

  function comboRecentSpreadScore(combo, recent) {
    let score = 0;
    recent.forEach((prev, idx) => {
      const factor = recent.length - idx;
      score += comboImmediateDiff(combo, prev) * factor;
      if (combo.s === prev.s) score -= 16 * factor;
      if (combo.o === prev.o) score -= 10 * factor;
      if (combo.b === prev.b) score -= 10 * factor;
      if (combo.c === prev.c) score -= 8 * factor;
    });
    return score;
  }

  function buildVariantSequence(parts) {
    const combos = [];
    parts.subjects.forEach((subject, sIdx) => {
      parts.openings.forEach((opening, oIdx) => {
        parts.bodies.forEach((body, bIdx) => {
          parts.ctas.forEach((cta, cIdx) => {
            combos.push({
              s: sIdx + 1,
              o: oIdx + 1,
              b: bIdx + 1,
              c: cIdx + 1,
              subject,
              body: buildVariantBody(opening, body, cta),
              code: `S${sIdx + 1}-O${oIdx + 1}-B${bIdx + 1}-C${cIdx + 1}`,
            });
          });
        });
      });
    });
    if (!combos.length) return [];
    const remaining = combos.slice().sort((a, b) => lexComboKey(a).localeCompare(lexComboKey(b)));
    const seq = [remaining.shift()];
    while (remaining.length) {
      const prev = seq[seq.length - 1];
      const recent = seq.slice(-4);
      let bestIdx = 0;
      let bestScore = -Infinity;
      let bestKey = '';
      for (let i = 0; i < remaining.length; i += 1) {
        const cand = remaining[i];
        const immediate = comboImmediateDiff(cand, prev);
        const recentSpread = comboRecentSpreadScore(cand, recent);
        const numericSpread = Math.abs(cand.s - prev.s) + Math.abs(cand.o - prev.o) + Math.abs(cand.b - prev.b) + Math.abs(cand.c - prev.c);
        const score = immediate * 1000 + recentSpread * 10 + numericSpread;
        const key = lexComboKey(cand);
        if (score > bestScore || (score === bestScore && key < bestKey)) {
          bestScore = score;
          bestIdx = i;
          bestKey = key;
        }
      }
      seq.push(remaining.splice(bestIdx, 1)[0]);
    }
    return seq;
  }

  function getTemplatePartsFromConfig(cfg) {
    return {
      namePlaceholder: cfg.namePlaceholder,
      subjects: parseVariantItems(cfg.variantSubjectsRaw, DEFAULT_VARIANT_SUBJECTS),
      openings: parseVariantItems(cfg.variantOpeningsRaw, DEFAULT_VARIANT_OPENINGS),
      bodies: parseVariantItems(cfg.variantBodiesRaw, DEFAULT_VARIANT_BODIES),
      ctas: parseVariantItems(cfg.variantCtasRaw, DEFAULT_VARIANT_CTAS),
    };
  }

  function getResolvedTemplate(cfg, task) {
    if (cfg.templateMode === 'single') {
      return {
        code: 'SINGLE',
        subject: String(cfg.subject || '').split(cfg.namePlaceholder).join(task.name),
        body: normalizeFormalBody(String(cfg.body || '').split(cfg.namePlaceholder).join(task.name)),
      };
    }
    const runtime = cfg.variantRuntime;
    if (!runtime || !runtime.sequence.length) throw new Error('随机抽取库没有可用组合。');
    const combo = runtime.sequence[STATE.variantCursor % runtime.sequence.length];
    return {
      code: combo.code,
      subject: String(combo.subject || '').split(cfg.namePlaceholder).join(task.name),
      body: normalizeFormalBody(String(combo.body || '').split(cfg.namePlaceholder).join(task.name)),
    };
  }

  function persistVariantCursor(cursor, signature) {
    STATE.variantCursor = cursor;
    STATE.variantSignature = signature || '';
    saveState();
  }

  function prepareConfigForBatch(cfg) {
    if (cfg.templateMode !== 'variants') return cfg;
    const parts = getTemplatePartsFromConfig(cfg);
    if (!parts.subjects.length) throw new Error('随机抽取 Subject 至少需要 1 条。');
    if (!parts.openings.length) throw new Error('随机抽取 Opening 至少需要 1 条。');
    if (!parts.bodies.length) throw new Error('随机抽取 Body 至少需要 1 条。');
    if (!parts.ctas.length) throw new Error('随机抽取 CTA 至少需要 1 条。');
    const signature = buildVariantSignature(parts);
    const runtime = {
      parts,
      signature,
      sequence: buildVariantSequence(parts),
    };
    if (!runtime.sequence.length) throw new Error('随机抽取库没有生成可用组合。');
    if (STATE.variantSignature !== signature) {
      persistVariantCursor(0, signature);
    } else {
      STATE.variantSignature = signature;
    }
    cfg.variantRuntime = runtime;
    return cfg;
  }

  function syncTemplateModeUI() {
    if (!STATE.els.templateMode) return;
    const mode = STATE.els.templateMode.value || DEFAULT_TEMPLATE_MODE;
    if (STATE.els.singleTemplateSection) STATE.els.singleTemplateSection.style.display = mode === 'single' ? 'grid' : 'none';
    if (STATE.els.variantTemplateSection) STATE.els.variantTemplateSection.style.display = mode === 'variants' ? 'grid' : 'none';
    refreshVariantSummary();
  }

  function randomBetween(minMs, maxMs) {

    if (maxMs <= minMs) return minMs;
    return Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        tasks: STATE.els.tasks.value,
        stepTimeout: STATE.els.stepTimeout.value,
        confirmDelay: STATE.els.confirmDelay.value,
        nextIntervalMin: STATE.els.nextIntervalMin.value,
        nextIntervalMax: STATE.els.nextIntervalMax.value,
        demoMode: STATE.els.demo.checked,
        demoExtra: STATE.els.demoExtra.value,
        templateMode: STATE.els.templateMode.value,
        subject: STATE.els.subject.value,
        body: STATE.els.body.value,
        namePlaceholder: STATE.els.namePlaceholder.value,
        variantSubjects: STATE.els.variantSubjects.value,
        variantOpenings: STATE.els.variantOpenings.value,
        variantBodies: STATE.els.variantBodies.value,
        variantCtas: STATE.els.variantCtas.value,
        variantCursor: STATE.variantCursor,
        variantSignature: STATE.variantSignature,
        left: STATE.els.root.style.left || '',
        top: STATE.els.root.style.top || '',
        settingsLeft: STATE.els.settings.style.left || '',
        settingsTop: STATE.els.settings.style.top || '',
        panelState: STATE.els.root.dataset.panelState || 'expanded',
      }));
    } catch (e) {
      console.warn('saveState failed:', e);
    }
  }

  function loadState() {
    try {
      const x = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      STATE.els.tasks.value = x.tasks || '';
      STATE.els.stepTimeout.value = x.stepTimeout || '25';
      STATE.els.confirmDelay.value = x.confirmDelay || '0.5';
      STATE.els.nextIntervalMin.value = x.nextIntervalMin || '8';
      STATE.els.nextIntervalMax.value = x.nextIntervalMax || '12';
      STATE.els.demo.checked = !!x.demoMode;
      STATE.els.demoExtra.value = x.demoExtra || '2';
      STATE.els.templateMode.value = x.templateMode || DEFAULT_TEMPLATE_MODE;
      STATE.els.subject.value = x.subject || DEFAULT_SUBJECT;
      STATE.els.body.value = x.body || DEFAULT_BODY;
      STATE.els.namePlaceholder.value = x.namePlaceholder || DEFAULT_PLACEHOLDER;
      setStoredVariantItems(STATE.els.variantSubjects, parseVariantItems(x.variantSubjects, DEFAULT_VARIANT_SUBJECTS));
      setStoredVariantItems(STATE.els.variantOpenings, parseVariantItems(x.variantOpenings, DEFAULT_VARIANT_OPENINGS));
      setStoredVariantItems(STATE.els.variantBodies, parseVariantItems(x.variantBodies, DEFAULT_VARIANT_BODIES));
      setStoredVariantItems(STATE.els.variantCtas, parseVariantItems(x.variantCtas, DEFAULT_VARIANT_CTAS));
      STATE.variantCursor = Number.isFinite(Number(x.variantCursor)) ? Math.max(0, Number(x.variantCursor)) : 0;
      STATE.variantSignature = String(x.variantSignature || '');
      if (x.left) { STATE.els.root.style.left = x.left; STATE.els.root.style.right = 'auto'; }
      if (x.top) { STATE.els.root.style.top = x.top; STATE.els.root.style.bottom = 'auto'; }
      if (x.settingsLeft) { STATE.els.settings.style.left = x.settingsLeft; STATE.els.settings.style.right = 'auto'; }
      if (x.settingsTop) { STATE.els.settings.style.top = x.settingsTop; STATE.els.settings.style.bottom = 'auto'; }
      setPanelMinimized((x.panelState || '') === 'minimized');
      syncTemplateModeUI();
      refreshVariantSummary();
    } catch (e) {
      console.warn(e);
    }
  }

  function injectStyle() {

    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID}{position:fixed;right:24px;bottom:24px;width:430px;max-height:84vh;background:#08101d;color:#f8fafc;border:1px solid #3a4555;box-shadow:0 0 0 1px rgba(130,170,255,.08),8px 8px 0 rgba(0,0,0,.2);z-index:2147483644;font:12px/1.45 Inter,Segoe UI,Arial,sans-serif;display:flex;flex-direction:column;overflow:hidden;border-radius:0}
      #${PANEL_ID}[data-panel-state=minimized] .orp-body{display:none}
      #${PANEL_ID}[data-panel-state=expanded] .orp-minibar{display:none}
      #${PANEL_ID},#${PANEL_ID} *{box-sizing:border-box}
      #${PANEL_ID} button,#${PANEL_ID} input,#${PANEL_ID} textarea{font:inherit;border-radius:0}
      #${PANEL_ID} .orp-header{display:grid;grid-template-columns:1fr auto;align-items:center;gap:8px;padding:8px 10px;background:#050b14;border-bottom:1px solid #3a4555;user-select:none;cursor:move}
      #${PANEL_ID} .orp-title-wrap{display:grid;gap:2px}
      #${PANEL_ID} .orp-title{font-size:14px;font-weight:700;letter-spacing:.02em}
      #${PANEL_ID} .orp-meta{font-size:11px;color:#7f90a7}
      #${PANEL_ID} .orp-header-actions{display:flex;gap:6px}
      #${PANEL_ID} .orp-icon-btn,#${PANEL_ID} .orp-btn{border:1px solid #49576d;background:#0e1624;color:#f8fafc;cursor:pointer;min-height:32px;padding:0 10px;transition:background .12s ease,color .12s ease,border-color .12s ease}
      #${PANEL_ID} .orp-icon-btn{width:30px;padding:0}
      #${PANEL_ID} .orp-toggle-btn{min-width:72px}
      #${PANEL_ID} .orp-icon-btn:hover,#${PANEL_ID} .orp-btn:hover{background:#f8fafc;color:#08101d}
      #${PANEL_ID} .orp-btn.primary{background:#12233b;border-color:#8ab4ff;color:#cfe0ff}
      #${PANEL_ID} .orp-btn.warn{background:#31150e;border-color:#ff925c;color:#ffd0b8}
      #${PANEL_ID} .orp-body{padding:10px;display:grid;gap:10px;background:#08101d;overflow:auto}
      #${PANEL_ID} .orp-status{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:center;color:#c7d2e4}
      #${PANEL_ID} .orp-minibar{padding:8px 10px;border-bottom:1px solid #3a4555;background:#08101d;display:grid;gap:6px}
      #${PANEL_ID} .orp-minibar-row{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:center;color:#c7d2e4}
      #${PANEL_ID} .orp-tag{display:inline-block;padding:2px 8px;border:1px solid #49576d;background:#09111c}
      #${PANEL_ID} .orp-block{display:grid;gap:6px}
      #${PANEL_ID} .orp-label{font-weight:700;color:#d7e0ef;letter-spacing:.03em}
      #${PANEL_ID} .orp-input,#${PANEL_ID} .orp-textarea{width:100%;border:1px solid #49576d;background:#040913;color:#f8fafc;padding:7px 8px;outline:none}
      #${PANEL_ID} .orp-input:focus,#${PANEL_ID} .orp-textarea:focus{border-color:#8ab4ff}
      #${PANEL_ID} .orp-textarea{min-height:140px;resize:vertical}
      #${PANEL_ID} .orp-toolbar{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
      #${PANEL_ID} .orp-hint{color:#8ea0b9;font-size:11px}
      #${PANEL_ID} .orp-log{height:245px;overflow:auto;background:#040913;border:1px solid #49576d;padding:8px;white-space:pre-wrap;word-break:break-word;user-select:text;-webkit-user-select:text;cursor:text}
      #${PANEL_ID} .orp-log *{user-select:text;-webkit-user-select:text;cursor:text}
      #${PANEL_ID} .orp-log-line{padding:0 0 4px;color:#d8e2f2}
      #${PANEL_ID} .orp-log-line.warn{color:#ffd37a}
      #${PANEL_ID} .orp-log-line.error{color:#ff8b8b}
      #${PANEL_ID} .orp-log-line.success{color:#92e6a7}
      #${PANEL_ID} .orp-log-time{color:#7f90a7}
      #${PANEL_ID} .orp-log-step{color:#8ab4ff;font-weight:700}
      #${PANEL_ID} .orp-log-email{color:#7dd3fc;font-weight:700}
      #${PANEL_ID} .orp-log-name{color:#f9f871;font-weight:700}
      #${PANEL_ID} .orp-log-time-strong{color:#86efac;font-weight:700}
      #${PANEL_ID} .orp-settings-mask{position:fixed;inset:0;background:rgba(4,9,19,.6);z-index:2147483645;display:none}
      #${PANEL_ID} .orp-settings-mask.open{display:block}
      #${PANEL_ID} .orp-settings{position:fixed;right:24px;bottom:24px;width:620px;max-height:86vh;overflow:auto;background:#08101d;color:#f8fafc;border:1px solid #3a4555;box-shadow:0 0 0 1px rgba(130,170,255,.08),8px 8px 0 rgba(0,0,0,.2);z-index:2147483646;display:none;border-radius:0}
      #${PANEL_ID} .orp-settings.open{display:block}
      #${PANEL_ID} .orp-settings-header{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;padding:8px 10px;background:#050b14;border-bottom:1px solid #3a4555;cursor:move;user-select:none}
      #${PANEL_ID} .orp-settings-body{padding:10px;display:grid;gap:10px}
      #${PANEL_ID} .orp-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
      #${PANEL_ID} .orp-grid-3{display:grid;grid-template-columns:1.1fr 1fr 1fr;gap:8px}
      #${PANEL_ID} .orp-grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}
      #${PANEL_ID} .orp-stack{display:grid;gap:8px}
      #${PANEL_ID} .orp-section{border:1px solid #3a4555;background:#0a121f;padding:9px;display:grid;gap:8px}
      #${PANEL_ID} .orp-section-title{font-weight:700;color:#cfe0ff;border-bottom:1px solid #293446;padding-bottom:6px}
      #${PANEL_ID} .orp-settings-footer{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:center;padding:0 10px 10px}
      #${PANEL_ID} .orp-checkbox{display:flex;align-items:center;gap:8px}
      #${PANEL_ID} .orp-demo-box{position:fixed;z-index:2147483646;border:2px solid #f9f871;pointer-events:none;display:none;box-shadow:0 0 0 99999px rgba(0,0,0,.18)}
      #${PANEL_ID} .orp-demo-box.open{display:block}
      #${PANEL_ID} .orp-demo-label{position:absolute;left:-2px;top:-24px;background:#f9f871;color:#08101d;padding:2px 6px;font-weight:700;white-space:nowrap;border:1px solid #08101d}
      #${PANEL_ID} .orp-rpa-state{min-width:108px;text-align:center;font-weight:700}
      #${PANEL_ID} .orp-toast{position:fixed;right:24px;top:24px;min-width:280px;max-width:min(420px,92vw);padding:12px 14px;border:1px solid #8f6b08;background:#fff7db;color:#5a4304;box-shadow:8px 8px 0 rgba(0,0,0,.16);z-index:2147483647;opacity:0;visibility:hidden;pointer-events:none;transform:translateY(-6px);transition:opacity .28s ease,transform .28s ease,visibility .28s ease}
      #${PANEL_ID} .orp-toast.open{opacity:1;visibility:visible;transform:translateY(0);pointer-events:auto}
      #${PANEL_ID} .orp-toast[data-toast-type=warn]{border-color:#d6a400;background:#fff7db;color:#5a4304}
      #${PANEL_ID} .orp-toast[data-toast-type=success]{border-color:#118847;background:#edfdf3;color:#0b5a2b}
      #${PANEL_ID} .orp-toast[data-toast-type=error]{border-color:#d14a4a;background:#fff1f1;color:#842525}
      #${PANEL_ID} .orp-toast-title{font-weight:800;margin-bottom:4px}
      #${PANEL_ID} .orp-toast-text{line-height:1.45;white-space:pre-wrap}
      #${PANEL_ID} .orp-modal{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:rgba(4,9,19,.72);z-index:2147483648}
      #${PANEL_ID} .orp-modal.open{display:flex}
      #${PANEL_ID} .orp-modal-card{width:min(560px,92vw);background:#f7f9fc;color:#08101d;border:2px solid #08101d;box-shadow:10px 10px 0 rgba(0,0,0,.28);padding:16px;display:grid;gap:10px}
      #${PANEL_ID} .orp-modal[data-modal-type=success] .orp-modal-card{background:#edfdf3;border-color:#118847;color:#0b5a2b}
      #${PANEL_ID} .orp-modal[data-modal-type=error] .orp-modal-card{background:#fff1f1;border-color:#d14a4a;color:#842525}
      #${PANEL_ID} .orp-modal[data-modal-type=warn] .orp-modal-card{background:#fff7db;border-color:#d6a400;color:#5a4304}
      #${PANEL_ID} .orp-modal-head{display:grid;grid-template-columns:auto 1fr;gap:10px;align-items:start}
      #${PANEL_ID} .orp-modal-symbol{font-size:28px;line-height:1}
      #${PANEL_ID} .orp-modal-title{font-size:18px;font-weight:800}
      #${PANEL_ID} .orp-modal-msg{font-size:14px;line-height:1.5;white-space:pre-wrap}
      #${PANEL_ID} .orp-modal-actions{display:flex;justify-content:flex-end}
      #${PANEL_ID} .orp-hidden-store{display:none !important}
      #${PANEL_ID} .orp-random-summary{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:8px}
      #${PANEL_ID} .orp-random-card{border:1px solid #3a4555;background:#040913;padding:8px;display:grid;gap:4px;min-height:64px}
      #${PANEL_ID} .orp-random-card-label{font-size:10px;color:#8ea0b9;letter-spacing:.04em;text-transform:uppercase}
      #${PANEL_ID} .orp-random-card-value{font-size:16px;font-weight:800;color:#f8fafc;line-height:1.2;word-break:break-word}
      #${PANEL_ID} .orp-random-actions{display:flex;justify-content:flex-end;align-items:center}
      #${PANEL_ID} .orp-variant-editor-mask{position:fixed;inset:0;background:rgba(4,9,19,.7);z-index:2147483647;display:none}
      #${PANEL_ID} .orp-variant-editor-mask.open{display:block}
      #${PANEL_ID} .orp-variant-editor{position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);width:min(1260px,88vw);height:min(84vh,920px);background:#08101d;color:#f8fafc;border:1px solid #3a4555;box-shadow:0 0 0 1px rgba(130,170,255,.08),10px 10px 0 rgba(0,0,0,.24);z-index:2147483648;display:none;grid-template-rows:auto 1fr auto;border-radius:0}
      #${PANEL_ID} .orp-variant-editor.open{display:grid}
      #${PANEL_ID} .orp-variant-editor-header{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;padding:10px 12px;background:#050b14;border-bottom:1px solid #3a4555}
      #${PANEL_ID} .orp-variant-editor-body{padding:12px;display:grid;grid-template-columns:1fr 1fr;gap:12px;overflow:auto;background:#08101d}
      #${PANEL_ID} .orp-variant-card{border:1px solid #3a4555;background:#0a121f;display:grid;grid-template-rows:auto 1fr;min-height:280px}
      #${PANEL_ID} .orp-variant-card-head{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;padding:9px 10px;background:#06101d;border-bottom:1px solid #293446}
      #${PANEL_ID} .orp-variant-card-title{font-size:13px;font-weight:800;letter-spacing:.04em;color:#dce8ff}
      #${PANEL_ID} .orp-variant-list{padding:10px;display:grid;gap:8px;align-content:start;overflow:auto}
      #${PANEL_ID} .orp-variant-row{display:grid;grid-template-columns:56px 1fr 30px;gap:8px;align-items:start;padding:8px;border:1px solid #293446;background:#040913}
      #${PANEL_ID} .orp-variant-code{display:flex;align-items:center;justify-content:center;min-height:38px;border:1px solid #49576d;background:#09111c;color:#cfe0ff;font-weight:800;letter-spacing:.03em}
      #${PANEL_ID} .orp-variant-entry{width:100%;border:1px solid #49576d;background:#040913;color:#f8fafc;padding:8px;outline:none;resize:vertical;min-height:76px;line-height:1.45}
      #${PANEL_ID} .orp-variant-entry.subject{min-height:38px}
      #${PANEL_ID} .orp-variant-entry:focus{border-color:#8ab4ff}
      #${PANEL_ID} .orp-variant-remove{width:30px;padding:0;min-height:30px;align-self:start}
      #${PANEL_ID} .orp-variant-editor-footer{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:center;padding:10px 12px;border-top:1px solid #3a4555;background:#050b14}
      #${PANEL_ID} .orp-variant-editor-footer .orp-hint{font-size:10px}
      #${PANEL_ID} .orp-tiny{font-size:10px;color:#8ea0b9;letter-spacing:.03em}
      @media (max-width: 1080px){#${PANEL_ID} .orp-random-summary{grid-template-columns:repeat(3,minmax(0,1fr))}#${PANEL_ID} .orp-variant-editor{width:min(96vw,1260px)}#${PANEL_ID} .orp-variant-editor-body{grid-template-columns:1fr}}
      @media (max-width: 760px){#${PANEL_ID} .orp-random-summary{grid-template-columns:repeat(2,minmax(0,1fr))}#${PANEL_ID} .orp-settings{width:min(96vw,620px)}#${PANEL_ID} .orp-variant-editor-footer{grid-template-columns:1fr 1fr}#${PANEL_ID} .orp-variant-editor-footer .orp-hint{grid-column:1 / -1;order:3}}
      #${PANEL_ID} input[type=number]::-webkit-outer-spin-button,#${PANEL_ID} input[type=number]::-webkit-inner-spin-button{appearance:none;-webkit-appearance:none;margin:0}
      #${PANEL_ID} input[type=number]{-moz-appearance:textfield;appearance:textfield}
    `;
    document.head.appendChild(style);
  }

  function createUI() {
    const root = document.createElement('div');
    root.id = PANEL_ID;
    root.dataset.panelState = 'expanded';
    root.innerHTML = `
      <div class="orp-header">
        <div class="orp-title-wrap">
          <div class="orp-title">Outlook 定时发信面板 v${META.version}</div>
          <div class="orp-meta">作者：${META.author}</div>
        </div>
        <div class="orp-header-actions">
          <button class="orp-btn orp-toggle-btn" data-action="minimize" title="最小化">最小化</button>
          <button class="orp-icon-btn" data-action="close" title="关闭">×</button>
        </div>
      </div>
      <div class="orp-minibar">
        <div class="orp-minibar-row">
          <div>状态：<span class="orp-tag" id="orpMiniStatus">待机</span></div>
          <div>进度：<span id="orpMiniProgress">0 / 0</span></div>
          <div>RPA：<span class="orp-tag orp-rpa-state" id="orpMiniRpaStatus" data-rpa-status="idle" data-detail="" data-error-code="">idle</span></div>
        </div>
      </div>
      <div class="orp-body">
        <div class="orp-status">
          <div>状态：<span class="orp-tag" id="orpStatus">待机</span></div>
          <div>进度：<span id="orpProgress">0 / 0</span></div>
          <div>RPA：<span class="orp-tag orp-rpa-state" id="orpRpaStatus" data-rpa-status="idle" data-detail="" data-error-code="">idle</span></div>
        </div>
        <div class="orp-block">
          <div class="orp-label">批量任务</div>
          <textarea class="orp-textarea" id="orpTasks" placeholder="邮箱1;邮箱2|姓名|日期|时间&#10;zhangsan@example.com;zhangsan.work@example.com|张三|2026/03/27|09:00&#10;zhangsan@example.com|张三|2026-03-27|09:00"></textarea>
          <div class="orp-hint">多个邮箱只支持英文分号 ; 分隔。日期支持 YYYY-MM-DD 或 YYYY/MM/DD，写入 Outlook 时统一输出 YYYY/MM/DD。若当前处于草稿详情编辑态，程序将拒绝启动并弹窗报错。</div>
        </div>
        <div class="orp-toolbar">
          <button class="orp-btn primary" id="orpStart">开始运行</button>
          <button class="orp-btn" id="orpPause">暂停</button>
          <button class="orp-btn warn" id="orpStop">停止</button>
          <button class="orp-btn" id="orpSettingsBtn">设置</button>
        </div>
        <div class="orp-block">
          <div class="orp-label">日志</div>
          <div class="orp-log" id="orpLog"></div>
        </div>
      </div>
      <div class="orp-settings-mask" id="orpSettingsMask"></div>
      <div class="orp-settings" id="orpSettings">
        <div class="orp-settings-header">
          <div>
            <div class="orp-title">设置</div>
            <div class="orp-meta">Outlook 定时发信面板 v${META.version} · 作者：${META.author}</div>
          </div>
          <button class="orp-icon-btn" id="orpSettingsClose">×</button>
        </div>
        <div class="orp-settings-body">
          <div class="orp-section">
            <div class="orp-section-title">执行设置</div>
            <div class="orp-grid-3">
              <div class="orp-block"><div class="orp-label">步骤超时（秒）</div><input class="orp-input" id="orpStepTimeout" type="text" inputmode="numeric" value="25"></div>
              <div class="orp-block"><div class="orp-label">系统确认邮箱时间（s）</div><input class="orp-input" id="orpConfirmDelay" type="text" inputmode="decimal" value="0.5"></div>
              <div class="orp-block"><div class="orp-label">演示额外暂停（秒）</div><input class="orp-input" id="orpDemoExtra" type="text" inputmode="decimal" value="2" disabled></div>
            </div>
            <div class="orp-grid-2">
              <div class="orp-block"><div class="orp-label">下一封发送间隔最小值（秒）</div><input class="orp-input" id="orpNextIntervalMin" type="text" inputmode="numeric" value="8"></div>
              <div class="orp-block"><div class="orp-label">下一封发送间隔最大值（秒）</div><input class="orp-input" id="orpNextIntervalMax" type="text" inputmode="numeric" value="12"></div>
            </div>
            <label class="orp-checkbox"><input id="orpDemoMode" type="checkbox"> 自动化发送演示</label>
            <div class="orp-hint">打开后，会用方框高亮相关元素，并在正常暂停基础上额外暂停设定秒数。关闭后恢复正常速度。若只填最小值，则按固定值等待下一封；联系人之间等待最小不得低于 5 秒。</div>
          </div>
          <div class="orp-section">
            <div class="orp-section-title">模板设置</div>
            <div class="orp-grid-2">
              <div class="orp-block"><div class="orp-label">模板模式</div><select class="orp-input" id="orpTemplateMode"><option value="variants">随机抽取模式</option><option value="single">单模板模式</option></select></div>
              <div class="orp-block"><div class="orp-label">姓名占位符</div><input class="orp-input" id="orpNamePlaceholder" type="text" value="${DEFAULT_PLACEHOLDER}"></div>
            </div>
            <div class="orp-stack" id="orpSingleTemplateSection">
              <div class="orp-grid-2">
                <div class="orp-block"><div class="orp-label">主题模板</div><input class="orp-input" id="orpSubject" type="text" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" name="orp_subject_template"></div>
                <div class="orp-block"><div class="orp-label">单模板说明</div><div class="orp-hint">保持原有单模板工作流。发送时始终使用当前主题模板与正文模板。</div></div>
              </div>
              <div class="orp-block"><div class="orp-label">正文模板</div><textarea class="orp-textarea" id="orpBody" style="min-height:138px" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false"></textarea></div>
            </div>
            <div class="orp-stack" id="orpVariantTemplateSection">
              <div class="orp-random-summary">
                <div class="orp-random-card"><div class="orp-random-card-label">Subject</div><div class="orp-random-card-value" id="orpVariantCountSubjects">0</div></div>
                <div class="orp-random-card"><div class="orp-random-card-label">Opening</div><div class="orp-random-card-value" id="orpVariantCountOpenings">0</div></div>
                <div class="orp-random-card"><div class="orp-random-card-label">Body</div><div class="orp-random-card-value" id="orpVariantCountBodies">0</div></div>
                <div class="orp-random-card"><div class="orp-random-card-label">CTA</div><div class="orp-random-card-value" id="orpVariantCountCtas">0</div></div>
                <div class="orp-random-card"><div class="orp-random-card-label">Next</div><div class="orp-random-card-value" id="orpVariantNextCode">--</div></div>
                <div class="orp-random-card"><div class="orp-random-card-label">Combos</div><div class="orp-random-card-value" id="orpVariantTotalCombos">0</div></div>
              </div>
              <div class="orp-random-actions">
                <button class="orp-btn" id="orpOpenVariantEditor" type="button">编辑随机抽取库</button>
              </div>
              <textarea class="orp-hidden-store" id="orpVariantSubjects"></textarea>
              <textarea class="orp-hidden-store" id="orpVariantOpenings"></textarea>
              <textarea class="orp-hidden-store" id="orpVariantBodies"></textarea>
              <textarea class="orp-hidden-store" id="orpVariantCtas"></textarea>
            </div>
          </div>
        </div>
        <div class="orp-settings-footer">
          <div class="orp-hint">建议先用 1 条任务做演示验证。</div>
          <button class="orp-btn" id="orpReset">恢复默认</button>
          <button class="orp-btn primary" id="orpSaveSettings">保存并关闭</button>
        </div>
      </div>
      <div class="orp-variant-editor-mask" id="orpVariantEditorMask"></div>
      <div class="orp-variant-editor" id="orpVariantEditor">
        <div class="orp-variant-editor-header">
          <div>
            <div class="orp-title">随机抽取库</div>
            <div class="orp-meta">逐条编辑 · 保存后从下一批开始生效</div>
          </div>
          <button class="orp-icon-btn" id="orpVariantEditorClose" type="button">×</button>
        </div>
        <div class="orp-variant-editor-body">
          <div class="orp-variant-card" data-group="subjects">
            <div class="orp-variant-card-head">
              <div class="orp-variant-card-title">Subject</div>
              <button class="orp-btn" id="orpAddVariantSubjects" type="button">添加</button>
            </div>
            <div class="orp-variant-list" id="orpVariantEditorSubjects"></div>
          </div>
          <div class="orp-variant-card" data-group="openings">
            <div class="orp-variant-card-head">
              <div class="orp-variant-card-title">Opening</div>
              <button class="orp-btn" id="orpAddVariantOpenings" type="button">添加</button>
            </div>
            <div class="orp-variant-list" id="orpVariantEditorOpenings"></div>
          </div>
          <div class="orp-variant-card" data-group="bodies">
            <div class="orp-variant-card-head">
              <div class="orp-variant-card-title">Body</div>
              <button class="orp-btn" id="orpAddVariantBodies" type="button">添加</button>
            </div>
            <div class="orp-variant-list" id="orpVariantEditorBodies"></div>
          </div>
          <div class="orp-variant-card" data-group="ctas">
            <div class="orp-variant-card-head">
              <div class="orp-variant-card-title">CTA</div>
              <button class="orp-btn" id="orpAddVariantCtas" type="button">添加</button>
            </div>
            <div class="orp-variant-list" id="orpVariantEditorCtas"></div>
          </div>
        </div>
        <div class="orp-variant-editor-footer">
          <div class="orp-hint">编号会按当前非空条目自动重排。</div>
          <button class="orp-btn" id="orpVariantEditorCancel" type="button">取消</button>
          <button class="orp-btn primary" id="orpVariantEditorSave" type="button">保存</button>
        </div>
      </div>
      <div class="orp-demo-box" id="orpDemoBox"><div class="orp-demo-label" id="orpDemoLabel"></div></div>
      <div class="orp-toast" id="orpStatusToast" data-toast-type="" data-rpa-status="">
        <div class="orp-toast-title">状态提示</div>
        <div class="orp-toast-text"></div>
      </div>
      <div class="orp-modal" id="orpPageModal" data-modal-type="warn" data-rpa-status="" data-error-code="">
        <div class="orp-modal-card">
          <div class="orp-modal-head">
            <div class="orp-modal-symbol">⚠</div>
            <div>
              <div class="orp-modal-title">状态提示</div>
              <div class="orp-modal-msg"></div>
            </div>
          </div>
          <div class="orp-modal-actions"><button id="orpModalOk">确定</button></div>
        </div>
      </div>
    `;
    document.body.appendChild(root);
    STATE.els = {
      root,
      status: root.querySelector('#orpStatus'),
      miniStatus: root.querySelector('#orpMiniStatus'),
      progress: root.querySelector('#orpProgress'),
      miniProgress: root.querySelector('#orpMiniProgress'),
      rpaStatus: root.querySelector('#orpRpaStatus'),
      miniRpaStatus: root.querySelector('#orpMiniRpaStatus'),
      minToggle: root.querySelector('[data-action="minimize"]'),
      tasks: root.querySelector('#orpTasks'),
      log: root.querySelector('#orpLog'),
      start: root.querySelector('#orpStart'),
      pause: root.querySelector('#orpPause'),
      stop: root.querySelector('#orpStop'),
      settingsBtn: root.querySelector('#orpSettingsBtn'),
      settingsMask: root.querySelector('#orpSettingsMask'),
      settings: root.querySelector('#orpSettings'),
      settingsClose: root.querySelector('#orpSettingsClose'),
      stepTimeout: root.querySelector('#orpStepTimeout'),
      confirmDelay: root.querySelector('#orpConfirmDelay'),
      demoExtra: root.querySelector('#orpDemoExtra'),
      nextIntervalMin: root.querySelector('#orpNextIntervalMin'),
      nextIntervalMax: root.querySelector('#orpNextIntervalMax'),
      demo: root.querySelector('#orpDemoMode'),
      templateMode: root.querySelector('#orpTemplateMode'),
      subject: root.querySelector('#orpSubject'),
      namePlaceholder: root.querySelector('#orpNamePlaceholder'),
      body: root.querySelector('#orpBody'),
      variantSubjects: root.querySelector('#orpVariantSubjects'),
      variantOpenings: root.querySelector('#orpVariantOpenings'),
      variantBodies: root.querySelector('#orpVariantBodies'),
      variantCtas: root.querySelector('#orpVariantCtas'),
      variantCountSubjects: root.querySelector('#orpVariantCountSubjects'),
      variantCountOpenings: root.querySelector('#orpVariantCountOpenings'),
      variantCountBodies: root.querySelector('#orpVariantCountBodies'),
      variantCountCtas: root.querySelector('#orpVariantCountCtas'),
      variantNextCode: root.querySelector('#orpVariantNextCode'),
      variantTotalCombos: root.querySelector('#orpVariantTotalCombos'),
      openVariantEditor: root.querySelector('#orpOpenVariantEditor'),
      variantEditorMask: root.querySelector('#orpVariantEditorMask'),
      variantEditor: root.querySelector('#orpVariantEditor'),
      variantEditorClose: root.querySelector('#orpVariantEditorClose'),
      variantEditorCancel: root.querySelector('#orpVariantEditorCancel'),
      variantEditorSave: root.querySelector('#orpVariantEditorSave'),
      variantEditorSubjects: root.querySelector('#orpVariantEditorSubjects'),
      variantEditorOpenings: root.querySelector('#orpVariantEditorOpenings'),
      variantEditorBodies: root.querySelector('#orpVariantEditorBodies'),
      variantEditorCtas: root.querySelector('#orpVariantEditorCtas'),
      addVariantSubjects: root.querySelector('#orpAddVariantSubjects'),
      addVariantOpenings: root.querySelector('#orpAddVariantOpenings'),
      addVariantBodies: root.querySelector('#orpAddVariantBodies'),
      addVariantCtas: root.querySelector('#orpAddVariantCtas'),
      singleTemplateSection: root.querySelector('#orpSingleTemplateSection'),
      variantTemplateSection: root.querySelector('#orpVariantTemplateSection'),
      reset: root.querySelector('#orpReset'),
      saveSettings: root.querySelector('#orpSaveSettings'),
    };
    STATE.demoBox = root.querySelector('#orpDemoBox');
    STATE.toast = root.querySelector('#orpStatusToast');
    STATE.modal = root.querySelector('#orpPageModal');
    STATE.rpaStatus = root.querySelector('#orpRpaStatus');
    STATE.els.subject.value = DEFAULT_SUBJECT;
    STATE.els.body.value = DEFAULT_BODY;
    STATE.els.templateMode.value = DEFAULT_TEMPLATE_MODE;
    STATE.els.namePlaceholder.value = DEFAULT_PLACEHOLDER;
    setStoredVariantItems(STATE.els.variantSubjects, DEFAULT_VARIANT_SUBJECTS);
    setStoredVariantItems(STATE.els.variantOpenings, DEFAULT_VARIANT_OPENINGS);
    setStoredVariantItems(STATE.els.variantBodies, DEFAULT_VARIANT_BODIES);
    setStoredVariantItems(STATE.els.variantCtas, DEFAULT_VARIANT_CTAS);
    refreshVariantSummary();
  }

  function makeDraggable(node, handle) {
    let dragging = false;
    let sx = 0; let sy = 0; let sl = 0; let st = 0;
    handle.addEventListener('mousedown', (e) => {
      if (e.target.closest('button,input,textarea,label')) return;
      dragging = true;
      const r = node.getBoundingClientRect();
      sx = e.clientX; sy = e.clientY; sl = r.left; st = r.top;
      document.body.style.userSelect = 'none';
    });
    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      node.style.left = `${clamp(sl + e.clientX - sx, 0, window.innerWidth - 260)}px`;
      node.style.top = `${clamp(st + e.clientY - sy, 0, window.innerHeight - 120)}px`;
      node.style.right = 'auto';
      node.style.bottom = 'auto';
    });
    document.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      document.body.style.userSelect = '';
      saveState();
    });
  }

  function getVariantDraftLabel(groupKey, idx) {
    const meta = VARIANT_GROUPS.find((x) => x.key === groupKey);
    return `${meta ? meta.prefix : '?'}${idx + 1}`;
  }

  function autoGrowVariantField(el) {
    if (!el || el.tagName !== 'TEXTAREA') return;
    const minHeight = Number(el.dataset.minHeight || 38);
    el.style.height = '0px';
    el.style.height = `${Math.max(minHeight, el.scrollHeight)}px`;
  }

  function getVariantDraftSource() {
    const parts = getVariantPartsFromFields();
    return {
      subjects: parts.subjects.length ? parts.subjects.slice() : [''],
      openings: parts.openings.length ? parts.openings.slice() : [''],
      bodies: parts.bodies.length ? parts.bodies.slice() : [''],
      ctas: parts.ctas.length ? parts.ctas.slice() : [''],
    };
  }

  function renderVariantEditorGroup(groupKey) {
    if (!STATE.variantDraft) return;
    const containerMap = {
      subjects: STATE.els.variantEditorSubjects,
      openings: STATE.els.variantEditorOpenings,
      bodies: STATE.els.variantEditorBodies,
      ctas: STATE.els.variantEditorCtas,
    };
    const container = containerMap[groupKey];
    if (!container) return;
    const values = STATE.variantDraft[groupKey] || [];
    if (!values.length) values.push('');
    container.innerHTML = '';
    values.forEach((value, idx) => {
      const row = document.createElement('div');
      row.className = 'orp-variant-row';
      const tag = document.createElement('div');
      tag.className = 'orp-variant-code';
      tag.textContent = getVariantDraftLabel(groupKey, idx);
      const field = document.createElement('textarea');
      field.className = 'orp-variant-entry';
      if (groupKey === 'subjects') field.classList.add('subject');
      field.value = String(value || '');
      field.rows = groupKey === 'subjects' ? 2 : 4;
      field.dataset.minHeight = groupKey === 'subjects' ? '38' : '76';
      field.addEventListener('input', () => {
        STATE.variantDraft[groupKey][idx] = field.value;
        autoGrowVariantField(field);
      });
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'orp-icon-btn orp-variant-remove';
      remove.textContent = '×';
      remove.title = '删除';
      remove.addEventListener('click', () => {
        STATE.variantDraft[groupKey].splice(idx, 1);
        if (!STATE.variantDraft[groupKey].length) STATE.variantDraft[groupKey].push('');
        renderVariantEditorGroup(groupKey);
      });
      row.appendChild(tag);
      row.appendChild(field);
      row.appendChild(remove);
      container.appendChild(row);
      autoGrowVariantField(field);
    });
  }

  function renderVariantEditor() {
    renderVariantEditorGroup('subjects');
    renderVariantEditorGroup('openings');
    renderVariantEditorGroup('bodies');
    renderVariantEditorGroup('ctas');
  }

  function openVariantEditor() {
    STATE.variantDraft = getVariantDraftSource();
    STATE.els.variantEditorMask.classList.add('open');
    STATE.els.variantEditor.classList.add('open');
    renderVariantEditor();
  }

  function closeVariantEditor() {
    STATE.variantDraft = null;
    STATE.els.variantEditorMask.classList.remove('open');
    STATE.els.variantEditor.classList.remove('open');
  }

  function addVariantDraftRow(groupKey) {
    if (!STATE.variantDraft) STATE.variantDraft = getVariantDraftSource();
    STATE.variantDraft[groupKey].push('');
    renderVariantEditorGroup(groupKey);
    const containerMap = {
      subjects: STATE.els.variantEditorSubjects,
      openings: STATE.els.variantEditorOpenings,
      bodies: STATE.els.variantEditorBodies,
      ctas: STATE.els.variantEditorCtas,
    };
    const fields = containerMap[groupKey]?.querySelectorAll('textarea');
    const last = fields && fields[fields.length - 1];
    if (last) last.focus();
  }

  function saveVariantEditor() {
    if (!STATE.variantDraft) return;
    setStoredVariantItems(STATE.els.variantSubjects, STATE.variantDraft.subjects);
    setStoredVariantItems(STATE.els.variantOpenings, STATE.variantDraft.openings);
    setStoredVariantItems(STATE.els.variantBodies, STATE.variantDraft.bodies);
    setStoredVariantItems(STATE.els.variantCtas, STATE.variantDraft.ctas);
    refreshVariantSummary();
    saveState();
    closeVariantEditor();
    logInfo('随机抽取库已保存。');
  }

  function resetDefaults() {
    STATE.els.stepTimeout.value = '25';
    STATE.els.confirmDelay.value = '0.5';
    STATE.els.demo.checked = false;
    STATE.els.demoExtra.value = '2';
    STATE.els.demoExtra.disabled = !STATE.els.demo.checked;
    STATE.els.nextIntervalMin.value = '8';
    STATE.els.nextIntervalMax.value = '12';
    STATE.els.templateMode.value = DEFAULT_TEMPLATE_MODE;
    STATE.els.subject.value = DEFAULT_SUBJECT;
    STATE.els.namePlaceholder.value = DEFAULT_PLACEHOLDER;
    STATE.els.body.value = DEFAULT_BODY;
    setStoredVariantItems(STATE.els.variantSubjects, DEFAULT_VARIANT_SUBJECTS);
    setStoredVariantItems(STATE.els.variantOpenings, DEFAULT_VARIANT_OPENINGS);
    setStoredVariantItems(STATE.els.variantBodies, DEFAULT_VARIANT_BODIES);
    setStoredVariantItems(STATE.els.variantCtas, DEFAULT_VARIANT_CTAS);
    STATE.variantCursor = 0;
    STATE.variantSignature = '';
    syncTemplateModeUI();
    refreshVariantSummary();
    saveState();
    logInfo('设置已恢复默认。');
  }

  function openSettings() {
    refreshVariantSummary();
    STATE.els.settingsMask.classList.add('open');
    STATE.els.settings.classList.add('open');
  }

  function closeSettings() {
    closeVariantEditor();
    STATE.els.settingsMask.classList.remove('open');
    STATE.els.settings.classList.remove('open');
  }

  function syncIntervalInputs() {
    const rawMin = Number(STATE.els.nextIntervalMin.value || 0);
    const safeMin = Number.isFinite(rawMin) ? Math.max(5, rawMin) : 5;
    const rawMaxText = String(STATE.els.nextIntervalMax.value || '').trim();
    const rawMax = rawMaxText === '' ? safeMin : Number(rawMaxText);
    const safeMax = Number.isFinite(rawMax) ? Math.max(safeMin, rawMax) : safeMin;
    STATE.els.nextIntervalMin.value = String(safeMin);
    STATE.els.nextIntervalMax.value = String(safeMax);
  }

  function bindUI() {
    STATE.els.minToggle.addEventListener('click', () => {
      setPanelMinimized(!(STATE.els.root.dataset.panelState === 'minimized'));
      saveState();
    });
    STATE.els.root.querySelector('[data-action="close"]').addEventListener('click', () => {
      hideStatusToast();
      hidePageModal();
      closeVariantEditor();
      STATE.els.root.remove();
      const style = document.getElementById(STYLE_ID);
      if (style) style.remove();
    });
    STATE.els.start.addEventListener('click', startRun);
    STATE.els.pause.addEventListener('click', pauseRun);
    STATE.els.stop.addEventListener('click', stopRun);
    STATE.els.settingsBtn.addEventListener('click', openSettings);
    STATE.els.settingsClose.addEventListener('click', closeSettings);
    STATE.els.settingsMask.addEventListener('click', closeSettings);
    STATE.els.saveSettings.addEventListener('click', () => { refreshVariantSummary(); saveState(); closeSettings(); logInfo('设置已保存。'); });
    STATE.els.reset.addEventListener('click', resetDefaults);
    STATE.modal.querySelector('#orpModalOk').addEventListener('click', hidePageModal);
    STATE.els.openVariantEditor.addEventListener('click', openVariantEditor);
    STATE.els.variantEditorClose.addEventListener('click', closeVariantEditor);
    STATE.els.variantEditorCancel.addEventListener('click', closeVariantEditor);
    STATE.els.variantEditorMask.addEventListener('click', closeVariantEditor);
    STATE.els.variantEditorSave.addEventListener('click', saveVariantEditor);
    STATE.els.addVariantSubjects.addEventListener('click', () => addVariantDraftRow('subjects'));
    STATE.els.addVariantOpenings.addEventListener('click', () => addVariantDraftRow('openings'));
    STATE.els.addVariantBodies.addEventListener('click', () => addVariantDraftRow('bodies'));
    STATE.els.addVariantCtas.addEventListener('click', () => addVariantDraftRow('ctas'));
    const syncDemoExtra = () => { STATE.els.demoExtra.disabled = !STATE.els.demo.checked; };
    STATE.els.demo.addEventListener('change', syncDemoExtra);
    STATE.els.templateMode.addEventListener('change', () => { syncTemplateModeUI(); saveState(); });
    STATE.els.namePlaceholder.addEventListener('input', refreshVariantSummary);
    STATE.els.namePlaceholder.addEventListener('change', refreshVariantSummary);
    syncDemoExtra();
    syncTemplateModeUI();
    syncIntervalInputs();
    [
      STATE.els.tasks, STATE.els.stepTimeout, STATE.els.confirmDelay, STATE.els.nextIntervalMin,
      STATE.els.nextIntervalMax, STATE.els.demo, STATE.els.subject, STATE.els.namePlaceholder, STATE.els.body,
    ].forEach((el) => {
      el.addEventListener('input', saveState);
      el.addEventListener('change', saveState);
    });
    [STATE.els.nextIntervalMin, STATE.els.nextIntervalMax].forEach((el) => {
      el.addEventListener('blur', () => {
        syncIntervalInputs();
        saveState();
      });
    });
    makeDraggable(STATE.els.root, STATE.els.root.querySelector('.orp-header'));
    makeDraggable(STATE.els.settings, STATE.els.settings.querySelector('.orp-settings-header'));
  }

  function waitConfig() {
    const stepTimeoutMs = Math.max(5000, Number(STATE.els.stepTimeout.value || 25) * 1000);
    const confirmDelayMs = Math.max(100, Math.round(Number(STATE.els.confirmDelay.value || 0.5) * 1000));
    const minSec = Math.max(5, Number(STATE.els.nextIntervalMin.value || 0));
    const rawMax = String(STATE.els.nextIntervalMax.value || '').trim();
    const maxSec = rawMax === '' ? minSec : Math.max(minSec, Number(rawMax));
    return {
      stepTimeoutMs,
      confirmDelayMs,
      nextIntervalMinMs: Math.round(minSec * 1000),
      nextIntervalMaxMs: Math.round((maxSec >= minSec ? maxSec : minSec) * 1000),
      demoMode: !!STATE.els.demo.checked,
      demoExtraMs: !!STATE.els.demo.checked ? Math.max(0, Math.round(Number(STATE.els.demoExtra.value || 2) * 1000)) : 0,
      templateMode: STATE.els.templateMode.value || DEFAULT_TEMPLATE_MODE,
      subject: STATE.els.subject.value,
      body: STATE.els.body.value,
      namePlaceholder: STATE.els.namePlaceholder.value || DEFAULT_PLACEHOLDER,
      variantSubjectsRaw: STATE.els.variantSubjects.value,
      variantOpeningsRaw: STATE.els.variantOpenings.value,
      variantBodiesRaw: STATE.els.variantBodies.value,
      variantCtasRaw: STATE.els.variantCtas.value,
      tasksRaw: STATE.els.tasks.value,
    };
  }

  function validateConfig(cfg) {
    if (!cfg.namePlaceholder) throw new Error('姓名占位符不能为空。');
    if (cfg.templateMode === 'single') {
      if (!cfg.subject) throw new Error('主题模板不能为空。');
      if (!cfg.body) throw new Error('正文模板不能为空。');
      return;
    }
    const parts = getTemplatePartsFromConfig(cfg);
    if (!parts.subjects.length) throw new Error('随机抽取 Subject 至少需要 1 条。');
    if (!parts.openings.length) throw new Error('随机抽取 Opening 至少需要 1 条。');
    if (!parts.bodies.length) throw new Error('随机抽取 Body 至少需要 1 条。');
    if (!parts.ctas.length) throw new Error('随机抽取 CTA 至少需要 1 条。');
  }

  function parseTasks(raw) {
    const lines = String(raw || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
    const tasks = [];
    lines.forEach((line, idx) => {
      const parts = line.split('|').map((x) => x.trim());
      if (parts.length < 4) throw new Error(`第 ${idx + 1} 行格式错误，应为：邮箱1;邮箱2|姓名|日期|时间`);
      const emails = (parts[0] || '').replace(/；/g, ';').split(';').map((x) => x.trim()).filter(Boolean);
      if (!emails.length) throw new Error(`第 ${idx + 1} 行没有有效邮箱。`);
      const invalid = emails.find((x) => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x));
      if (invalid) throw new Error(`第 ${idx + 1} 行邮箱格式错误：${invalid}`);
      const date = normalizeTaskDate(parts[2] || '');
      const time = parts[3] || '';
      if (!date) throw new Error(`第 ${idx + 1} 行日期格式错误，应为 YYYY-MM-DD 或 YYYY/MM/DD。`);
      if (!/^\d{2}:\d{2}$/.test(time)) throw new Error(`第 ${idx + 1} 行时间格式错误，应为 HH:MM。`);
      tasks.push({ lineNumber: idx + 1, raw: line, emails, name: parts[1] || '', date, time });
    });
    return tasks;
  }

  async function maybePause() {

    while (STATE.paused && !STATE.stopRequested) await sleep(200);
    if (STATE.stopRequested) throw new Error('已停止。');
  }

  async function waitFor(fn, timeoutMs, label, interval = 180) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (STATE.stopRequested) throw new Error('已停止。');
      const x = fn();
      if (x) return x;
      await sleep(interval);
    }
    throw new Error(`等待超时：${label}`);
  }

  async function waitGone(fn, timeoutMs, label, interval = 180) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (STATE.stopRequested) throw new Error('已停止。');
      if (!fn()) return true;
      await sleep(interval);
    }
    throw new Error(`等待未消失：${label}`);
  }

  async function demoFocus(el, label, cfg) {
    if (!cfg.demoMode || !el || !STATE.demoBox) return;
    const r = el.getBoundingClientRect();
    const box = STATE.demoBox;
    box.style.left = `${r.left - 4}px`;
    box.style.top = `${r.top - 4}px`;
    box.style.width = `${r.width + 8}px`;
    box.style.height = `${r.height + 8}px`;
    box.querySelector('#orpDemoLabel').textContent = label;
    box.classList.add('open');
    await sleep(cfg.demoExtraMs);
  }

  function hideDemoFocus() {
    if (STATE.demoBox) STATE.demoBox.classList.remove('open');
  }

  function findButtons(root = document) {
    return qsaVisible('button, [role="button"], [role="menuitem"]', root);
  }

  function findButtonByText(targets, root = document, mode = 'eq') {
    return findButtons(root).find((el) => {
      const txt = text(el);
      return mode === 'contains' ? includesAny(txt, targets) : eqAny(txt, targets);
    }) || null;
  }

  function findButtonContainsDeep(targets, root = document) {
    return findButtons(root).find((el) => {
      if (includesAny(text(el), targets)) return true;
      return Array.from(el.querySelectorAll('div,span')).some((node) => includesAny(text(node), targets));
    }) || null;
  }

  function findNewMail() {
    const primary = findButtonByText(LANG.newMail, document, 'eq')
      || qsaVisible('button, [role="button"]', document).find((el) => eqAny(el.getAttribute('aria-label') || '', LANG.newMail) || includesAny(text(el), LANG.newMail) || includesAny(el.getAttribute('title') || '', LANG.newMail))
      || null;
    if (!primary) return { primary: null, menu: null };
    const group = primary.closest('[data-automation-type="RibbonSplitButton"], [class*="SplitButton"], [class*="splitButton"]') || primary.parentElement;
    const menu = group ? qsaVisible('button, [role="button"]', group).find((el) => el !== primary) || null : null;
    return { primary, menu };
  }

  function findRecipient() {
    const labelBtn = qsaVisible('button, [role="button"]', document).find((el) => {
      return el.id === 'recipient-well-label-to'
        || eqAny(text(el), LANG.recipientLabel)
        || eqAny(el.getAttribute('aria-label') || '', LANG.recipientLabel);
    }) || null;

    let input = null;
    let shell = null;

    if (labelBtn) {
      let node = labelBtn;
      for (let i = 0; node && i < 8; i += 1, node = node.parentElement) {
        const found = qsaVisible('input[type="text"], textarea, [contenteditable="true"]', node).find((el) => !el.closest('[role="dialog"]')) || null;
        if (found) {
          input = found;
          shell = node;
          break;
        }
      }
      shell = shell || labelBtn.closest('[data-app-section], [role="main"], form, section, div') || labelBtn.parentElement;
    }

    if (!input) {
      input = qsaVisible('[contenteditable="true"][aria-label], input[type="text"]', document).find((el) => {
        if (el.closest('[role="dialog"]')) return false;
        const aria = low(el.getAttribute('aria-label') || '');
        const role = low(el.getAttribute('role') || '');
        return aria.includes('to') || aria.includes('收件') || (role === 'textbox' && !!labelBtn);
      }) || null;
    }

    shell = shell || input?.parentElement || null;
    return { labelBtn, shell, input };
  }

  function findSendPrimary() {
    return qsaVisible('button, [role="button"]', document).find((el) => {
      if (el.closest('[role="dialog"]')) return false;
      return eqAny(text(el), LANG.send) || eqAny(el.getAttribute('aria-label') || '', LANG.send);
    }) || null;
  }

  function findSendMenu(sendPrimary) {
    if (!sendPrimary) return null;
    const sendRect = sendPrimary.getBoundingClientRect();
    const sameRow = (el) => {
      const r = el.getBoundingClientRect();
      return Math.abs(r.top - sendRect.top) <= 32 && r.left >= sendRect.right - 12 && r.left <= sendRect.right + 220;
    };
    const candidates = qsaVisible('button, [role="button"]', document).filter((el) => el !== sendPrimary && !el.closest('[role="dialog"]'));
    return candidates
      .map((el) => {
        let score = 0;
        const aria = norm(el.getAttribute('aria-label') || '');
        const title = norm(el.getAttribute('title') || '');
        const cls = norm(el.className || '');
        if (sameRow(el)) score += 8;
        if (el.getAttribute('aria-haspopup') === 'menu') score += 8;
        if (includesAny(aria, LANG.moreSendOptions) || includesAny(title, LANG.moreSendOptions)) score += 9;
        if (/MenuButton|SplitButton__menuButton|menuButton/i.test(cls)) score += 5;
        return { el, score };
      })
      .filter((x) => x.score > 0)
      .sort((a, b) => b.score - a.score)[0]?.el || null;
  }

  function getCompose() {
    const recipient = findRecipient();
    const subject = qsaVisible('input', document).find((el) => !el.closest('[role="dialog"]') && (eqAny(el.getAttribute('aria-label') || '', LANG.subjectLabel) || eqAny(el.getAttribute('placeholder') || '', LANG.subjectPlaceholder))) || null;
    const body = qsaVisible('[role="textbox"], [contenteditable="true"]', document).find((el) => !el.closest('[role="dialog"]') && (eqAny(el.getAttribute('aria-label') || '', LANG.bodyLabel) || el.getAttribute('data-ms-editor') === 'true')) || null;
    const sendPrimary = findSendPrimary();
    const sendMenu = findSendMenu(sendPrimary);
    return { recipient, subject, body, sendPrimary, sendMenu };
  }

  function getDialogs() {
    return qsaVisible('[role="dialog"]', document);
  }

  function getCustomDialog() {
    return getDialogs().find((dlg) => dlg.querySelector('input[id^="datePicker-input"], input[role="combobox"][aria-haspopup="dialog"]') && dlg.querySelector('input.fui-TimePicker__input, .fui-TimePicker input[role="combobox"]')) || null;
  }

  function findCustomTimeButton(dlg) {
    return findButtonContainsDeep(LANG.customTime, dlg);
  }

  function getScheduleDialog() {
    return getDialogs().find((dlg) => {
      if (dlg === getCustomDialog()) return false;
      return /schedule send|计划发送/i.test(text(dlg)) || !!findCustomTimeButton(dlg);
    }) || null;
  }

  function stage() {
    if (getCustomDialog()) return 'scheduleCustom';
    if (getScheduleDialog()) return 'schedule';
    const c = getCompose();
    if (c.recipient.input && c.subject && c.body && c.sendPrimary) return 'compose';
    if (findNewMail().primary) return 'main';
    return 'unknown';
  }

  function composeEmpty(c) {
    const shellText = norm(text(c.recipient.shell || c.recipient.input || c.recipient.labelBtn || '')).replace(/\b(收件人|抄送|密件抄送|to|cc|bcc)\b/gi, '').trim();
    const recipientValue = c.recipient.input && 'value' in c.recipient.input ? norm(c.recipient.input.value || '') : '';
    const subjectValue = norm(c.subject?.value || '');
    const bodyValue = norm(text(c.body || ''));
    return !shellText && !recipientValue && !subjectValue && !bodyValue;
  }

  async function cancelDialogs(timeoutMs) {
    for (let i = 0; i < 3; i += 1) {
      const custom = getCustomDialog();
      if (custom) {
        const cancelBtn = findButtonByText(LANG.cancel, custom, 'eq');
        if (cancelBtn) {
          click(cancelBtn);
          await waitGone(() => getCustomDialog(), timeoutMs, '自定义时间弹窗关闭');
          await sleep(220);
          continue;
        }
      }
      const schedule = getScheduleDialog();
      if (schedule) {
        const cancelBtn = findButtonByText(LANG.cancel, schedule, 'eq');
        if (cancelBtn) {
          click(cancelBtn);
          await waitGone(() => getScheduleDialog(), timeoutMs, '计划发送弹层关闭');
          await sleep(220);
          continue;
        }
      }
      break;
    }
  }

  async function ensureReady(cfg) {
    let s = stage();
    if (s === 'scheduleCustom' || s === 'schedule') {
      logWarn('检测到计划发送弹窗残留，先安全返回。');
      await cancelDialogs(cfg.stepTimeoutMs);
      s = stage();
    }
    if (s === 'compose') {
      const c = getCompose();
      if (!c.recipient.input || !c.subject || !c.body) throw new Error('当前处于异常编辑态，关键输入区不完整。');
      if (composeEmpty(c)) return 'compose';
      throw new Error('当前处于草稿详情编辑态或非空写信界面，已统一定义为禁止启动。请先关闭或清空当前草稿，再重新开始。');
    }
    if (s === 'main') return 'main';
    if (findNewMail().primary && !getCompose().subject && !getCompose().body) return 'main';
    throw new Error('当前页面不是可安全启动的主界面，已拒绝运行。');
  }

  async function getComposeReady(timeoutMs) {
    return waitFor(() => {
      const c = getCompose();
      return c.recipient.input && c.subject && c.body && c.sendPrimary ? c : null;
    }, timeoutMs, '写信界面就绪');
  }

  async function getBodyReady(timeoutMs) {
    return waitFor(() => {
      const c = getCompose();
      return c.body ? c : null;
    }, timeoutMs, '正文区域就绪');
  }

  async function clickNewMail(timeoutMs, cfg) {
    const nm = await waitFor(() => findNewMail().primary, timeoutMs, '新邮件按钮');
    logStep(1, '点击“新邮件”...');
    await demoFocus(nm, '新建邮件', cfg);
    click(nm);
    hideDemoFocus();
    await getComposeReady(timeoutMs);
    await sleep(220);
  }

  async function writeRecipients(task, cfg) {
    logStep(2, `逐个写入收件人，共 ${task.emails.length} 个...`);
    for (let i = 0; i < task.emails.length; i += 1) {
      await maybePause();
      const c = await getComposeReady(cfg.stepTimeoutMs);
      const box = c.recipient.input;
      const email = task.emails[i];
      if (!box) throw new Error('未找到收件人输入框。');
      await demoFocus(box, `收件人 ${i + 1}`, cfg);
      if ('value' in box) setInputValue(box, email);
      else {
        box.innerHTML = '';
        document.execCommand('insertText', false, email);
        fireInput(box);
      }
      hideDemoFocus();
      await sleep(60);
      await demoFocus(box, '补分号', cfg);
      typeSemicolon(box);
      hideDemoFocus();
      logHtml(`<span class="orp-log-step">Step.2:</span> 已输入邮箱 ${i + 1}/${task.emails.length}：<span class="orp-log-email">${esc(email)}</span>`);
      logInfo(`固定等待 ${cfg.confirmDelayMs} ms，让 Outlook 自行解析当前邮箱...`);
      await sleep(cfg.confirmDelayMs);
    }
  }

  async function writeSubject(task, cfg, template) {
    logStep(3, '写入主题...');
    const c = await getComposeReady(cfg.stepTimeoutMs);
    await demoFocus(c.subject, '主题', cfg);
    c.subject.focus();
    setInputValue(c.subject, template.subject);
    c.subject.dispatchEvent(new Event('blur', { bubbles: true }));
    hideDemoFocus();
    await sleep(180);
  }

  async function writeBody(task, cfg, template) {
    logStep(4, '写入正文...');
    const c = await getBodyReady(cfg.stepTimeoutMs);
    await demoFocus(c.body, '正文', cfg);
    c.body.focus();
    prependComposeBody(c.body, template.body);
    hideDemoFocus();
    await sleep(220);
  }

  async function clickSendMenu(timeoutMs, cfg) {
    logStep(5, '展开发送下拉菜单...');
    const c = await getComposeReady(timeoutMs);
    const menuBtn = c.sendMenu || findSendMenu(c.sendPrimary);
    if (!menuBtn) throw new Error('未找到发送下拉按钮。');
    await demoFocus(menuBtn, '发送下拉', cfg);
    click(menuBtn);
    hideDemoFocus();
    await waitFor(() => qsaVisible('[role="menu"]', document).find((m) => findButtonByText(LANG.scheduleSend, m, 'contains')), timeoutMs, '发送下拉菜单出现');
  }

  async function clickScheduleSend(timeoutMs, cfg) {
    logStep(6, '选择“计划发送”...');
    const menu = await waitFor(() => qsaVisible('[role="menu"]', document).find((m) => findButtonByText(LANG.scheduleSend, m, 'contains')), timeoutMs, '发送下拉菜单');
    const item = findButtonByText(LANG.scheduleSend, menu, 'contains');
    if (!item) throw new Error('未找到“计划发送”菜单项。');
    await demoFocus(item, '计划发送', cfg);
    click(item);
    hideDemoFocus();
    try {
      await waitFor(() => getScheduleDialog(), Math.min(timeoutMs, 4000), '计划发送弹层出现');
    } catch (err) {
      const direct = getDialogs().find((d) => /schedule send|计划发送/i.test(text(d)) || !!findCustomTimeButton(d));
      if (!direct) throw err;
      logWarn('已看到 Schedule send 弹层，但原检测未命中，已自动切换到兜底识别。');
    }
    await sleep(240);
  }

  async function clickCustomTime(timeoutMs, cfg) {
    logStep(7, '点击“自定义时间”...');
    const dlg = await waitFor(() => getScheduleDialog() || getDialogs().find((d) => !!findCustomTimeButton(d)), timeoutMs, '计划发送弹层');
    const btn = findCustomTimeButton(dlg) || findButtonByText(LANG.customTime, dlg, 'contains');
    if (!btn) throw new Error('未找到“自定义时间”按钮。');
    logInfo(`已定位到“自定义时间”按钮：${text(btn)}`);
    await demoFocus(btn, '自定义时间', cfg);
    click(btn);
    hideDemoFocus();
    await waitFor(() => getCustomDialog(), timeoutMs, '设置自定义日期和时间弹窗出现');
    await sleep(240);
  }

  async function getCustomReady(timeoutMs) {
    return waitFor(() => {
      const dlg = getCustomDialog();
      if (!dlg) return null;
      const dateInput = dlg.querySelector('input[id^="datePicker-input"], input[role="combobox"][aria-haspopup="dialog"]');
      const timeInput = dlg.querySelector('input.fui-TimePicker__input, .fui-TimePicker input[role="combobox"]');
      const sendBtn = findButtonByText(LANG.send, dlg, 'eq');
      const cancelBtn = findButtonByText(LANG.cancel, dlg, 'eq');
      return dateInput && timeInput && sendBtn && cancelBtn ? { dlg, dateInput, timeInput, sendBtn, cancelBtn } : null;
    }, timeoutMs, '自定义时间弹窗就绪');
  }

  async function setDateTime(task, cfg) {
    logStep(8, `设置计划发送时间 ${task.date} ${task.time} ...`, ` <span class="orp-log-time-strong">${esc(task.date)} ${esc(task.time)}</span>`);
    const x = await getCustomReady(cfg.stepTimeoutMs);
    await demoFocus(x.dateInput, '日期', cfg);
    x.dateInput.focus();
    setInputValue(x.dateInput, task.date);
    x.dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
    hideDemoFocus();
    await sleep(180);
    await demoFocus(x.timeInput, '时间', cfg);
    x.timeInput.focus();
    setInputValue(x.timeInput, task.time);
    x.timeInput.dispatchEvent(new Event('blur', { bubbles: true }));
    fireKey(x.timeInput, 'Enter', 'Enter');
    hideDemoFocus();
    await sleep(280);
  }

  async function clickFinalSend(timeoutMs, cfg) {
    logStep(9, '点击最终发送...');
    const btn = await waitFor(() => {
      const dlg = getCustomDialog();
      if (!dlg) return null;
      const b = findButtonByText(LANG.send, dlg, 'eq');
      return b && !b.disabled && b.getAttribute('aria-disabled') !== 'true' ? b : null;
    }, timeoutMs, '最终发送按钮可点击');
    await demoFocus(btn, '最终发送', cfg);
    click(btn);
    hideDemoFocus();
    await waitGone(() => getCustomDialog(), timeoutMs, '自定义时间弹窗关闭');
    await sleep(280);
  }

  async function waitTaskSettled(timeoutMs) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (STATE.stopRequested) throw new Error('已停止。');
      if (stage() === 'main') return true;
      await sleep(220);
    }
    return true;
  }

  async function processTask(task, cfg, template) {
    const ready = await ensureReady(cfg);
    if (ready === 'main') await clickNewMail(cfg.stepTimeoutMs, cfg);
    await maybePause();
    await writeRecipients(task, cfg);
    await maybePause();
    await writeSubject(task, cfg, template);
    await maybePause();
    await writeBody(task, cfg, template);
    await maybePause();
    await clickSendMenu(cfg.stepTimeoutMs, cfg);
    await maybePause();
    await clickScheduleSend(cfg.stepTimeoutMs, cfg);
    await maybePause();
    await clickCustomTime(cfg.stepTimeoutMs, cfg);
    await maybePause();
    await setDateTime(task, cfg);
    await maybePause();
    await clickFinalSend(cfg.stepTimeoutMs, cfg);
    await waitTaskSettled(Math.min(cfg.stepTimeoutMs, 12000));
    logSuccess(`第 ${task.lineNumber} 行已完成。`);
  }

  async function runLoop(cfg) {
    while (STATE.currentIndex < STATE.tasks.length) {
      await maybePause();
      const task = STATE.tasks[STATE.currentIndex];
      const template = getResolvedTemplate(cfg, task);
      setProgress(STATE.currentIndex + 1, STATE.tasks.length);
      logHtml('━━━━━━━━━━━━━━━━━━━━━━━━', '');
      logTaskStart(task, template.code);
      logInfo(`本封使用模板组合：${template.code}`);
      await processTask(task, cfg, template);
      STATE.currentIndex += 1;
      STATE.successCount += 1;
      if (cfg.templateMode === 'variants' && cfg.variantRuntime) {
        persistVariantCursor((STATE.variantCursor + 1) % cfg.variantRuntime.sequence.length, cfg.variantRuntime.signature);
      }
      setProgress(STATE.currentIndex, STATE.tasks.length);
      if (STATE.currentIndex < STATE.tasks.length) {
        const waitMs = randomBetween(cfg.nextIntervalMinMs, cfg.nextIntervalMaxMs);
        setStatus('等待下一封');
        setRpaStatus('waiting-next', `next:${waitMs}`);
        logInfo(`下一封等待 ${waitMs} ms ...`);
        showStatusToast('warn', '当前联系人已完成，等待下一封邮件发送...', WAIT_TOAST_MS);
        await sleep(waitMs);
        if (!STATE.stopRequested && !STATE.paused) {
          setStatus('运行中');
          setRpaStatus('running');
        }
      }
    }
    STATE.running = false;
    setStatus('全部完成');
    setRpaStatus('success');
    logSuccess('全部任务已完成。');
    showPageModal('success', '批量任务已完成', `运行完成。
成功：${STATE.successCount} 封
失败：0 封
总计：${STATE.tasks.length} 条`, { rpaStatus: 'success' });
  }

  async function startRun() {
    hidePageModal();
    hideStatusToast();
    hideDemoFocus();
    try {
      saveState();
      if (STATE.running) {
        logWarn('当前已有正在运行的任务。');
        return;
      }
      let cfg = waitConfig();
      validateConfig(cfg);
      cfg = prepareConfigForBatch(cfg);
      STATE.tasks = parseTasks(cfg.tasksRaw);
      STATE.currentIndex = 0;
      STATE.successCount = 0;
      STATE.stopRequested = false;
      STATE.paused = false;
      STATE.running = true;
      setStatus('运行中');
      setRpaStatus('running');
      setProgress(0, STATE.tasks.length);
      logInfo(`共解析到 ${STATE.tasks.length} 条任务。`);
      if (cfg.templateMode === 'variants' && cfg.variantRuntime) {
        logInfo(`随机抽取模式已锁定本批配置，共 ${cfg.variantRuntime.sequence.length} 组组合；当前起始序号：${STATE.variantCursor + 1}`);
      } else {
        logInfo('单模板模式已锁定本批配置。');
      }
      await runLoop(cfg);
    } catch (e) {
      STATE.running = false;
      STATE.paused = false;
      const msg = e.message || String(e);
      if (msg === '已停止。') {
        setStatus('已停止');
        setRpaStatus('stopped');
        logWarn('运行已停止。');
        return;
      }
      setStatus('报错停止');
      const errorCode = errorCodeFromMessage(msg);
      setRpaStatus('error', msg, errorCode);
      logError(`❌ ${msg}`);
      showPageModal('error', '运行错误', `${msg}

错误代码：${errorCode}`, { rpaStatus: 'error', errorCode });
    } finally {
      hideDemoFocus();
    }
  }

  function pauseRun() {

    if (!STATE.running) {
      logWarn('当前没有正在运行的任务。');
      return;
    }
    STATE.paused = !STATE.paused;
    setStatus(STATE.paused ? '已暂停' : '运行中');
    setRpaStatus(STATE.paused ? 'paused' : 'running');
    logInfo(STATE.paused ? '当前步骤完成后暂停。' : '继续运行。');
  }

  function stopRun() {
    STATE.stopRequested = true;
    STATE.paused = false;
    setStatus('停止中');
    setRpaStatus('stopping');
    hideStatusToast();
    logWarn('已请求停止。当前步骤结束后终止，不再进入下一封。');
  }

  function init() {

    injectStyle();
    createUI();
    bindUI();
    setPanelMinimized(false);
    loadState();
    syncIntervalInputs();
    saveState();
    setStatus('待机');
    setRpaStatus('idle');
    setProgress(0, 0);
    logInfo('面板已加载。v3.1.5 已更新：黄色等待提示改为显示 3 秒后再淡出；联系人之间的处理间隔新增 5 秒下限保护。');
  }

  init();
})();
