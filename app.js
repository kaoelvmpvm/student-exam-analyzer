(() => {
  'use strict';

  const APP_VERSION = '2.0.0-web';
  const WATERMARK = '板橋國中 高勳頊老師製作';
  const SHEET_RE = /^(\d{3})-([789])([AB])-0([123])$/;
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  const state = {
    workbook: null,
    groupKey: '',
    analyses: [],
    missing: [],
    detailId: '',
    sourceFile: ''
  };

  const el = (id) => document.getElementById(id);
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  // -------------------- UI --------------------

  const dropZone = el('drop-zone');
  const fileInput = el('file-input');
  const fileStatus = el('file-status');
  const querySection = el('query-section');
  const groupSelect = el('group-select');
  const queryInput = el('query-input');
  const analyzeBtn = el('analyze-btn');
  const clearBtn = el('clear-btn');
  const resultsSection = el('results-section');
  const resultsMeta = el('results-meta');
  const missingWarning = el('missing-warning');
  const overviewWrap = el('overview-wrap');
  const overviewBody = el('overview-body');
  const overviewFilter = el('overview-filter');
  const overviewSort = el('overview-sort');
  const detailSelectorWrap = el('detail-selector-wrap');
  const detailSelector = el('detail-selector');
  const studentDetail = el('student-detail');
  const batchSummary = el('batch-summary');
  const toast = el('toast');
  const printRoot = el('print-root');

  fileInput.addEventListener('change', () => {
    if (fileInput.files && fileInput.files[0]) loadWorkbook(fileInput.files[0]);
  });
  ['dragenter', 'dragover'].forEach(evt => dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  }));
  ['dragleave', 'drop'].forEach(evt => dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
  }));
  dropZone.addEventListener('drop', (e) => {
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) loadWorkbook(file);
  });
  groupSelect.addEventListener('change', () => {
    state.groupKey = groupSelect.value;
    clearResults();
  });
  queryInput.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') runAnalysis();
  });
  analyzeBtn.addEventListener('click', runAnalysis);
  clearBtn.addEventListener('click', () => {
    queryInput.value = '';
    queryInput.focus();
    clearResults();
  });
  detailSelector.addEventListener('change', () => {
    state.detailId = detailSelector.value;
    renderDetail();
  });
  overviewFilter.addEventListener('input', renderOverview);
  overviewSort.addEventListener('change', renderOverview);
  el('download-pdf-btn').addEventListener('click', () => exportCombinedPDF());
  el('download-zip-btn').addEventListener('click', () => exportIndividualZIP());
  el('print-btn').addEventListener('click', printReports);

  document.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-open-student]');
    if (!btn) return;
    const id = btn.getAttribute('data-open-student');
    if (!id) return;
    state.detailId = id;
    detailSelector.value = id;
    renderDetail();
    studentDetail.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  async function loadWorkbook(file) {
    if (!file || !/\.xlsx$/i.test(file.name)) {
      showToast('目前只支援 .xlsx 檔案。');
      return;
    }
    if (!window.JSZip) {
      showToast('必要元件載入失敗，請確認 vendor/jszip.min.js 存在。');
      return;
    }

    setBusy(true, `正在解析 ${file.name}…`);
    clearResults();
    try {
      const wb = await parseXLSX(file);
      state.workbook = wb;
      state.sourceFile = file.name;
      renderFileStatus(wb);
      populateGroups(wb);
      querySection.classList.add('enabled');
      groupSelect.disabled = false;
      queryInput.disabled = false;
      analyzeBtn.disabled = false;
      clearBtn.disabled = false;
      queryInput.focus();
      showToast(`已讀取 ${wb.groups.length} 個年級資料群組。`);
    } catch (err) {
      console.error(err);
      state.workbook = null;
      fileStatus.classList.remove('hidden');
      fileStatus.innerHTML = `<div class="status-title">讀取失敗</div><div class="status-meta">${escapeHTML(err.message || String(err))}</div>`;
      querySection.classList.remove('enabled');
      groupSelect.disabled = true;
      queryInput.disabled = true;
      analyzeBtn.disabled = true;
      clearBtn.disabled = true;
    } finally {
      setBusy(false);
    }
  }

  function renderFileStatus(wb) {
    const allExamCount = wb.groups.reduce((sum, g) => sum + g.exams.length, 0);
    const warnings = wb.warnings || [];
    fileStatus.classList.remove('hidden');
    fileStatus.innerHTML = `
      <div class="status-title">✓ 已讀取：${escapeHTML(wb.fileName)}</div>
      <div class="status-meta">偵測到 ${wb.groups.length} 個年級資料群組、共 ${allExamCount} 個正式段考分頁。公式映射會優先於 R 欄顯示標題。</div>
      ${warnings.length ? `<div class="status-warning">注意：${warnings.map(escapeHTML).join('；')}</div>` : ''}
    `;
  }

  function populateGroups(wb) {
    groupSelect.innerHTML = '';
    const groups = [...wb.groups].sort((a, b) => (Number(b.year) - Number(a.year)) || (Number(b.grade) - Number(a.grade)));
    groups.forEach(g => {
      const opt = document.createElement('option');
      opt.value = g.key;
      opt.textContent = `${g.label}｜${g.exams.length} 次段考｜${g.studentIds.length} 人`;
      groupSelect.appendChild(opt);
    });
    state.groupKey = groups[0]?.key || '';
    groupSelect.value = state.groupKey;
  }

  function runAnalysis() {
    if (!state.workbook) return showToast('請先匯入 Excel 成績檔。');
    const query = queryInput.value.trim();
    if (!query) return showToast('請輸入班級座號，例如 71501、715 或多人座號。');
    const group = getGroup(state.groupKey);
    if (!group) return showToast('找不到指定資料群組。');

    const { ids, missing } = resolveQuery(group, query);
    if (!ids.length) {
      state.analyses = [];
      state.missing = missing;
      showToast('沒有找到可分析的學生座號。');
      return;
    }

    state.analyses = ids.map(id => analyzeStudent(group, id));
    state.missing = missing;
    state.detailId = state.analyses[0].id;
    renderResults();
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderResults() {
    const n = state.analyses.length;
    const group = getGroup(state.groupKey);
    resultsSection.classList.remove('hidden');
    resultsMeta.textContent = `${group?.label || ''}｜${n} 位學生｜來源：${state.sourceFile}`;

    if (state.missing.length) {
      missingWarning.classList.remove('hidden');
      missingWarning.textContent = `以下輸入未找到資料：${state.missing.join('、')}`;
    } else {
      missingWarning.classList.add('hidden');
      missingWarning.textContent = '';
    }

    renderBatchSummary();
    renderDetailSelector();
    renderOverview();
    renderDetail();
  }

  function renderBatchSummary() {
    const list = state.analyses;
    if (!list.length) return batchSummary.classList.add('hidden');
    batchSummary.classList.remove('hidden');
    el('summary-count').textContent = list.length;
    el('summary-improved').textContent = list.filter(a => a.overallTrend.includes('進步')).length;
    el('summary-declined').textContent = list.filter(a => a.overallTrend.includes('退步')).length;
    el('summary-gap').textContent = list.filter(a => a.subjectGap >= 20).length;
  }

  function renderDetailSelector() {
    if (!state.analyses.length) return;
    detailSelector.innerHTML = state.analyses.map(a => `<option value="${escapeHTML(a.id)}">${escapeHTML(a.id)} ${escapeHTML(a.name || '（未命名）')}</option>`).join('');
    detailSelector.value = state.detailId;
    detailSelectorWrap.classList.toggle('hidden', state.analyses.length <= 1);
  }

  function renderOverview() {
    if (state.analyses.length <= 1) {
      overviewWrap.classList.add('hidden');
      return;
    }
    overviewWrap.classList.remove('hidden');
    const f = overviewFilter.value.trim().toLowerCase();
    let list = state.analyses.filter(a => !f || a.id.toLowerCase().includes(f) || (a.name || '').toLowerCase().includes(f));
    const sortMode = overviewSort.value;
    list = [...list].sort((a, b) => {
      if (sortMode === 'latest') {
        const ar = a.latestTotalOK ? a.latestTotalRank : 99999;
        const br = b.latestTotalOK ? b.latestTotalRank : 99999;
        return ar - br || Number(a.id) - Number(b.id);
      }
      if (sortMode === 'trend') {
        return b.overallDelta - a.overallDelta || Number(a.id) - Number(b.id);
      }
      return Number(a.id) - Number(b.id);
    });

    overviewBody.innerHTML = list.map(a => `
      <tr>
        <td><strong>${escapeHTML(a.id)}</strong></td>
        <td>${escapeHTML(a.name || '—')}</td>
        <td>${a.latestTotalOK ? `第 ${a.latestTotalRank} 名` : '—'}</td>
        <td>${a.bestTotalOK ? `第 ${a.bestTotalRank} 名` : '—'}</td>
        <td>${trendBadge(a.overallTrend)}</td>
        <td><span class="badge good">${escapeHTML(a.strongest || '—')}</span></td>
        <td><span class="badge bad">${escapeHTML(a.weakest || '—')}</span></td>
        <td>${a.subjectGap.toFixed(0)} 點</td>
        <td><button class="mini-btn" data-open-student="${escapeHTML(a.id)}">查看</button></td>
      </tr>
    `).join('');
  }

  function renderDetail() {
    const a = state.analyses.find(x => x.id === state.detailId) || state.analyses[0];
    if (!a) {
      studentDetail.innerHTML = '';
      return;
    }
    state.detailId = a.id;
    studentDetail.innerHTML = reportHTML(a, { forPrint: false });
  }

  function clearResults() {
    state.analyses = [];
    state.missing = [];
    state.detailId = '';
    resultsSection.classList.add('hidden');
    studentDetail.innerHTML = '';
    overviewBody.innerHTML = '';
  }

  function setBusy(busy, message = '') {
    document.body.style.cursor = busy ? 'progress' : '';
    if (busy && message) showToast(message, 60000);
    if (!busy && toast.dataset.long === '1') hideToast();
  }

  let toastTimer = null;
  function showToast(message, duration = 2600) {
    toast.textContent = message;
    toast.classList.remove('hidden');
    toast.dataset.long = duration > 10000 ? '1' : '0';
    clearTimeout(toastTimer);
    if (duration < 60000) toastTimer = setTimeout(hideToast, duration);
  }
  function hideToast() {
    toast.classList.add('hidden');
    toast.dataset.long = '0';
  }

  // -------------------- XLSX reader (browser / JSZip) --------------------

  async function parseXLSX(file) {
    const zip = await JSZip.loadAsync(file);
    const readText = async (name, required = true) => {
      const zf = zip.file(name);
      if (!zf) {
        if (required) throw new Error(`Excel 內缺少 ${name}`);
        return '';
      }
      return zf.async('string');
    };

    const workbookText = await readText('xl/workbook.xml');
    const relsText = await readText('xl/_rels/workbook.xml.rels');
    const wbDoc = parseXML(workbookText, 'workbook.xml');
    const relDoc = parseXML(relsText, 'workbook.xml.rels');

    const relMap = new Map();
    Array.from(relDoc.getElementsByTagName('Relationship')).forEach(rel => {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      if (id && target) relMap.set(id, normalizeXLPath(target));
    });

    const sharedStrings = [];
    const sharedText = await readText('xl/sharedStrings.xml', false);
    if (sharedText) {
      const sDoc = parseXML(sharedText, 'sharedStrings.xml');
      Array.from(sDoc.getElementsByTagName('si')).forEach(si => {
        const ts = Array.from(si.getElementsByTagName('t')).map(t => t.textContent || '');
        sharedStrings.push(ts.join(''));
      });
    }

    const groupsByKey = new Map();
    const warnings = [];
    const sheets = Array.from(wbDoc.getElementsByTagName('sheet'));

    for (const sheet of sheets) {
      const name = (sheet.getAttribute('name') || '').trim();
      const m = name.match(SHEET_RE);
      if (!m) continue;
      const rid = sheet.getAttribute('r:id') || sheet.getAttributeNS(REL_NS, 'id');
      const target = relMap.get(rid);
      if (!target) {
        warnings.push(`找不到工作表關聯：${name}`);
        continue;
      }
      try {
        const sheetText = await readText(target);
        const exam = parseExamSheet(name, sheetText, sharedStrings, m[1], m[2], m[3], m[4]);
        if (exam.mappingWarnings?.length) warnings.push(...exam.mappingWarnings.map(w => `${name}：${w}`));
        const key = `${m[1]}-${m[2]}`;
        if (!groupsByKey.has(key)) {
          groupsByKey.set(key, {
            key,
            label: `${m[1]}學年度 ${gradeChinese(m[2])}年級`,
            year: m[1],
            grade: m[2],
            exams: [],
            studentIds: [],
            classes: [],
            subjects: []
          });
        }
        groupsByKey.get(key).exams.push(exam);
      } catch (err) {
        warnings.push(`${name}：${err.message}`);
      }
    }

    if (!groupsByKey.size) throw new Error('找不到符合命名規則的正式段考分頁（例如 114-7A-01）。');

    const groups = Array.from(groupsByKey.values());
    groups.forEach(g => {
      g.exams.sort((a, b) => (a.term.localeCompare(b.term)) || (a.examNo - b.examNo));
      const ids = new Set();
      const classes = new Set();
      const subjects = [];
      const seenSubjects = new Set();
      g.exams.forEach(ex => {
        ex.students.forEach((_, id) => {
          ids.add(id);
          if (id.length >= 3) classes.add(id.slice(0, 3));
        });
        ex.subjects.forEach(s => {
          if (!seenSubjects.has(s.key)) {
            subjects.push(s.key);
            seenSubjects.add(s.key);
          }
        });
      });
      g.studentIds = Array.from(ids).sort((a, b) => Number(a) - Number(b));
      g.classes = Array.from(classes).sort();
      g.subjects = subjects;
      if (g.exams.length !== 6) warnings.push(`${g.label} 偵測到 ${g.exams.length} 次正式段考（完整學年一般為 6 次）。`);
    });

    return { fileName: file.name, groups, warnings, version: APP_VERSION };
  }

  function parseXML(text, name) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error(`${name} XML 解析失敗。`);
    return doc;
  }

  function normalizeXLPath(target) {
    target = String(target).replace(/\\/g, '/');
    if (target.startsWith('/')) return target.replace(/^\/+/, '');
    if (target.startsWith('xl/')) return cleanPath(target);
    return cleanPath(`xl/${target}`);
  }

  function cleanPath(p) {
    const out = [];
    p.split('/').forEach(part => {
      if (!part || part === '.') return;
      if (part === '..') out.pop();
      else out.push(part);
    });
    return out.join('/');
  }

  function parseExamSheet(name, sheetText, shared, year, grade, term, examNoText) {
    const doc = parseXML(sheetText, name);
    const rowEls = Array.from(doc.getElementsByTagName('row'));
    if (!rowEls.length) throw new Error('工作表沒有資料。');

    const rows = [];
    const formulas = [];
    rowEls.forEach(rowEl => {
      const row = {};
      const frow = {};
      Array.from(rowEl.getElementsByTagName('c')).forEach(c => {
        const ref = c.getAttribute('r') || '';
        const col = cellCol(ref);
        if (!col) return;
        row[col] = cellValue(c, shared);
        const f = c.getElementsByTagName('f')[0];
        if (f && (f.textContent || '').trim()) frow[col] = (f.textContent || '').trim();
      });
      rows.push(row);
      formulas.push(frow);
    });

    let headerIdx = -1;
    let header = null;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      if (containsValue(rows[i], '班級座號') && containsValue(rows[i], '姓名')) {
        headerIdx = i;
        header = rows[i];
        break;
      }
    }
    if (headerIdx < 0) throw new Error('找不到「班級座號／姓名」標題列。');

    const colByHead = {};
    Object.entries(header).forEach(([col, raw]) => {
      const h = String(raw || '').trim();
      if (!h) return;
      if (!colByHead[h] || excelColNumber(col) < excelColNumber(colByHead[h])) colByHead[h] = col;
    });

    const idCol = colByHead['班級座號'];
    const nameCol = colByHead['姓名'];
    if (!idCol || !nameCol) throw new Error('缺少必要欄位。');

    const rankScoreMap = inferRankScoreColumns(formulas, headerIdx);
    const subjects = detectSubjects(colByHead, rankScoreMap);
    if (subjects.length < 3) throw new Error('可辨識的單科欄位少於 3 科。');

    const students = new Map();
    const rankMax = {};
    const rankCount = {};
    let totalMax = 0;
    let totalCount = 0;
    const totalCol = colByHead['校排'];
    const classCol = colByHead['班排'];
    const avgCol = colByHead['平均'];

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const id = normalizeID(r[idCol]);
      if (id.length < 4 || !/^\d+$/.test(id)) continue;
      const rec = {
        id,
        name: String(r[nameCol] || '').trim(),
        scores: {}, scoreOK: {}, ranks: {}, rankOK: {},
        totalRank: 0, totalOK: false,
        classRank: 0, classOK: false,
        average: 0, averageOK: false
      };

      const total = parseRank(r[totalCol]);
      if (total.ok) {
        rec.totalRank = total.value; rec.totalOK = true;
        totalMax = Math.max(totalMax, total.value); totalCount++;
      }
      const cls = parseRank(r[classCol]);
      if (cls.ok) { rec.classRank = cls.value; rec.classOK = true; }
      const avg = parseNumber(r[avgCol]);
      if (avg.ok) { rec.average = avg.value; rec.averageOK = true; }

      subjects.forEach(s => {
        const score = parseNumber(r[s.scoreCol]);
        if (score.ok) { rec.scores[s.key] = score.value; rec.scoreOK[s.key] = true; }
        const rank = parseRank(r[s.rankCol]);
        if (rank.ok) {
          rec.ranks[s.key] = rank.value; rec.rankOK[s.key] = true;
          rankMax[s.key] = Math.max(rankMax[s.key] || 0, rank.value);
          rankCount[s.key] = (rankCount[s.key] || 0) + 1;
        }
      });

      const hasData = rec.name || rec.totalOK || rec.classOK || rec.averageOK || Object.keys(rec.scoreOK).length;
      if (hasData) students.set(id, rec);
    }

    if (!students.size) throw new Error('沒有辨識到學生資料。');
    totalMax = Math.max(totalMax, totalCount);
    Object.keys(rankMax).forEach(k => { rankMax[k] = Math.max(rankMax[k], rankCount[k] || 0); });

    const mappingWarnings = [];
    const mismatched = subjects.filter(s => s.rankHead && s.rankHead.endsWith('R') && s.rankHead[0] !== s.label[0]);
    if (mismatched.length) {
      mappingWarnings.push(`偵測到單科 R 欄標題與 RANK() 公式來源不一致（${mismatched.map(s => `${s.rankHead}→${s.label}`).join('、')}），已依公式實際引用科目修正。`);
    }

    const examNo = Number(examNoText);
    return {
      sheetName: name,
      year, grade, term, examNo,
      label: `${termLabel(term)}第${examChinese(examNo)}次`,
      shortLabel: `${grade}${term === 'A' ? '上' : '下'}${examNo}`,
      subjects, students, rankMax, totalMax, mappingWarnings
    };
  }

  function cellValue(cell, shared) {
    const type = cell.getAttribute('t') || '';
    const vEl = cell.getElementsByTagName('v')[0];
    if (type === 's') {
      const idx = Number((vEl?.textContent || '').trim());
      return Number.isInteger(idx) && idx >= 0 && idx < shared.length ? shared[idx] : '';
    }
    if (type === 'inlineStr') {
      return Array.from(cell.getElementsByTagName('t')).map(t => t.textContent || '').join('');
    }
    if (type === 'str') return vEl?.textContent || '';
    return vEl?.textContent || '';
  }

  function cellCol(ref) {
    const m = String(ref || '').match(/^[A-Za-z]+/);
    return m ? m[0].toUpperCase() : '';
  }
  function excelColNumber(col) {
    let n = 0;
    for (const ch of String(col).toUpperCase()) {
      if (ch < 'A' || ch > 'Z') break;
      n = n * 26 + ch.charCodeAt(0) - 64;
    }
    return n;
  }
  function containsValue(obj, value) {
    return Object.values(obj).some(x => String(x || '').trim() === value);
  }
  function normalizeID(v) {
    const s = String(v ?? '').trim();
    if (!s) return '';
    const n = Number(s);
    if (Number.isFinite(n) && Math.abs(n - Math.round(n)) < 1e-8) return String(Math.round(n));
    return s;
  }
  function parseNumber(v) {
    const s = String(v ?? '').trim();
    if (!s) return { ok: false, value: 0 };
    const n = Number(s);
    return Number.isFinite(n) ? { ok: true, value: n } : { ok: false, value: 0 };
  }
  function parseRank(v) {
    const p = parseNumber(v);
    if (!p.ok || p.value <= 0) return { ok: false, value: 0 };
    return { ok: true, value: Math.round(p.value) };
  }

  function inferRankScoreColumns(formulas, headerIdx) {
    const out = {};
    const re = /RANK(?:\.EQ)?\(\s*\$?([A-Z]+)\$?\d+\s*[,;]/i;
    for (let i = headerIdx + 1; i < Math.min(formulas.length, headerIdx + 40); i++) {
      Object.entries(formulas[i]).forEach(([rankCol, f]) => {
        if (out[rankCol]) return;
        const m = String(f).toUpperCase().match(re);
        if (m) out[rankCol] = m[1];
      });
    }
    return out;
  }

  function detectSubjects(colByHead, rankScoreMap) {
    const candidates = [
      ['國文', '國文', ['國文'], ['國R', '國文R']],
      ['英文', '英文', ['英文'], ['英R', '英文R']],
      ['數學', '數學', ['數學'], ['數R', '數學R']],
      ['社會', '社會', ['社會'], ['社R', '社會R']],
      ['生物', '生物', ['生物'], ['生R', '生物R', '自R']],
      ['理化', '理化', ['理化'], ['理R', '理化R', '自R']],
      ['自然', '自然', ['自然'], ['自R', '自然R']]
    ];
    const headerByCol = {};
    Object.entries(colByHead).forEach(([h, c]) => { headerByCol[c] = h; });
    const usedRank = new Set();
    const out = [];

    candidates.forEach(([key, label, scoreAliases, rankAliases]) => {
      const [scoreCol, scoreHead] = firstHeader(colByHead, scoreAliases);
      if (!scoreCol) return;
      let rankCol = '', rankHead = '';
      for (const [rCol, sCol] of Object.entries(rankScoreMap)) {
        if (sCol === scoreCol && !usedRank.has(rCol)) {
          rankCol = rCol; rankHead = headerByCol[rCol] || ''; break;
        }
      }
      if (!rankCol) [rankCol, rankHead] = firstHeader(colByHead, rankAliases);
      if (!rankCol) {
        const first = label[0];
        for (const [h, col] of Object.entries(colByHead)) {
          if (h.endsWith('R') && !usedRank.has(col) && h[0] === first) {
            rankCol = col; rankHead = h; break;
          }
        }
      }
      if (!rankCol || usedRank.has(rankCol)) return;
      usedRank.add(rankCol);
      out.push({ key, label, scoreCol, rankCol, scoreHead, rankHead });
    });
    return out;
  }

  function firstHeader(map, aliases) {
    for (const a of aliases) if (map[a]) return [map[a], a];
    return ['', ''];
  }

  // -------------------- Analysis --------------------

  function getGroup(key) {
    return state.workbook?.groups.find(g => g.key === key) || null;
  }

  function resolveQuery(group, query) {
    const tokens = String(query).split(/[\s,，;；、|]+/).map(x => x.trim()).filter(Boolean);
    const studentSet = new Set(group.studentIds);
    const seen = new Set();
    const ids = [];
    const missing = [];
    const add = id => { if (!seen.has(id)) { seen.add(id); ids.push(id); } };

    tokens.forEach(tok => {
      if (/^\d{3}$/.test(tok)) {
        const found = group.studentIds.filter(id => id.startsWith(tok));
        if (found.length) found.forEach(add); else missing.push(`${tok}（班級）`);
        return;
      }
      if (/^\d+-\d+$/.test(tok)) {
        let [a, b] = tok.split('-').map(Number);
        if (a > b) [a, b] = [b, a];
        if (b - a <= 100) {
          let found = 0;
          for (let n = a; n <= b; n++) {
            const id = String(n);
            if (studentSet.has(id)) { add(id); found++; }
          }
          if (!found) missing.push(tok);
          return;
        }
      }
      const id = normalizeID(tok);
      if (studentSet.has(id)) add(id); else missing.push(tok);
    });
    return { ids, missing };
  }

  function analyzeStudent(group, id) {
    const a = {
      id, name: '', groupLabel: group.label,
      exams: [], subjects: [], strongest: '', weakest: '',
      overallTrend: '資料不足', overallDelta: 0, overallVolatility: '資料不足',
      latestTotalRank: 0, latestTotalOK: false,
      bestTotalRank: 0, bestTotalOK: false,
      subjectGap: 0, summary: [], warnings: []
    };
    const nameCount = new Map();
    const totalSeries = [];
    const subjectSeries = {};
    const subjectRanks = {};
    const latestSubjectRank = {};

    group.exams.forEach(ex => {
      const rec = ex.students.get(id);
      const er = {
        label: ex.label, shortLabel: ex.shortLabel, sheetName: ex.sheetName,
        average: null, classRank: null, totalRank: null, totalMax: ex.totalMax, totalIndex: null,
        ranks: {}, rankMax: {}, scores: {}
      };
      if (!rec) {
        a.warnings.push(`${ex.label}找不到資料`);
        ex.subjects.forEach(s => { er.ranks[s.key] = null; er.rankMax[s.key] = ex.rankMax[s.key] || 0; er.scores[s.key] = null; });
        a.exams.push(er);
        return;
      }
      if (rec.name) nameCount.set(rec.name, (nameCount.get(rec.name) || 0) + 1);
      if (rec.averageOK) er.average = rec.average;
      if (rec.classOK) er.classRank = rec.classRank;
      if (rec.totalOK) {
        er.totalRank = rec.totalRank;
        const idx = rankIndex(rec.totalRank, ex.totalMax);
        if (idx !== null) { er.totalIndex = idx; totalSeries.push(idx); }
        if (!a.bestTotalOK || rec.totalRank < a.bestTotalRank) { a.bestTotalRank = rec.totalRank; a.bestTotalOK = true; }
        a.latestTotalRank = rec.totalRank; a.latestTotalOK = true;
      }
      ex.subjects.forEach(s => {
        er.rankMax[s.key] = ex.rankMax[s.key] || 0;
        if (rec.rankOK[s.key]) {
          const v = rec.ranks[s.key];
          er.ranks[s.key] = v;
          latestSubjectRank[s.key] = v;
          (subjectRanks[s.key] ||= []).push(v);
          const idx = rankIndex(v, ex.rankMax[s.key]);
          if (idx !== null) (subjectSeries[s.key] ||= []).push(idx);
        } else er.ranks[s.key] = null;
        er.scores[s.key] = rec.scoreOK[s.key] ? rec.scores[s.key] : null;
      });
      a.exams.push(er);
    });

    let bestName = '', bestCount = -1;
    nameCount.forEach((count, name) => { if (count > bestCount) { bestName = name; bestCount = count; } });
    a.name = bestName;
    if (nameCount.size > 1) a.warnings.push('不同段考中的姓名不完全一致，請核對座號。');

    [a.overallTrend, a.overallDelta] = trendOf(totalSeries);
    a.overallVolatility = volatilityLabel(totalSeries);

    group.subjects.forEach(key => {
      const xs = subjectSeries[key] || [];
      if (!xs.length) return;
      const [trend, delta] = trendOf(xs);
      a.subjects.push({
        key, label: key,
        avgIndex: mean(xs),
        avgRank: mean(subjectRanks[key] || []),
        latestRank: latestSubjectRank[key] || 0,
        latestOK: !!latestSubjectRank[key],
        trend, delta,
        volatility: volatilityLabel(xs),
        count: xs.length
      });
    });

    if (a.subjects.length) {
      let best = a.subjects[0], worst = a.subjects[0];
      a.subjects.slice(1).forEach(s => {
        if (s.avgIndex > best.avgIndex) best = s;
        if (s.avgIndex < worst.avgIndex) worst = s;
      });
      a.strongest = best.key;
      a.weakest = worst.key;
      a.subjectGap = best.avgIndex - worst.avgIndex;
    }
    a.summary = buildSummary(a);
    return a;
  }

  function rankIndex(rank, max) {
    if (!(rank > 0) || !(max > 0)) return null;
    if (max === 1) return 100;
    return clamp(100 * (max - rank) / (max - 1), 0, 100);
  }
  function mean(xs) { return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0; }
  function stddev(xs) {
    if (xs.length < 2) return 0;
    const m = mean(xs);
    return Math.sqrt(xs.reduce((s, x) => s + (x - m) ** 2, 0) / xs.length);
  }
  function trendOf(xs) {
    if (xs.length < 2) return ['資料不足', 0];
    const split = Math.max(1, Math.floor(xs.length / 2));
    const d = mean(xs.slice(split)) - mean(xs.slice(0, split));
    if (d >= 10) return ['明顯進步', d];
    if (d >= 4) return ['穩定進步', d];
    if (d <= -10) return ['明顯退步', d];
    if (d <= -4) return ['略有退步', d];
    return ['大致穩定', d];
  }
  function volatilityLabel(xs) {
    if (xs.length < 2) return '資料不足';
    const sd = stddev(xs);
    if (sd >= 12) return '波動較大';
    if (sd >= 6) return '有些波動';
    return '表現穩定';
  }
  function topPct(rank, max) {
    if (!(rank > 0) || !(max > 0)) return '';
    return `前${Math.max(1, Math.round(100 * rank / max))}%`;
  }
  function buildSummary(a) {
    const out = [];
    const valid = a.exams.filter(e => e.totalRank != null);
    if (valid.length >= 2) {
      const first = valid[0], last = valid[valid.length - 1];
      out.push(`整體校排由第 ${first.totalRank} 名（約${topPct(first.totalRank, first.totalMax)}）到第 ${last.totalRank} 名（約${topPct(last.totalRank, last.totalMax)}），趨勢判定為「${a.overallTrend}」。`);
    } else out.push('整體校排資料不足，無法完整判讀六次趨勢。');

    const best = a.subjects.find(s => s.key === a.strongest);
    const worst = a.subjects.find(s => s.key === a.weakest);
    if (best) out.push(`相對優勢科目為${best.label}：六次平均相對能力指數 ${best.avgIndex.toFixed(0)}${best.latestOK ? `，最近一次單科校排第 ${best.latestRank} 名` : ''}。`);
    if (worst) out.push(`相對較弱科目為${worst.label}：六次平均相對能力指數 ${worst.avgIndex.toFixed(0)}${worst.latestOK ? `，最近一次單科校排第 ${worst.latestRank} 名` : ''}；建議搭配原始分數與題型再判斷原因。`);
    if (a.subjectGap < 8) out.push('五科相對表現差距小，整體較為均衡。');
    else if (a.subjectGap >= 20) out.push('科目間差距較明顯，後續複習可優先處理相對弱科，同時維持優勢科。');
    return out;
  }

  // -------------------- HTML report + SVG charts --------------------

  function reportHTML(a, { forPrint = false } = {}) {
    const subjectHeaders = a.subjects.map(s => s.key);
    const rankRows = a.exams.map(e => `
      <tr>
        <td>${escapeHTML(e.shortLabel)}</td>
        ${subjectHeaders.map(k => `<td>${e.ranks[k] ?? '—'}</td>`).join('')}
        <td><strong>${e.totalRank ?? '—'}</strong></td>
        <td>${e.classRank ?? '—'}</td>
        <td>${e.average != null ? Number(e.average).toFixed(1) : '—'}</td>
      </tr>
    `).join('');
    const subjectRows = a.subjects.map(s => {
      const cls = s.key === a.strongest ? 'good' : s.key === a.weakest ? 'bad' : '';
      const marker = s.key === a.strongest ? '▲ ' : s.key === a.weakest ? '▼ ' : '';
      return `
        <tr>
          <td class="subject-name ${cls}">${marker}${escapeHTML(s.label)}</td>
          <td>${s.avgIndex.toFixed(0)}</td>
          <td>${s.avgRank.toFixed(1)}</td>
          <td>${s.latestOK ? s.latestRank : '—'}</td>
          <td class="${trendClass(s.trend)}">${escapeHTML(s.trend)}</td>
          <td>${escapeHTML(s.volatility)}</td>
        </tr>`;
    }).join('');

    return `
      <article class="student-report${forPrint ? ' print-report' : ''}">
        <div class="report-header">
          <div>
            <p class="eyebrow">INDIVIDUAL PERFORMANCE REPORT</p>
            <h3>${escapeHTML(a.name || '學生')}｜六次段考學習表現</h3>
            <p>${escapeHTML(a.groupLabel)}</p>
          </div>
          <div class="student-id-tag">${escapeHTML(a.id)}</div>
        </div>
        <div class="report-body">
          <div class="kpi-grid">
            <div class="kpi"><span>最近一次總校排</span><strong>${a.latestTotalOK ? `第 ${a.latestTotalRank} 名` : '資料不足'}</strong></div>
            <div class="kpi"><span>六次最佳總校排</span><strong>${a.bestTotalOK ? `第 ${a.bestTotalRank} 名` : '資料不足'}</strong></div>
            <div class="kpi"><span>整體趨勢</span><strong class="${trendClass(a.overallTrend)}">${escapeHTML(a.overallTrend)}</strong></div>
            <div class="kpi good"><span>優勢 / 相對弱勢</span><strong><span class="subject-name good">${escapeHTML(a.strongest || '—')}</span> / <span class="subject-name bad">${escapeHTML(a.weakest || '—')}</span></strong></div>
          </div>

          <ul class="summary-lines">${a.summary.map(x => `<li>${escapeHTML(x)}</li>`).join('')}</ul>

          <div class="chart-grid">
            <div class="chart-card">
              <h4>總校排趨勢</h4>
              <p>圖線往上代表相對排名改善；點位標示當次實際校排。</p>
              ${trendSVG(a)}
            </div>
            <div class="chart-card">
              <h4>五科相對能力雷達圖</h4>
              <p>以單科校排換算為 0–100 相對能力指數後取六次平均。</p>
              ${radarSVG(a)}
            </div>
          </div>

          <section class="report-section">
            <h4>六次段考排名紀錄</h4>
            <div class="report-table-wrap">
              <table class="report-table">
                <thead><tr><th>段考</th>${subjectHeaders.map(k => `<th>${escapeHTML(shortSubject(k))}R</th>`).join('')}<th>總校排</th><th>班排</th><th>平均</th></tr></thead>
                <tbody>${rankRows}</tbody>
              </table>
            </div>
          </section>

          <section class="report-section">
            <h4>科目能力與趨勢</h4>
            <div class="report-table-wrap">
              <table class="report-table">
                <thead><tr><th>科目</th><th>六次平均能力</th><th>六次平均校排</th><th>最近單科校排</th><th>趨勢</th><th>穩定度</th></tr></thead>
                <tbody>${subjectRows}</tbody>
              </table>
            </div>
          </section>

          ${a.warnings.length ? `<div class="report-note">資料提醒：${a.warnings.map(escapeHTML).join('；')}</div>` : ''}
          <div class="report-note">判讀原則：雷達圖反映「相對排名位置」，不是原始分數。優勢科以六次平均相對能力最高者標示，弱勢科以最低者標示；若科目差距很小，應視為整體均衡而非能力優劣的絕對判定。</div>
          <div class="report-watermark">${WATERMARK}</div>
        </div>
      </article>
    `;
  }

  function trendSVG(a) {
    const W = 620, H = 320, L = 56, R = 24, T = 30, B = 58;
    const plotW = W - L - R, plotH = H - T - B;
    const xs = a.exams.map((_, i) => L + (a.exams.length === 1 ? plotW / 2 : i * plotW / (a.exams.length - 1)));
    const y = v => T + (100 - v) * plotH / 100;
    let grid = '';
    [0, 25, 50, 75, 100].forEach(v => {
      const yy = y(v);
      grid += `<line x1="${L}" y1="${yy}" x2="${W - R}" y2="${yy}" stroke="#dbe2e7" stroke-width="1"/>`;
      grid += `<text x="${L - 8}" y="${yy + 4}" text-anchor="end" font-size="11" fill="#7a8791">${v}</text>`;
    });
    let path = '';
    let open = false;
    a.exams.forEach((e, i) => {
      if (e.totalIndex == null) { open = false; return; }
      path += `${open ? ' L' : ' M'} ${xs[i]} ${y(e.totalIndex)}`;
      open = true;
    });
    const points = a.exams.map((e, i) => {
      const label = `<text x="${xs[i]}" y="${H - 27}" text-anchor="middle" font-size="11" fill="#596772">${escapeSVG(e.shortLabel)}</text>`;
      if (e.totalIndex == null) return label;
      return `${label}<circle cx="${xs[i]}" cy="${y(e.totalIndex)}" r="5" fill="#214e67"/><text x="${xs[i]}" y="${y(e.totalIndex) - 10}" text-anchor="middle" font-size="11" font-weight="700" fill="#1e2b35">第${e.totalRank}</text>`;
    }).join('');
    return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="總校排趨勢圖">
      ${grid}
      <text x="16" y="${T + plotH / 2}" transform="rotate(-90 16 ${T + plotH / 2})" text-anchor="middle" font-size="11" fill="#7a8791">相對能力指數</text>
      <path d="${path}" fill="none" stroke="#214e67" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
      ${points}
    </svg>`;
  }

  function radarSVG(a) {
    const subjects = a.subjects;
    if (subjects.length < 3) return '<div class="report-note">可用科目少於 3 科，無法繪製雷達圖。</div>';
    const W = 500, H = 330, cx = 250, cy = 166, radius = 112;
    const n = subjects.length;
    const pt = (i, ratio) => {
      const ang = -Math.PI / 2 + i * 2 * Math.PI / n;
      return [cx + Math.cos(ang) * radius * ratio, cy + Math.sin(ang) * radius * ratio];
    };
    const polygon = ratio => subjects.map((_, i) => pt(i, ratio).join(',')).join(' ');
    let rings = '';
    [0.25, 0.5, 0.75, 1].forEach(r => rings += `<polygon points="${polygon(r)}" fill="none" stroke="#dce3e8" stroke-width="1"/>`);
    let axes = '';
    subjects.forEach((s, i) => {
      const [x, y] = pt(i, 1);
      const [lx, ly] = pt(i, 1.23);
      const clsColor = s.key === a.strongest ? '#1f7a4d' : s.key === a.weakest ? '#b64040' : '#43515c';
      axes += `<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#dce3e8"/><text x="${lx}" y="${ly}" text-anchor="middle" dominant-baseline="middle" font-size="13" font-weight="700" fill="${clsColor}">${escapeSVG(s.label)} ${s.avgIndex.toFixed(0)}</text>`;
    });
    const dataPts = subjects.map((s, i) => pt(i, clamp(s.avgIndex / 100, 0, 1)).join(',')).join(' ');
    const dots = subjects.map((s, i) => {
      const [x, y] = pt(i, clamp(s.avgIndex / 100, 0, 1));
      const c = s.key === a.strongest ? '#1f7a4d' : s.key === a.weakest ? '#b64040' : '#214e67';
      return `<circle cx="${x}" cy="${y}" r="5" fill="${c}" stroke="#fff" stroke-width="2"/>`;
    }).join('');
    return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="五科相對能力雷達圖">
      ${rings}${axes}
      <polygon points="${dataPts}" fill="rgba(33,78,103,.18)" stroke="#214e67" stroke-width="3" stroke-linejoin="round"/>
      ${dots}
    </svg>`;
  }

  function trendBadge(t) {
    if (t.includes('進步')) return `<span class="badge good">↑ ${escapeHTML(t)}</span>`;
    if (t.includes('退步')) return `<span class="badge bad">↓ ${escapeHTML(t)}</span>`;
    return `<span class="badge neutral">→ ${escapeHTML(t)}</span>`;
  }
  function trendClass(t) {
    if (t.includes('進步')) return 'trend-up';
    if (t.includes('退步')) return 'trend-down';
    return 'trend-flat';
  }
  function shortSubject(k) {
    return ({ 國文: '國', 英文: '英', 數學: '數', 社會: '社', 生物: '生', 理化: '理', 自然: '自' })[k] || k;
  }

  // -------------------- Print --------------------

  function printReports() {
    if (!state.analyses.length) return showToast('目前沒有可列印的分析結果。');
    printRoot.innerHTML = state.analyses.map(a => `<section class="print-page">${reportHTML(a, { forPrint: true })}</section>`).join('');
    requestAnimationFrame(() => setTimeout(() => window.print(), 80));
  }

  // -------------------- Direct PDF (canvas -> JPEG -> minimal PDF) --------------------

  async function exportCombinedPDF() {
    if (!state.analyses.length) return showToast('目前沒有可輸出的分析結果。');
    const useColor = el('pdf-color-toggle').checked;
    try {
      setBusy(true, `正在產生 ${state.analyses.length} 頁 PDF…`);
      const blob = await createPDFBlob(state.analyses, useColor, (i, n) => showToast(`正在產生 PDF：${i}/${n}`, 60000));
      downloadBlob(blob, `學生六次段考分析_${dateStamp()}.pdf`);
      showToast('PDF 已產生。');
    } catch (err) {
      console.error(err);
      showToast(`PDF 產生失敗：${err.message}`);
    } finally { setBusy(false); }
  }

  async function exportIndividualZIP() {
    if (!state.analyses.length) return showToast('目前沒有可輸出的分析結果。');
    const useColor = el('pdf-color-toggle').checked;
    try {
      setBusy(true, `正在建立 ${state.analyses.length} 份個別 PDF…`);
      const zip = new JSZip();
      for (let i = 0; i < state.analyses.length; i++) {
        const a = state.analyses[i];
        showToast(`建立個別 PDF：${i + 1}/${state.analyses.length}`, 60000);
        const blob = await createPDFBlob([a], useColor);
        zip.file(safeFileName(`${a.id}_${a.name || '學生'}_段考分析.pdf`), await blob.arrayBuffer());
        await idleYield();
      }
      const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      downloadBlob(zipBlob, `學生個別PDF_${dateStamp()}.zip`);
      showToast('個別 PDF ZIP 已產生。');
    } catch (err) {
      console.error(err);
      showToast(`ZIP 產生失敗：${err.message}`);
    } finally { setBusy(false); }
  }

  async function createPDFBlob(analyses, useColor, progress) {
    if (document.fonts?.ready) await document.fonts.ready;
    const pages = [];
    for (let i = 0; i < analyses.length; i++) {
      if (progress) progress(i + 1, analyses.length);
      const canvas = renderStudentCanvas(analyses[i], useColor);
      const jpeg = dataURLToBytes(canvas.toDataURL('image/jpeg', 0.92));
      pages.push({ jpeg, width: canvas.width, height: canvas.height });
      await idleYield();
    }
    const bytes = buildImagePDF(pages);
    return new Blob([bytes], { type: 'application/pdf' });
  }

  function renderStudentCanvas(a, useColor) {
    const W = 1240, H = 1754;
    const canvas = document.createElement('canvas');
    canvas.width = W; canvas.height = H;
    const c = canvas.getContext('2d');
    const C = useColor ? {
      ink: '#17212b', muted: '#65717d', line: '#d3dae0', soft: '#f5f7f8', primary: '#214e67',
      good: '#1f7a4d', bad: '#b64040', goodSoft: '#edf7f1', badSoft: '#fbefef'
    } : {
      ink: '#111111', muted: '#555555', line: '#bdbdbd', soft: '#f3f3f3', primary: '#222222',
      good: '#111111', bad: '#555555', goodSoft: '#f5f5f5', badSoft: '#eeeeee'
    };
    c.fillStyle = '#ffffff'; c.fillRect(0, 0, W, H);
    c.textBaseline = 'alphabetic';
    const font = (size, weight = 400) => `${weight} ${size}px "Microsoft JhengHei", "Noto Sans TC", "PingFang TC", Arial, sans-serif`;
    const margin = 68;
    let y = 66;

    // Header
    c.fillStyle = C.muted; c.font = font(17, 700); c.fillText('INDIVIDUAL PERFORMANCE REPORT', margin, y);
    y += 44;
    c.fillStyle = C.ink; c.font = font(34, 800); c.fillText(`${a.name || '學生'}｜六次段考學習表現`, margin, y);
    c.font = font(20, 400); c.fillStyle = C.muted; c.fillText(a.groupLabel, margin, y + 31);
    roundRect(c, W - margin - 150, 66, 150, 50, 12); c.fillStyle = C.primary; c.fill();
    c.fillStyle = '#fff'; c.font = font(22, 800); c.textAlign = 'center'; c.fillText(a.id, W - margin - 75, 99); c.textAlign = 'left';
    y += 70;
    c.strokeStyle = C.line; c.lineWidth = 2; c.beginPath(); c.moveTo(margin, y); c.lineTo(W - margin, y); c.stroke();
    y += 24;

    // KPI row
    const kpis = [
      ['最近一次總校排', a.latestTotalOK ? `第 ${a.latestTotalRank} 名` : '資料不足', ''],
      ['六次最佳總校排', a.bestTotalOK ? `第 ${a.bestTotalRank} 名` : '資料不足', ''],
      ['整體趨勢', a.overallTrend, a.overallTrend.includes('進步') ? 'good' : a.overallTrend.includes('退步') ? 'bad' : ''],
      ['優勢 / 相對弱勢', `${a.strongest || '—'} / ${a.weakest || '—'}`, 'good']
    ];
    const gap = 14, boxW = (W - 2 * margin - 3 * gap) / 4, boxH = 92;
    kpis.forEach((k, i) => {
      const x = margin + i * (boxW + gap);
      roundRect(c, x, y, boxW, boxH, 12);
      c.fillStyle = k[2] === 'good' ? C.goodSoft : k[2] === 'bad' ? C.badSoft : '#fff'; c.fill();
      c.strokeStyle = C.line; c.stroke();
      c.fillStyle = C.muted; c.font = font(15, 400); c.fillText(k[0], x + 16, y + 27);
      c.fillStyle = k[2] === 'good' ? C.good : k[2] === 'bad' ? C.bad : C.ink; c.font = font(23, 800); c.fillText(k[1], x + 16, y + 62);
    });
    y += boxH + 18;

    // Summary
    c.font = font(16, 400);
    const summaryLines = [];
    a.summary.slice(0, 4).forEach(line => {
      wrapText(c, `• ${line}`, W - 2 * margin - 36).slice(0, 2).forEach(t => summaryLines.push(t));
    });
    const summaryH = Math.max(110, 34 + summaryLines.length * 22);
    roundRect(c, margin, y, W - 2 * margin, summaryH, 12); c.fillStyle = C.soft; c.fill();
    c.fillStyle = C.ink; c.font = font(16, 400);
    let sy = y + 29;
    summaryLines.forEach(t => { c.fillText(t, margin + 18, sy); sy += 22; });
    y += summaryH + 20;

    // Charts titles
    const chartGap = 20, chartW = (W - 2 * margin - chartGap) / 2, chartH = 340;
    drawPanel(c, margin, y, chartW, chartH, C);
    drawPanel(c, margin + chartW + chartGap, y, chartW, chartH, C);
    c.fillStyle = C.ink; c.font = font(19, 800); c.fillText('總校排趨勢', margin + 18, y + 31);
    c.fillText('五科相對能力雷達圖', margin + chartW + chartGap + 18, y + 31);
    c.fillStyle = C.muted; c.font = font(13, 400);
    c.fillText('往上代表相對排名改善', margin + 18, y + 53);
    c.fillText('六次單科校排換算 0–100 後平均', margin + chartW + chartGap + 18, y + 53);
    drawTrendCanvas(c, a, margin + 18, y + 67, chartW - 36, chartH - 83, C, font);
    drawRadarCanvas(c, a, margin + chartW + chartGap + 18, y + 67, chartW - 36, chartH - 83, C, font, useColor);
    y += chartH + 24;

    // Rankings table
    c.fillStyle = C.ink; c.font = font(20, 800); c.fillText('六次段考排名紀錄', margin, y + 20); y += 34;
    const subj = a.subjects.map(s => s.key);
    const rankCols = ['段考', ...subj.map(shortSubject), '總校排', '班排', '平均'];
    const rankData = a.exams.map(e => [e.shortLabel, ...subj.map(k => e.ranks[k] ?? '—'), e.totalRank ?? '—', e.classRank ?? '—', e.average != null ? Number(e.average).toFixed(1) : '—']);
    y = drawCanvasTable(c, rankCols, rankData, margin, y, W - 2 * margin, 46, C, font, { headerFont: 15, cellFont: 15 });
    y += 22;

    // Subject table
    c.fillStyle = C.ink; c.font = font(20, 800); c.fillText('科目能力與趨勢', margin, y + 20); y += 34;
    const subjectData = a.subjects.map(s => [
      `${s.key === a.strongest ? '▲ ' : s.key === a.weakest ? '▼ ' : ''}${s.label}`,
      s.avgIndex.toFixed(0), s.avgRank.toFixed(1), s.latestOK ? s.latestRank : '—', s.trend, s.volatility
    ]);
    y = drawCanvasTable(c, ['科目', '平均能力', '平均校排', '最近校排', '趨勢', '穩定度'], subjectData, margin, y, W - 2 * margin, 43, C, font, {
      headerFont: 14, cellFont: 14,
      cellColor: (row, col, val) => col === 0 && String(val).startsWith('▲') ? C.good : col === 0 && String(val).startsWith('▼') ? C.bad : C.ink
    });

    // Notes + watermark
    const noteY = Math.min(H - 150, y + 22);
    c.fillStyle = C.muted; c.font = font(13, 400);
    const note = '判讀原則：雷達圖呈現相對排名位置，不等同原始分數；優弱科為六次平均相對位置的比較，應配合原始分數、題型與學習歷程解讀。';
    wrapText(c, note, W - 2 * margin).slice(0, 2).forEach((line, i) => c.fillText(line, margin, noteY + i * 20));
    c.strokeStyle = C.line; c.beginPath(); c.moveTo(margin, H - 82); c.lineTo(W - margin, H - 82); c.stroke();
    c.fillStyle = C.muted; c.font = font(14, 500); c.textAlign = 'center'; c.fillText(WATERMARK, W / 2, H - 45); c.textAlign = 'left';
    return canvas;
  }

  function drawPanel(c, x, y, w, h, C) {
    roundRect(c, x, y, w, h, 14); c.fillStyle = '#fff'; c.fill(); c.strokeStyle = C.line; c.lineWidth = 1.5; c.stroke();
  }

  function drawTrendCanvas(c, a, x, y, w, h, C, font) {
    const L = 42, R = 12, T = 24, B = 42;
    const pw = w - L - R, ph = h - T - B;
    const yy = v => y + T + (100 - v) * ph / 100;
    c.font = font(11, 400); c.fillStyle = C.muted; c.textAlign = 'right';
    [0,25,50,75,100].forEach(v => {
      const gy = yy(v); c.strokeStyle = C.line; c.lineWidth = 1; c.beginPath(); c.moveTo(x + L, gy); c.lineTo(x + w - R, gy); c.stroke(); c.fillText(String(v), x + L - 7, gy + 4);
    });
    const px = i => x + L + (a.exams.length <= 1 ? pw/2 : i * pw / (a.exams.length - 1));
    c.strokeStyle = C.primary; c.lineWidth = 3; c.beginPath(); let open = false;
    a.exams.forEach((e, i) => {
      if (e.totalIndex == null) { open = false; return; }
      const xx = px(i), yv = yy(e.totalIndex);
      if (!open) c.moveTo(xx, yv); else c.lineTo(xx, yv); open = true;
    }); c.stroke();
    c.textAlign = 'center';
    a.exams.forEach((e, i) => {
      const xx = px(i);
      c.fillStyle = C.muted; c.font = font(11, 400); c.fillText(e.shortLabel, xx, y + h - 12);
      if (e.totalIndex == null) return;
      const yv = yy(e.totalIndex); c.fillStyle = C.primary; c.beginPath(); c.arc(xx, yv, 5, 0, Math.PI*2); c.fill();
      c.fillStyle = C.ink; c.font = font(11, 700); c.fillText(`第${e.totalRank}`, xx, yv - 10);
    });
    c.textAlign = 'left';
  }

  function drawRadarCanvas(c, a, x, y, w, h, C, font, useColor) {
    const s = a.subjects;
    if (s.length < 3) { c.fillStyle = C.muted; c.font = font(14); c.fillText('可用科目少於 3 科', x + 20, y + 50); return; }
    const cx = x + w/2, cy = y + h/2 + 5, r = Math.min(w, h) * .34, n = s.length;
    const point = (i, ratio) => {
      const ang = -Math.PI/2 + i * Math.PI*2/n;
      return [cx + Math.cos(ang)*r*ratio, cy + Math.sin(ang)*r*ratio];
    };
    c.strokeStyle = C.line; c.lineWidth = 1;
    [0.25,0.5,0.75,1].forEach(rr => {
      c.beginPath(); s.forEach((_, i) => { const [px,py]=point(i,rr); i ? c.lineTo(px,py) : c.moveTo(px,py); }); c.closePath(); c.stroke();
    });
    s.forEach((sub, i) => { const [px,py] = point(i,1); c.beginPath(); c.moveTo(cx,cy); c.lineTo(px,py); c.stroke(); });
    c.beginPath();
    s.forEach((sub, i) => { const [px,py] = point(i, clamp(sub.avgIndex/100,0,1)); i ? c.lineTo(px,py) : c.moveTo(px,py); }); c.closePath();
    c.fillStyle = useColor ? 'rgba(33,78,103,.16)' : 'rgba(0,0,0,.08)'; c.fill(); c.strokeStyle = C.primary; c.lineWidth = 3; c.stroke();
    c.textAlign = 'center'; c.textBaseline = 'middle';
    s.forEach((sub, i) => {
      const [px,py] = point(i, clamp(sub.avgIndex/100,0,1));
      c.fillStyle = sub.key === a.strongest ? C.good : sub.key === a.weakest ? C.bad : C.primary; c.beginPath(); c.arc(px,py,5,0,Math.PI*2); c.fill();
      const [lx,ly] = point(i,1.2); c.fillStyle = sub.key === a.strongest ? C.good : sub.key === a.weakest ? C.bad : C.ink; c.font = font(12, 700);
      const prefix = sub.key === a.strongest ? '▲' : sub.key === a.weakest ? '▼' : '';
      c.fillText(`${prefix}${sub.label} ${sub.avgIndex.toFixed(0)}`, lx, ly);
    });
    c.textAlign = 'left'; c.textBaseline = 'alphabetic';
  }

  function drawCanvasTable(c, headers, rows, x, y, w, rowH, C, font, opts = {}) {
    const cols = headers.length;
    const cw = w / cols;
    c.fillStyle = C.soft; c.fillRect(x, y, w, rowH);
    c.strokeStyle = C.line; c.lineWidth = 1;
    c.font = font(opts.headerFont || 14, 700); c.textAlign = 'center'; c.textBaseline = 'middle'; c.fillStyle = C.ink;
    headers.forEach((h, i) => c.fillText(String(h), x + cw*(i+.5), y + rowH/2));
    for (let r = 0; r < rows.length; r++) {
      const yy = y + rowH*(r+1);
      c.fillStyle = '#fff'; c.fillRect(x, yy, w, rowH);
      c.font = font(opts.cellFont || 14, 400);
      rows[r].forEach((v, col) => {
        c.fillStyle = opts.cellColor ? opts.cellColor(r, col, v) : C.ink;
        c.fillText(String(v), x + cw*(col+.5), yy + rowH/2);
      });
    }
    const totalH = rowH * (rows.length + 1);
    c.strokeStyle = C.line; c.strokeRect(x, y, w, totalH);
    for (let i=1;i<cols;i++) { c.beginPath(); c.moveTo(x+cw*i,y); c.lineTo(x+cw*i,y+totalH); c.stroke(); }
    for (let r=1;r<=rows.length;r++) { c.beginPath(); c.moveTo(x,y+rowH*r); c.lineTo(x+w,y+rowH*r); c.stroke(); }
    c.textAlign = 'left'; c.textBaseline = 'alphabetic';
    return y + totalH;
  }

  function buildImagePDF(pages) {
    const PAGE_W = 595.28, PAGE_H = 841.89;
    const objectCount = 2 + pages.length * 3;
    const objects = new Array(objectCount + 1);
    objects[1] = asciiBytes('<< /Type /Catalog /Pages 2 0 R >>');
    const kids = pages.map((_, i) => `${3 + i*3} 0 R`).join(' ');
    objects[2] = asciiBytes(`<< /Type /Pages /Kids [${kids}] /Count ${pages.length} >>`);

    pages.forEach((p, i) => {
      const pageObj = 3 + i*3, imageObj = pageObj + 1, contentObj = pageObj + 2;
      objects[pageObj] = asciiBytes(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /XObject << /Im${i} ${imageObj} 0 R >> >> /Contents ${contentObj} 0 R >>`);
      const imgHead = asciiBytes(`<< /Type /XObject /Subtype /Image /Width ${p.width} /Height ${p.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`);
      const imgTail = asciiBytes('\nendstream');
      objects[imageObj] = concatBytes([imgHead, p.jpeg, imgTail]);
      const content = `q\n${PAGE_W} 0 0 ${PAGE_H} 0 0 cm\n/Im${i} Do\nQ`;
      const cb = asciiBytes(content);
      objects[contentObj] = concatBytes([asciiBytes(`<< /Length ${cb.length} >>\nstream\n`), cb, asciiBytes('\nendstream')]);
    });

    const chunks = [];
    let offset = 0;
    const push = b => { chunks.push(b); offset += b.length; };
    push(Uint8Array.from([0x25,0x50,0x44,0x46,0x2d,0x31,0x2e,0x34,0x0a,0x25,0xe2,0xe3,0xcf,0xd3,0x0a]));
    const offsets = new Array(objectCount + 1).fill(0);
    for (let n = 1; n <= objectCount; n++) {
      offsets[n] = offset;
      push(asciiBytes(`${n} 0 obj\n`));
      push(objects[n]);
      push(asciiBytes('\nendobj\n'));
    }
    const xrefOffset = offset;
    let xref = `xref\n0 ${objectCount + 1}\n0000000000 65535 f \n`;
    for (let n=1;n<=objectCount;n++) xref += `${String(offsets[n]).padStart(10,'0')} 00000 n \n`;
    xref += `trailer\n<< /Size ${objectCount + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
    push(asciiBytes(xref));
    return concatBytes(chunks);
  }

  // -------------------- Helpers --------------------

  function gradeChinese(g) { return ({'7':'七','8':'八','9':'九'})[g] || g; }
  function termLabel(t) { return t === 'A' ? '上學期' : '下學期'; }
  function examChinese(n) { return ({1:'一',2:'二',3:'三'})[n] || String(n); }
  function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
  function escapeHTML(s) { return String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'})[c]); }
  function escapeSVG(s) { return escapeHTML(s); }
  function dateStamp() {
    const d = new Date();
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  }
  function safeFileName(s) { return String(s).replace(/[\\/:*?"<>|]/g, '_'); }
  function idleYield() { return new Promise(resolve => setTimeout(resolve, 0)); }
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }
  function dataURLToBytes(url) {
    const base64 = url.split(',')[1];
    const bin = atob(base64);
    const out = new Uint8Array(bin.length);
    for (let i=0;i<bin.length;i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function asciiBytes(s) {
    const out = new Uint8Array(s.length);
    for (let i=0;i<s.length;i++) out[i] = s.charCodeAt(i) & 0xff;
    return out;
  }
  function concatBytes(arrays) {
    const len = arrays.reduce((s,a)=>s+a.length,0);
    const out = new Uint8Array(len); let off = 0;
    arrays.forEach(a => { out.set(a, off); off += a.length; });
    return out;
  }
  function roundRect(c, x, y, w, h, r) {
    r = Math.min(r, w/2, h/2);
    c.beginPath(); c.moveTo(x+r,y); c.arcTo(x+w,y,x+w,y+h,r); c.arcTo(x+w,y+h,x,y+h,r); c.arcTo(x,y+h,x,y,r); c.arcTo(x,y,x+w,y,r); c.closePath();
  }
  function wrapText(c, text, maxWidth) {
    const chars = Array.from(String(text));
    const lines = []; let line = '';
    for (const ch of chars) {
      const test = line + ch;
      if (line && c.measureText(test).width > maxWidth) { lines.push(line); line = ch; }
      else line = test;
    }
    if (line) lines.push(line);
    return lines;
  }

  // Expose a small read-only core API for diagnostics and automated validation.
  if (typeof window !== 'undefined') {
    window.StudentExamAnalyzerCore = Object.freeze({ parseXLSX, analyzeStudent, resolveQuery, rankIndex, reportHTML, buildImagePDF });
  }

  console.info(`Student Exam Analyzer Web ${APP_VERSION}`);
})();
