// ========================================
// 魚館＜手話検定シミュレーター＞ - script.js
// ========================================

(function () {
  'use strict';

  // --- State ---
  const state = {
    screen: 'top',    // top, select, question, complete
    mode: null,       // 'shuwa' or 'writing'
    loading: true,
    shuwaData: [],
    writingData: [],
    // Select
    selectedYear: '',
    selectedLevel: '',
    selectedType: '',
    randomYear: null,
    yearOpts: [],
    levelOpts: [],
    typeOpts: [],
    // Question
    filteredQuestions: [],
    usedQuestions: [],
    currentQuestion: null,
    storyQuestions: [],
    selectedAnswer: null,
    storyAnswers: {},
    showResult: false,
    answered: 0,
    correct: 0,
    streak: 0,
    maxStreak: 0,
    // Countdown
    countdownTime: 0,
    countdownTotal: 0,
    countdownActive: false,
    countdownDone: false,
    // Essay
    essayText: '',
    essayDone: false,
    // Stats (persisted)
    totalStudied: 0,
    totalCorrect: 0,
    streakDays: 0,
  };

  let countdownInterval = null;

  // --- DOM refs ---
  const $ = (id) => document.getElementById(id);

  // --- State update helper ---
  function setState(updates) {
    Object.assign(state, updates);
    render();
  }

  // --- Stats persistence ---
  function loadStats() {
    try {
      const s = JSON.parse(localStorage.getItem('shuwa_stats') || '{}');
      const today = new Date().toDateString();
      const lastDay = s.lastDay || '';
      let streakDays = s.streakDays || 0;
      if (lastDay && lastDay !== today) {
        const diff = (new Date(today) - new Date(lastDay)) / 86400000;
        if (diff > 1) streakDays = 0;
      }
      state.totalStudied = s.totalStudied || 0;
      state.totalCorrect = s.totalCorrect || 0;
      state.streakDays = streakDays;
    } catch (e) { /* ignore */ }
  }

  function saveStats(addStudied, addCorrect) {
    const today = new Date().toDateString();
    const ts = state.totalStudied + addStudied;
    const tc = state.totalCorrect + addCorrect;
    let sd = state.streakDays;
    try {
      const prev = JSON.parse(localStorage.getItem('shuwa_stats') || '{}');
      if (prev.lastDay !== today) {
        sd = (prev.lastDay && (new Date(today) - new Date(prev.lastDay)) / 86400000 <= 1) ? sd + 1 : 1;
      }
      localStorage.setItem('shuwa_stats', JSON.stringify({
        totalStudied: ts, totalCorrect: tc, streakDays: sd, lastDay: today
      }));
    } catch (e) { /* ignore */ }
    setState({ totalStudied: ts, totalCorrect: tc, streakDays: sd });
  }

  // --- Excel loading ---
  async function loadExcel() {
    try {
      const resp = await fetch('excel/shuwa_exam_question_list_ver0.04.xlsx');
      const ab = await resp.arrayBuffer();
      const wb = XLSX.read(ab, { type: 'array' });
      const shuwa = XLSX.utils.sheet_to_json(wb.Sheets['shuwa'] || {});
      const writing = XLSX.utils.sheet_to_json(wb.Sheets['writing'] || {});
      setState({ shuwaData: shuwa, writingData: writing, loading: false });
    } catch (e) {
      console.error('Excel load error:', e);
      setState({ loading: false });
    }
  }

  // --- Data helpers ---
  function getData() {
    return state.mode === 'shuwa' ? state.shuwaData : state.writingData;
  }

  function populateYears() {
    const data = getData();
    let years = [...new Set(data.map(d => d.Year))].filter(v => v != null && `${v}`.trim() !== '');
    years.sort((a, b) => Number(b) - Number(a));
    return years;
  }

  function populateLevels(year) {
    const data = getData();
    const filtered = year === 'none' ? data : data.filter(d => `${d.Year}` === `${year}` || `${d.Year}` === 'unknown');
    return [...new Set(filtered.map(d => d.Level))].filter(v => v != null && `${v}`.trim() !== '');
  }

  function populateTypes(year, level) {
    const data = getData();
    let filtered;
    if (year === 'none') {
      filtered = data.filter(d => `${d.Level}` === `${level}`);
    } else {
      filtered = data.filter(d => `${d.Year}` === `${year}` && `${d.Level}` === `${level}`);
    }
    return [...new Set(filtered.map(d => d.Type))].filter(v => v != null && `${v}`.trim() !== '');
  }

  function selectRandomYear() {
    const data = getData();
    const allYears = [...new Set(data.map(d => d.Year))].filter(v => v != null && `${v}` !== 'unknown');
    return allYears[Math.floor(Math.random() * allYears.length)];
  }

  function filterQuestions() {
    const { selectedYear, selectedLevel, selectedType, randomYear } = state;
    const data = getData();
    const yr = selectedYear === 'none' ? randomYear : selectedYear;
    return data.filter(d =>
      (`${d.Year}` === `${yr}` || `${d.Year}` === 'unknown') &&
      `${d.Level}` === `${selectedLevel}` &&
      `${d.Type}` === `${selectedType}`
    );
  }

  function getRemainingCount() {
    const filtered = filterQuestions();
    const used = state.usedQuestions;
    if (state.selectedType === 'ストーリー') {
      const allNums = new Set(filtered.map(d => d['ストーリーナンバリング']));
      const usedNums = new Set(used.map(d => d['ストーリーナンバリング']));
      return [...allNums].filter(n => !usedNums.has(n)).length;
    }
    return filtered.filter(q => !used.includes(q)).length;
  }

  function getTotalCount() {
    const filtered = filterQuestions();
    if (state.selectedType === 'ストーリー') {
      return new Set(filtered.map(d => d['ストーリーナンバリング'])).size;
    }
    return filtered.length;
  }

  function pickNextQuestion() {
    const filtered = filterQuestions();
    const used = state.usedQuestions;
    const remaining = filtered.filter(q => !used.includes(q));
    if (remaining.length === 0) return null;

    const q = remaining[Math.floor(Math.random() * remaining.length)];

    if (q.Type === 'ストーリー') {
      const stNum = q['ストーリーナンバリング'];
      const storyQs = filtered
        .filter(d => d['ストーリーナンバリング'] === stNum)
        .sort((a, b) => a['#'] - b['#'])
        .slice(0, 3);
      Object.assign(state, {
        currentQuestion: q, storyQuestions: storyQs,
        usedQuestions: [...used, ...storyQs],
        selectedAnswer: null, storyAnswers: {}, showResult: false,
        countdownActive: false, countdownDone: false,
        essayDone: false, essayText: ''
      });
    } else {
      Object.assign(state, {
        currentQuestion: q, storyQuestions: [],
        usedQuestions: [...used, q],
        selectedAnswer: null, storyAnswers: {}, showResult: false,
        countdownActive: false, countdownDone: false,
        essayDone: false, essayText: ''
      });
    }
    render();
    return q;
  }

  // --- Countdown ---
  function startCountdown() {
    const level = state.selectedLevel;
    const total = (level === '5' || level === '4' || level === '3' || level === 5 || level === 4 || level === 3) ? 60 : 120;
    state.countdownTime = total;
    state.countdownTotal = total;
    state.countdownActive = true;
    state.countdownDone = false;
    render();

    if (countdownInterval) clearInterval(countdownInterval);
    countdownInterval = setInterval(() => {
      state.countdownTime -= 1;
      if (state.countdownTime <= 0) {
        clearInterval(countdownInterval);
        countdownInterval = null;
        state.countdownTime = 0;
        state.countdownActive = false;
        state.countdownDone = true;
      }
      render();
    }, 1000);
  }

  function stopCountdown() {
    if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
    setState({ countdownActive: false, countdownDone: true });
  }

  function cancelCountdown() {
    if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
    setState({ countdownActive: false, countdownDone: false, countdownTime: 0, countdownTotal: 0 });
  }

  // --- Answer submission ---
  function submitAnswer() {
    const q = state.currentQuestion;
    const correctIdx = parseInt(q.Answer) - 1;
    const correct = state.selectedAnswer === correctIdx;
    const streak = correct ? state.streak + 1 : 0;
    const maxStreak = Math.max(streak, state.maxStreak);
    setState({
      showResult: true,
      answered: state.answered + 1,
      correct: state.correct + (correct ? 1 : 0),
      streak, maxStreak
    });
  }

  function submitStoryAnswers() {
    const { storyQuestions, storyAnswers } = state;
    let correctCount = 0;
    storyQuestions.forEach((sq, i) => {
      if (parseInt(sq.Answer) - 1 === storyAnswers[i]) correctCount++;
    });
    const allCorrect = correctCount === storyQuestions.length;
    const streak = allCorrect ? state.streak + 1 : 0;
    const maxStreak = Math.max(streak, state.maxStreak);
    setState({
      showResult: true,
      answered: state.answered + 1,
      correct: state.correct + (allCorrect ? 1 : 0),
      streak, maxStreak
    });
  }

  function submitFillBlank(uniqueMatches) {
    const q = state.currentQuestion;
    const selectedStr = uniqueMatches.map((_, i) =>
      (state.storyAnswers[i] || '').split('.')[0].trim()
    ).join(',');
    const correct = selectedStr === `${q.Answer}`;
    const streak = correct ? state.streak + 1 : 0;
    const maxStreak = Math.max(streak, state.maxStreak);
    setState({
      showResult: true,
      answered: state.answered + 1,
      correct: state.correct + (correct ? 1 : 0),
      streak, maxStreak
    });
  }

  function nextQuestion() {
    const remaining = getRemainingCount();
    if (remaining <= 0) {
      saveStats(state.answered, state.correct);
      setState({ screen: 'complete' });
      return;
    }
    pickNextQuestion();
  }

  // --- Theme colors helper ---
  function getTheme() {
    const isS = state.mode === 'shuwa';
    return {
      accent: isS ? '#E8937E' : '#7A9DE8',
      accentDark: isS ? '#C26A54' : '#5471B8',
      accentMid: isS ? 'rgba(158,84,64,.6)' : 'rgba(62,86,148,.6)',
      accentBg: isS ? '#FFF5F1' : '#F1F4FF',
      bandBg: isS ? 'linear-gradient(135deg,#FFEAE3,#F4B5A5)' : 'linear-gradient(135deg,#E3EAFF,#A5BDF4)',
      btnBg: isS ? 'linear-gradient(135deg,#E8937E,#D97A66)' : 'linear-gradient(135deg,#7A9DE8,#5C82D0)',
      btnShadow: isS ? 'rgba(232,147,126,.25)' : 'rgba(122,157,232,.25)',
    };
  }

  // ========================================
  // RENDER
  // ========================================
  function render() {
    const screens = ['screen-top', 'screen-select', 'screen-question', 'screen-complete'];
    const activeId = 'screen-' + state.screen;
    screens.forEach(id => {
      const el = $(id);
      el.style.display = id === activeId ? 'flex' : 'none';
    });

    // Footer: hide TOP link on top screen
    const footerLink = $('footer-top-link');
    if (footerLink) footerLink.style.display = state.screen === 'top' ? 'none' : 'inline-block';

    if (state.screen === 'top') renderTop();
    else if (state.screen === 'select') renderSelect();
    else if (state.screen === 'question') renderQuestion();
    else if (state.screen === 'complete') renderComplete();
  }

  // --- TOP ---
  function renderTop() {
    const ts = state.totalStudied || 0;
    const acc = ts > 0 ? Math.round((state.totalCorrect || 0) / ts * 100) + '%' : '-';
    const sd = state.streakDays > 0 ? state.streakDays + '日' : '-';

    $('stat-total').textContent = ts;
    $('stat-accuracy').textContent = acc;
    $('stat-streak').textContent = sd;
    $('loading-msg').style.display = state.loading ? 'block' : 'none';
  }

  // --- SELECT ---
  function renderSelect() {
    const t = getTheme();

    // Band
    $('select-band').style.background = t.bandBg;
    $('select-icon').textContent = state.mode === 'shuwa' ? '実' : '筆';
    $('select-icon').style.color = t.accent;
    $('select-title').textContent = (state.mode === 'shuwa' ? '実技' : '筆記') + 'コース';
    $('select-title').style.color = t.accentDark;
    $('select-subtitle').style.color = t.accentMid;
    $('btn-back').style.color = t.accentMid;

    // Step circles
    $('step-1').style.background = t.accent;
    $('step-2').style.background = state.selectedYear ? t.accent : '#C5C5C5';
    $('step-2').className = 'step-circle' + (state.selectedYear ? '' : ' disabled');
    $('step-3').style.background = state.selectedLevel ? t.accent : '#C5C5C5';
    $('step-3').className = 'step-circle' + (state.selectedLevel ? '' : ' disabled');

    // Year dropdown
    const yearSel = $('year-select');
    yearSel.style.borderColor = state.selectedYear ? t.accent : '#E0DCD6';
    // Populate options (keep first two fixed)
    while (yearSel.options.length > 2) yearSel.remove(2);
    state.yearOpts.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      yearSel.appendChild(opt);
    });
    yearSel.value = state.selectedYear;

    // Level dropdown
    const levelSel = $('level-select');
    levelSel.disabled = !state.selectedYear;
    levelSel.style.borderColor = state.selectedLevel ? t.accent : '#E0DCD6';
    levelSel.style.opacity = state.selectedYear ? '1' : '0.5';
    levelSel.style.color = state.selectedYear ? '#2D2A26' : 'rgba(45,42,38,.35)';
    while (levelSel.options.length > 1) levelSel.remove(1);
    state.levelOpts.forEach(l => {
      const opt = document.createElement('option');
      opt.value = l;
      opt.textContent = l;
      levelSel.appendChild(opt);
    });
    levelSel.value = state.selectedLevel;

    // Type dropdown
    const typeSel = $('type-select');
    typeSel.disabled = !state.selectedLevel;
    typeSel.style.borderColor = state.selectedType ? t.accent : '#E0DCD6';
    typeSel.style.opacity = state.selectedLevel ? '1' : '0.5';
    typeSel.style.color = state.selectedLevel ? '#2D2A26' : 'rgba(45,42,38,.35)';
    while (typeSel.options.length > 1) typeSel.remove(1);
    state.typeOpts.forEach(tp => {
      const opt = document.createElement('option');
      opt.value = tp;
      opt.textContent = tp;
      typeSel.appendChild(opt);
    });
    typeSel.value = state.selectedType;

    // Info badge & start button
    const allSelected = state.selectedYear && state.selectedLevel && state.selectedType;
    const infoEl = $('question-info');
    const startEl = $('btn-start');
    if (allSelected) {
      const count = filterQuestions().length;
      infoEl.style.display = 'block';
      infoEl.style.background = t.accentBg;
      infoEl.style.color = t.accentDark;
      $('question-count').textContent = count + '問';

      startEl.style.display = 'block';
      startEl.style.background = t.btnBg;
      startEl.style.boxShadow = '0 4px 14px ' + t.btnShadow;
    } else {
      infoEl.style.display = 'none';
      startEl.style.display = 'none';
    }
  }

  // --- QUESTION ---
  function renderQuestion() {
    const t = getTheme();
    const header = $('question-header');
    header.style.background = t.bandBg;

    const isS = state.mode === 'shuwa';
    const modeLabel = isS ? '実技' : '筆記';
    $('progress-label').textContent = `${modeLabel} ・ ${state.selectedLevel} ・ ${state.selectedType}`;
    $('progress-label').style.color = t.accentDark;

    const remaining = getRemainingCount();
    const total = getTotalCount();
    const done = total - remaining;
    const unit = state.selectedType === 'ストーリー' ? 'セット' : '問';
    $('remaining-label').textContent = `残り ${remaining}${unit}`;
    $('remaining-label').style.color = t.accentMid;

    const pct = total > 0 ? Math.round(done / total * 100) : 0;
    $('progress-bar-fill').style.width = pct + '%';
    $('progress-text').textContent = `${done}/${total}`;
    $('progress-text').style.color = t.accentDark;

    // Build question body
    const body = $('question-body');
    const q = state.currentQuestion;

    // Preserve YouTube iframe to avoid reload flicker (only if same video)
    let existingVideo = body.querySelector('.video-embed');
    if (existingVideo) {
      const iframe = existingVideo.querySelector('iframe');
      const currentRef = iframe ? iframe.src : '';
      if (q && q.Reference && currentRef.includes(q.Reference)) {
        existingVideo.remove();
      } else {
        existingVideo = null;
      }
    }
    body.innerHTML = '';
    if (!q) return;

    if (q.Type === 'スピーチ') buildSpeechUI(body, q, t);
    else if (q.Type === '小論文') buildEssayUI(body, q, t);
    else if (q.Type === 'ストーリー') buildStoryUI(body, q, t, existingVideo);
    else if (q.Type === '穴埋め形式') buildFillBlankUI(body, q, t);
    else if (['基本単語', '短文', '4択問題'].includes(q.Type)) buildMultipleChoiceUI(body, q, t, existingVideo);
    else body.textContent = '問題タイプが不明です: ' + q.Type;
  }

  // --- Build Speech UI ---
  function buildSpeechUI(container, q, t) {
    appendEl(container, 'div', { className: 'q-label' }, 'スピーチテーマ');
    appendEl(container, 'div', { className: 'q-theme' }, q.Theme);
    appendEl(container, 'div', {
      style: "font:400 11px/1.5 'Noto Sans JP',sans-serif;color:rgba(45,42,38,.45);text-align:center;margin-top:-18px;margin-bottom:20px"
    }, '※実際には面接官に対してスピーチをします。本サイトでは1人で話してみてください。');

    if (!state.countdownActive && !state.countdownDone) {
      const btn = appendEl(container, 'div', {
        className: 'btn-action',
        style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow}`
      });
      appendEl(btn, 'div', { className: 'btn-action-text' }, 'カウントダウン開始');
      btn.onclick = () => startCountdown();
    }

    if (state.countdownActive) {
      const wrap = appendEl(container, 'div', { className: 'timer-wrap' });
      const svgWrap = appendEl(wrap, 'div', { className: 'timer-svg-wrap' });

      const pct = state.countdownTime / state.countdownTotal;
      const r = 70, circ = 2 * Math.PI * r;
      svgWrap.innerHTML = `
        <svg width="160" height="160" viewBox="0 0 160 160" style="transform:rotate(-90deg)">
          <circle cx="80" cy="80" r="${r}" fill="none" stroke="#F0EDE8" stroke-width="8"/>
          <circle cx="80" cy="80" r="${r}" fill="none" stroke="${t.accent}" stroke-width="8"
            stroke-dasharray="${circ}" stroke-dashoffset="${circ * (1 - pct)}"
            stroke-linecap="round" style="transition:stroke-dashoffset 1s linear"/>
        </svg>
        <div class="timer-center">
          <div class="timer-number" style="color:${t.accentDark}">${state.countdownTime}</div>
          <div class="timer-unit">秒</div>
        </div>`;

      const btns = appendEl(wrap, 'div', { className: 'timer-buttons' });
      const cancelBtn = appendEl(btns, 'div', { className: 'btn-outline', style: 'flex:1' }, 'キャンセル');
      cancelBtn.onclick = () => cancelCountdown();
      const endBtn = appendEl(btns, 'div', {
        className: 'btn-action',
        style: `flex:1;background:${t.btnBg};font:600 12px/1 'Noto Sans JP',sans-serif;color:#fff`
      }, '終了する');
      endBtn.onclick = () => stopCountdown();
    }

    if (state.countdownDone) {
      if (q.Reference) buildVideo(container, q.Reference);
      if (q.Detail) buildDetail(container, q.Detail);
      buildNextButton(container, t);
    }
  }

  // --- Build Essay UI ---
  function buildEssayUI(container, q, t) {
    appendEl(container, 'div', {
      style: "font:600 15px/1.5 'Noto Sans JP',sans-serif;color:#2D2A26;margin-bottom:6px"
    }, '小論文テーマ');
    appendEl(container, 'div', {
      style: "font:400 13px/1.6 'Noto Sans JP',sans-serif;color:rgba(45,42,38,.7);margin-bottom:16px"
    }, q.Question || q.Theme);
    appendEl(container, 'div', {
      className: 'essay-guide',
      style: `background:${t.accentBg};color:${t.accentDark}`
    }, '600字〜800字で記述してください');

    const ta = document.createElement('textarea');
    ta.className = 'essay-textarea';
    ta.rows = 10;
    ta.value = state.essayText;
    ta.placeholder = 'ここに小論文を書いてください（どこにも送信されません）';
    ta.disabled = state.essayDone;
    ta.style.border = `2px solid ${state.essayDone ? '#E0DCD6' : t.accent}`;
    ta.oninput = (e) => {
      state.essayText = e.target.value;
      const countEl = container.querySelector('.essay-count');
      if (countEl) countEl.textContent = state.essayText.length + ' 字';
    };
    container.appendChild(ta);

    appendEl(container, 'div', {
      className: 'essay-count', style: `color:${t.accentDark}`
    }, (state.essayText || '').length + ' 字');

    if (!state.essayDone) {
      const btn = appendEl(container, 'div', {
        className: 'btn-action',
        style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow}`
      });
      appendEl(btn, 'div', { className: 'btn-action-text' }, '完了');
      btn.onclick = () => setState({ essayDone: true, showResult: true });
    } else {
      if (q.Detail) buildDetail(container, q.Detail);
      buildNextButton(container, t);
    }
  }

  // --- Build Story UI ---
  function buildStoryUI(container, q, t, existingVideo) {
    const { storyQuestions, storyAnswers, showResult } = state;
    if (storyQuestions.length === 0) return;

    if (q.Reference) {
      if (existingVideo) container.appendChild(existingVideo);
      else buildVideo(container, q.Reference);
    }

    storyQuestions.forEach((sq, qi) => {
      const wrap = appendEl(container, 'div', { style: 'margin-bottom:16px' });
      const header = appendEl(wrap, 'div', { className: 'story-q-header' });
      const num = appendEl(header, 'div', { className: 'story-q-num', style: `background:${t.accent}` });
      num.textContent = qi + 1;
      header.appendChild(document.createTextNode(sq.Theme));

      const optsWrap = appendEl(wrap, 'div', { className: 'story-opts' });
      const opts = (sq.Option || '').split('\n').filter(o => o.trim());
      const answered = storyAnswers[qi] !== undefined;

      opts.forEach((opt, oi) => {
        const selected = storyAnswers[qi] === oi;
        const isCorrect = showResult && parseInt(sq.Answer) === oi + 1;
        const isWrong = showResult && selected && parseInt(sq.Answer) !== oi + 1;

        let bg = '#F8F6F3', border = 'transparent', textColor = '#2D2A26';
        let radioBorder = '#D4D0CA', radioInner = '';
        if (selected && !showResult) {
          bg = t.accentBg; border = t.accent; textColor = t.accentDark;
          radioBorder = t.accent; radioInner = t.accent;
        }
        if (isCorrect) { bg = '#E8F5E9'; border = '#5BAE6A'; textColor = '#2E7D32'; radioBorder = '#5BAE6A'; radioInner = '#5BAE6A'; }
        if (isWrong) { bg = '#FFEBEE'; border = '#E57373'; textColor = '#C62828'; radioBorder = '#E57373'; radioInner = '#E57373'; }

        const item = appendEl(optsWrap, 'div', {
          className: 'story-option',
          style: `background:${bg};border-color:${border};color:${textColor};cursor:${(answered && showResult) ? 'default' : 'pointer'}`
        });

        const radio = appendEl(item, 'div', { className: 'story-radio', style: `border-color:${radioBorder}` });
        if (selected || isCorrect) {
          const inner = appendEl(radio, 'div', { className: 'story-radio-inner' });
          inner.style.display = 'block';
          inner.style.background = radioInner || t.accent;
        }
        item.appendChild(document.createTextNode(opt));

        if (!(answered && showResult)) {
          item.onclick = () => {
            state.storyAnswers = { ...state.storyAnswers, [qi]: oi };
            render();
          };
        }
      });
    });

    const allAnswered = storyQuestions.every((_, i) => storyAnswers[i] !== undefined);
    if (!showResult && allAnswered) {
      const btn = appendEl(container, 'div', {
        className: 'btn-action',
        style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow};margin-top:8px`
      });
      appendEl(btn, 'div', { className: 'btn-action-text' }, 'すべて回答する');
      btn.onclick = () => submitStoryAnswers();
    }

    if (showResult) {
      if (q.Detail) buildDetail(container, q.Detail);
      buildNextButton(container, t);
    }
  }

  // --- Build Multiple Choice UI ---
  function buildMultipleChoiceUI(container, q, t, existingVideo) {
    if (q.Reference) {
      if (existingVideo) container.appendChild(existingVideo);
      else buildVideo(container, q.Reference);
    }

    const qText = q.Type === '4択問題' ? q.Question : q.Theme;
    appendEl(container, 'div', { className: 'q-text' }, qText);

    const opts = (q.Option || '').split('\n').filter(o => o.trim());
    const correctIdx = parseInt(q.Answer) - 1;
    const { showResult, selectedAnswer } = state;

    const listEl = appendEl(container, 'div', { className: 'option-list' });

    const items = [];
    opts.forEach((opt, oi) => {
      const selected = selectedAnswer === oi;
      const isCorrect = showResult && oi === correctIdx;
      const isWrong = showResult && selected && oi !== correctIdx;

      let bg = '#F8F6F3', border = 'transparent', textColor = '#2D2A26';
      let radioBorder = '#D4D0CA', radioInner = '', fontWeight = '400';
      if (selected && !showResult) {
        bg = t.accentBg; border = t.accent; textColor = t.accentDark;
        radioBorder = t.accent; radioInner = t.accent; fontWeight = '500';
      }
      if (isCorrect) { bg = '#E8F5E9'; border = '#5BAE6A'; textColor = '#2E7D32'; radioBorder = '#5BAE6A'; radioInner = '#5BAE6A'; }
      if (isWrong) { bg = '#FFEBEE'; border = '#E57373'; textColor = '#C62828'; radioBorder = '#E57373'; radioInner = '#E57373'; }

      const item = appendEl(listEl, 'div', {
        className: 'option-item' + (showResult ? ' disabled' : ''),
        style: `background:${bg};border-color:${border}`
      });
      items.push(item);

      const radio = appendEl(item, 'div', { className: 'option-radio', style: `border-color:${radioBorder}` });
      if (selected || isCorrect) {
        const inner = appendEl(radio, 'div', { className: 'option-radio-inner' });
        inner.style.display = 'block';
        inner.style.background = radioInner || t.accent;
      }

      appendEl(item, 'span', {
        className: 'option-text',
        style: `color:${textColor};font-weight:${fontWeight}`
      }, opt);
    });

    // Answer button area (inserted after option list, before result feedback)
    const answerArea = appendEl(container, 'div', { className: 'answer-area' });

    if (!showResult) {
      items.forEach((item, oi) => {
        item.onclick = () => {
          state.selectedAnswer = oi;
          // Update option styles without re-rendering
          items.forEach((el, j) => {
            const isSel = j === oi;
            el.style.background = isSel ? t.accentBg : '#F8F6F3';
            el.style.borderColor = isSel ? t.accent : 'transparent';
            const radio = el.querySelector('.option-radio');
            radio.style.borderColor = isSel ? t.accent : '#D4D0CA';
            let inner = radio.querySelector('.option-radio-inner');
            if (isSel) {
              if (!inner) { inner = appendEl(radio, 'div', { className: 'option-radio-inner' }); }
              inner.style.display = 'block';
              inner.style.background = t.accent;
            } else if (inner) {
              inner.style.display = 'none';
            }
            const text = el.querySelector('.option-text');
            text.style.color = isSel ? t.accentDark : '#2D2A26';
            text.style.fontWeight = isSel ? '500' : '400';
          });
          // Show answer button
          if (!answerArea.querySelector('.btn-action')) {
            const btn = appendEl(answerArea, 'div', {
              className: 'btn-action',
              style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow}`
            });
            appendEl(btn, 'div', { className: 'btn-action-text' }, '回答する');
            btn.onclick = () => submitAnswer();
          }
        };
      });
    }

    if (!showResult && selectedAnswer !== null) {
      const btn = appendEl(answerArea, 'div', {
        className: 'btn-action',
        style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow}`
      });
      appendEl(btn, 'div', { className: 'btn-action-text' }, '回答する');
      btn.onclick = () => submitAnswer();
    }

    if (showResult) {
      const correct = selectedAnswer === correctIdx;
      // Result feedback
      const fb = appendEl(container, 'div', {
        className: 'result-feedback ' + (correct ? 'correct' : 'wrong')
      });
      const icon = appendEl(fb, 'div', {
        className: 'result-feedback-icon ' + (correct ? 'correct' : 'wrong')
      }, correct ? '✓' : '✗');
      const fbText = appendEl(fb, 'div');
      appendEl(fbText, 'div', {
        className: 'result-feedback-title',
        style: `color:${correct ? '#2E7D32' : '#C62828'}`
      }, correct ? '正解です' : '不正解です');
      if (!correct) {
        appendEl(fbText, 'div', {
          className: 'result-feedback-sub',
          style: 'color:rgba(198,40,40,.6)'
        }, '正解: ' + (opts[correctIdx] || ''));
      }

      if (q.Detail) buildDetail(container, q.Detail);

      // Mini stats
      const statsRow = appendEl(container, 'div', { className: 'mini-stats' });
      const s1 = appendEl(statsRow, 'div', { className: 'mini-stat', style: `background:${t.accentBg}` });
      appendEl(s1, 'div', { className: 'mini-stat-val', style: `color:${t.accent}` }, '' + state.streak);
      appendEl(s1, 'div', { className: 'mini-stat-label' }, '連続正解');
      const s2 = appendEl(statsRow, 'div', { className: 'mini-stat', style: 'background:#E8F5E9' });
      const curAcc = state.answered > 0 ? Math.round(state.correct / state.answered * 100) + '%' : '-';
      appendEl(s2, 'div', { className: 'mini-stat-val', style: 'color:#5BAE6A' }, curAcc);
      appendEl(s2, 'div', { className: 'mini-stat-label' }, '現在の正答率');

      buildNextButton(container, t);
    }
  }

  // --- Build Fill-Blank UI ---
  function buildFillBlankUI(container, q, t) {
    appendEl(container, 'div', { className: 'q-text-fill' }, q.Question);

    const opts = (q.Option || '').split('\n').filter(o => o.trim());
    const matches = [...(q.Question || '').matchAll(/[ア-オ]/g)];
    const uniqueMatches = [...new Set(matches.map(m => m[0]))];

    uniqueMatches.forEach((m, mi) => {
      const grp = appendEl(container, 'div', { className: 'fill-select-group' });
      appendEl(grp, 'div', { className: 'fill-select-label', style: `color:${t.accentDark}` }, m + ' の選択');

      const sel = document.createElement('select');
      sel.className = 'fill-select';
      sel.style.border = `2px solid ${t.accent}`;
      sel.disabled = state.showResult;

      const defOpt = document.createElement('option');
      defOpt.value = '';
      defOpt.textContent = '選択してください';
      sel.appendChild(defOpt);

      opts.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.trim();
        opt.textContent = o.trim();
        sel.appendChild(opt);
      });

      sel.value = state.storyAnswers[mi] || '';
      sel.onchange = (e) => {
        state.storyAnswers = { ...state.storyAnswers, [mi]: e.target.value };
        render();
      };
      grp.appendChild(sel);
    });

    const allFilled = uniqueMatches.every((_, i) => state.storyAnswers[i]);
    if (!state.showResult && allFilled) {
      const btn = appendEl(container, 'div', {
        className: 'btn-action',
        style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow};margin-top:12px`
      });
      appendEl(btn, 'div', { className: 'btn-action-text' }, '回答する');
      btn.onclick = () => submitFillBlank(uniqueMatches);
    }

    if (state.showResult) {
      const selectedStr = uniqueMatches.map((_, i) =>
        (state.storyAnswers[i] || '').split('.')[0].trim()
      ).join(',');
      const correct = selectedStr === `${q.Answer}`;
      appendEl(container, 'div', {
        className: 'fill-result ' + (correct ? 'correct' : 'wrong')
      }, correct ? '正解です' : '不正解です');

      if (q.Detail) buildDetail(container, q.Detail);
      buildNextButton(container, t);
    }
  }

  // --- Shared builders ---
  function buildVideo(container, ref) {
    const wrap = appendEl(container, 'div', { className: 'video-embed' });
    wrap.innerHTML = `<iframe src="https://www.youtube.com/embed/${ref}" allow="accelerometer;autoplay;clipboard-write;encrypted-media;gyroscope;picture-in-picture" allowfullscreen></iframe>`;
  }

  function buildDetail(container, detail) {
    const box = appendEl(container, 'div', { className: 'detail-box' });
    appendEl(box, 'div', { className: 'detail-title' }, '解説');
    appendEl(box, 'div', { className: 'detail-text' }, detail);
  }

  function buildNextButton(container, t) {
    const btn = appendEl(container, 'div', {
      className: 'btn-action',
      style: `background:${t.btnBg};box-shadow:0 3px 10px ${t.btnShadow}`
    });
    appendEl(btn, 'div', { className: 'btn-action-text' }, '次の問題へ');
    btn.onclick = () => nextQuestion();
    setTimeout(() => {
      const max = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
      window.scrollTo({ top: max, behavior: 'smooth' });
    }, 600);
  }

  function appendEl(parent, tag, attrs, text) {
    const el = document.createElement(tag);
    if (attrs) {
      if (attrs.className) el.className = attrs.className;
      if (attrs.style) el.setAttribute('style', attrs.style);
      if (attrs.id) el.id = attrs.id;
    }
    if (text !== undefined) el.textContent = text;
    parent.appendChild(el);
    return el;
  }

  // --- COMPLETE ---
  function renderComplete() {
    const t = getTheme();
    const { answered, correct, maxStreak } = state;
    const isS = state.mode === 'shuwa';
    const modeLabel = isS ? '実技' : '筆記';

    $('complete-subtitle').textContent = `${modeLabel} ・ ${state.selectedLevel} ・ ${state.selectedType}を完了しました`;

    // Score circle
    const pct = answered > 0 ? Math.round(correct / answered * 100) : 0;
    const r = 52, circ = 2 * Math.PI * r;
    const offset = circ * (1 - pct / 100);
    $('score-circle').innerHTML = `
      <svg width="120" height="120" viewBox="0 0 120 120" style="transform:rotate(-90deg)">
        <circle cx="60" cy="60" r="${r}" fill="none" stroke="#F0EDE8" stroke-width="8"/>
        <circle cx="60" cy="60" r="${r}" fill="none" stroke="#5BAE6A" stroke-width="8"
          stroke-dasharray="${circ}" stroke-dashoffset="${offset}" stroke-linecap="round"/>
      </svg>
      <div class="score-center">
        <div class="score-pct">${pct}%</div>
        <div class="score-label">正答率</div>
      </div>`;

    $('result-correct').textContent = correct + '問';
    $('result-wrong').textContent = (answered - correct) + '問';
    $('result-maxstreak').textContent = maxStreak + '問';

    // Retry button style
    $('btn-retry').style.background = t.btnBg;
    $('btn-retry').style.boxShadow = '0 3px 10px ' + t.btnShadow;
  }

  // ========================================
  // EVENT BINDINGS
  // ========================================
  function bindEvents() {
    // Top page buttons
    $('btn-jitsugi').onclick = () => {
      const data = state.shuwaData;
      if (!data || data.length === 0) { loadExcel(); return; }
      let years = [...new Set(data.map(d => d.Year))].filter(v => v != null && `${v}`.trim() !== '' && `${v}` !== 'unknown');
      years.sort((a, b) => Number(b) - Number(a));
      setState({
        mode: 'shuwa', screen: 'select',
        yearOpts: years, levelOpts: [], typeOpts: [],
        selectedYear: '', selectedLevel: '', selectedType: ''
      });
    };

    $('btn-hikki').onclick = () => {
      const data = state.writingData;
      if (!data || data.length === 0) { loadExcel(); return; }
      let years = [...new Set(data.map(d => d.Year))].filter(v => v != null && `${v}`.trim() !== '' && `${v}` !== 'unknown');
      years.sort((a, b) => Number(b) - Number(a));
      setState({
        mode: 'writing', screen: 'select',
        yearOpts: years, levelOpts: [], typeOpts: [],
        selectedYear: '', selectedLevel: '', selectedType: ''
      });
    };

    // Select page
    $('btn-back').onclick = goTop;

    $('year-select').onchange = (e) => {
      const val = e.target.value;
      let randomYear = null;
      if (val === 'none') randomYear = selectRandomYear();
      const levels = val ? populateLevels(val === 'none' ? `${randomYear}` : val) : [];
      setState({
        selectedYear: val, randomYear,
        selectedLevel: '', selectedType: '',
        levelOpts: levels, typeOpts: []
      });
    };

    $('level-select').onchange = (e) => {
      const val = e.target.value;
      const { selectedYear, randomYear } = state;
      const data = getData();
      const yr = selectedYear === 'none' ? randomYear : selectedYear;
      const filtered = data.filter(d => (`${d.Year}` === `${yr}` || `${d.Year}` === 'unknown') && `${d.Level}` === `${val}`);
      const types = [...new Set(filtered.map(d => d.Type))].filter(v => v != null && `${v}`.trim() !== '');
      setState({ selectedLevel: val, selectedType: '', typeOpts: types });
    };

    $('type-select').onchange = (e) => {
      setState({ selectedType: e.target.value });
    };

    $('btn-start').onclick = () => {
      Object.assign(state, {
        screen: 'question', usedQuestions: [],
        answered: 0, correct: 0, streak: 0, maxStreak: 0
      });
      render();
      pickNextQuestion();
    };

    // Complete page
    $('btn-retry').onclick = () => {
      Object.assign(state, {
        screen: 'question', usedQuestions: [],
        answered: 0, correct: 0, streak: 0, maxStreak: 0,
        currentQuestion: null
      });
      render();
      pickNextQuestion();
    };

    $('btn-home').onclick = goTop;

    // Question page back button
    $('btn-back-question').onclick = () => {
      if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
      setState({
        screen: 'select',
        currentQuestion: null,
        usedQuestions: [],
        answered: 0, correct: 0, streak: 0, maxStreak: 0
      });
    };

    // Footer TOP link
    $('footer-top-link').onclick = goTop;
  }

  function goTop() {
    if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
    setState({
      screen: 'top', mode: null,
      selectedYear: '', selectedLevel: '', selectedType: '',
      usedQuestions: [], currentQuestion: null,
      answered: 0, correct: 0, streak: 0, maxStreak: 0
    });
  }

  // ========================================
  // INIT
  // ========================================
  function init() {
    loadStats();
    bindEvents();
    render();
    loadExcel();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
