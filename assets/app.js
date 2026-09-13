(() => {
  'use strict';

  const DAYS = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  const DAY_SHORT = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
  const DAY_FULL = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  const DB_NAME = 'teacher-schedule-db';
  const STORE = 'files';
  const DEFAULT_FILE = 'files/schedule.xlsx';

  const state = {
    lessons: [],
    selectedDay: defaultDayIndex(),
    teacher: localStorage.getItem('teacherName') || 'Обеднин',
    file: null,
    fileName: '',
    updatedAt: null,
    view: 'day'
  };

  const $ = (selector) => document.querySelector(selector);
  const els = {
    todayLabel: $('#todayLabel'), summaryKicker: $('#summaryKicker'), summaryTitle: $('#summaryTitle'),
    summaryText: $('#summaryText'), summaryNumber: $('#summaryNumber'), dayTabs: $('#dayTabs'),
    selectedDate: $('#selectedDate'), selectedDay: $('#selectedDay'), lessonCount: $('#lessonCount'),
    lessonList: $('#lessonList'), weekSection: $('#weekSection'), weekList: $('#weekList'),
    manageSection: $('#manageSection'), fileInput: $('#fileInput'), uploadZone: $('#uploadZone'),
    fileName: $('#fileName'), fileMeta: $('#fileMeta'), downloadFile: $('#downloadFile'),
    teacherInput: $('#teacherInput'), applyTeacher: $('#applyTeacher'), openUpload: $('#openUpload'), toast: $('#toast')
  };

  function currentDayIndex() {
    const day = new Date().getDay();
    return day >= 1 && day <= 6 ? day - 1 : -1;
  }

  function defaultDayIndex() {
    const current = currentDayIndex();
    return current >= 0 ? current : 0;
  }

  function normalize(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
  }

  function normalizeTime(value) {
    return normalize(value).replace(/[–—]/g, '-').replace(/\./g, ':').replace(/\s/g, '');
  }

  function className(value) {
    return normalize(value).replace(/\s/g, '').toUpperCase();
  }

  function isClass(value) {
    return /^\d{1,2}\s*[а-яёa-z]$/i.test(normalize(value));
  }

  function parseCell(raw, teacher) {
    const text = normalize(raw);
    const pos = text.toLocaleLowerCase('ru').indexOf(teacher.toLocaleLowerCase('ru'));
    const before = pos >= 0 ? text.slice(0, pos).trim() : text;
    const after = pos >= 0 ? text.slice(pos + teacher.length).trim() : '';
    let subject = before.replace(/[\\/;,\-]+$/g, '').trim() || 'Занятие';
    if (/^обзр$/i.test(subject)) subject = 'ОБЗР';
    else subject = subject.charAt(0).toUpperCase() + subject.slice(1);
    return { subject, room: after.replace(/^[,;:\-]+/, '').trim(), raw: text };
  }

  function parseWorkbook(buffer, teacher) {
    if (!window.XLSX) throw new Error('Не удалось загрузить модуль чтения XLSX');
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
    const lessons = [];

    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '' });
      if (!rows.length) return;

      let headerRow = -1;
      let bestScore = 0;
      rows.slice(0, 12).forEach((row, index) => {
        const score = row.filter(isClass).length + row.filter((cell) => /звонки/i.test(normalize(cell))).length * 2;
        if (score > bestScore) { bestScore = score; headerRow = index; }
      });
      if (headerRow < 0 || bestScore < 2) return;

      const headers = rows[headerRow];
      const timeColumns = headers.map((value, index) => /звонки/i.test(normalize(value)) ? index : -1).filter((index) => index >= 0);
      let currentDay = '';

      for (let r = headerRow + 1; r < rows.length; r += 1) {
        const row = rows[r];
        const foundDay = row.map(normalize).find((cell) => DAYS.includes(cell.toLocaleLowerCase('ru')));
        if (foundDay) currentDay = foundDay.toLocaleLowerCase('ru');
        if (!currentDay) continue;

        headers.forEach((header, c) => {
          if (!isClass(header)) return;
          const raw = normalize(row[c]);
          if (!raw || !raw.toLocaleLowerCase('ru').includes(teacher.toLocaleLowerCase('ru'))) return;
          const timeCol = [...timeColumns].reverse().find((col) => col < c);
          if (timeCol == null) return;
          const time = normalizeTime(row[timeCol]);
          const number = normalize(row[timeCol - 1]);
          if (!time) return;
          const details = parseCell(raw, teacher);
          lessons.push({
            id: `${sheetName}-${r}-${c}`,
            sheet: sheetName,
            day: currentDay,
            dayIndex: DAYS.indexOf(currentDay),
            className: className(header),
            number: number.replace(/\.0$/, ''),
            time,
            ...details
          });
        });
      }
    });

    return lessons.sort((a, b) => a.dayIndex - b.dayIndex || minutes(a.time) - minutes(b.time) || a.className.localeCompare(b.className, 'ru'));
  }

  function minutes(time) {
    const start = String(time).split('-')[0];
    const [h, m] = start.split(':').map(Number);
    return Number.isFinite(h) && Number.isFinite(m) ? h * 60 + m : 9999;
  }

  function endMinutes(time) {
    const end = String(time).split('-')[1] || '';
    const [h, m] = end.split(':').map(Number);
    return Number.isFinite(h) && Number.isFinite(m) ? h * 60 + m : minutes(time) + 40;
  }

  function lessonCard(lesson, markNow = false) {
    const [start, end] = lesson.time.split('-');
    const nowClass = markNow && isCurrentLesson(lesson) ? ' is-now' : '';
    return `<article class="lesson-card${nowClass}">
      <div class="lesson-time"><strong>${escapeHtml(start || lesson.time)}</strong><span>${escapeHtml(end || '')}</span></div>
      <div class="lesson-main">
        <strong>${escapeHtml(lesson.subject)}</strong>
        <div class="lesson-meta">
          <span class="pill">${escapeHtml(lesson.className)} класс</span>
          ${lesson.room ? `<span class="pill room">Каб. ${escapeHtml(lesson.room)}</span>` : ''}
        </div>
      </div>
      <span class="lesson-number">${lesson.number ? `${escapeHtml(lesson.number)} ур.` : ''}</span>
    </article>`;
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
  }

  function isCurrentLesson(lesson) {
    if (lesson.dayIndex !== currentDayIndex()) return false;
    const now = new Date();
    const current = now.getHours() * 60 + now.getMinutes();
    return current >= minutes(lesson.time) && current <= endMinutes(lesson.time);
  }

  function render() {
    els.teacherInput.value = state.teacher;
    renderTabs();
    renderDay();
    renderWeek();
    renderSummary();
    renderFile();
  }

  function renderTabs() {
    els.dayTabs.innerHTML = DAY_SHORT.map((label, index) => {
      const hasLessons = state.lessons.some((lesson) => lesson.dayIndex === index);
      return `<button class="day-tab${state.selectedDay === index ? ' is-active' : ''}${hasLessons ? '' : ' is-empty'}" data-day="${index}" type="button">${label}</button>`;
    }).join('');
  }

  function renderDay() {
    const lessons = state.lessons.filter((lesson) => lesson.dayIndex === state.selectedDay);
    const isToday = state.selectedDay === currentDayIndex();
    els.selectedDate.textContent = isToday ? 'Сегодня' : 'День недели';
    els.selectedDay.textContent = DAY_FULL[state.selectedDay];
    els.lessonCount.textContent = countLabel(lessons.length);
    els.lessonList.innerHTML = lessons.length
      ? lessons.map((lesson) => lessonCard(lesson, isToday)).join('')
      : `<div class="empty-state"><strong>Занятий нет</strong><span>${state.lessons.length ? 'На этот день ничего не найдено.' : 'Загрузите таблицу XLSX, чтобы увидеть расписание.'}</span></div>`;
  }

  function renderWeek() {
    els.weekList.innerHTML = DAYS.map((day, index) => {
      const lessons = state.lessons.filter((lesson) => lesson.dayIndex === index);
      if (!lessons.length) return '';
      return `<section class="week-day">
        <div class="week-day-title"><strong>${DAY_FULL[index]}</strong><span>${countLabel(lessons.length)}</span></div>
        <div class="lesson-list">${lessons.map((lesson) => lessonCard(lesson, false)).join('')}</div>
      </section>`;
    }).join('') || `<div class="empty-state"><strong>Расписание пустое</strong><span>Загрузите XLSX во вкладке «Файл».</span></div>`;
  }

  function renderSummary() {
    const currentDay = currentDayIndex();
    const today = state.lessons.filter((lesson) => lesson.dayIndex === currentDay);
    const now = today.find(isCurrentLesson);
    const currentMinutes = new Date().getHours() * 60 + new Date().getMinutes();
    const next = today.find((lesson) => minutes(lesson.time) > currentMinutes);
    els.todayLabel.textContent = new Intl.DateTimeFormat('ru-RU', { weekday: 'long', day: 'numeric', month: 'long' }).format(new Date());
    els.summaryNumber.textContent = today.length || '—';
    if (!state.lessons.length) {
      els.summaryKicker.textContent = 'Нет данных';
      els.summaryTitle.textContent = 'Загрузите расписание';
      els.summaryText.textContent = 'Подойдёт файл Excel в формате XLSX';
    } else if (now) {
      els.summaryKicker.textContent = 'Сейчас идёт';
      els.summaryTitle.textContent = `${now.subject} · ${now.className}`;
      els.summaryText.textContent = `${now.time}${now.room ? ` · кабинет ${now.room}` : ''}`;
    } else if (next) {
      els.summaryKicker.textContent = 'Следующее занятие';
      els.summaryTitle.textContent = `${next.subject} · ${next.className}`;
      els.summaryText.textContent = `${next.time}${next.room ? ` · кабинет ${next.room}` : ''}`;
    } else if (today.length) {
      els.summaryKicker.textContent = 'На сегодня всё';
      els.summaryTitle.textContent = countLabel(today.length);
      els.summaryText.textContent = 'Следующее расписание — в ближайший учебный день';
    } else {
      els.summaryKicker.textContent = 'Сегодня';
      els.summaryTitle.textContent = 'Занятий нет';
      els.summaryText.textContent = 'Можно посмотреть всю неделю';
    }
  }

  function renderFile() {
    els.fileName.textContent = state.fileName || 'Файл не загружен';
    els.fileMeta.textContent = state.updatedAt
      ? `Обновлено ${new Intl.DateTimeFormat('ru-RU', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }).format(new Date(state.updatedAt))} · найдено ${state.lessons.length}`
      : '—';
    els.downloadFile.disabled = !state.file;
  }

  function countLabel(count) {
    const last = count % 10;
    const lastTwo = count % 100;
    const word = last === 1 && lastTwo !== 11 ? 'занятие' : last >= 2 && last <= 4 && !(lastTwo >= 12 && lastTwo <= 14) ? 'занятия' : 'занятий';
    return `${count} ${word}`;
  }

  async function handleFile(file, silent = false) {
    if (!file || !/\.xlsx?$/i.test(file.name)) {
      showToast('Выберите файл Excel в формате XLSX');
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const lessons = parseWorkbook(buffer, state.teacher);
      if (!lessons.length) throw new Error(`Фамилия «${state.teacher}» не найдена в расписании`);
      state.lessons = lessons;
      state.file = new Blob([buffer], { type: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      state.fileName = file.name;
      state.updatedAt = Date.now();
      await saveToDb({ id: 'current', blob: state.file, name: state.fileName, updatedAt: state.updatedAt });
      persistParsed();
      render();
      if (!silent) { switchView('day'); showToast(`Готово: найдено ${lessons.length} занятий`); }
    } catch (error) {
      console.error(error);
      showToast(error.message || 'Не удалось прочитать файл');
    } finally {
      els.fileInput.value = '';
    }
  }

  function persistParsed() {
    localStorage.setItem('teacherName', state.teacher);
    localStorage.setItem('parsedSchedule', JSON.stringify({ lessons: state.lessons, fileName: state.fileName, updatedAt: state.updatedAt }));
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: 'id' });
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function saveToDb(record) {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(record);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  }

  async function loadFromDb() {
    const db = await openDb();
    const record = await new Promise((resolve, reject) => {
      const request = db.transaction(STORE, 'readonly').objectStore(STORE).get('current');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return record;
  }

  async function initialize() {
    const cached = localStorage.getItem('parsedSchedule');
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        state.lessons = Array.isArray(parsed.lessons) ? parsed.lessons : [];
        state.fileName = parsed.fileName || '';
        state.updatedAt = parsed.updatedAt || null;
      } catch (_) { /* use bundled file below */ }
    }
    try {
      const saved = await loadFromDb();
      if (saved?.blob) {
        state.file = saved.blob;
        state.fileName = saved.name;
        state.updatedAt = saved.updatedAt;
        if (!state.lessons.length) await handleFile(new File([saved.blob], saved.name), true);
      } else if (!state.lessons.length) {
        const response = await fetch(DEFAULT_FILE);
        if (response.ok) await handleFile(new File([await response.blob()], 'schedule.xlsx'), true);
      }
    } catch (error) {
      console.warn('Initial schedule was not loaded', error);
    }
    render();
  }

  function switchView(view) {
    state.view = view;
    const isDay = view === 'day';
    $('.schedule-section').hidden = !isDay;
    els.dayTabs.hidden = !isDay;
    els.summaryCard.hidden = !isDay;
    els.weekSection.hidden = view !== 'week';
    els.manageSection.hidden = view !== 'manage';
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('is-active', item.dataset.view === view));
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  let toastTimer;
  function showToast(message) {
    els.toast.textContent = message;
    els.toast.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove('is-visible'), 2800);
  }

  els.dayTabs.addEventListener('click', (event) => {
    const button = event.target.closest('[data-day]');
    if (!button) return;
    state.selectedDay = Number(button.dataset.day);
    renderTabs();
    renderDay();
  });
  document.querySelectorAll('.nav-item').forEach((button) => button.addEventListener('click', () => switchView(button.dataset.view)));
  els.openUpload.addEventListener('click', () => switchView('manage'));
  els.uploadZone.addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', () => handleFile(els.fileInput.files[0]));
  ['dragenter', 'dragover'].forEach((type) => els.uploadZone.addEventListener(type, (event) => { event.preventDefault(); els.uploadZone.classList.add('is-dragging'); }));
  ['dragleave', 'drop'].forEach((type) => els.uploadZone.addEventListener(type, (event) => { event.preventDefault(); els.uploadZone.classList.remove('is-dragging'); }));
  els.uploadZone.addEventListener('drop', (event) => handleFile(event.dataTransfer.files[0]));
  els.applyTeacher.addEventListener('click', async () => {
    const nextTeacher = normalize(els.teacherInput.value);
    if (!nextTeacher) return showToast('Введите фамилию преподавателя');
    if (!state.file) return showToast('Сначала загрузите файл расписания');
    const previous = state.teacher;
    state.teacher = nextTeacher;
    const lessons = parseWorkbook(await state.file.arrayBuffer(), state.teacher);
    if (!lessons.length) { state.teacher = previous; els.teacherInput.value = previous; return showToast(`Фамилия «${nextTeacher}» не найдена`); }
    state.lessons = lessons;
    persistParsed();
    render();
    showToast(`Показано расписание: ${state.teacher}`);
  });
  els.downloadFile.addEventListener('click', () => {
    if (!state.file) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(state.file);
    link.download = state.fileName || 'schedule.xlsx';
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  });

  initialize();
})();
