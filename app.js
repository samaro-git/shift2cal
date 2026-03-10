(function () {
  'use strict';

  // Fallback, wenn dienstzeiten.json per fetch nicht geladen werden kann (z. B. bei Öffnen über file://)
  var DIENSTZEITEN_FALLBACK = '[{"id":"OF","label":"Oberarzt- Frühdienst","start":"07:54:00","end":"16:24:00"},{"id":"RF","label":"R- Frühdienst","start":"07:24:00","end":"15:54:00"},{"id":"F","label":"Frühdienst","start":"07:24:00","end":"15:54:00"},{"id":"F1","label":"Frühdienst","start":"07:24:00","end":"15:54:00"},{"id":"F2","label":"Frühdienst","start":"07:24:00","end":"15:54:00"},{"id":"F3","label":"Frühdienst","start":"09:24:00","end":"17:54:00"},{"id":"KF","label":"Später Frühdienst","start":"10:54:00","end":"19:24:00"},{"id":"FW","label":"kurzer Frühdienst Wochenende","start":"07:24:00","end":"13:24:00"},{"id":"RW","label":"R-Dienst Wochenende","start":"10:54:00","end":"19:24:00"},{"id":"Z1","label":"Zwischendienst","start":"11:54:00","end":"20:24:00"},{"id":"Z2","label":"Zwischendienst","start":"13:24:00","end":"21:54:00"},{"id":"ZW","label":"Zwischendienst Wochenende","start":"12:54:00","end":"21:24:00"},{"id":"OS","label":"Oberarzt- Spätdienst","start":"15:54:00","end":"00:24:00"},{"id":"RS","label":"R- Spätdienst","start":"14:54:00","end":"23:24:00"},{"id":"S","label":"Spätdienst","start":"14:54:00","end":"23:24:00"},{"id":"S1","label":"Spätdienst","start":"14:54:00","end":"23:24:00"},{"id":"S2","label":"Spätdienst","start":"14:54:00","end":"23:24:00"},{"id":"KS","label":"Später Spätdienst","start":"16:54:00","end":"01:24:00"},{"id":"RN","label":"R- Nachtdienst","start":"22:24:00","end":"08:24:00"},{"id":"N","label":"Nachtdienst","start":"22:24:00","end":"08:24:00"},{"id":"N1","label":"Nachtdienst","start":"22:24:00","end":"08:24:00"},{"id":"N2","label":"Nachtdienst","start":"22:24:00","end":"08:24:00"},{"id":"AVD1","label":"AvD1 DIM","start":"20:00:00","end":"08:00:00"},{"id":"AVD2","label":"AvD2 DIM","start":"20:00:00","end":"08:00:00"},{"id":"AVD3","label":"AvD3 DIM","start":"20:00:00","end":"08:00:00"},{"id":"FTU","label":"Fast-Track-Unit","start":"08:00:00","end":"16:45:00"},{"id":"TF","label":"Trauma- Frühdienst","start":"07:24:00","end":"15:54:00"},{"id":"TS","label":"Trauma- Spätdienst","start":"11:30:00","end":"20:00:00"},{"id":"ITW","label":"Intensivtransport","start":"08:54:00","end":"19:06:00"},{"id":"NEFF","label":"NEF- Frühdienst","start":"07:30:00","end":"16:00:00"},{"id":"NEFS","label":"NEF- Spätdienst","start":"15:09:00","end":"23:39:00"},{"id":"NEFW","label":"NEF- Dienst Wochenende","start":"07:54:00","end":"19:06:00"},{"id":"Heli","label":"Christoph54","start":"07:30:00","end":"16:00:00"},{"id":"B","label":"Büro/ Orga","start":"08:00:00","end":"16:30:00"},{"id":"SZ","label":"Schulungszentrum","start":"08:00:00","end":"16:30:00"},{"id":"FB","label":"Fortbildung","start":"08:00:00","end":"16:30:00"},{"id":"DR","label":"Dienstreise","start":"08:00:00","end":"16:30:00"},{"id":"OA Ruf","label":"Oberarzt- Rufdienst","start":"00:24:00","end":"07:54:00"},{"id":"GF","label":"Geplantes Frei","start":null,"end":null}]';

  function buildDienstzeitenLookup(list) {
    var map = {};
    (list || []).forEach(function (e) {
      var id = (e.id || '').trim();
      if (id) map[id] = { start: e.start || '', end: e.end || '' };
    });
    return map;
  }

  var dienstzeitenLookup = {};
  var dienstzeitenList = [];

  var MAX_DAYS = 31;

  function buildIsoDate(year, month, day) {
    var m = String(month).padStart(2, '0');
    var d = String(day).padStart(2, '0');
    return year + '-' + m + '-' + d;
  }

  function formatDateDE(isoDate) {
    if (!isoDate || typeof isoDate !== 'string') return '';
    var p = isoDate.split('-');
    if (p.length !== 3 || !p[0] || !p[1] || !p[2]) return '';
    return parseInt(p[2], 10) + '.' + parseInt(p[1], 10) + '.' + p[0];
  }

  function getSelectedYearMonth() {
    var yEl = document.getElementById('yearSelect');
    var mEl = document.getElementById('monthSelect');
    var y = yEl ? parseInt(yEl.value, 10) : new Date().getFullYear();
    var m = mEl ? parseInt(mEl.value, 10) : null;
    return { year: isNaN(y) ? new Date().getFullYear() : y, month: (m >= 1 && m <= 12) ? m : null };
  }

  function monthNumberFromSheetName(sheetName) {
    var s = (sheetName || '').trim();
    if (!s) return null;
    var names = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
    for (var i = 0; i < names.length; i++) {
      if (s.indexOf(names[i]) !== -1) return i + 1;
    }
    return null;
  }

  function syncMonthFromSheet() {
    var sheetName = sheetDropdown && sheetDropdown.value ? sheetDropdown.value : '';
    var monthNum = monthNumberFromSheetName(sheetName);
    var monthEl = document.getElementById('monthSelect');
    if (monthEl && monthNum != null) monthEl.value = String(monthNum);
  }

  /** Liest Zellenwert als Tageszahl (1–31); Zahl oder String "1" etc. */
  function cellAsDay(val) {
    if (val == null) return null;
    if (typeof val === 'number' && !isNaN(val) && val >= 1 && val <= 31) return val;
    var s = String(val).trim();
    if (s === '') return null;
    var n = parseInt(s, 10);
    return (n >= 1 && n <= 31) ? n : null;
  }

  /** Findet die Zeile und Spalte, in der die Tagesfolge 1, 2, 3, 4, 5, … beginnt. */
  function findDayRowAndFirstColumn(data) {
    var maxRows = Math.min(25, data.length);
    for (var r = 0; r < maxRows; r++) {
      var row = data[r];
      if (!row || !row.length) continue;
      var maxC = Math.min(50, row.length - 5);
      for (var c = 0; c < maxC; c++) {
        if (cellAsDay(row[c]) === 1 && cellAsDay(row[c + 1]) === 2 &&
            cellAsDay(row[c + 2]) === 3 && cellAsDay(row[c + 3]) === 4 &&
            cellAsDay(row[c + 4]) === 5) {
          return { dayRow: r, firstDayCol: c };
        }
      }
    }
    return null;
  }

  /** Findet die Namensspalte (Format "Nachname, Vorname") nur links von firstDayCol. */
  function detectNameColumn(data, firstDayCol) {
    var headerWords = /^(name|mitarbeiter|person)$/i;
    var namePattern = /^[^,]+,\s*.+$/;
    var maxCol = firstDayCol != null ? Math.max(0, firstDayCol) : 10;
    if (maxCol === 0) return 0;
    if (data[0] && data[0].length < maxCol) maxCol = Math.min(maxCol, data[0].length);
    var bestCol = 0;
    var bestCount = -1;
    for (var c = 0; c < maxCol; c++) {
      var count = 0;
      for (var r = 0; r < data.length; r++) {
        var cell = data[r][c] != null ? String(data[r][c]).trim() : '';
        if (!cell || headerWords.test(cell)) continue;
        if (namePattern.test(cell)) count++;
      }
      if (count > bestCount) {
        bestCount = count;
        bestCol = c;
      }
    }
    return bestCol;
  }

  function parseSheetData(data) {
    var result = { names: [], dateKeys: [], rows: [] };
    if (!data.length) return result;

    var selected = getSelectedYearMonth();
    var year = selected.year;
    var month = selected.month;
    if (month == null) month = new Date().getMonth() + 1;

    var dateKeys = [];
    for (var day = 1; day <= MAX_DAYS; day++) {
      dateKeys.push(buildIsoDate(year, month, day));
    }

    var dayInfo = findDayRowAndFirstColumn(data);
    var firstDayCol = (dayInfo && dayInfo.firstDayCol >= 0) ? dayInfo.firstDayCol : 1;
    var dayRow = (dayInfo && dayInfo.dayRow >= 0) ? dayInfo.dayRow : -1;

    var nameCol = detectNameColumn(data, firstDayCol);
    var names = [];
    var rows = [];
    var headerWords = /^(name|mitarbeiter|person)$/i;
    for (var r = 0; r < data.length; r++) {
      if (r === dayRow) continue;
      var row = data[r];
      var name = (row[nameCol] != null ? String(row[nameCol]) : '').trim();
      if (!name || headerWords.test(name)) continue;
      names.push(name);
      var entries = [];
      for (var d = 0; d < MAX_DAYS; d++) {
        var col = firstDayCol + d;
        if (col >= row.length) break;
        var isoDate = dateKeys[d];
        var kuerzel = (row[col] != null ? String(row[col]) : '').trim();
        entries.push({ isoDate: isoDate, kuerzel: kuerzel });
      }
      rows.push({ name: name, entries: entries });
    }
    result.names = names;
    result.dateKeys = dateKeys;
    result.rows = rows;
    return result;
  }

  function parseSheetFromWorkbook(wb, sheetName) {
    var sheet = wb.Sheets[sheetName];
    if (!sheet) return { names: [], dateKeys: [], rows: [] };
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    return parseSheetData(data);
  }

  function parseXLS(buffer) {
    var wb = XLSX.read(buffer, { type: 'array', cellDates: false, raw: false });
    var first = wb.SheetNames[0];
    return parseSheetFromWorkbook(wb, first);
  }

  function getZeiten(kuerzel) {
    var k = (kuerzel || '').trim();
    var z = dienstzeitenLookup[k];
    if (z) return z;
    for (var key in dienstzeitenLookup) {
      if (key.trim() === k) return dienstzeitenLookup[key];
    }
    if (k.slice(-1) === '!') {
      var base = k.slice(0, -1).trim();
      z = dienstzeitenLookup[base];
      if (z) return z;
      for (var key2 in dienstzeitenLookup) {
        if (key2.trim() === base) return dienstzeitenLookup[key2];
      }
    }
    return { start: '', end: '' };
  }

  function getDienstLabel(kuerzel) {
    var k = (kuerzel || '').trim();
    for (var i = 0; i < dienstzeitenList.length; i++) {
      if ((dienstzeitenList[i].id || '').trim() === k) return dienstzeitenList[i].label || '';
    }
    if (k.slice(-1) === '!') {
      var base = k.slice(0, -1).trim();
      for (var j = 0; j < dienstzeitenList.length; j++) {
        if ((dienstzeitenList[j].id || '').trim() === base) return dienstzeitenList[j].label || '';
      }
    }
    return '';
  }

  function refreshDienstzeitenLookup() {
    dienstzeitenLookup = buildDienstzeitenLookup(dienstzeitenList);
  }

  function formatTimeForICS(t) {
    if (!t || !t.length) return '';
    var parts = t.split(':');
    return (parts[0] || '00') + (parts[1] || '00') + (parts[2] || '00');
  }

  function parseTimeToMinutes(t) {
    if (!t || !t.length) return 0;
    var parts = t.split(':').map(Number);
    return (parts[0] || 0) * 60 + (parts[1] || 0) + (parts[2] || 0) / 60;
  }

  function buildICS(events, formatLabel) {
    var lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Dienstplan-Kalender//DE',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH'
    ];
    events.forEach(function (ev) {
      var startTime = formatTimeForICS(ev.start);
      var endTime = formatTimeForICS(ev.end);
      var startM = parseTimeToMinutes(ev.start);
      var endM = parseTimeToMinutes(ev.end);
      var endIso = ev.isoDate;
      if (endM > 0 && endM < startM) {
        var d = new Date(ev.isoDate + 'T12:00:00');
        d.setDate(d.getDate() + 1);
        endIso = d.toISOString().slice(0, 10);
      }
      var startDt = ev.isoDate.replace(/-/g, '') + 'T' + startTime;
      var endDt = endIso.replace(/-/g, '') + 'T' + endTime;
      var summary = (ev.kuerzel || '') + (ev.label ? ' (' + ev.label + ')' : '');
      lines.push('BEGIN:VEVENT');
      lines.push('DTSTART:' + startDt);
      lines.push('DTEND:' + endDt);
      lines.push('SUMMARY:' + summary);
      lines.push('UID:' + ev.uid);
      lines.push('DTSTAMP:' + new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z');
      lines.push('END:VEVENT');
    });
    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }

  function oneEventICS(ev, formatLabel) {
    return buildICS([ev], formatLabel);
  }

  function download(content, filename, mime) {
    mime = mime || 'text/calendar;charset=utf-8';
    var blob = new Blob([content], { type: mime });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function uid() {
    return 'd-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  }

  function escapeHtml(str) {
    if (!str) return '';
    var div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  var state = {
    workbook: null,
    parsed: null,
    selectedName: null
  };

  var uploadZone = document.getElementById('uploadZone');
  var fileInput = document.getElementById('fileInput');
  var uploadStatus = document.getElementById('uploadStatus');
  var sheetSelectorContainer = document.getElementById('sheetSelectorContainer');
  var sheetDropdown = document.getElementById('sheetDropdown');
  var nameSelect = document.getElementById('nameSelect');
  var btnCreate = document.getElementById('btnCreate');
  var errorMsg = document.getElementById('errorMsg');
  var tableWrap = document.getElementById('tableWrap');
  var tableBody = document.getElementById('tableBody');

  function showError(msg) {
    errorMsg.textContent = msg || '';
    errorMsg.classList.toggle('hidden', !msg);
  }

  function fillNameSelect() {
    nameSelect.innerHTML = '<option value="">— Person wählen —</option>';
    if (state.parsed && state.parsed.names) {
      state.parsed.names.forEach(function (n) {
        var opt = document.createElement('option');
        opt.value = n;
        opt.textContent = n;
        nameSelect.appendChild(opt);
      });
    }
    nameSelect.disabled = !state.parsed || !state.parsed.names.length;
    btnCreate.disabled = nameSelect.disabled;
  }

  function onSheetSelected() {
    if (!state.workbook || !sheetDropdown.value) return;
    try {
      syncMonthFromSheet();
      state.parsed = parseSheetFromWorkbook(state.workbook, sheetDropdown.value);
      fillNameSelect();
      tableWrap.classList.add('hidden');
    } catch (err) {
      showError('Fehler beim Auswerten des Blatts: ' + err.message);
    }
  }

  function handleFile(file) {
    if (!file || (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls'))) {
      showError('Bitte eine .xlsx- oder .xls-Datei wählen.');
      return;
    }
    showError('');
    uploadStatus.textContent = 'Lade Datei …';
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        state.workbook = XLSX.read(e.target.result, { type: 'array', cellDates: false, raw: false });
        var names = state.workbook.SheetNames;
        if (!names || !names.length) {
          showError('Keine Tabellenblätter in der Datei gefunden.');
          uploadStatus.textContent = 'Fehler';
          return;
        }
        uploadStatus.innerHTML = escapeHtml(file.name) + ' <span class="status-success">&gt; Datei geladen.</span>';
        sheetDropdown.innerHTML = '';
        names.forEach(function (name) {
          var opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          sheetDropdown.appendChild(opt);
        });
        sheetSelectorContainer.classList.remove('hidden');
        syncMonthFromSheet();
        state.parsed = parseSheetFromWorkbook(state.workbook, sheetDropdown.value);
        fillNameSelect();
        tableWrap.classList.add('hidden');
      } catch (err) {
        showError('Fehler beim Lesen der Datei: ' + err.message);
        uploadStatus.textContent = 'Fehler';
        sheetSelectorContainer.classList.add('hidden');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  if (sheetDropdown) {
    sheetDropdown.addEventListener('change', onSheetSelected);
  }

  uploadZone.addEventListener('dragover', function (e) {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  });
  uploadZone.addEventListener('dragleave', function () {
    uploadZone.classList.remove('dragover');
  });
  uploadZone.addEventListener('drop', function (e) {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    var f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  fileInput.addEventListener('change', function () {
    var f = fileInput.files[0];
    if (f) handleFile(f);
  });

  function buildTableForPerson(name) {
    var row = state.parsed.rows.find(function (r) { return r.name === name; });
    if (!row) return;
    tableBody.innerHTML = '';
    row.entries.forEach(function (ent) {
      var zeiten = getZeiten(ent.kuerzel);
      var start = zeiten.start || '—';
      var end = zeiten.end || '—';
      if (start.length > 8) start = start.slice(0, 8);
      if (end.length > 8) end = end.slice(0, 8);

      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td>' + formatDateDE(ent.isoDate) + '</td>' +
        '<td>' + (ent.kuerzel || '—') + '</td>' +
        '<td>' + start + '</td>' +
        '<td>' + end + '</td>' +
        '<td class="export-cell"></td>';

      var exportCell = tr.querySelector('.export-cell');
      var ev = {
        isoDate: ent.isoDate,
        kuerzel: ent.kuerzel,
        label: getDienstLabel(ent.kuerzel),
        start: zeiten.start,
        end: zeiten.end,
        uid: uid()
      };
      var baseName = 'Dienst_' + (ent.isoDate || '') + '_' + (ent.kuerzel || '').replace(/\s+/g, '_');

      var btnIcs = document.createElement('button');
      btnIcs.type = 'button';
      btnIcs.className = 'btn-export btn-ics';
      btnIcs.textContent = '.ics - Download';
      btnIcs.title = 'iCalendar-Datei herunterladen';
      btnIcs.addEventListener('click', function () {
        download(oneEventICS(ev, '.ics'), baseName + '.ics');
      });
      exportCell.appendChild(btnIcs);

      tableBody.appendChild(tr);
    });
    tableWrap.classList.remove('hidden');
  }

  btnCreate.addEventListener('click', function () {
    var name = nameSelect.value;
    if (!name || !state.parsed) return;
    buildTableForPerson(name);
  });

  function applyDienstzeitenData(arr) {
    dienstzeitenList = Array.isArray(arr) ? arr.slice() : [];
    refreshDienstzeitenLookup();
    renderDienstzeitenTable();
  }

  function fetchDienstzeitenJson(noCache) {
    var url = 'dienstzeiten.json';
    if (noCache) url += '?_=' + Date.now();
    return fetch(url, { cache: 'no-store', method: 'GET' })
      .then(function (r) {
        if (!r.ok) throw new Error('Netzwerk: ' + r.status);
        return r.json();
      })
      .then(function (data) {
        if (!Array.isArray(data)) throw new Error('Ungültiges Format');
        return data;
      });
  }

  function loadDienstzeitenList() {
    fetchDienstzeitenJson(true)
      .then(function (arr) {
        applyDienstzeitenData(arr);
      })
      .catch(function () {
        try {
          applyDienstzeitenData(JSON.parse(DIENSTZEITEN_FALLBACK));
        } catch (e) {
          applyDienstzeitenData([]);
        }
      });
  }

  var dienstzeitenHeader = document.getElementById('dienstzeitenHeader');
  var dienstzeitenBody = document.getElementById('dienstzeitenBody');
  var dienstzeitenBodyRows = document.getElementById('dienstzeitenBodyRows');

  dienstzeitenHeader.addEventListener('click', function () {
    dienstzeitenHeader.classList.toggle('open');
    dienstzeitenBody.classList.toggle('open');
    dienstzeitenHeader.setAttribute('aria-expanded', dienstzeitenBody.classList.contains('open'));
  });

  var btnReloadDienstzeiten = document.getElementById('btnReloadDienstzeiten');
  if (btnReloadDienstzeiten) {
    btnReloadDienstzeiten.addEventListener('click', function () {
      btnReloadDienstzeiten.disabled = true;
      btnReloadDienstzeiten.textContent = 'Lade …';
      fetchDienstzeitenJson(true)
        .then(function (arr) {
          applyDienstzeitenData(arr);
        })
        .catch(function () {
          try {
            applyDienstzeitenData(JSON.parse(DIENSTZEITEN_FALLBACK));
          } catch (e) {
            applyDienstzeitenData([]);
          }
        })
        .then(function () {
          btnReloadDienstzeiten.textContent = 'dienstzeiten.json neu laden';
          btnReloadDienstzeiten.disabled = false;
        });
    });
  }

  function renderDienstzeitenTable() {
    dienstzeitenBodyRows.innerHTML = '';
    dienstzeitenList.forEach(function (entry) {
      var tr = document.createElement('tr');
      var tdId = document.createElement('td');
      tdId.textContent = entry.id || '';
      var tdLabel = document.createElement('td');
      tdLabel.textContent = entry.label || '';
      var tdStart = document.createElement('td');
      tdStart.textContent = entry.start || '';
      var tdEnd = document.createElement('td');
      tdEnd.textContent = entry.end || '';
      tr.appendChild(tdId);
      tr.appendChild(tdLabel);
      tr.appendChild(tdStart);
      tr.appendChild(tdEnd);
      dienstzeitenBodyRows.appendChild(tr);
    });
  }

  function escapeAttr(str) {
    if (str == null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  (function initYearSelect() {
    var sel = document.getElementById('yearSelect');
    if (!sel) return;
    var current = new Date().getFullYear();
    for (var y = current - 2; y <= current + 4; y++) {
      var opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      if (y === current) opt.selected = true;
      sel.appendChild(opt);
    }
  })();

  (function setDefaultMonth() {
    var sel = document.getElementById('monthSelect');
    if (!sel) return;
    var m = new Date().getMonth() + 1;
    sel.value = String(m);
  })();

  loadDienstzeitenList();
})();
