(function() {
  'use strict';
  // M365 Bulk User CSV Generator v4.6 — ported as dashboard module (User management > Bulk generator).
  // Standalone behavior preserved: no dashboard links, no persistence, fresh config per page load.
  // Uses the dashboard's vendored SheetJS instead of the inline copy.
  var UG_HTML = `<div class="container">
  <h1>📋 Microsoft 365 Bulk User CSV Generator</h1>
  <div class="sub">Offline tool to convert school enrollment Excel files into M365 bulk import CSVs.</div>
  <div class="ug-tenant-row" style="display:flex;gap:10px;align-items:end;margin-bottom:6px;">
    <div><label for="ug-tenant">Tenant code</label><input id="ug-tenant" placeholder="BSCS" style="width:110px;text-transform:uppercase;"></div>
    <div><label for="ug-month">Month</label><input id="ug-month" placeholder="SEP" style="width:80px;text-transform:uppercase;"></div>
    <div class="sub" style="margin:0;padding-bottom:8px;">Used as filename prefix, e.g. BSCS_SEP_</div>
  </div>

  <!-- SECTION 1: INPUT DATA -->
  <div class="card">
    <h2>1. Input Data</h2>
    <div class="checkbox-row" style="gap:24px;">
      <label style="display:flex;align-items:center;gap:6px;font-weight:600;margin:0;">
        <input type="radio" name="inputMode" value="excel" checked style="width:auto;"> 📄 Upload Excel File
      </label>
      <label style="display:flex;align-items:center;gap:6px;font-weight:600;margin:0;">
        <input type="radio" name="inputMode" value="manual" style="width:auto;"> ⌨️ Manual Entry (best for &lt; 10 users)
      </label>
    </div>

    <!-- Excel panel -->
    <div id="excelPanel" style="margin-top:16px;">
      <input type="file" id="fileInput" accept=".xlsx,.xls">
      <div class="progress" id="progress"><div class="progress-bar" id="progressBar"></div></div>
      <div class="status" id="fileStatus"></div>

      <!-- Sheet selection (shown only when the workbook has more than one usable sheet) -->
      <div id="sheetSelectWrap" style="display:none; margin-top:14px; padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:6px;">
        <label style="margin-bottom:8px;">Sheets to Process</label>
        <div style="font-size:11px; color:#6b7280; margin-bottom:8px;">
          This file has multiple sheets. Tick the ones you want to include. Column mapping is auto-detected from the <strong>first selected</strong> sheet.
        </div>
        <div class="checkbox-row" style="margin-bottom:8px;">
          <input type="checkbox" id="sheetSelectAll" checked>
          <label for="sheetSelectAll" style="margin:0;">Select / deselect all sheets</label>
        </div>
        <div id="sheetCheckboxes" style="display:flex; flex-wrap:wrap; gap:8px 18px;"></div>
      </div>
    </div>

    <!-- Manual entry panel -->
    <div id="manualPanel" style="display:none; margin-top:16px;">
      <div style="font-size:12px; color:#6b7280; margin-bottom:10px; line-height:1.5;">
        Type users in directly — ideal for fewer than 10. <strong>First Name</strong> and <strong>Last Name</strong> are always required.
        <strong>Student Number</strong> is required only when the selected Email Format (Section 2) actually uses it — so email schemes that don't rely on a student number will work without one.
        Grade Level, Job Title, and Display Name are optional.
      </div>

      <!-- Bulk paste from Excel / CSV -->
      <div style="margin-bottom:14px; padding:12px 14px; background:#eff6ff; border:1px solid #bfdbfe; border-radius:6px;">
        <label for="manualPasteBox" style="margin-bottom:6px;">⇩ Bulk Paste from Excel / CSV</label>
        <div style="font-size:11px; color:#6b7280; margin-bottom:8px; line-height:1.6;">
          Copy the rows from your spreadsheet (or CSV) and paste them below, then click <strong>Parse &amp; Fill</strong>. Tab-separated (Excel/Sheets) and comma-separated (CSV) both work.
          <br><strong>Include your header row</strong> (e.g. <code>FIRST NAME* &nbsp; LAST NAME* &nbsp; DISPLAY NAME* &nbsp; JOB TITLE*</code>) and columns are matched <strong>by name</strong> — so the column order can differ from the table below and missing columns (like Student Number) are just left blank. Without a header row, columns are read in this order: <code>First → Last → Student Number → Grade → Job Title → Display Name</code>.
          <br>Parsing <strong>replaces</strong> the rows in the table. Tip: you can also click any cell and press <strong>Ctrl+V</strong> to drop a block straight onto the grid.
        </div>
        <textarea id="manualPasteBox" rows="5" placeholder="Paste rows here — e.g. copy the columns straight out of Excel or Google Sheets (header row included is fine)." style="width:100%; font-family:'Consolas','Courier New',monospace; font-size:12px; padding:8px 10px; border:1px solid #d1d5db; border-radius:5px; resize:vertical;"></textarea>
        <div style="margin-top:8px;">
          <button type="button" id="manualParseBtn">⇩ Parse &amp; Fill</button>
          <button type="button" id="manualPasteClearBtn" style="background:#6b7280;">Clear Box</button>
        </div>
        <div class="status" id="manualPasteStatus"></div>
      </div>

      <div class="table-wrap" style="max-height:none; overflow:visible;">
        <table id="manualTable">
          <thead><tr>
            <th>First Name <span style="color:#dc2626">*</span></th>
            <th>Last Name <span style="color:#dc2626">*</span></th>
            <th>Student Number</th>
            <th>Grade Level</th>
            <th>Job Title</th>
            <th>Display Name</th>
            <th></th>
          </tr></thead>
          <tbody id="manualBody"></tbody>
        </table>
      </div>
      <button type="button" id="manualAddBtn" style="margin-top:10px;">+ Add User</button>
      <span id="manualCount" style="font-size:12px; color:#6b7280; margin-left:8px;"></span>
    </div>
  </div>

  <!-- SECTION 2: CONFIGURATION -->
  <div class="card">
    <h2>2. Configuration</h2>
    <div class="grid">
      <div>
        <label>Domain</label>
        <input type="text" id="domain" placeholder="e.g. school.edu.ph">
      </div>
      <div>
        <label>Password (temporary)</label>
        <input type="text" id="password" placeholder="e.g. Welcome2025!">
      </div>
      <div>
        <label>Job Title (default/fallback)</label>
        <input type="text" id="jobTitle" value="Student">
      </div>
      <div>
        <label>Email Format</label>
        <select id="emailFormat">
          <option value="1">1. GivenName.Surname</option>
          <option value="2">2. Surname.GivenName.StudentNumber</option>
          <option value="3" selected="">3. GivenName.Surname.Last4Digits</option>
          <option value="4">4. Surname.GivenName</option>
          <option value="5">5. Surname.GivenName.Last4Digits</option>
          <option value="6">6. StudentNumber (with hyphen)</option>
          <option value="7">7. StudentNumber (no hyphen)</option>
          <option value="8">8. GivenName_Surname.Last4Digits</option>
          <option value="9">9. Custom Format...</option>
        </select>
      </div>
      <div>
        <label>Display Name Format</label>
        <select id="displayNameFormat">
          <option value="0" selected="">0. Use mapped Display Name column (fallback: GivenName Surname)</option>
          <option value="1">1. First Name Last Name</option>
          <option value="2">2. Last Name, First Name</option>
          <option value="3">3. First Name Last Name (ALL CAPS)</option>
          <option value="4">4. Last Name, First Name (ALL CAPS)</option>
          <option value="5">5. Custom Format (build from columns)...</option>
        </select>
      </div>
    </div>
    <div class="checkbox-row">
      <input type="checkbox" id="dedupeNames">
      <label for="dedupeNames" style="margin:0;">Also remove duplicate names (FirstName + LastName), not just UPN/StudentNumber</label>
    </div>

    <!-- Custom Format panel (shown only when Email Format = 9) -->
    <div id="customFormatPanel" style="display:none; margin-top:14px; padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:6px;">
      <label for="customFormatInput">Custom Pattern</label>
      <input type="text" id="customFormatInput" placeholder="e.g. {given}_{surname}{last4}">
      <div style="font-size:11px; color:#6b7280; margin-top:6px; line-height:1.5;">
        Available tokens: <code>{given}</code> <code>{surname}</code> <code>{last4}</code>
        <code>{studentnumber}</code> <code>{studentnumberhyphen}</code> <code>{giveninitial}</code> <code>{surnameinitial}</code>.
        You can mix in literal characters like <code>.</code> <code>_</code> <code>-</code>.
        Example: <code>{given}_{surname}{last4}</code> → <strong id="customFormatExample">juan_delacruz6908</strong>
      </div>
    </div>

    <!-- Custom Display Name panel (shown only when Display Name Format = 5) -->
    <div id="dnCustomPanel" style="display:none; margin-top:14px; padding:12px 14px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:6px;">
      <label style="margin-bottom:8px;">Display Name Parts (in order)</label>
      <div style="font-size:11px; color:#6b7280; margin-bottom:8px;">
        Pick any column from your file for each part, or choose "Literal text" to insert fixed characters (spaces, commas, etc.). Parts are joined in the order shown.
      </div>
      <div id="dnPartsContainer"></div>
      <button type="button" id="dnAddPartBtn" style="margin-top:4px;">+ Add Part</button>
      <div class="checkbox-row">
        <input type="checkbox" id="dnCustomAllCaps">
        <label for="dnCustomAllCaps" style="margin:0;">Convert result to ALL CAPS</label>
      </div>
      <div class="checkbox-row">
        <input type="checkbox" id="dnCustomProperCase">
        <label for="dnCustomProperCase" style="margin:0;">Force Proper Case (fixes ALL-CAPS/lowercase source data — ignored if ALL CAPS above is checked)</label>
      </div>
      <div style="font-size:12px; margin-top:8px;">
        Preview: <strong id="dnPreview">(upload a file and add parts above)</strong>
      </div>
    </div>
  </div>

  <!-- SECTION 2b: COLUMN MAPPING -->
  <div class="card" id="mappingCard" style="display:none;">
    <h2>3. Map Columns</h2>
    <div style="font-size:12px; color:#6b7280; margin-bottom:10px;">
      Choose which column in your Excel file holds each field. Auto-detected from the first sheet's header row — adjust if needed. Header-row selections will be skipped during processing.
    </div>
    <div class="grid">
      <div>
        <label>Student Number <span style="color:#dc2626">*</span></label>
        <select id="mapStudentNumber"></select>
      </div>
      <div>
        <label>First Name (GivenName) <span style="color:#dc2626">*</span></label>
        <select id="mapFirstName"></select>
      </div>
      <div>
        <label>Last Name (Surname) <span style="color:#dc2626">*</span></label>
        <select id="mapLastName"></select>
      </div>
      <div>
        <label>Display Name</label>
        <select id="mapDisplayName"></select>
      </div>
      <div>
        <label>Grade Level (Department)</label>
        <select id="mapDepartment"></select>
      </div>
      <div>
        <label>Job Title <span style="color:#6b7280;font-weight:400;">(optional - overrides fixed value below)</span></label>
        <select id="mapJobTitle"></select>
      </div>
      <div>
        <label>Header Row Index (0-based)</label>
        <input type="number" id="headerRowIdx" value="0" min="0" max="20">
      </div>
    </div>
  </div>

  <!-- SECTION 3: PROCESS -->
  <div class="card">
    <h2>4. Process</h2>
    <button id="processBtn" disabled="">▶ Process Excel</button>
    <button id="clearBtn">✕ Clear</button>
    <div class="status" id="procStatus"></div>
  </div>

  <!-- SECTION 4: VALIDATION / STATS -->
  <div class="card" id="statsCard" style="display:none;">
    <h2>5. Summary</h2>
    <div class="stats" id="stats"></div>
    <div id="errorsWrap" style="display:none;">
      <strong style="color:#991b1b; font-size:13px;">Validation Issues:</strong>
      <div class="errors-box" id="errorsBox"></div>
    </div>
  </div>

  <!-- SECTION 5b: SKIPPED ROWS AUDIT -->
  <div class="card" id="skippedCard" style="display:none;">
    <h2>5b. Skipped Rows (Audit)</h2>
    <div style="font-size:12px; color:#6b7280; margin-bottom:10px;">
      Every row excluded from the output, with the reason it was skipped, for auditing against the source file.
    </div>
    <button id="downloadSkippedBtn">⬇ Download Skipped Rows (CSV)</button>
    <div class="table-wrap" style="margin-top:10px;">
      <table id="skippedTable"></table>
    </div>
  </div>

  <!-- SECTION 6: CLASSIFY LICENSE GROUPS -->
  <div class="card" id="classifyCard" style="display:none;">
    <h2>6. Classify License Groups (Faculty vs Student)</h2>
    <div style="font-size:12px; color:#6b7280; margin-bottom:10px;">
      Every distinct Job Title/Role found in your file is listed below. Choose whether accounts with that role
      should get a <strong>Faculty</strong> or <strong>Student</strong> license. Teaching-sounding roles are
      pre-selected as Faculty; everything else (Registrar, Guidance, Nurse, Cashier, etc.) defaults to Student —
      adjust any row, then click Apply.
    </div>
    <div class="table-wrap" style="max-height:300px;">
      <table id="classifyTable">
        <thead><tr><th>Role / Job Title</th><th># Users</th><th>License Group</th></tr></thead>
        <tbody id="classifyBody"></tbody>
      </table>
    </div>
    <button id="applyClassificationBtn" style="margin-top:10px;">✓ Apply Classification</button>
    <div class="status" id="classifyStatus"></div>
  </div>

  <!-- SECTION 7: PREVIEW + EXPORT -->
  <div class="card" id="previewCard" style="display:none;">
    <h2>7. Preview &amp; Export</h2>
    <button id="downloadBtn" class="success">⬇ Download CSV (O365_Users_Export.csv)</button>
    <div class="table-wrap">
      <table id="previewTable"></table>
    </div>
  </div>
</div>`;

  var _booted = false;

  // Tenant/month filename prefix, e.g. BSCS_SEP_ (same convention as SOC exports).
  function ugFilePrefix() {
    var t = (document.getElementById('ug-tenant') || {}).value || '';
    var m = (document.getElementById('ug-month') || {}).value || '';
    t = t.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    m = m.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!t && !m) return '';
    return (t || 'NA') + '_' + (m || 'NA') + '_';
  }

  function ugBoot() {
  /* =========================================================
     M365 Bulk User CSV Generator - core logic
     ========================================================= */

  let workbook = null;       // parsed XLSX workbook
  let outputRows = [];       // final dataset (array of objects)
  let validationErrors = []; // collected per-row error strings
  let skippedRows = [];      // audit trail: one entry per row excluded from output
  let headerColumnOptionsHtml = '<option value="literal">-- Literal text --</option>'; // populated once a file is loaded

  const HEADERS = [
    'GivenName','Surname','DisplayName','UserPrincipalName',
    'MailNickname','Password','UsageLocation','JobTitle','Department','LicenseGroup'
  ];

  /* ---------- Utility helpers ---------- */
  const $ = id => document.getElementById(id);
  const setStatus = (el, msg, cls) => { el.className = 'status ' + cls; el.textContent = msg; };

  // Strip accents/diacritics, remove spaces, keep only [a-z0-9.]
  function sanitizeForEmail(s) {
    return (s || '')
      .toString()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove diacritics
      .toLowerCase()
      .replace(/\s+/g, '')                              // remove spaces
      .replace(/[^a-z0-9.]/g, '');                      // strip special chars except .
  }

  // Student number: remove hyphens (2026-06908 -> 202606908)
  function cleanStudentNumber(sn) {
    return (sn || '').toString().trim().replace(/-/g, '');
  }

  // Last 4 digits of cleaned student number
  function last4(sn) {
    const c = cleanStudentNumber(sn);
    return c.slice(-4);
  }

  // Build UPN local-part from chosen format
  function buildLocalPart(given, surname, studentNum, fmt, customFormat) {
    const g = sanitizeForEmail(given);
    const s = sanitizeForEmail(surname);
    const sn = cleanStudentNumber(studentNum);
    switch (fmt) {
      case '1': return `${g}.${s}`;
      case '2': return `${s}.${g}.${sn}`;
      case '3': return `${g}.${s}.${last4(sn)}`;
      case '4': return `${s}.${g}`;
      case '5': return `${s}.${g}.${last4(sn)}`;
      case '6': return (studentNum || '').toString().trim().toLowerCase().replace(/\s+/g, '');
      case '7': return sn.toLowerCase();
      case '8': return `${g}_${s}${last4(sn)}`;
      case '9': return applyCustomFormat(customFormat, given, surname, studentNum);
      default:  return `${g}.${s}`;
    }
  }

  // Replace {tokens} in a user-defined custom pattern, then sanitize the result.
  // Supported tokens (case-insensitive): {given} {surname} {last4} {studentnumber}
  // {studentnumberhyphen} {giveninitial} {surnameinitial}
  function applyCustomFormat(pattern, given, surname, studentNum) {
    const g  = sanitizeForEmail(given);
    const s  = sanitizeForEmail(surname);
    const sn = cleanStudentNumber(studentNum);
    const snHyphen = (studentNum || '').toString().trim().toLowerCase().replace(/\s+/g, '');
    const tokens = {
      given: g, firstname: g,
      surname: s, lastname: s,
      last4: last4(sn), sn4: last4(sn),
      studentnumber: sn, sn: sn,
      studentnumberhyphen: snHyphen, snhyphen: snHyphen,
      giveninitial: g.charAt(0), firstinitial: g.charAt(0),
      surnameinitial: s.charAt(0), lastinitial: s.charAt(0)
    };
    const raw = (pattern || '').replace(/\{([a-zA-Z0-9_]+)\}/g, (match, key) => {
      const k = key.toLowerCase();
      return tokens.hasOwnProperty(k) ? tokens[k] : match;
    });
    // Keep only email-safe local-part characters (letters, digits, dot, underscore, hyphen)
    return raw.toLowerCase().replace(/[^a-z0-9._-]/g, '');
  }

  // Build the DisplayName from the chosen Display Name Format.
  // dnPartsSpec (only used for fmt '5'): array of { type:'literal', text } or { type:'column', colIndex }
  // Normalize a name to standard Proper Case regardless of source casing
  // (handles ALL CAPS, all lowercase, and mixed case source data).
  // "DELA CRUZ" -> "Dela Cruz", "o'brien" -> "O'Brien", "santos-reyes" -> "Santos-Reyes"
  function toProperCase(str) {
    return (str || '')
      .toString()
      .toLowerCase()
      .replace(/(^|[\s,.'"()-])(\p{L})/gu, (m, sep, letter) => sep + letter.toUpperCase());
  }

  // Build the DisplayName from the chosen Display Name Format.
  // dnPartsSpec (only used for fmt '5'): array of { type:'literal', text } or { type:'column', colIndex }
  function buildDisplayName(fmt, given, surname, mappedDisplay, row, get, dnPartsSpec, allCaps, properCase) {
    const g = toProperCase(given);
    const s = toProperCase(surname);
    switch (fmt) {
      case '0': return mappedDisplay || `${g} ${s}`;
      case '1': return `${g} ${s}`;
      case '2': return `${s}, ${g}`;
      case '3': return `${given} ${surname}`.toUpperCase();
      case '4': return `${surname}, ${given}`.toUpperCase();
      case '5': {
        let result = (dnPartsSpec || []).map(p => p.type === 'literal' ? p.text : get(row, p.colIndex)).join('');
        if (!result.trim()) result = mappedDisplay || `${g} ${s}`;
        if (allCaps) return result.toUpperCase();
        if (properCase) return toProperCase(result);
        return result;
      }
      default: return mappedDisplay || `${g} ${s}`;
    }
  }

  // CSV-escape a single field
  function csvEscape(v) {
    if (v == null) return '';
    const s = String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  // Default License Group guess from a Job Title/Role string.
  // Teaching-sounding roles -> Faculty, everything else (Registrar, Guidance, Nurse, Cashier, etc.) -> Student.
  function classifyRole(role) {
    return /teach|instructor|faculty|professor/i.test(role || '') ? 'Faculty' : 'Student';
  }

  /* ---------- Sheet + input-mode helpers ---------- */
  let lastMappedSheet = null; // tracks which sheet the column mapping was last built from

  // Usable sheets = every sheet except the placeholder "uploading" sheet.
  function getUsableSheets() {
    return workbook ? workbook.SheetNames.filter(n => n.toLowerCase() !== 'uploading') : [];
  }

  // Currently ticked sheets (falls back to all usable sheets when no selector is shown).
  function getSelectedSheets() {
    const usable = getUsableSheets();
    const boxes = [...document.querySelectorAll('#sheetCheckboxes .sheetChk')];
    if (!boxes.length) return usable;
    const chosen = boxes.filter(b => b.checked).map(b => b.value);
    return usable.filter(n => chosen.includes(n));
  }

  // Which email formats actually consume the student number.
  function formatNeedsStudentNumber(fmt, customFormat) {
    if (['2', '3', '5', '6', '7', '8'].includes(fmt)) return true;
    if (fmt === '9') return /\{(last4|sn4|studentnumber|sn|studentnumberhyphen|snhyphen)\}/i.test(customFormat || '');
    return false;
  }

  function currentInputMode() {
    const el = document.querySelector('input[name="inputMode"]:checked');
    return el ? el.value : 'excel';
  }

  // Build the per-sheet checkbox list (only when there is more than one usable sheet).
  function buildSheetSelector(sheets) {
    const wrap = $('sheetSelectWrap');
    const box  = $('sheetCheckboxes');
    if (sheets.length <= 1) { wrap.style.display = 'none'; box.innerHTML = ''; return; }
    wrap.style.display = 'block';
    box.innerHTML = sheets.map(n => {
      const safeVal = n.replace(/"/g, '&quot;');
      const safeTxt = n.replace(/</g, '&lt;');
      return `<label style="display:flex; align-items:center; gap:6px; font-weight:400; font-size:13px; margin:0;">
        <input type="checkbox" class="sheetChk" value="${safeVal}" checked style="width:auto;"> ${safeTxt}
      </label>`;
    }).join('');
    $('sheetSelectAll').checked = true;
    box.querySelectorAll('.sheetChk').forEach(chk => chk.addEventListener('change', onSheetSelectionChange));
  }

  // When the sheet selection changes, keep "select all" in sync and re-map columns
  // off the first selected sheet (only if that first sheet actually changed).
  function onSheetSelectionChange() {
    const sel = getSelectedSheets();
    const all = getUsableSheets();
    $('sheetSelectAll').checked = (all.length > 0 && sel.length === all.length);
    const first = sel[0] || null;
    if (first && first !== lastMappedSheet) buildColumnMapping(first);
    if ($('displayNameFormat').value === '5') updateDnPreview();
  }

  $('sheetSelectAll').addEventListener('change', () => {
    const checked = $('sheetSelectAll').checked;
    document.querySelectorAll('#sheetCheckboxes .sheetChk').forEach(c => { c.checked = checked; });
    onSheetSelectionChange();
  });

  /* ---------- Input mode (Excel vs Manual) wiring ---------- */
  function onModeChange() {
    const mode = currentInputMode();
    $('excelPanel').style.display  = mode === 'excel'  ? 'block' : 'none';
    $('manualPanel').style.display = mode === 'manual' ? 'block' : 'none';
    $('processBtn').textContent = mode === 'manual' ? '▶ Process Users' : '▶ Process Excel';
    if (mode === 'manual') {
      if (!document.querySelector('#manualBody tr')) addManualRow();
      updateManualCount();
      $('mappingCard').style.display = 'none'; // manual entry doesn't use Excel column mapping
      $('processBtn').disabled = false;
    } else {
      $('mappingCard').style.display = workbook ? 'block' : 'none';
      $('processBtn').disabled = !workbook;
    }
  }
  document.querySelectorAll('input[name="inputMode"]').forEach(r => r.addEventListener('change', onModeChange));

  /* ---------- Manual entry rows ---------- */
  // Field order of the manual table (also the positional paste order).
  const MANUAL_FIELDS = ['first', 'last', 'sn', 'grade', 'job', 'display'];

  function addManualRow(data) {
    const tb = $('manualBody');
    const tr = document.createElement('tr');
    const cell = (f, ph, w) =>
      `<td><input type="text" data-f="${f}" placeholder="${ph || ''}" style="min-width:${w || 110}px; padding:5px 7px;"></td>`;
    tr.innerHTML =
      cell('first', 'Juan', 110) +
      cell('last', 'Dela Cruz', 120) +
      cell('sn', '2026-06908', 110) +
      cell('grade', 'e.g. Grade 7', 100) +
      cell('job', 'e.g. Student', 110) +
      cell('display', '(optional)', 130) +
      `<td><button type="button" class="manualRemove" title="Remove row" style="background:#dc2626; padding:5px 9px; margin:0;">✕</button></td>`;
    tb.appendChild(tr);
    // Seed initial values (used by bulk-paste / grid-paste)
    if (data && typeof data === 'object') {
      MANUAL_FIELDS.forEach(f => {
        if (data[f] != null) { const el = tr.querySelector(`[data-f="${f}"]`); if (el) el.value = data[f]; }
      });
    }
    tr.querySelector('.manualRemove').addEventListener('click', () => { tr.remove(); updateManualCount(); });
    updateManualCount();
    return tr;
  }

  function updateManualCount() {
    const n = document.querySelectorAll('#manualBody tr').length;
    const el = $('manualCount');
    if (!el) return;
    let msg = `${n} user row${n === 1 ? '' : 's'}`;
    if (n > 10) msg += ' — manual entry is intended for fewer than 10; an Excel upload is easier past that.';
    el.textContent = msg;
    el.style.color = n > 10 ? '#b45309' : '#6b7280';
  }

  // Collect non-empty manual rows as raw records for the shared generator.
  function getManualRawRows() {
    const rows = [];
    document.querySelectorAll('#manualBody tr').forEach((tr, i) => {
      const val = f => (tr.querySelector(`[data-f="${f}"]`)?.value || '').trim();
      const first = val('first'), last = val('last'), sn = val('sn'),
            grade = val('grade'), jobCol = val('job'), display = val('display');
      if (!first && !last && !sn && !grade && !jobCol && !display) return; // skip fully blank rows
      rows.push({ sheet: 'Manual Entry', rowNum: i + 1, sn, first, last, display, grade, jobCol, rowArr: [], get: () => '' });
    });
    return rows;
  }

  $('manualAddBtn').addEventListener('click', () => addManualRow());

  /* ---------- Bulk paste (Manual mode) ---------- */
  // Header-name patterns per field. Used to detect a header row and to map
  // columns by name (so the pasted column ORDER doesn't have to match the table,
  // and missing columns like Student Number are simply left blank).
  const HEADER_PATTERNS = {
    first:   [/first\s*name/i, /given\s*name/i, /^first$/i],
    last:    [/last\s*name/i, /surname/i, /family\s*name/i, /^last$/i],
    sn:      [/student\s*(no|number|id)/i, /^lrn$/i, /^id$/i],
    grade:   [/grade\s*level/i, /^grade$/i, /year\s*level/i, /^level$/i, /section/i, /department/i],
    job:     [/job\s*title/i, /^role$/i, /position/i, /account\s*type/i, /user\s*type/i, /student\s*type/i],
    display: [/display\s*name/i, /full\s*name/i, /name\s*of\s*student/i, /^name$/i]
  };

  // Quote-aware splitter for one line, given a delimiter.
  function splitDelimitedLine(line, delim) {
    const out = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else inQ = false; }
        else cur += ch;
      } else {
        if (ch === '"') inQ = true;
        else if (ch === delim) { out.push(cur); cur = ''; }
        else cur += ch;
      }
    }
    out.push(cur);
    return out.map(c => c.trim());
  }

  // Turn a pasted blob into a 2D array of cells. Auto-detects tab / comma / semicolon.
  function parseDelimited(text) {
    const norm = (text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let lines = norm.split('\n');
    while (lines.length && lines[lines.length - 1].trim() === '') lines.pop(); // drop trailing blank lines
    lines = lines.filter(l => l.trim() !== '');
    if (!lines.length) return { rows: [], delimiter: '' };
    const sample = lines[0];
    let delim = '\t';
    if (sample.includes('\t')) delim = '\t';
    else if (sample.includes(',')) delim = ',';
    else if (sample.includes(';')) delim = ';';
    return { rows: lines.map(l => splitDelimitedLine(l, delim)), delimiter: delim };
  }

  // >=2 cells matching known field-header names => treat as a header row.
  function looksLikeHeader(cells) {
    let hits = 0;
    cells.forEach(c => {
      for (const f in HEADER_PATTERNS) {
        if (HEADER_PATTERNS[f].some(p => p.test(c))) { hits++; break; }
      }
    });
    return hits >= 2;
  }

  // Build a {field: columnIndex} map from a header row (first match wins; -1 if absent).
  function mapHeaderColumns(headerCells) {
    const map = {};
    MANUAL_FIELDS.forEach(f => { map[f] = -1; });
    headerCells.forEach((cell, idx) => {
      for (const f of MANUAL_FIELDS) {
        if (map[f] === -1 && HEADER_PATTERNS[f].some(p => p.test(cell))) { map[f] = idx; break; }
      }
    });
    return map;
  }

  // Parse the paste box and REPLACE the manual rows.
  function parseManualPaste() {
    const st = $('manualPasteStatus');
    const { rows } = parseDelimited($('manualPasteBox').value);
    if (!rows.length) { setStatus(st, '✗ Nothing to parse — paste your rows into the box first.', 'err'); return; }

    let dataRows = rows, colMap = null;
    if (looksLikeHeader(rows[0])) { colMap = mapHeaderColumns(rows[0]); dataRows = rows.slice(1); }
    if (!dataRows.length) { setStatus(st, '✗ Only a header row was found — no data rows to load.', 'err'); return; }

    // Warn if a header was present but First/Last couldn't be located by name
    if (colMap && (colMap.first === -1 || colMap.last === -1)) {
      setStatus(st, '✗ Could not find a "First Name" and "Last Name" column in the pasted header. Check the header names and try again.', 'err');
      return;
    }

    $('manualBody').innerHTML = ''; // REPLACE existing rows
    let missing = 0;
    dataRows.forEach(cells => {
      const pick = f => {
        const idx = colMap ? colMap[f] : MANUAL_FIELDS.indexOf(f);
        return (idx >= 0 && idx < cells.length) ? cells[idx] : '';
      };
      const vals = { first: pick('first'), last: pick('last'), sn: pick('sn'),
                     grade: pick('grade'), job: pick('job'), display: pick('display') };
      if (!vals.first || !vals.last) missing++;
      addManualRow(vals);
    });
    if (!document.querySelector('#manualBody tr')) addManualRow(); // never leave the table empty
    updateManualCount();

    const how = colMap ? 'matched by header names' : 'mapped by column order';
    let msg = `✓ Loaded ${dataRows.length} row(s) (${how}).`;
    if (missing) msg += ` ${missing} row(s) are missing a First or Last name — fix those before processing.`;
    setStatus(st, msg, missing ? 'warn' : 'ok');
  }

  $('manualParseBtn').addEventListener('click', parseManualPaste);
  $('manualPasteClearBtn').addEventListener('click', () => {
    $('manualPasteBox').value = '';
    $('manualPasteStatus').className = 'status';
    $('manualPasteStatus').textContent = '';
  });

  // Excel-style grid paste: Ctrl+V a multi-cell block onto any cell and it spills
  // across rows/columns from that cell, adding rows as needed. (No header skipping —
  // this mirrors Excel: whatever you copied is what lands.)
  $('manualBody').addEventListener('paste', e => {
    const target = e.target;
    if (!target || !target.matches || !target.matches('input[data-f]')) return;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text || !/[\t\n\r]/.test(text)) return; // single value -> let the browser paste normally
    e.preventDefault();
    const { rows } = parseDelimited(text);
    if (!rows.length) return;
    const startFieldIdx = MANUAL_FIELDS.indexOf(target.dataset.f);
    const startRowIdx = [...$('manualBody').querySelectorAll('tr')].indexOf(target.closest('tr'));
    rows.forEach((cells, r) => {
      let tr = $('manualBody').querySelectorAll('tr')[startRowIdx + r];
      if (!tr) tr = addManualRow();
      cells.forEach((val, c) => {
        const fieldIdx = startFieldIdx + c;
        if (fieldIdx < 0 || fieldIdx >= MANUAL_FIELDS.length) return;
        const el = tr.querySelector(`[data-f="${MANUAL_FIELDS[fieldIdx]}"]`);
        if (el) el.value = val;
      });
    });
    updateManualCount();
  });

  /* ---------- Custom Format panel wiring ---------- */
  function updateCustomFormatExample() {
    const pattern = $('customFormatInput').value.trim();
    const preview = pattern
      ? applyCustomFormat(pattern, 'Juan', 'Dela Cruz', '2026-06908')
      : '(enter a pattern above)';
    $('customFormatExample').textContent = preview;
  }
  $('emailFormat').addEventListener('change', () => {
    const isCustom = $('emailFormat').value === '9';
    $('customFormatPanel').style.display = isCustom ? 'block' : 'none';
    if (isCustom) updateCustomFormatExample();
  });
  $('customFormatInput').addEventListener('input', updateCustomFormatExample);

  /* ---------- Custom Display Name panel wiring ---------- */
  function addDnPart(initialCol, initialLiteral) {
    const container = $('dnPartsContainer');
    const row = document.createElement('div');
    row.className = 'dnPartRow';
    row.style.cssText = 'display:flex; gap:6px; align-items:center; margin-bottom:6px;';
    row.innerHTML = `
      <select class="dnPartSelect" style="flex:1;">${headerColumnOptionsHtml}</select>
      <input type="text" class="dnPartLiteral" placeholder="literal text (e.g. space or comma)" style="flex:1; display:none;">
      <button type="button" class="dnRemovePartBtn" style="background:#dc2626; padding:6px 10px; margin:0;">✕</button>
    `;
    container.appendChild(row);
    const sel = row.querySelector('.dnPartSelect');
    const lit = row.querySelector('.dnPartLiteral');
    if (initialLiteral !== undefined) {
      sel.value = 'literal';
      lit.value = initialLiteral;
    } else if (initialCol !== undefined && initialCol >= 0) {
      sel.value = String(initialCol);
    }
    lit.style.display = sel.value === 'literal' ? 'block' : 'none';
    sel.addEventListener('change', () => { lit.style.display = sel.value === 'literal' ? 'block' : 'none'; updateDnPreview(); });
    lit.addEventListener('input', updateDnPreview);
    row.querySelector('.dnRemovePartBtn').addEventListener('click', () => { row.remove(); updateDnPreview(); });
  }

  function getDnPartsSpec() {
    return [...document.querySelectorAll('#dnPartsContainer .dnPartRow')].map(row => {
      const sel = row.querySelector('.dnPartSelect');
      const lit = row.querySelector('.dnPartLiteral');
      if (sel.value === 'literal') return { type: 'literal', text: lit.value };
      return { type: 'column', colIndex: parseInt(sel.value, 10) };
    });
  }

  function updateDnPreview() {
    const previewEl = $('dnPreview');
    if (!previewEl) return;
    const parts = getDnPartsSpec();
    if (!parts.length) { previewEl.textContent = '(add at least one part above)'; return; }
    if (!workbook) { previewEl.textContent = '(upload a file to preview with real data)'; return; }
    const sheets = getSelectedSheets();
    if (!sheets.length) { previewEl.textContent = ''; return; }
    const ws = workbook.Sheets[sheets[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    const headerRowIdx = Math.max(0, parseInt($('headerRowIdx').value, 10) || 0);
    const sampleRow = rows.find((r, i) => i !== headerRowIdx && r.some(c => (c || '').toString().trim() !== '')) || [];
    const get = (row, idx) => idx >= 0 ? (row[idx] || '').toString().trim() : '';
    let result = parts.map(p => p.type === 'literal' ? p.text : get(sampleRow, p.colIndex)).join('');
    if ($('dnCustomAllCaps').checked) result = result.toUpperCase();
    else if ($('dnCustomProperCase').checked) result = toProperCase(result);
    previewEl.textContent = result || '(configure parts above)';
  }

  $('displayNameFormat').addEventListener('change', () => {
    const isCustom = $('displayNameFormat').value === '5';
    $('dnCustomPanel').style.display = isCustom ? 'block' : 'none';
    if (isCustom) {
      if (!document.querySelector('#dnPartsContainer .dnPartRow')) addDnPart();
      updateDnPreview();
    }
  });
  $('dnAddPartBtn').addEventListener('click', () => { addDnPart(); updateDnPreview(); });
  $('dnCustomAllCaps').addEventListener('change', updateDnPreview);
  $('dnCustomProperCase').addEventListener('change', updateDnPreview);

  /* ---------- File load ---------- */
  $('fileInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    const fs = $('fileStatus');
    setStatus(fs, 'Reading file...', 'info');
    try {
      const buf = await file.arrayBuffer();
      workbook = XLSX.read(buf, { type: 'array' });
      const sheets = workbook.SheetNames.filter(n => n.toLowerCase() !== 'uploading');
      if (!sheets.length) {
        setStatus(fs, '✗ No usable sheets found (all skipped or empty).', 'err');
        $('processBtn').disabled = true; return;
      }
      setStatus(fs, `✓ Loaded "${file.name}". ${sheets.length} usable sheet(s): ${sheets.join(', ')}${sheets.length > 1 ? ' — pick which to process below.' : ''}`, 'ok');
      buildSheetSelector(sheets);
      buildColumnMapping(sheets[0]);
      $('mappingCard').style.display = 'block';
      $('processBtn').disabled = false;
    } catch (err) {
      setStatus(fs, '✗ Failed to read file: ' + err.message, 'err');
      workbook = null;
      $('processBtn').disabled = true;
      $('mappingCard').style.display = 'none';
    }
  });

  /* ---------- Column mapping ---------- */
  // Convert 0-based index to Excel column letter (0->A, 25->Z, 26->AA)
  function colLetter(idx) {
    let s = '';
    idx = idx + 1;
    while (idx > 0) {
      const m = (idx - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      idx = Math.floor((idx - 1) / 26);
    }
    return s;
  }

  // Auto-pick a column by matching header text against patterns (returns index or -1)
  function autoPick(headers, patterns) {
    for (let i = 0; i < headers.length; i++) {
      const h = (headers[i] || '').toString().toLowerCase().trim();
      if (!h) continue;
      if (patterns.some(p => p.test(h))) return i;
    }
    return -1;
  }

  function buildColumnMapping(firstSheetName) {
    lastMappedSheet = firstSheetName;
    const ws = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if (!rows.length) return;

    // Find the widest row in the first 5 rows (to determine column count)
    let maxCols = 0;
    for (let i = 0; i < Math.min(5, rows.length); i++) maxCols = Math.max(maxCols, rows[i].length);

    // Use row 0 as header row by default; user can override via headerRowIdx
    const headerRow = rows[0] || [];

    // Build options: "A — Student Number" style
    const opts = ['<option value="-1">— Not in file —</option>'];
    for (let i = 0; i < maxCols; i++) {
      const label = (headerRow[i] || '').toString().trim() || '(empty)';
      opts.push(`<option value="${i}">${colLetter(i)} — ${label.replace(/</g,'&lt;')}</option>`);
    }
    const html = opts.join('');

    const fields = ['mapStudentNumber','mapFirstName','mapLastName','mapDisplayName','mapDepartment','mapJobTitle'];
    fields.forEach(id => { $(id).innerHTML = html; });

    // Auto-detect best matches based on common header names; fall back to original fixed positions
    const autoMap = {
      mapStudentNumber: autoPick(headerRow, [/student\s*(no|number|id)/i, /^lrn$/i, /^id$/i]),
      mapFirstName:     autoPick(headerRow, [/first\s*name/i, /given\s*name/i, /^first$/i]),
      mapLastName:      autoPick(headerRow, [/last\s*name/i, /surname/i, /family\s*name/i, /^last$/i]),
      mapDisplayName:   autoPick(headerRow, [/display\s*name/i, /full\s*name/i, /name\s*of\s*student/i, /^name$/i]),
      mapDepartment:    autoPick(headerRow, [/grade\s*level/i, /^grade$/i, /^level$/i, /year\s*level/i, /department/i, /section/i]),
      mapJobTitle:      autoPick(headerRow, [/job\s*title/i, /^role$/i, /position/i, /student\s*type/i, /account\s*type/i, /user\s*type/i])
    };
    // Fallback to original fixed positions (A,F,E,D,L) if no header match
    const fallback = { mapStudentNumber:0, mapFirstName:5, mapLastName:4, mapDisplayName:3, mapDepartment:11, mapJobTitle:-1 };
    let effectiveFirstIdx = -1, effectiveLastIdx = -1;
    Object.keys(autoMap).forEach(id => {
      const idx = autoMap[id] >= 0 ? autoMap[id] : (fallback[id] < maxCols ? fallback[id] : -1);
      $(id).value = String(idx);
      if (id === 'mapFirstName') effectiveFirstIdx = idx;
      if (id === 'mapLastName')  effectiveLastIdx = idx;
    });
    $('headerRowIdx').value = '0';

    // Build the column list used by the Custom Display Name part builder, then reseed it
    // with a sensible starting template ("Last, First") based on the auto-detected columns.
    const partOpts = ['<option value="literal">-- Literal text --</option>'];
    for (let i = 0; i < maxCols; i++) {
      const label = (headerRow[i] || '').toString().trim() || '(empty)';
      partOpts.push(`<option value="${i}">${colLetter(i)} — ${label.replace(/</g,'&lt;')}</option>`);
    }
    headerColumnOptionsHtml = partOpts.join('');
    $('dnPartsContainer').innerHTML = '';
    addDnPart(effectiveLastIdx >= 0 ? effectiveLastIdx : undefined);
    addDnPart(undefined, ', ');
    addDnPart(effectiveFirstIdx >= 0 ? effectiveFirstIdx : undefined);
    if ($('displayNameFormat').value === '5') updateDnPreview();
  }

  /* ---------- Shared record generator (used by both Excel and Manual modes) ----------
     rawRows: [{ sheet, rowNum, sn, first, last, display, grade, jobCol, rowArr, get }]
     cfg:     { domain, password, jobTitle, fmt, customFormat, displayNameFmt, dnPartsSpec,
                dnCustomAllCaps, dnCustomProperCase, dedupeNames, needsSN } */
  function generateOutput(rawRows, cfg) {
    outputRows = [];
    validationErrors = [];
    skippedRows = [];
    const seenUPN = new Set();
    const seenSN  = new Set();
    const seenName = new Set();

    rawRows.forEach(rr => {
      const { sheet, rowNum, sn, first, last, display, grade, jobCol, rowArr, get } = rr;

      // Required fields: First + Last always; Student Number only when the email format uses it.
      if (!first || !last || (cfg.needsSN && !sn)) {
        skippedRows.push({ Sheet: sheet, Row: rowNum,
          Reason: cfg.needsSN ? 'Missing Student Number, First Name, or Last Name' : 'Missing First Name or Last Name',
          StudentNumber: sn, GivenName: first, Surname: last, DisplayName: display, GeneratedEmail: '' });
        return;
      }
      // Skip if the student-number cell is literally the header text (stray Excel headers)
      if (/student\s*number/i.test(sn)) {
        skippedRows.push({ Sheet: sheet, Row: rowNum, Reason: 'Student Number cell contains header text',
          StudentNumber: sn, GivenName: first, Surname: last, DisplayName: display, GeneratedEmail: '' });
        return;
      }

      // Build UPN
      const localPart = buildLocalPart(first, last, sn, cfg.fmt, cfg.customFormat);
      const upn = `${localPart}@${cfg.domain}`;
      const mnk = localPart; // MailNickname == prefix

      // Dedupe by UPN or Student Number (cleaned)
      const snClean = cleanStudentNumber(sn);
      if (seenUPN.has(upn)) {
        skippedRows.push({ Sheet: sheet, Row: rowNum, Reason: 'Duplicate generated email (UPN) — already used by an earlier row',
          StudentNumber: sn, GivenName: first, Surname: last, DisplayName: display, GeneratedEmail: upn });
        return;
      }
      if (snClean && seenSN.has(snClean)) {
        skippedRows.push({ Sheet: sheet, Row: rowNum, Reason: 'Duplicate Student Number — already used by an earlier row',
          StudentNumber: sn, GivenName: first, Surname: last, DisplayName: display, GeneratedEmail: upn });
        return;
      }

      // Optional dedupe by name
      const nameKey = (first + '|' + last).toLowerCase();
      if (cfg.dedupeNames && seenName.has(nameKey)) {
        skippedRows.push({ Sheet: sheet, Row: rowNum, Reason: 'Duplicate Name (First+Last) — already used by an earlier row',
          StudentNumber: sn, GivenName: first, Surname: last, DisplayName: display, GeneratedEmail: upn });
        return;
      }

      // Grade Level is optional — warn (do not skip) when missing
      if (!grade) {
        validationErrors.push(`${sheet} row ${rowNum}: missing Grade Level (Department) for ${display || upn}`);
      }

      const rec = {
        GivenName: first,
        Surname:  last,
        DisplayName: buildDisplayName(cfg.displayNameFmt, first, last, display, rowArr, get, cfg.dnPartsSpec, cfg.dnCustomAllCaps, cfg.dnCustomProperCase),
        UserPrincipalName: upn,
        MailNickname: mnk,
        Password: cfg.password,
        UsageLocation: 'PH',
        JobTitle: jobCol || cfg.jobTitle,
        Department: grade
      };

      // Per-row validation
      if (/\s/.test(rec.UserPrincipalName)) validationErrors.push(`${sheet} row ${rowNum}: UPN contains spaces`);
      if (rec.UserPrincipalName !== rec.UserPrincipalName.toLowerCase()) validationErrors.push(`${sheet} row ${rowNum}: UPN not lowercase`);
      if (rec.MailNickname !== rec.UserPrincipalName.split('@')[0]) validationErrors.push(`${sheet} row ${rowNum}: MailNickname mismatch`);

      seenUPN.add(upn);
      if (snClean) seenSN.add(snClean);
      seenName.add(nameKey);
      outputRows.push(rec);
    });

    return { scanned: rawRows.length, skipped: skippedRows.length };
  }

  /* ---------- Process ---------- */
  $('processBtn').addEventListener('click', () => {
    const ps = $('procStatus');
    const mode = currentInputMode();

    // Shared configuration
    const domain   = $('domain').value.trim().toLowerCase();
    const password = $('password').value;
    const jobTitle = $('jobTitle').value.trim();
    const fmt      = $('emailFormat').value;
    const customFormat = $('customFormatInput').value.trim();
    const displayNameFmt = $('displayNameFormat').value;
    const dnPartsSpec = displayNameFmt === '5' ? getDnPartsSpec() : null;
    const dnCustomAllCaps = $('dnCustomAllCaps').checked;
    const dnCustomProperCase = $('dnCustomProperCase').checked;
    const dedupeNames = $('dedupeNames').checked;
    const needsSN = formatNeedsStudentNumber(fmt, customFormat);

    // Basic input validation
    if (!domain)   { setStatus(ps, '✗ Domain is required.', 'err'); return; }
    if (!password) { setStatus(ps, '✗ Password is required.', 'err'); return; }
    if (fmt === '9' && !customFormat) { setStatus(ps, '✗ Custom Pattern is required when Email Format is set to Custom.', 'err'); return; }
    if (displayNameFmt === '5' && (!dnPartsSpec || !dnPartsSpec.length)) { setStatus(ps, '✗ Add at least one part when Display Name Format is set to Custom.', 'err'); return; }

    const cfg = { domain, password, jobTitle, fmt, customFormat, displayNameFmt, dnPartsSpec, dnCustomAllCaps, dnCustomProperCase, dedupeNames, needsSN };

    let rawRows = [];
    let sourceLabel = '';

    if (mode === 'manual') {
      rawRows = getManualRawRows();
      if (!rawRows.length) { setStatus(ps, '✗ Add at least one user row with a First and Last name.', 'err'); return; }
      sourceLabel = `${rawRows.length} manual row(s)`;
      setStatus(ps, 'Processing manual entries...', 'info');
    } else {
      if (!workbook) { setStatus(ps, '✗ Upload an Excel file first.', 'err'); return; }
      const sheetsToProcess = getSelectedSheets();
      if (!sheetsToProcess.length) { setStatus(ps, '✗ Select at least one sheet to process.', 'err'); return; }

      // User-mapped column indices (-1 = not provided)
      const COL = {
        sn:      parseInt($('mapStudentNumber').value, 10),
        display: parseInt($('mapDisplayName').value, 10),
        last:    parseInt($('mapLastName').value, 10),
        first:   parseInt($('mapFirstName').value, 10),
        grade:   parseInt($('mapDepartment').value, 10),
        job:     parseInt($('mapJobTitle').value, 10)
      };
      const headerRowIdx = Math.max(0, parseInt($('headerRowIdx').value, 10) || 0);

      // Required-mapping guard. Student Number required only when the email format uses it.
      if (COL.first < 0 || COL.last < 0 || (needsSN && COL.sn < 0)) {
        setStatus(ps, needsSN
          ? '✗ Student Number, First Name, and Last Name columns must be mapped for the selected email format.'
          : '✗ First Name and Last Name columns must be mapped.', 'err');
        return;
      }

      setStatus(ps, 'Processing sheets...', 'info');
      $('progress').style.display = 'block';

      // Safe getter — returns '' if column not mapped
      const get = (row, idx) => idx >= 0 ? (row[idx] || '').toString().trim() : '';

      sheetsToProcess.forEach((sheetName, idx) => {
        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        rows.forEach((row, rIdx) => {
          if (rIdx === headerRowIdx) return; // skip the header row in every sheet
          rawRows.push({ sheet: sheetName, rowNum: rIdx + 1,
            sn: get(row, COL.sn), display: get(row, COL.display),
            last: get(row, COL.last), first: get(row, COL.first),
            grade: get(row, COL.grade), jobCol: get(row, COL.job),
            rowArr: row, get });
        });
        $('progressBar').style.width = Math.round(((idx + 1) / sheetsToProcess.length) * 100) + '%';
      });
      sourceLabel = `${sheetsToProcess.length} sheet(s)`;
    }

    const { scanned, skipped } = generateOutput(rawRows, cfg);

    setStatus(ps, `✓ Processed ${sourceLabel}. ${outputRows.length} user(s) ready, ${skipped} row(s) skipped.`, 'ok');
    renderStats(scanned, skipped);
    renderSkippedRows();

    // Default License Group classification (Faculty vs Student) - editable below before download
    outputRows.forEach(r => { r.LicenseGroup = classifyRole(r.JobTitle); });
    renderClassification();

    renderPreview();
    $('progress').style.display = 'none';
    $('progressBar').style.width = '0%';
  });

  /* ---------- Stats / Errors ---------- */
  function renderStats(scanned, skipped) {
    $('statsCard').style.display = 'block';
    $('stats').innerHTML = `
      <div class="stat"><div class="n">${scanned}</div><div class="l">Rows Scanned</div></div>
      <div class="stat"><div class="n">${outputRows.length}</div><div class="l">Users Generated</div></div>
      <div class="stat"><div class="n">${skipped}</div><div class="l">Rows Skipped</div></div>
      <div class="stat"><div class="n">${validationErrors.length}</div><div class="l">Validation Warnings</div></div>
    `;
    if (validationErrors.length) {
      $('errorsWrap').style.display = 'block';
      $('errorsBox').textContent = validationErrors.slice(0, 200).join('\n') +
        (validationErrors.length > 200 ? `\n... and ${validationErrors.length-200} more` : '');
    } else {
      $('errorsWrap').style.display = 'none';
    }
  }

  /* ---------- Skipped Rows Audit ---------- */
  const SKIPPED_HEADERS = ['Sheet','Row','Reason','StudentNumber','GivenName','Surname','DisplayName','GeneratedEmail'];

  function renderSkippedRows() {
    if (!skippedRows.length) { $('skippedCard').style.display = 'none'; return; }
    $('skippedCard').style.display = 'block';
    const t = $('skippedTable');
    let html = '<thead><tr>' + SKIPPED_HEADERS.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
    const cap = Math.min(skippedRows.length, 500);
    for (let i = 0; i < cap; i++) {
      html += '<tr>' + SKIPPED_HEADERS.map(h => `<td>${(skippedRows[i][h] ?? '').toString().replace(/</g,'&lt;')}</td>`).join('') + '</tr>';
    }
    html += '</tbody>';
    t.innerHTML = html;
    if (skippedRows.length > cap) {
      t.insertAdjacentHTML('beforeend', `<caption style="caption-side:bottom;color:#6b7280;font-size:11px;">Showing first ${cap} of ${skippedRows.length} skipped rows. Download the CSV for the full list.</caption>`);
    }
  }

  $('downloadSkippedBtn').addEventListener('click', () => {
    if (!skippedRows.length) return;
    const lines = [SKIPPED_HEADERS.join(',')];
    skippedRows.forEach(r => {
      lines.push(SKIPPED_HEADERS.map(h => csvEscape(r[h])).join(','));
    });
    const csv = '\uFEFF' + lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = ugFilePrefix() + 'Skipped_Rows_Audit.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  /* ---------- License Group Classification ---------- */
  function renderClassification() {
    // Tally distinct Job Title/Role values and a keyword-based default guess
    const counts = new Map(); // key: role text (trimmed, or '(blank)'), value: { count, guess }
    outputRows.forEach(r => {
      const role = (r.JobTitle || '').toString().trim() || '(blank)';
      if (!counts.has(role)) counts.set(role, { count: 0, guess: classifyRole(role) });
      counts.get(role).count++;
    });

    const body = $('classifyBody');
    let html = '';
    [...counts.keys()].sort((a, b) => a.localeCompare(b)).forEach(role => {
      const info = counts.get(role);
      html += `<tr>
        <td>${role.replace(/</g, '&lt;')}</td>
        <td>${info.count}</td>
        <td>
          <select class="roleGroupSelect" data-role="${role.replace(/"/g, '&quot;')}">
            <option value="Faculty" ${info.guess === 'Faculty' ? 'selected' : ''}>Faculty</option>
            <option value="Student" ${info.guess === 'Student' ? 'selected' : ''}>Student</option>
          </select>
        </td>
      </tr>`;
    });
    body.innerHTML = html;
    $('classifyCard').style.display = counts.size ? 'block' : 'none';
  }

  $('applyClassificationBtn').addEventListener('click', () => {
    const map = {};
    document.querySelectorAll('.roleGroupSelect').forEach(sel => {
      map[sel.dataset.role] = sel.value;
    });
    outputRows.forEach(r => {
      const role = (r.JobTitle || '').toString().trim() || '(blank)';
      r.LicenseGroup = map[role] || classifyRole(role);
    });
    renderPreview();
    setStatus($('classifyStatus'), `✓ Classification applied to ${outputRows.length} user(s).`, 'ok');
  });

  /* ---------- Preview ---------- */
  function renderPreview() {
    if (!outputRows.length) { $('previewCard').style.display = 'none'; return; }
    $('previewCard').style.display = 'block';
    const t = $('previewTable');
    let html = '<thead><tr>' + HEADERS.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
    // Show up to 500 rows in preview to keep DOM light
    const cap = Math.min(outputRows.length, 500);
    for (let i = 0; i < cap; i++) {
      html += '<tr>' + HEADERS.map(h => `<td>${(outputRows[i][h] ?? '').toString().replace(/</g,'&lt;')}</td>`).join('') + '</tr>';
    }
    html += '</tbody>';
    t.innerHTML = html;
    if (outputRows.length > cap) {
      t.insertAdjacentHTML('beforeend', `<caption style="caption-side:bottom;color:#6b7280;font-size:11px;">Showing first ${cap} of ${outputRows.length} rows.</caption>`);
    }
  }

  /* ---------- CSV Download ---------- */
  $('downloadBtn').addEventListener('click', () => {
    if (!outputRows.length) return;
    const lines = [HEADERS.join(',')];
    outputRows.forEach(r => {
      lines.push(HEADERS.map(h => csvEscape(r[h])).join(','));
    });
    // Prepend BOM for Excel UTF-8 compatibility
    const csv = '\uFEFF' + lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = ugFilePrefix() + 'O365_Users_Export.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  /* ---------- Clear ---------- */
  $('clearBtn').addEventListener('click', () => {
    workbook = null; outputRows = []; validationErrors = []; skippedRows = [];
    $('fileInput').value = '';
    $('fileStatus').className = 'status'; $('fileStatus').textContent = '';
    $('procStatus').className = 'status'; $('procStatus').textContent = '';
    $('statsCard').style.display = 'none';
    $('skippedCard').style.display = 'none';
    $('skippedTable').innerHTML = '';
    $('previewCard').style.display = 'none';
    $('mappingCard').style.display = 'none';
    $('classifyCard').style.display = 'none';
    $('classifyBody').innerHTML = '';
    $('classifyStatus').className = 'status'; $('classifyStatus').textContent = '';
    $('dnPartsContainer').innerHTML = '';
    headerColumnOptionsHtml = '<option value="literal">-- Literal text --</option>';
    // Reset sheet selector + manual entry + input mode
    lastMappedSheet = null;
    $('sheetSelectWrap').style.display = 'none';
    $('sheetCheckboxes').innerHTML = '';
    $('sheetSelectAll').checked = true;
    $('manualBody').innerHTML = '';
    $('manualPasteBox').value = '';
    $('manualPasteStatus').className = 'status'; $('manualPasteStatus').textContent = '';
    updateManualCount();
    const excelRadio = document.querySelector('input[name="inputMode"][value="excel"]');
    if (excelRadio) excelRadio.checked = true;
    onModeChange();
  });
  // Test hooks: expose pure logic functions.
  QBR._ugFns = { sanitizeForEmail, cleanStudentNumber, last4, buildLocalPart,
    applyCustomFormat, toProperCase, buildDisplayName, csvEscape, classifyRole };
  }

  function usergenInit() {
    if (_booted) return;
    var host = document.getElementById('um-generator');
    if (!host) return;
    host.innerHTML = UG_HTML;
    ugBoot();
    _booted = true;
  }

  window.QBR = window.QBR || {};
  QBR.usergenInit = usergenInit;
  // Test hooks (not part of the UI).
  QBR.usergenTest = function(name, fn) { QBR._ugTest = QBR._ugTest || {}; QBR._ugTest[name] = fn; };
  QBR.ugFilePrefix = ugFilePrefix;
})();
