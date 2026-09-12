// ============================================================
// CALENDARPICKER.GS — VISUAL CALENDAR DATE PICKER ENGINE
// ============================================================

/**
 * Opens a visual interactive calendar modal dialog to choose a date.
 *
 * @param {string} mode - 'tasks' (for task entry) or 'dsr' (for Daily Status Report)
 * @param {Date} [initialDate] - Pre-selected Date object (defaults to today)
 */
function openCalendarPicker(mode, initialDate) {
  const timezone = getTimezone();
  const date = initialDate || getToday();

  const currentYear = date.getFullYear();
  const currentMonth = date.getMonth();
  const currentDay = date.getDate();

  const dateInfo = {
    mode: mode || 'tasks',
    year: currentYear,
    monthIndex: currentMonth,
    day: currentDay,
    formattedDate: Utilities.formatDate(date, timezone, 'EEEE, dd MMMM yyyy'),
    isoDate: Utilities.formatDate(date, timezone, 'yyyy-MM-dd')
  };

  const htmlContent = getCalendarPickerHtml(dateInfo, mode);
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(410)
    .setHeight(530);

  const dialogTitle = mode === 'dsr'
    ? 'Choose Date for DSR'
    : 'Choose Date for Tasks';

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
}

/**
 * Client-callable function: Proceeds to the target workflow (Tasks or DSR)
 * once the user confirms a date from the calendar.
 *
 * @param {Object} payload - { mode: 'tasks' | 'dsr', dateStr: 'YYYY-MM-DD' }
 */
function proceedFromCalendar(payload) {
  if (!payload || !payload.dateStr) {
    return;
  }

  const parts = payload.dateStr.split('-');
  const year = parseInt(parts[0], 10);
  const monthIndex = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  const targetDate = new Date(year, monthIndex, day);

  if (payload.mode === 'tasks') {
    openTaskDialog(targetDate);
  } else if (payload.mode === 'dsr') {
    showDSRDialog(targetDate, false);
  }
}

/**
 * Generates HTML for the visual calendar picker modal.
 *
 * @param {Object} dateInfo
 * @param {string} mode
 * @returns {string} HTML content
 */
function getCalendarPickerHtml(dateInfo, mode) {
  const jsonDateInfo = JSON.stringify(dateInfo).replace(/<\/script/gi, '<\\/script');
  const actionButtonText = mode === 'dsr' ? '📋 Generate DSR' : '📝 Fill Tasks';
  const headerIcon = mode === 'dsr' ? '📊' : '⚡';
  const headerTitle = mode === 'dsr' ? 'Select Date for DSR' : 'Select Date for Tasks';

  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    html, body {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      overflow: hidden;
      background-color: #f8fafc;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      color: #1e293b;
      font-size: 13px;
      line-height: 1.4;
      user-select: none;
    }

    body {
      padding: 12px;
    }

    .cal-card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      padding: 14px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }

    .cal-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 10px;
      padding-bottom: 8px;
      border-bottom: 1px solid #f1f5f9;
    }

    .cal-header-title {
      display: flex;
      align-items: center;
      gap: 7px;
      font-size: 14px;
      font-weight: 700;
      color: #0f172a;
    }

    .cal-badge-mode {
      font-size: 11px;
      font-weight: 700;
      padding: 2px 8px;
      border-radius: 9999px;
      background: #ede9fe;
      color: #6366f1;
    }

    /* Month Navigation */
    .month-nav-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 12px;
    }

    .nav-btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 28px;
      height: 28px;
      border-radius: 6px;
      border: 1px solid #e2e8f0;
      background: #ffffff;
      color: #475569;
      cursor: pointer;
      font-size: 12px;
      font-weight: 700;
      transition: all 0.15s ease-in-out;
    }

    .nav-btn:hover {
      background: #f1f5f9;
      color: #0f172a;
      border-color: #cbd5e1;
    }

    .month-year-label {
      font-size: 14px;
      font-weight: 700;
      color: #0f172a;
    }

    /* Days of Week */
    .weekdays-row {
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      gap: 4px;
      text-align: center;
      margin-bottom: 6px;
    }

    .weekday-label {
      font-size: 11px;
      font-weight: 700;
      color: #64748b;
      padding: 4px 0;
      text-transform: uppercase;
      letter-spacing: 0.03em;
    }

    .weekday-label.weekend {
      color: #ef4444;
    }

    /* Days Grid */
    .days-grid {
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      gap: 3px;
      margin-bottom: 12px;
    }

    .day-cell {
      height: 32px;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 500;
      color: #1e293b;
      cursor: pointer;
      transition: all 0.12s ease-in-out;
      border: 1px solid transparent;
      position: relative;
    }

    .day-cell:hover {
      background-color: #f1f5f9;
      color: #0f172a;
    }

    .day-cell.other-month {
      color: #cbd5e1;
    }

    .day-cell.weekend:not(.other-month):not(.selected) {
      color: #b91c1c;
    }

    .day-cell.today {
      border-color: #6366f1;
      font-weight: 700;
      color: #4f46e5;
    }

    .day-cell.selected {
      background-color: #4f46e5 !important;
      color: #ffffff !important;
      font-weight: 700;
      box-shadow: 0 1px 3px rgba(79, 70, 229, 0.4);
    }

    /* Selected Date Display */
    .selected-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: #f1f5f9;
      border-radius: 8px;
      padding: 7px 12px;
      margin-bottom: 12px;
      font-size: 12px;
    }

    .selected-bar-title {
      color: #64748b;
      font-weight: 600;
    }

    .selected-bar-date {
      color: #0f172a;
      font-weight: 700;
    }

    /* Footer Buttons */
    .cal-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      padding: 7px 12px;
      font-size: 12px;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.15s ease-in-out;
      outline: none;
      border: 1px solid transparent;
    }

    .btn-primary {
      background-color: #4f46e5;
      color: #ffffff;
      flex: 1.4;
      box-shadow: 0 1px 2px rgba(79, 70, 229, 0.3);
    }

    .btn-primary:hover {
      background-color: #4338ca;
    }

    .btn-primary:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .btn-today {
      background: #ffffff;
      color: #4f46e5;
      border-color: #c7d2fe;
      flex: 0.8;
    }

    .btn-today:hover {
      background: #eef2ff;
    }

    .btn-cancel {
      background: #ffffff;
      color: #64748b;
      border-color: #e2e8f0;
      flex: 0.8;
    }

    .btn-cancel:hover {
      background: #f1f5f9;
      color: #0f172a;
    }

    .hint-text {
      text-align: center;
      font-size: 11px;
      color: #94a3b8;
      margin-top: 8px;
    }
  </style>
</head>
<body>
  <div class="cal-card">
    <!-- Header -->
    <div class="cal-header">
      <div class="cal-header-title">
        <span>${headerIcon}</span>
        <span>${headerTitle}</span>
      </div>
      <span class="cal-badge-mode">${mode.toUpperCase()}</span>
    </div>

    <!-- Month Navigation -->
    <div class="month-nav-bar">
      <button class="nav-btn" onclick="prevMonth()" title="Previous Month">◀</button>
      <div class="month-year-label" id="monthYearLabel"></div>
      <button class="nav-btn" onclick="nextMonth()" title="Next Month">▶</button>
    </div>

    <!-- Days of Week (Mon - Sun) -->
    <div class="weekdays-row">
      <span class="weekday-label">Mo</span>
      <span class="weekday-label">Tu</span>
      <span class="weekday-label">We</span>
      <span class="weekday-label">Th</span>
      <span class="weekday-label">Fr</span>
      <span class="weekday-label weekend">Sa</span>
      <span class="weekday-label weekend">Su</span>
    </div>

    <!-- Calendar Days Grid -->
    <div class="days-grid" id="daysGrid"></div>

    <!-- Selected Date Preview -->
    <div class="selected-bar">
      <span class="selected-bar-title">Selected:</span>
      <span class="selected-bar-date" id="selectedDateText"></span>
    </div>

    <!-- Footer Actions -->
    <div class="cal-footer">
      <button class="btn btn-today" onclick="selectToday()">Today</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
      <button id="proceedBtn" class="btn btn-primary" onclick="confirmSelection()">
        ${actionButtonText}
      </button>
    </div>

    <div class="hint-text">💡 Tip: Double-click any date to select instantly</div>
  </div>

  <script>
    const config = ${jsonDateInfo};

    const MONTH_NAMES = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];

    // Today's exact date values
    const todayObj = new Date();
    const todayYear = todayObj.getFullYear();
    const todayMonth = todayObj.getMonth();
    const todayDay = todayObj.getDate();

    // Currently viewed month/year in calendar
    let viewYear = config.year;
    let viewMonth = config.monthIndex;

    // Currently selected target date
    let selectedYear = config.year;
    let selectedMonth = config.monthIndex;
    let selectedDay = config.day;

    window.onload = function() {
      renderCalendar();
      updateSelectedDateDisplay();
    };

    function prevMonth() {
      viewMonth--;
      if (viewMonth < 0) {
        viewMonth = 11;
        viewYear--;
      }
      renderCalendar();
    }

    function nextMonth() {
      viewMonth++;
      if (viewMonth > 11) {
        viewMonth = 0;
        viewYear++;
      }
      renderCalendar();
    }

    function selectToday() {
      viewYear = todayYear;
      viewMonth = todayMonth;
      selectedYear = todayYear;
      selectedMonth = todayMonth;
      selectedDay = todayDay;
      renderCalendar();
      updateSelectedDateDisplay();
    }

    function renderCalendar() {
      document.getElementById('monthYearLabel').innerText =
        MONTH_NAMES[viewMonth] + ' ' + viewYear;

      const grid = document.getElementById('daysGrid');
      grid.innerHTML = '';

      // First day of current view month (0=Sun, 1=Mon, ..., 6=Sat)
      const firstDayDate = new Date(viewYear, viewMonth, 1);
      // Convert Sunday (0) to 6, Monday (1) to 0, etc. (Monday start)
      const startOffset = (firstDayDate.getDay() + 6) % 7;

      const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
      const prevMonthDays = new Date(viewYear, viewMonth, 0).getDate();

      // 1. Previous month trailing days
      for (let i = startOffset - 1; i >= 0; i--) {
        const d = prevMonthDays - i;
        const cell = document.createElement('div');
        cell.className = 'day-cell other-month';
        cell.innerText = d;
        cell.onclick = (function(prevM, prevY, dayNum) {
          return function() {
            viewMonth = prevM;
            viewYear = prevY;
            selectDate(prevY, prevM, dayNum);
          };
        })(viewMonth === 0 ? 11 : viewMonth - 1, viewMonth === 0 ? viewYear - 1 : viewYear, d);
        grid.appendChild(cell);
      }

      // 2. Current month days
      for (let d = 1; d <= daysInMonth; d++) {
        const cell = document.createElement('div');
        cell.className = 'day-cell';
        cell.innerText = d;

        // Day of week index for weekend highlighting (Monday=0 ... Sunday=6)
        const dayOfWeek = (startOffset + (d - 1)) % 7;
        if (dayOfWeek === 5 || dayOfWeek === 6) {
          cell.classList.add('weekend');
        }

        // Today highlight
        if (viewYear === todayYear && viewMonth === todayMonth && d === todayDay) {
          cell.classList.add('today');
        }

        // Selected highlight
        if (viewYear === selectedYear && viewMonth === selectedMonth && d === selectedDay) {
          cell.classList.add('selected');
        }

        cell.onclick = (function(y, m, dayNum) {
          return function() {
            selectDate(y, m, dayNum);
          };
        })(viewYear, viewMonth, d);

        cell.ondblclick = (function(y, m, dayNum) {
          return function() {
            selectDate(y, m, dayNum);
            confirmSelection();
          };
        })(viewYear, viewMonth, d);

        grid.appendChild(cell);
      }

      // 3. Next month leading days to complete the rows
      const totalCells = startOffset + daysInMonth;
      const nextMonthSlots = (7 - (totalCells % 7)) % 7;
      for (let d = 1; d <= nextMonthSlots; d++) {
        const cell = document.createElement('div');
        cell.className = 'day-cell other-month';
        cell.innerText = d;
        cell.onclick = (function(nextM, nextY, dayNum) {
          return function() {
            viewMonth = nextM;
            viewYear = nextY;
            selectDate(nextY, nextM, dayNum);
          };
        })(viewMonth === 11 ? 0 : viewMonth + 1, viewMonth === 11 ? viewYear + 1 : viewYear, d);
        grid.appendChild(cell);
      }
    }

    function selectDate(y, m, d) {
      selectedYear = y;
      selectedMonth = m;
      selectedDay = d;
      renderCalendar();
      updateSelectedDateDisplay();
    }

    function updateSelectedDateDisplay() {
      const d = new Date(selectedYear, selectedMonth, selectedDay);
      const options = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
      document.getElementById('selectedDateText').innerText =
        d.toLocaleDateString('en-GB', options);
    }

    function confirmSelection() {
      const proceedBtn = document.getElementById('proceedBtn');
      proceedBtn.disabled = true;
      proceedBtn.innerText = '⏳ Opening...';

      const mm = (selectedMonth + 1 < 10 ? '0' : '') + (selectedMonth + 1);
      const dd = (selectedDay < 10 ? '0' : '') + selectedDay;
      const dateStr = selectedYear + '-' + mm + '-' + dd;

      google.script.run
        .withFailureHandler(function(err) {
          proceedBtn.disabled = false;
          proceedBtn.innerText = '${actionButtonText}';
          alert('Error: ' + (err.message || err));
        })
        .proceedFromCalendar({
          mode: config.mode,
          dateStr: dateStr
        });
    }

    // Keyboard support: Enter confirms selection
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        confirmSelection();
      }
    });
  </script>
</body>
</html>`;
}
