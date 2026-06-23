// Oasis Attendance Automation Web Dashboard Controller
// Single Source of Truth Architecture:
// Raw Files -> Execute Engine once -> Results Object in memory -> Views (Preview, Stats, Exceptions, Excel Export)

(function() {
  // UI Element References
  const attendanceInput = document.getElementById('attendance-file');
  const shiftsInput = document.getElementById('shifts-file');
  const attendanceZone = document.getElementById('attendance-zone');
  const shiftsZone = document.getElementById('shifts-zone');
  
  // Checklist Elements
  const chkAttLoaded = document.getElementById('chk-att-loaded');
  const chkAttRange = document.getElementById('chk-att-range');
  const chkAttNames = document.getElementById('chk-att-names');
  const chkAttStruct = document.getElementById('chk-att-struct');
  
  const chkShiftLoaded = document.getElementById('chk-shift-loaded');
  const chkShiftRules = document.getElementById('chk-shift-rules');
  const chkShiftStruct = document.getElementById('chk-shift-struct');
  
  const chkSysReady = document.getElementById('chk-sys-ready');

  // Metrics Display
  const valQuality = document.getElementById('val-quality');
  const badgeQuality = document.getElementById('badge-quality');
  const subtextQuality = document.getElementById('subtext-quality');
  
  const valCompliance = document.getElementById('val-compliance');
  const valLateness = document.getElementById('val-lateness');
  const valEarlyDep = document.getElementById('val-early-dep');
  const valMissing = document.getElementById('val-missing');
  const auditRangeLbl = document.getElementById('audit-range-lbl');

  // Department Insights
  const insHighLate = document.getElementById('insight-high-late');
  const insHighOt = document.getElementById('insight-high-ot');
  const insBestComp = document.getElementById('insight-best-comp');

  // Actions & Console
  const generateBtn = document.getElementById('generate-btn');
  const consoleLogBox = document.getElementById('console-log-box');
  const clearConsoleBtn = document.getElementById('clear-console');

  // Tabs & Search
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabPanes = document.querySelectorAll('.tab-pane');
  const tableSearch = document.getElementById('table-search');

  // Tables
  const tblSessions = document.getElementById('tbl-sessions').querySelector('tbody');
  const tblExceptions = document.getElementById('tbl-exceptions').querySelector('tbody');
  const tblEmployees = document.getElementById('tbl-employees').querySelector('tbody');
  const tblRules = document.getElementById('tbl-rules').querySelector('tbody');
  const tblRaw = document.getElementById('tbl-raw').querySelector('tbody');

  // Exceptions Sub-filters
  const subFilterBtns = document.querySelectorAll('.sub-filter-btn');

  // Global State
  let rawAttendanceData = null;
  let rawShiftsData = null;
  let auditResults = null; // The Single Source of Truth results object
  let activeTab = 'tab-sessions';
  let activeSubFilter = 'all';
  let searchQuery = '';

  // Setup drag-and-drop
  setupDragAndDrop(attendanceZone, attendanceInput, handleAttendanceFile);
  setupDragAndDrop(shiftsZone, shiftsInput, handleShiftsFile);

  // Setup tabs
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      tabPanes.forEach(p => p.classList.remove('active'));
      
      btn.classList.add('active');
      activeTab = btn.getAttribute('data-tab');
      document.getElementById(activeTab).classList.add('active');
      
      renderActiveTab();
    });
  });

  // Setup sub-filters for Exceptions tab
  subFilterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      subFilterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeSubFilter = btn.getAttribute('data-filter');
      renderExceptionsTable();
    });
  });

  // Setup search filter
  tableSearch.addEventListener('input', (e) => {
    searchQuery = e.target.value.toLowerCase();
    renderActiveTab();
  });

  // Setup clear console log
  clearConsoleBtn.addEventListener('click', () => {
    consoleLogBox.innerHTML = '';
    logSystem("Console log cleared.");
  });

  // Setup main report generation click handler
  generateBtn.addEventListener('click', async () => {
    if (!auditResults) return;

    generateBtn.disabled = true;
    document.querySelector('.btn-spinner').classList.remove('hidden');
    document.querySelector('.btn-text').innerText = 'Generating Workbook...';
    tableSearch.disabled = true;

    logSystem("Report generation initiated.");
    await new Promise(r => setTimeout(r, 80)); // Yield to paint

    try {
      // Call Excel download layer with the exact results object (Single Source of Truth)
      await window.AttendanceEngine.downloadExcel(auditResults, logEngine);
      logSuccess("Report compiled and downloaded successfully!");
    } catch (err) {
      logError(`Excel generation failed: ${err.message}`);
      console.error(err);
    } finally {
      generateBtn.disabled = false;
      document.querySelector('.btn-spinner').classList.add('hidden');
      document.querySelector('.btn-text').innerText = 'Generate & Download Report';
      tableSearch.disabled = false;
    }
  });

  // Drag & drop framework
  function setupDragAndDrop(zone, input, handler) {
    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.classList.add('dragover');
    });

    zone.addEventListener('dragleave', () => {
      zone.classList.remove('dragover');
    });

    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        input.files = e.dataTransfer.files;
        handler(e.dataTransfer.files[0]);
      }
    });

    input.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handler(e.target.files[0]);
      }
    });
  }

  // Logger utilities
  function writeLog(message, type = 'info') {
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.innerText = `[${new Date().toLocaleTimeString()}] ${message}`;
    consoleLogBox.appendChild(entry);
    consoleLogBox.scrollTop = consoleLogBox.scrollHeight;
  }

  function logSystem(msg) { writeLog(`[SYSTEM] ${msg}`, 'system-msg'); }
  function logInfo(msg) { writeLog(`[INFO] ${msg}`, 'info'); }
  function logSuccess(msg) { writeLog(`[SUCCESS] ${msg}`, 'success'); }
  function logWarning(msg) { writeLog(`[WARNING] ${msg}`, 'warning'); }
  function logError(msg) { writeLog(`[ERROR] ${msg}`, 'error'); }

  function logEngine(msg, type = 'info') {
    if (msg.startsWith('Warning:') || msg.startsWith('  ·')) {
      writeLog(msg, 'warning');
    } else if (msg.includes('success') || msg.includes('ready') || msg.includes('complete')) {
      writeLog(msg, 'success');
    } else {
      writeLog(msg, type);
    }
  }

  // File parsing inputs
  function handleAttendanceFile(file) {
    logSystem(`Loading attendance punches: ${file.name}...`);
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        const rows = getSheetRows(sheet);
        
        if (rows && rows.length > 0) {
          // File structural verification
          const first = rows[0];
          const hasName = 'Name' in first;
          const hasTime = 'Time' in first;
          const hasState = 'State' in first;
          
          if (!hasName || !hasTime || !hasState) {
            throw new Error("Invalid table columns. Expected Name, Time, and State headers.");
          }

          rawAttendanceData = rows;
          
          // UI Zone adjustments
          document.getElementById('attendance-name').innerText = file.name;
          document.getElementById('attendance-helper').innerText = `${rows.length.toLocaleString()} logs parsed`;
          attendanceZone.classList.add('loaded');
          
          // Update Checklist Items
          chkAttLoaded.className = 'checklist-item success';
          chkAttStruct.className = 'checklist-item success';
          
          logSuccess(`Attendance file verified: ${rows.length.toLocaleString()} log entries loaded.`);
          
          runCoreProcessingEngine();
        } else {
          throw new Error("Spreadsheet appears to have no data rows.");
        }
      } catch (err) {
        logError(`Failed parsing attendance file: ${err.message}`);
        rawAttendanceData = null;
        auditResults = null;
        attendanceZone.classList.remove('loaded');
        document.getElementById('attendance-name').innerText = 'attendance.xls';
        document.getElementById('attendance-helper').innerText = 'Drop biometric logs (.xls)';
        chkAttLoaded.className = 'checklist-item pending';
        chkAttRange.className = 'checklist-item pending';
        chkAttNames.className = 'checklist-item pending';
        chkAttStruct.className = 'checklist-item pending';
        resetDashboardData();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleShiftsFile(file) {
    logSystem(`Loading shifts manifest: ${file.name}...`);
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        let employeesSheet = workbook.SheetNames.find(n => n === 'Employees' || n === 'Shifts');
        let configsSheet = workbook.SheetNames.find(n => n === 'Shifts_Config');
        let rulesSheet = workbook.SheetNames.find(n => n === 'Department_Rules');
        
        if (!employeesSheet) {
          employeesSheet = workbook.SheetNames[0];
        }

        const empRows = getSheetRows(workbook.Sheets[employeesSheet]);
        const configRows = configsSheet ? getSheetRows(workbook.Sheets[configsSheet]) : [];
        const rulesRows = rulesSheet ? getSheetRows(workbook.Sheets[rulesSheet]) : [];
        
        if (empRows && empRows.length > 0) {
          // File structural check
          const first = empRows[0];
          const hasName = 'Name' in first;
          const hasDept = 'Department' in first;
          if (!hasName || !hasDept) {
            throw new Error("Employees directory missing Name or Department columns.");
          }

          rawShiftsData = {
            employees: empRows,
            configs: configRows,
            rules: rulesRows
          };

          // UI adjustments
          document.getElementById('shifts-name').innerText = file.name;
          document.getElementById('shifts-helper').innerText = `${empRows.length} employees, ${configRows.length} configs`;
          shiftsZone.classList.add('loaded');

          // Update checklist items
          chkShiftLoaded.className = 'checklist-item success';
          chkShiftRules.className = 'checklist-item success';
          chkShiftStruct.className = 'checklist-item success';

          logSuccess(`Shifts file verified: ${empRows.length} staff records, ${configRows.length} shift templates loaded.`);
          
          runCoreProcessingEngine();
        } else {
          throw new Error("Employees sheet is blank or missing headers.");
        }
      } catch (err) {
        logError(`Failed parsing shifts file: ${err.message}`);
        rawShiftsData = null;
        auditResults = null;
        shiftsZone.classList.remove('loaded');
        document.getElementById('shifts-name').innerText = 'shifts.xlsx';
        document.getElementById('shifts-helper').innerText = 'Drop employee shifts (.xlsx)';
        chkShiftLoaded.className = 'checklist-item pending';
        chkShiftRules.className = 'checklist-item pending';
        chkShiftStruct.className = 'checklist-item pending';
        resetDashboardData();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Parses worksheet cells to structured JS objects
  function getSheetRows(sheet) {
    if (!sheet) return [];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (data.length === 0) return [];
    
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(data.length, 5); i++) {
      const row = data[i];
      if (row && row.some(cell => {
        const s = String(cell || '').trim().toLowerCase();
        return s === 'name' || s === 'shift_id' || s === 'department' || s === 'allowed_shifts';
      })) {
        headerRowIndex = i;
        break;
      }
    }
    
    const headers = data[headerRowIndex].map(h => String(h || '').trim());
    const rows = [];
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const rowData = data[i];
      if (!rowData || rowData.length === 0) continue;
      const rowObj = {};
      headers.forEach((h, colIdx) => {
        if (h) rowObj[h] = rowData[colIdx];
      });
      if (Object.values(rowObj).some(v => v !== undefined && v !== '')) {
        rows.push(rowObj);
      }
    }
    return rows;
  }

  // -------------------------------------------------------------
  // SINGLE SOURCE OF TRUTH PIPELINE
  // -------------------------------------------------------------
  function runCoreProcessingEngine() {
    if (!rawAttendanceData || !rawShiftsData) {
      validateChecklist();
      return;
    }

    logSystem("Executing production pairing engine...");
    
    try {
      // 1. Call calculations engine ONCE
      const results = window.AttendanceEngine.calculate(rawAttendanceData, rawShiftsData, logEngine);
      
      // 2. Capture Result Object in memory (Single Source of Truth)
      auditResults = results;
      
      logSuccess(`Engine execution complete: paired ${results.sessions.length.toLocaleString()} shifts.`);

      // 3. Update Validations Checklist
      updateChecklistVerification(results);

      // 4. Compute Operational Analytics & Department Rankings
      computeOperationalAnalytics(results);

      // 5. Render previews & exceptions
      renderActiveTab();
      
    } catch (err) {
      logError(`Engine calculation error: ${err.message}`);
      console.error(err);
      auditResults = null;
      resetDashboardData();
    }
  }

  // Reset operational stats when files are unloaded/error
  function resetDashboardData() {
    valQuality.innerText = '-';
    badgeQuality.className = 'metric-badge hidden';
    badgeQuality.innerText = '';
    subtextQuality.innerText = 'Based on exceptions rate';
    valCompliance.innerText = '-';
    valLateness.innerText = '-';
    valEarlyDep.innerText = '-';
    valMissing.innerText = '-';
    
    auditRangeLbl.innerText = 'Select files to inspect range';
    insHighLate.innerText = '-';
    insHighOt.innerText = '-';
    insBestComp.innerText = '-';

    tableSearch.disabled = true;
    generateBtn.disabled = true;
    chkSysReady.className = 'checklist-item pending';

    // Clear previews
    tblSessions.innerHTML = '<tr class="empty-row"><td colspan="10">No sessions parsed. Upload biometric log and shift manifest files in the sidebar.</td></tr>';
    tblExceptions.innerHTML = '<tr class="empty-row"><td colspan="5">No exceptions parsed. Upload files to verify anomalies.</td></tr>';
    tblEmployees.innerHTML = '<tr class="empty-row"><td colspan="4">No employee database parsed. Upload shifts file.</td></tr>';
    tblRules.innerHTML = '<tr class="empty-row"><td colspan="4">No shift configs parsed. Upload shifts file.</td></tr>';
    tblRaw.innerHTML = '<tr class="empty-row"><td colspan="4">No biometric logs parsed. Upload attendance file.</td></tr>';
  }

  // Validation checklists
  function validateChecklist() {
    if (rawAttendanceData) {
      chkAttLoaded.className = 'checklist-item success';
      chkAttStruct.className = 'checklist-item success';
    } else {
      chkAttLoaded.className = 'checklist-item pending';
      chkAttRange.className = 'checklist-item pending';
      chkAttNames.className = 'checklist-item pending';
      chkAttStruct.className = 'checklist-item pending';
    }

    if (rawShiftsData) {
      chkShiftLoaded.className = 'checklist-item success';
      chkShiftRules.className = 'checklist-item success';
      chkShiftStruct.className = 'checklist-item success';
    } else {
      chkShiftLoaded.className = 'checklist-item pending';
      chkShiftRules.className = 'checklist-item pending';
      chkShiftStruct.className = 'checklist-item pending';
    }

    chkSysReady.className = 'checklist-item pending';
    generateBtn.disabled = true;
  }

  function updateChecklistVerification(results) {
    chkAttLoaded.className = 'checklist-item success';
    chkAttStruct.className = 'checklist-item success';
    chkShiftLoaded.className = 'checklist-item success';
    chkShiftRules.className = 'checklist-item success';
    chkShiftStruct.className = 'checklist-item success';

    // Verify date range
    if (results.sessions && results.sessions.length > 0) {
      chkAttRange.className = 'checklist-item success';
    } else {
      chkAttRange.className = 'checklist-item pending';
    }

    // Verify unmatched/lookups names
    if (results.unmatchedEmployees && results.unmatchedEmployees.size > 0) {
      chkAttNames.className = 'checklist-item warning';
      chkAttNames.querySelector('.check-lbl').innerText = `Lookups mapped (${results.unmatchedEmployees.size} unmatched)`;
    } else {
      chkAttNames.className = 'checklist-item success';
      chkAttNames.querySelector('.check-lbl').innerText = 'Employee names mapped';
    }

    // System clearance
    chkSysReady.className = 'checklist-item success';
    generateBtn.disabled = false;
  }

  // Analytics calculator using output object
  function computeOperationalAnalytics(results) {
    const { sessions, employeeSummary, statusSummary } = results;

    if (!sessions || sessions.length === 0) return;

    // Date range label
    let minD = new Date();
    let maxD = new Date(0);
    sessions.forEach(s => {
      if (s.Date.getTime() < minD.getTime()) minD = s.Date;
      if (s.Date.getTime() > maxD.getTime()) maxD = s.Date;
    });
    const formatDateStr = (d) => `${d.getDate()} ${d.toLocaleString('en-US', { month: 'short' })} ${d.getFullYear()}`;
    auditRangeLbl.innerText = `Audit logs cover: ${formatDateStr(minD)} to ${formatDateStr(maxD)}`;

    // 1. Compliance Rate
    const totalScheduled = sessions.filter(s => s.DayAttendance !== 'Rest Day').length;
    const fullDays = sessions.filter(s => s.DayAttendance === 'Full Day').length;
    const complianceRate = totalScheduled > 0 ? (fullDays / totalScheduled) * 100 : 0;
    valCompliance.innerText = `${complianceRate.toFixed(1)}%`;

    // 2. Lateness Rate (shared with report export)
    const validCins = sessions.filter(s => s.CheckIn && s.DayAttendance !== 'Rest Day').length;
    const lateDays = sessions.filter(s => s.LateMinutes > 0 && s.DayAttendance !== 'Rest Day').length;
    const latenessRate = results.summaryMetrics?.latenessRate ?? (validCins > 0 ? (lateDays / validCins) * 100 : 0);
    valLateness.innerText = `${latenessRate.toFixed(1)}%`;

    // 3. Early Departure Rate
    const earlyDeps = sessions.filter(s => s.DayAttendance === 'Early Departure').length;
    const earlyDepRate = totalScheduled > 0 ? (earlyDeps / totalScheduled) * 100 : 0;
    valEarlyDep.innerText = `${earlyDepRate.toFixed(1)}%`;

    // 4. Missing Punch Events (Missing Check-in or Check-out)
    const missingPunches = sessions.filter(s => s.MissingCheckin === 'Yes' || s.MissingCheckout === 'Yes').length;
    valMissing.innerText = missingPunches.toLocaleString();

    // 5. Overtime Rate
    // Overtime is generated if actual HoursWorked > expectedShiftHours
    let otCount = 0;
    let totalCouts = 0;
    sessions.forEach(s => {
      if (s.CheckIn && s.CheckOut && s.DayAttendance !== 'Rest Day') {
        totalCouts++;
        const expected = getShiftExpectedHours(s.Shift);
        if (s.HoursWorked > expected) {
          otCount++;
        }
      }
    });
    const otRate = totalCouts > 0 ? (otCount / totalCouts) * 100 : 0;

    // 6. Employees Requiring Review count
    // Anomalies are: unmatched profile, punches on rest day, missing clock-in/out, or late > 60 mins (Very Late)
    const reviewSet = new Set();
    
    // Unmatched employees
    results.unmatchedEmployees.forEach(name => reviewSet.add(name));
    
    // Sessions anomalies
    sessions.forEach(s => {
      const isAnomalous = s.DayAttendance === 'Anomalous' || 
                           s.MissingCheckin === 'Yes' || 
                           s.MissingCheckout === 'Yes' || 
                           s.Shift === 'Rest Day' || 
                           s.LateMinutes > 60 ||
                           (s.DataQualityFlag && s.DataQualityFlag !== '');
      if (isAnomalous) {
        reviewSet.add(s.Employee);
      }
    });

    // 7. Attendance Quality Score
    // Calculate a deduction-based score starting at 100
    let score = 100;
    
    // Deduct for missing punches rate (Weight: 35%)
    const missingRate = totalScheduled > 0 ? (missingPunches / totalScheduled) : 0;
    score -= (missingRate * 35);
    
    // Deduct for non-compliance rate (Weight: 25%)
    const nonCompliance = 100 - complianceRate;
    score -= (nonCompliance / 100 * 25);
    
    // Deduct for unmatched employees (Weight: 15% max)
    if (rawShiftsData && rawShiftsData.employees.length > 0) {
      const unmatchedRatio = results.unmatchedEmployees.size / rawShiftsData.employees.length;
      score -= Math.min(15, unmatchedRatio * 100 * 2);
    }
    
    // Deduct for lateness rate (Weight: 15% max)
    score -= (latenessRate / 100 * 15);
    
    // Deduct for exceptions count (Weight: 10% max)
    const exceptionsCount = missingPunches + results.unmatchedEmployees.size + results.duplicates.length;
    score -= Math.min(10, exceptionsCount * 0.2);

    score = Math.max(0, Math.min(100, score));
    valQuality.innerText = `${score.toFixed(1)}%`;
    
    // Update Score badge
    badgeQuality.className = 'metric-badge';
    if (score >= 95) {
      badgeQuality.innerText = 'Excellent';
      badgeQuality.classList.add('excellent');
    } else if (score >= 85) {
      badgeQuality.innerText = 'Good';
      badgeQuality.classList.add('good');
    } else if (score >= 70) {
      badgeQuality.innerText = 'Attention';
      badgeQuality.classList.add('warning');
    } else {
      badgeQuality.innerText = 'Critical';
      badgeQuality.classList.add('critical');
    }
    subtextQuality.innerText = `${reviewSet.size} employees require review`;

    // 8. Department Insights Ranking Calculations
    const deptStats = {};
    sessions.forEach(s => {
      const dept = s.Department || 'UNKNOWN';
      if (!deptStats[dept]) {
        deptStats[dept] = {
          name: dept,
          totalScheduled: 0,
          fullDays: 0,
          validCins: 0,
          lateDays: 0,
          validCouts: 0,
          otDays: 0
        };
      }
      const stat = deptStats[dept];
      if (s.DayAttendance !== 'Rest Day') {
        stat.totalScheduled++;
        if (s.DayAttendance === 'Full Day') stat.fullDays++;
        if (s.CheckIn) {
          stat.validCins++;
          if (s.LateMinutes > 0) stat.lateDays++;
        }
        if (s.CheckIn && s.CheckOut) {
          stat.validCouts++;
          const expected = getShiftExpectedHours(s.Shift);
          if (s.HoursWorked > expected) {
            stat.otDays++;
          }
        }
      }
    });

    let highestLatenessDept = '-';
    let maxLatenessVal = -1;
    let highestOtDept = '-';
    let maxOtVal = -1;
    let bestComplianceDept = '-';
    let maxCompVal = -1;

    Object.keys(deptStats).forEach(key => {
      const stat = deptStats[key];
      if (stat.totalScheduled < 3) return; // Ignore small sample departments

      const lateness = stat.validCins > 0 ? (stat.lateDays / stat.validCins) * 100 : 0;
      const compliance = stat.totalScheduled > 0 ? (stat.fullDays / stat.totalScheduled) * 100 : 0;
      const ot = stat.validCouts > 0 ? (stat.otDays / stat.validCouts) * 100 : 0;

      if (lateness > maxLatenessVal) {
        maxLatenessVal = lateness;
        highestLatenessDept = `${stat.name} (${lateness.toFixed(1)}%)`;
      }
      if (ot > maxOtVal) {
        maxOtVal = ot;
        highestOtDept = `${stat.name} (${ot.toFixed(1)}%)`;
      }
      if (compliance > maxCompVal) {
        maxCompVal = compliance;
        bestComplianceDept = `${stat.name} (${compliance.toFixed(1)}%)`;
      }
    });

    insHighLate.innerText = highestLatenessDept;
    insHighOt.innerText = highestOtDept;
    insBestComp.innerText = bestComplianceDept;
  }

  // Estimates expected hours for a shift
  function getShiftExpectedHours(shiftStr) {
    if (!shiftStr || shiftStr === 'Rest Day') return 9;
    const m = shiftStr.match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (m) {
      const sh = parseInt(m[1], 10);
      const sm = parseInt(m[2], 10);
      const eh = parseInt(m[3], 10);
      const em = parseInt(m[4], 10);
      
      let diff = (eh + em/60) - (sh + sm/60);
      if (diff <= 0) {
        diff += 24; // Overnight carry-over
      }
      return diff;
    }
    return 9; // Default
  }

  // Previews display router
  function renderActiveTab() {
    if (activeTab === 'tab-sessions') {
      renderSessionsTable();
    } else if (activeTab === 'tab-exceptions') {
      renderExceptionsTable();
    } else if (activeTab === 'tab-employees') {
      renderEmployeesTable();
    } else if (activeTab === 'tab-rules') {
      renderRulesTable();
    } else if (activeTab === 'tab-raw') {
      renderRawTable();
    }
  }

  function renderSessionsTable() {
    tblSessions.innerHTML = '';
    
    if (!auditResults || !auditResults.sessions || auditResults.sessions.length === 0) {
      tblSessions.innerHTML = '<tr class="empty-row"><td colspan="10">No sessions parsed. Upload files to verify paired sessions.</td></tr>';
      return;
    }

    const filtered = auditResults.sessions.filter(s => {
      if (!searchQuery) return true;
      return s.Employee.toLowerCase().includes(searchQuery) ||
             (s.Department && s.Department.toLowerCase().includes(searchQuery)) ||
             s.Shift.toLowerCase().includes(searchQuery) ||
             s.AttendanceStatus.toLowerCase().includes(searchQuery);
    });

    if (filtered.length === 0) {
      tblSessions.innerHTML = '<tr class="empty-row"><td colspan="10">No matching sessions found.</td></tr>';
      return;
    }

    const renderList = filtered.slice(0, 100);
    renderList.forEach(s => {
      const tr = document.createElement('tr');

      const tdDate = document.createElement('td');
      tdDate.innerText = s.Date.toLocaleDateString('en-GB');

      const tdEmp = document.createElement('td');
      tdEmp.innerText = s.Employee;

      const tdDept = document.createElement('td');
      tdDept.innerText = s.Department || '-';

      const tdShift = document.createElement('td');
      tdShift.innerText = s.Shift;

      const tdCin = document.createElement('td');
      tdCin.innerText = s.CheckIn ? s.CheckIn.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }) : '-';

      const tdCout = document.createElement('td');
      tdCout.innerText = s.CheckOut ? s.CheckOut.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }) : '-';

      const tdLate = document.createElement('td');
      tdLate.innerText = s.LateMinutes > 0 ? s.LateMinutes : '0';

      const tdHours = document.createElement('td');
      tdHours.innerText = s.HoursWorked !== null ? s.HoursWorked.toFixed(2) : '-';

      const tdStatus = document.createElement('td');
      const badge = document.createElement('span');
      badge.className = 'status-badge';
      const statusText = s.AttendanceStatus;
      badge.innerText = statusText;
      if (statusText.startsWith('On Time')) {
        badge.classList.add('on-time');
      } else if (statusText.startsWith('Slightly Late')) {
        badge.classList.add('warning');
      } else if (statusText.startsWith('Late') || statusText.startsWith('Very Late')) {
        badge.classList.add('late');
      } else {
        badge.classList.add('muted');
      }
      tdStatus.appendChild(badge);

      const tdFlags = document.createElement('td');
      if (s.DataQualityFlag) {
        const span = document.createElement('span');
        span.className = 'anomaly-flag';
        span.innerText = s.DataQualityFlag;
        tdFlags.appendChild(span);
      } else {
        tdFlags.innerText = '-';
      }

      tr.appendChild(tdDate);
      tr.appendChild(tdEmp);
      tr.appendChild(tdDept);
      tr.appendChild(tdShift);
      tr.appendChild(tdCin);
      tr.appendChild(tdCout);
      tr.appendChild(tdLate);
      tr.appendChild(tdHours);
      tr.appendChild(tdStatus);
      tr.appendChild(tdFlags);

      tblSessions.appendChild(tr);
    });

    if (filtered.length > 100) {
      addTruncationRow(tblSessions, 10, filtered.length);
    }
  }

  function renderExceptionsTable() {
    tblExceptions.innerHTML = '';

    if (!auditResults) {
      tblExceptions.innerHTML = '<tr class="empty-row"><td colspan="5">No exceptions parsed. Upload files to verify anomalies.</td></tr>';
      return;
    }

    const exceptions = [];

    // 1. Missing Punches
    auditResults.sessions.forEach(s => {
      if (s.MissingCheckin === 'Yes') {
        exceptions.push({
          date: s.Date,
          dateStr: s.Date.toLocaleDateString('en-GB'),
          employee: s.Employee,
          department: s.Department,
          type: 'missing',
          typeLabel: 'Missing Punch',
          details: 'Missing check-in record (unknown lateness)'
        });
      }
      if (s.MissingCheckout === 'Yes') {
        exceptions.push({
          date: s.Date,
          dateStr: s.Date.toLocaleDateString('en-GB'),
          employee: s.Employee,
          department: s.Department,
          type: 'missing',
          typeLabel: 'Missing Punch',
          details: 'Missing check-out record'
        });
      }
    });

    // 2. Unknown Employees (biometric log names not in shifts roster)
    auditResults.unmatchedEmployees.forEach(name => {
      exceptions.push({
        date: null,
        dateStr: 'N/A',
        employee: name,
        department: 'UNKNOWN',
        type: 'unknown',
        typeLabel: 'Unknown Employee',
        details: 'Attendance punch matches no registered employee profile'
      });
    });

    // 3. Rest Day punches
    auditResults.sessions.forEach(s => {
      if (s.Shift === 'Rest Day') {
        exceptions.push({
          date: s.Date,
          dateStr: s.Date.toLocaleDateString('en-GB'),
          employee: s.Employee,
          department: s.Department,
          type: 'rest-day',
          typeLabel: 'Rest Day Punch',
          details: 'Biometric punches recorded on designated rest day'
        });
      }
    });

    // 4. Duplicate Records Removed
    auditResults.duplicates.forEach(d => {
      exceptions.push({
        date: d.Time,
        dateStr: d.Time.toLocaleString('en-GB'),
        employee: d.Name,
        department: d.Department,
        type: 'duplicates',
        typeLabel: 'Duplicate Punch',
        details: `Duplicate log state [${d.State}] removed by pairing engine`
      });
    });

    // 5. General Anomalies (DataQualityFlag or Anomalous sessions)
    auditResults.sessions.forEach(s => {
      if (s.DayAttendance === 'Anomalous' || (s.DataQualityFlag && s.DataQualityFlag !== '')) {
        exceptions.push({
          date: s.Date,
          dateStr: s.Date.toLocaleDateString('en-GB'),
          employee: s.Employee,
          department: s.Department,
          type: 'anomalies',
          typeLabel: 'Attendance Anomaly',
          details: s.DataQualityFlag || 'Anomalous punch patterns detected'
        });
      }
    });

    // Filter by category
    let filtered = exceptions;
    if (activeSubFilter !== 'all') {
      filtered = exceptions.filter(e => e.type === activeSubFilter);
    }

    // Filter by text search
    if (searchQuery) {
      filtered = filtered.filter(e => {
        return e.employee.toLowerCase().includes(searchQuery) ||
               e.details.toLowerCase().includes(searchQuery) ||
               e.typeLabel.toLowerCase().includes(searchQuery);
      });
    }

    if (filtered.length === 0) {
      tblExceptions.innerHTML = '<tr class="empty-row"><td colspan="5">No exceptions found matching current filters.</td></tr>';
      return;
    }

    const renderList = filtered.slice(0, 100);
    renderList.forEach(e => {
      const tr = document.createElement('tr');

      const tdDate = document.createElement('td');
      tdDate.innerText = e.dateStr;

      const tdEmp = document.createElement('td');
      tdEmp.innerText = e.employee;

      const tdDept = document.createElement('td');
      tdDept.innerText = e.department || '-';

      const tdType = document.createElement('td');
      const badge = document.createElement('span');
      badge.className = 'status-badge';
      badge.innerText = e.typeLabel;
      if (e.type === 'missing') {
        badge.classList.add('warning');
      } else if (e.type === 'unknown' || e.type === 'anomalies') {
        badge.classList.add('late');
      } else {
        badge.classList.add('muted');
      }
      tdType.appendChild(badge);

      const tdDetails = document.createElement('td');
      tdDetails.innerText = e.details;

      tr.appendChild(tdDate);
      tr.appendChild(tdEmp);
      tr.appendChild(tdDept);
      tr.appendChild(tdType);
      tr.appendChild(tdDetails);

      tblExceptions.appendChild(tr);
    });

    if (filtered.length > 100) {
      addTruncationRow(tblExceptions, 5, filtered.length);
    }
  }

  function renderEmployeesTable() {
    tblEmployees.innerHTML = '';

    if (!rawShiftsData || !rawShiftsData.employees || rawShiftsData.employees.length === 0) {
      tblEmployees.innerHTML = '<tr class="empty-row"><td colspan="4">No employees parsed. Select Shifts File.</td></tr>';
      return;
    }

    const filtered = rawShiftsData.employees.filter(row => {
      if (!searchQuery) return true;
      return String(row.Name || '').toLowerCase().includes(searchQuery) ||
             String(row.Department || '').toLowerCase().includes(searchQuery) ||
             String(row.Rest_Day || '').toLowerCase().includes(searchQuery) ||
             String(row.Override_Shift || row.Shift || '').toLowerCase().includes(searchQuery);
    });

    if (filtered.length === 0) {
      tblEmployees.innerHTML = '<tr class="empty-row"><td colspan="4">No matching records found.</td></tr>';
      return;
    }

    const renderList = filtered.slice(0, 100);
    renderList.forEach(row => {
      const tr = document.createElement('tr');
      
      const tdName = document.createElement('td');
      tdName.innerText = row.Name || '-';
      
      const tdDept = document.createElement('td');
      tdDept.innerText = row.Department || '-';
      
      const tdShift = document.createElement('td');
      tdShift.innerText = row.Override_Shift || row.Shift || row.SHIFT || '-';
      
      const tdRest = document.createElement('td');
      tdRest.innerText = row.Rest_Day || '-';
      
      tr.appendChild(tdName);
      tr.appendChild(tdDept);
      tr.appendChild(tdShift);
      tr.appendChild(tdRest);
      
      tblEmployees.appendChild(tr);
    });

    if (filtered.length > 100) {
      addTruncationRow(tblEmployees, 4, filtered.length);
    }
  }

  function renderRulesTable() {
    tblRules.innerHTML = '';

    if (!rawShiftsData || !rawShiftsData.configs || rawShiftsData.configs.length === 0) {
      tblRules.innerHTML = '<tr class="empty-row"><td colspan="4">No shift configs parsed. Select Shifts File.</td></tr>';
      return;
    }

    const filtered = rawShiftsData.configs.filter(row => {
      if (!searchQuery) return true;
      return String(row.Shift_ID || '').toLowerCase().includes(searchQuery) ||
             String(row.Start_Time || '').toLowerCase().includes(searchQuery) ||
             String(row.End_Time || '').toLowerCase().includes(searchQuery);
    });

    if (filtered.length === 0) {
      tblRules.innerHTML = '<tr class="empty-row"><td colspan="4">No matching configs found.</td></tr>';
      return;
    }

    filtered.forEach(row => {
      const tr = document.createElement('tr');
      
      const tdId = document.createElement('td');
      tdId.innerText = row.Shift_ID || '-';
      
      const tdStart = document.createElement('td');
      tdStart.innerText = row.Start_Time || '-';
      
      const tdEnd = document.createElement('td');
      tdEnd.innerText = row.End_Time || '-';
      
      const tdNight = document.createElement('td');
      tdNight.innerText = row.Is_Night_Shift || '-';
      
      tr.appendChild(tdId);
      tr.appendChild(tdStart);
      tr.appendChild(tdEnd);
      tr.appendChild(tdNight);
      
      tblRules.appendChild(tr);
    });
  }

  function renderRawTable() {
    tblRaw.innerHTML = '';

    if (!rawAttendanceData || rawAttendanceData.length === 0) {
      tblRaw.innerHTML = '<tr class="empty-row"><td colspan="4">No logs parsed. Select Attendance File.</td></tr>';
      return;
    }

    const filtered = rawAttendanceData.filter(row => {
      if (!searchQuery) return true;
      return String(row.Name || '').toLowerCase().includes(searchQuery) ||
             String(row.State || '').toLowerCase().includes(searchQuery) ||
             String(row.Department || '').toLowerCase().includes(searchQuery);
    });

    if (filtered.length === 0) {
      tblRaw.innerHTML = '<tr class="empty-row"><td colspan="4">No matching attendance records found.</td></tr>';
      return;
    }

    const renderList = filtered.slice(0, 100);
    renderList.forEach(row => {
      const tr = document.createElement('tr');
      
      const tdName = document.createElement('td');
      tdName.innerText = row.Name || '-';
      
      const tdTime = document.createElement('td');
      let tStr = '-';
      const d = window.AttendanceEngine.parseExcelDate(row.Time);
      if (d) {
        tStr = d.toLocaleString('en-GB');
      } else {
        tStr = String(row.Time || '-');
      }
      tdTime.innerText = tStr;
      
      const tdState = document.createElement('td');
      tdState.innerText = row.State || '-';
      
      const tdDept = document.createElement('td');
      tdDept.innerText = row.Department || '-';
      
      tr.appendChild(tdName);
      tr.appendChild(tdTime);
      tr.appendChild(tdState);
      tr.appendChild(tdDept);
      
      tblRaw.appendChild(tr);
    });

    if (filtered.length > 100) {
      addTruncationRow(tblRaw, 4, filtered.length);
    }
  }

  function addTruncationRow(tbody, colspan, total) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row';
    tr.innerHTML = `<td colspan="${colspan}" style="padding: 10px; color: var(--text-muted); font-size: 11px;">Showing first 100 of ${total.toLocaleString()} rows. Use search box to filter.</td>`;
    tbody.appendChild(tr);
  }

  function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }
})();
