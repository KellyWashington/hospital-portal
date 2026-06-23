// Oasis Attendance Automation Core Processing Engine
// Decoupled into a clean "Single Source of Truth" architecture:
// 1. calculate(): Executes all pairing and analysis rules, returning a results object.
// 2. downloadExcel(): Takes the results object and builds/downloads the formatted Excel workbook.
// Both functions preserve the exact rules and layouts of the production system.

(function() {
  function normalizeName(name) {
    return String(name || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
  }

  function parseExcelDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
      // Excel serial dates
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      return date;
    }
    if (typeof val === 'string') {
      const s = val.trim();
      // Parse DD/MM/YYYY HH:MM:SS or DD/MM/YYYY HH:MM
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
      if (m) {
        const day = parseInt(m[1], 10);
        const month = parseInt(m[2], 10) - 1;
        const year = parseInt(m[3], 10);
        const hour = parseInt(m[4], 10);
        const minute = parseInt(m[5], 10);
        const second = m[6] ? parseInt(m[6], 10) : 0;
        return new Date(year, month, day, hour, minute, second);
      }
      const parsed = Date.parse(s);
      if (!isNaN(parsed)) return new Date(parsed);
    }
    return null;
  }

  function getShiftExpectedHours(shiftStr) {
    if (!shiftStr || shiftStr === 'Rest Day') return 9;
    const m = shiftStr.match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (m) {
      const sh = parseInt(m[1], 10);
      const sm = parseInt(m[2], 10);
      const eh = parseInt(m[3], 10);
      const em = parseInt(m[4], 10);
      let start = sh * 60 + sm;
      let end = eh * 60 + em;
      if (end <= start) end += 24 * 60;
      return Math.max(0, (end - start) / 60);
    }
    return 9;
  }

  function computeSummaryMetrics(sessions, statusSummary, unmatchedEmployees, duplicates) {
    const totalScheduled = sessions.filter(s => s.DayAttendance !== 'Rest Day').length;
    const fullDays = sessions.filter(s => s.DayAttendance === 'Full Day').length;
    const validCins = sessions.filter(s => s.CheckIn && s.DayAttendance !== 'Rest Day').length;
    const lateDays = sessions.filter(s => s.LateMinutes > 0 && s.DayAttendance !== 'Rest Day').length;
    const earlyDepDays = sessions.filter(s => s.DayAttendance === 'Early Departure').length;

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

    const missingPunches = sessions.filter(s => s.MissingCheckin === 'Yes' || s.MissingCheckout === 'Yes').length;
    const complianceRate = totalScheduled > 0 ? (fullDays / totalScheduled) * 100 : 0;
    const latenessRate = validCins > 0 ? (lateDays / validCins) * 100 : 0;
    const earlyDepRate = totalScheduled > 0 ? (earlyDepDays / totalScheduled) * 100 : 0;
    const otRate = totalCouts > 0 ? (otCount / totalCouts) * 100 : 0;

    const reviewSet = new Set([
      ...(unmatchedEmployees || []),
      ...sessions
        .filter(s => (
          s.DayAttendance === 'Anomalous' ||
          s.MissingCheckin === 'Yes' ||
          s.MissingCheckout === 'Yes' ||
          s.Shift === 'Rest Day' ||
          s.LateMinutes > 60 ||
          (s.DataQualityFlag && s.DataQualityFlag !== '')
        ))
        .map(s => s.Employee)
    ]);

    const totalEmployees = statusSummary.length;
    const totalDaysAttended = statusSummary.reduce((sum, s) => sum + (s['Days Attended'] || 0), 0);
    const totalMissingCin = statusSummary.reduce((sum, s) => sum + (s['Missing C/In (Unknown Lateness)'] || 0), 0);
    const totalValidCinDays = statusSummary.reduce((sum, s) => sum + (s['Valid C/In Days'] || 0), 0);
    const totalLateDays = statusSummary.reduce((sum, s) => sum + (s['Late Days'] || 0), 0);

    return {
      totalEmployees,
      totalScheduled,
      fullDays,
      validCins,
      lateDays,
      earlyDepDays,
      missingPunches,
      otCount,
      totalCouts,
      complianceRate,
      latenessRate,
      earlyDepRate,
      otRate,
      reviewCount: reviewSet.size,
      totalDaysAttended,
      totalMissingCin,
      totalValidCinDays,
      totalLateDays
    };
  }

  function calculate(attendanceData, shiftsData, log) {
    log("Initializing attendance processing engine...");

    // Excluded employee list
    const EXCLUDED_EMPLOYEES = new Set(["wickliff ondiba"]);

    // 1. Prepare shifts lookups
    log("Mapping employee shifts and departments...");
    
    // Auto-detect shift column name from headers
    const empSample = shiftsData.employees[0] || {};
    let shiftCol = 'Override_Shift';
    for (let candidate of ["Override_Shift", "Shift", "SHIFT"]) {
      if (candidate in empSample) {
        shiftCol = candidate;
        break;
      }
    }
    
    const shiftMap = {};
    const deptMap = {};
    const restDayMap = {};
    const shiftNormNames = new Set();
    
    shiftsData.employees.forEach(row => {
      const rawName = row.Name;
      const norm = normalizeName(rawName);
      if (norm) {
        shiftMap[norm] = String(row[shiftCol] || '').trim();
        deptMap[norm] = String(row.Department || '').trim();
        restDayMap[norm] = String(row.Rest_Day || '').trim();
        shiftNormNames.add(norm);
      }
    });

    // Load shifts config
    log("Loading Shift Config definitions...");
    const shiftsConfigLookup = {};
    shiftsData.configs.forEach(row => {
      const sid = String(row.Shift_ID || '').trim();
      if (sid) {
        shiftsConfigLookup[sid] = {
          Shift_ID: sid,
          Start_Time: String(row.Start_Time || '').trim(),
          End_Time: String(row.End_Time || '').trim(),
          Is_Night_Shift: String(row.Is_Night_Shift || '').trim().toLowerCase() === 'yes' || 
                          String(row.Is_Night_Shift || '').trim().toLowerCase() === 'true' ||
                          String(row.Is_Night_Shift || '') === '1'
        };
      }
    });

    // Load department rules
    log("Loading Department Shift Pools...");
    const deptRules = {};
    shiftsData.rules.forEach(row => {
      const deptLabel = String(row.Department || '').trim();
      const allowed = String(row.Allowed_Shifts || '')
        .split(',')
        .map(s => s.trim())
        .filter(s => s);
      if (deptLabel) {
        deptRules[deptLabel] = allowed;
      }
    });

    // 2. Prepare attendance logs
    log("Cleaning attendance logs...");
    let rawPunches = [];
    let unmatchedEmployees = new Set();
    
    attendanceData.forEach(row => {
      const rawName = row.Name;
      const norm = normalizeName(rawName);
      if (!norm) return;

      if (EXCLUDED_EMPLOYEES.has(norm)) return; // Exclude

      const timeVal = parseExcelDate(row.Time);
      if (!timeVal) return;

      const state = String(row.State || '').trim();
      if (state !== 'C/In' && state !== 'C/Out') return;

      rawPunches.push({
        Name: String(rawName).trim(),
        norm_name: norm,
        Time: timeVal,
        State: state,
        Department: row.Department
      });

      if (!shiftNormNames.has(norm)) {
        unmatchedEmployees.add(rawName);
      }
    });

    if (unmatchedEmployees.size > 0) {
      log(`Warning: ${unmatchedEmployees.size} employee(s) in attendance have no shift config (defaulting to 08:00-17:00):`, 'warning');
      Array.from(unmatchedEmployees).sort().forEach(u => log(`  · '${u}'`, 'warning'));
    }

    // Deduplicate punches (exact same employee, time, state)
    const seenPunches = new Set();
    const duplicates = [];
    rawPunches = rawPunches.filter(p => {
      const key = `${p.norm_name}_${p.Time.getTime()}_${p.State}`;
      if (seenPunches.has(key)) {
        duplicates.push(p);
        return false;
      }
      seenPunches.add(key);
      return true;
    });

    // Night-capable check
    const NIGHT_CAPABLE_DEPTS = new Set([
      "NURSES", "NURSE", "NURSE/PA", "PA", "HOUSEKEEPING( PA)", "HOUSEKEEPING(JANITORS)",
      "LABORATORY", "FRONT OFFICE", "THEATRE", "CLINICAL", "RADIOLOGY", "PHARMACY"
    ]);

    function employeeIsNightCapable(normName) {
      const dept = String(deptMap[normName] || '').trim().toUpperCase();
      const shiftTxt = String(shiftMap[normName] || '').trim().toLowerCase().replace(/\s/g, '');
      
      if (NIGHT_CAPABLE_DEPTS.has(dept)) return true;
      if (shiftTxt.includes('night') || shiftTxt.startsWith('6pm-') || shiftTxt.startsWith('18:00-') || shiftTxt.startsWith('18:30-') || shiftTxt.startsWith('6:30pm-') || shiftTxt.startsWith('630pm-')) {
        return true;
      }
      
      // Fallback: Infer from actual punch patterns
      if (!shiftNormNames.has(normName)) {
        const empPunches = rawPunches.filter(p => p.norm_name === normName);
        const totalCins = empPunches.filter(p => p.State === 'C/In').length;
        const eveningCins = empPunches.filter(p => p.State === 'C/In' && p.Time.getHours() >= 17 && p.Time.getHours() <= 23).length;
        const morningCouts = empPunches.filter(p => p.State === 'C/Out' && p.Time.getHours() <= 10).length;
        
        if (totalCins > 0 && (eveningCins / totalCins) >= 0.5) return true;
        if (morningCouts >= 2) return true;
      }
      return false;
    }

    function classifyPunch(row, isNightCapable) {
      const h = row.Time.getHours();
      const state = row.State;
      const tStr = `${String(h).padStart(2,'0')}:${String(row.Time.getMinutes()).padStart(2,'0')}`;
      
      if (isNightCapable) {
        return { intent: state === 'C/In' ? 'arrival' : 'departure', corrected: false, note: '' };
      } else {
        if (h < 13) {
          const corrected = state !== 'C/In';
          return {
            intent: 'arrival',
            corrected,
            note: corrected ? `arrival corrected (was C/Out at ${tStr})` : ''
          };
        } else {
          const corrected = state !== 'C/Out';
          return {
            intent: 'departure',
            corrected,
            note: corrected ? `departure corrected (was C/In at ${tStr})` : ''
          };
        }
      }
    }

    // 3. Process sessions grouping per employee
    log("Grouping punches into sessions using Intent-Aware Pairing...");
    const sessions = [];
    const PAIR_WINDOW = 16 * 60 * 60 * 1000; // 16 hours
    const DOUBLE_TAP_WINDOW = 5 * 60 * 1000; // 5 mins

    // Group punches by employee
    const punchesByEmp = {};
    rawPunches.forEach(p => {
      if (!punchesByEmp[p.norm_name]) punchesByEmp[p.norm_name] = [];
      punchesByEmp[p.norm_name].push(p);
    });

    Object.keys(punchesByEmp).forEach(empNorm => {
      const grp = punchesByEmp[empNorm].sort((a,b) => a.Time.getTime() - b.Time.getTime());
      const empName = grp[0].Name;
      const nightCap = employeeIsNightCapable(empNorm);
      const departmentRaw = String(deptMap[empNorm] || '').trim();

      // Step 1: Classify intent
      let classified = grp.map(row => {
        const cl = classifyPunch(row, nightCap);
        return {
          Time: row.Time,
          State: row.State,
          intent: cl.intent,
          corrected: cl.corrected,
          note: cl.note
        };
      });

      // Step 2a: Discard same-second C/In when C/Out exists
      const sameSecondOutTimes = new Set();
      classified.forEach(p => {
        if (p.intent === 'departure') sameSecondOutTimes.add(p.Time.getTime());
      });
      classified = classified.filter(p => !(p.intent === 'arrival' && sameSecondOutTimes.has(p.Time.getTime())));

      // Step 2b: Collapse double taps (within 5 minutes)
      let deduped = [];
      const usedSet = new Set();
      const n = classified.length;
      for (let i = 0; i < n; i++) {
        if (usedSet.has(i)) continue;
        usedSet.add(i);
        const cluster = [classified[i]];
        for (let j = i + 1; j < n; j++) {
          if (usedSet.has(j)) continue;
          if ((classified[j].Time.getTime() - cluster[cluster.length - 1].Time.getTime()) <= DOUBLE_TAP_WINDOW) {
            cluster.push(classified[j]);
            usedSet.add(j);
          } else {
            break;
          }
        }

        const arrivals = cluster.filter(p => p.intent === 'arrival');
        const departures = cluster.filter(p => p.intent === 'departure');
        if (arrivals.length > 0) {
          deduped.push(arrivals.reduce((min, p) => p.Time.getTime() < min.Time.getTime() ? p : min, arrivals[0]));
        }
        if (departures.length > 0) {
          deduped.push(departures.reduce((max, p) => p.Time.getTime() > max.Time.getTime() ? p : max, departures[0]));
        }
      }
      deduped.sort((a,b) => a.Time.getTime() - b.Time.getTime());

      // Step 2c: Collapse same-day all-arrival sequences
      const byDate = {};
      deduped.forEach(p => {
        const dateStr = p.Time.toLocaleDateString('en-GB'); // DD/MM/YYYY
        if (!byDate[dateStr]) byDate[dateStr] = [];
        byDate[dateStr].push(p);
      });

      let dedupedClean = [];
      Object.keys(byDate).sort().forEach(dStr => {
        const punches = byDate[dStr];
        const arrivalsToday = punches.filter(p => p.intent === 'arrival');
        const departuresToday = punches.filter(p => p.intent === 'departure');
        if (arrivalsToday.length > 1 && departuresToday.length === 0) {
          const first = arrivalsToday.reduce((min, p) => p.Time.getTime() < min.Time.getTime() ? p : min, arrivalsToday[0]);
          const strayTimes = arrivalsToday
            .filter(p => p !== first)
            .map(p => `${String(p.Time.getHours()).padStart(2,'0')}:${String(p.Time.getMinutes()).padStart(2,'0')}`)
            .join(', ');
          
          const firstCopy = { ...first };
          firstCopy.note = `${firstCopy.note || ''}; stray re-tap(s) discarded: ${strayTimes}`.replace(/^;\s*/, '');
          dedupedClean.push(firstCopy);
        } else {
          dedupedClean.push(...punches);
        }
      });
      dedupedClean.sort((a,b) => a.Time.getTime() - b.Time.getTime());
      deduped = dedupedClean;

      // Step 2d: Collapse consecutive orphan departures (day-only)
      if (!nightCap) {
        const collapsed = [];
        let i = 0;
        while (i < deduped.length) {
          const curr = deduped[i];
          if (curr.intent === 'departure') {
            const run = [curr];
            let j = i + 1;
            while (j < deduped.length && deduped[j].intent === 'departure') {
              run.push(deduped[j]);
              j++;
            }

            const byDateRun = {};
            run.forEach(p => {
              const dStr = p.Time.toLocaleDateString('en-GB');
              if (!byDateRun[dStr]) byDateRun[dStr] = [];
              byDateRun[dStr].push(p);
            });

            const collapsedRun = [];
            Object.keys(byDateRun).sort().forEach(d => {
              const maxP = byDateRun[d].reduce((max, p) => p.Time.getTime() > max.Time.getTime() ? p : max, byDateRun[d][0]);
              collapsedRun.push(maxP);
            });

            const hasPrecedingArrival = (collapsed.length > 0 && collapsed[collapsed.length - 1].intent === 'arrival');
            if (hasPrecedingArrival) {
              collapsed.push(collapsedRun[0]);
              if (collapsedRun.length > 1) {
                collapsed.push(...collapsedRun.slice(1));
              }
            } else {
              collapsed.push(...collapsedRun);
            }
            i = j;
          } else {
            collapsed.push(curr);
            i++;
          }
        }
        deduped = collapsed;
      }

      // Step 3: Pair arrivals to departures
      const n_d = deduped.length;
      const usedPair = new Set();
      const pairedSessions = [];

      for (let i = 0; i < n_d; i++) {
        if (usedPair.has(i)) continue;
        const curr = deduped[i];
        if (curr.intent !== 'arrival') continue;

        let best_j = null;
        for (let j = i + 1; j < n_d; j++) {
          if (usedPair.has(j)) continue;
          const nxt = deduped[j];
          const diff = nxt.Time.getTime() - curr.Time.getTime();
          if (diff > PAIR_WINDOW) break;
          if (nxt.intent === 'departure') {
            best_j = j;
            break; // greedy first
          }
        }

        const flags = [];
        if (curr.corrected) flags.push(curr.note);

        if (best_j !== null) {
          const dep = deduped[best_j];
          if (dep.corrected) flags.push(dep.note);
          usedPair.add(i);
          usedPair.add(best_j);
          pairedSessions.push({
            CheckIn: curr.Time,
            CheckOut: dep.Time,
            Flags: flags.join('; ')
          });
        } else {
          usedPair.add(i);
          flags.push("Missing C/Out");
          pairedSessions.push({
            CheckIn: curr.Time,
            CheckOut: null,
            Flags: flags.join('; ')
          });
        }
      }

      // Step 4: Orphan departures
      for (let i = 0; i < n_d; i++) {
        if (usedPair.has(i)) continue;
        const curr = deduped[i];
        const flags = [];
        if (curr.corrected) flags.push(curr.note);
        flags.push("Missing C/In");
        pairedSessions.push({
          CheckIn: null,
          CheckOut: curr.Time,
          Flags: flags.join('; ')
        });
      }

      // Step 5: Discard phantom arrivals inside a completed pair
      const pairedIntervals = pairedSessions
        .filter(s => s.CheckIn && s.CheckOut)
        .map(s => ({ lo: s.CheckIn.getTime(), hi: s.CheckOut.getTime() }));

      const cleanedSessions = [];
      pairedSessions.forEach(s => {
        if (!s.CheckOut && s.CheckIn) {
          const cinVal = s.CheckIn.getTime();
          const isPhantom = pairedIntervals.some(interval => cinVal > interval.lo && cinVal < interval.hi);
          if (isPhantom) return; // Discard phantom
        }
        cleanedSessions.push(s);
      });

      cleanedSessions.sort((a,b) => {
        const aTime = a.CheckIn ? a.CheckIn.getTime() : a.CheckOut.getTime();
        const bTime = b.CheckIn ? b.CheckIn.getTime() : b.CheckOut.getTime();
        return aTime - bTime;
      });

      // Step 6: Process session dates and assign configs
      cleanedSessions.forEach(s => {
        const cin = s.CheckIn;
        const cout = s.CheckOut;
        let sessionDate = null;

        if (cin) {
          sessionDate = new Date(cin.getFullYear(), cin.getMonth(), cin.getDate());
        } else {
          sessionDate = new Date(cout.getFullYear(), cout.getMonth(), cout.getDate());
          if (nightCap && cout.getHours() < 13) {
            // Night shift carry-over
            const weekdayConfigs = getAllowedShiftsForEmployee(empNorm, empName, null);
            const temp = getClosestShift(cout, weekdayConfigs, true);
            if (temp && temp.Is_Night_Shift) {
              sessionDate.setDate(sessionDate.getDate() - 1);
            }
          }
        }

        // Load specific allowed shifts for this date
        const allowedConfigs = getAllowedShiftsForEmployee(empNorm, empName, sessionDate);

        // Rest Day check
        if (allowedConfigs.length === 0) {
          sessions.push({
            Employee: empName,
            Date: sessionDate,
            Shift: 'Rest Day',
            CheckIn: cin,
            CheckOut: cout,
            LateMinutes: 0,
            AttendanceStatus: 'Rest Day — Not Scheduled',
            HoursWorked: null,
            DayAttendance: 'Rest Day',
            MissingCheckin: cin ? 'No' : 'Yes',
            MissingCheckout: cout ? 'No' : 'Yes',
            NightShift: 'No',
            DataQualityFlag: s.Flags ? `${s.Flags}; Punched on rest day` : 'Punched on rest day',
            Department: departmentRaw
          });
          return;
        }

        // Assign closest shift config
        let assigned = null;
        if (cin) {
          assigned = getClosestShift(cin, allowedConfigs, false);
        } else {
          assigned = getClosestShift(cout, allowedConfigs, true);
        }

        const isNightSession = assigned.Is_Night_Shift;
        const shiftName = `${assigned.Start_Time}-${assigned.End_Time}`;

        // Calculate lateness
        let lateMins = 0;
        let latenessStatus = '';
        if (cin) {
          const lateness = calculateLatenessMins(cin, assigned);
          lateMins = lateness.lateMins;
          latenessStatus = cout ? lateness.status : `${lateness.status} (Missing C/Out)`;
        } else {
          latenessStatus = "Unknown — no C/In recorded";
        }

        // Calculate hours worked
        let hoursWorked = null;
        if (cin && cout) {
          hoursWorked = Math.round(((cout.getTime() - cin.getTime()) / 3600000) * 100) / 100;
        }

        // Determine day attendance category
        let dayAttendance = 'Anomalous';
        let isFullDay = false;
        let isEarlyDep = false;

        if (cin && cout) {
          const isSaturday = sessionDate.getDay() === 6; // 6 = Saturday
          let startH, startM, endH, endM;

          if (isSaturday && !isNightSession) {
            // Saturday standard half-day rule
            startH = 8; startM = 0;
            endH = 14; endM = 0;
          } else {
            [startH, startM] = assigned.Start_Time.split(':').map(Number);
            [endH, endM] = assigned.End_Time.split(':').map(Number);
          }

          let shiftStartDt = new Date(sessionDate.getFullYear(), sessionDate.getMonth(), sessionDate.getDate(), startH, startM, 0);
          let shiftEndDt = new Date(sessionDate.getFullYear(), sessionDate.getMonth(), sessionDate.getDate(), endH, endM, 0);

          if (isNightSession) {
            shiftEndDt.setDate(shiftEndDt.getDate() + 1);
            if (cin.getHours() < 13) {
              shiftStartDt.setDate(shiftStartDt.getDate() - 1);
            }
          }

          const cinOffsetMins = (cin.getTime() - shiftStartDt.getTime()) / 60000;
          const coutOffsetMins = (cout.getTime() - shiftEndDt.getTime()) / 60000;

          const cinInWindow = cinOffsetMins >= -180 && cinOffsetMins <= 180;
          const leftEarly = coutOffsetMins < -90;

          if (cinInWindow && !leftEarly) {
            isFullDay = true;
          } else if (cinInWindow && leftEarly) {
            isEarlyDep = true;
          }
        }

        if (isFullDay) {
          dayAttendance = 'Full Day';
        } else if (isEarlyDep) {
          dayAttendance = 'Early Departure';
        } else if (!cin || !cout) {
          dayAttendance = 'Incomplete';
        } else {
          dayAttendance = 'Anomalous';
        }

        sessions.push({
          Employee: empName,
          Date: sessionDate,
          Shift: shiftName,
          CheckIn: cin,
          CheckOut: cout,
          LateMinutes: lateMins,
          AttendanceStatus: latenessStatus,
          HoursWorked: hoursWorked,
          DayAttendance: dayAttendance,
          MissingCheckin: cin ? 'No' : 'Yes',
          MissingCheckout: cout ? 'No' : 'Yes',
          NightShift: isNightSession ? 'Yes' : 'No',
          DataQualityFlag: s.Flags,
          Department: departmentRaw
        });
      });
    });

    // Helpers for allowed shifts and closest shift matching
    function getAllowedShiftsForEmployee(normName, empName, sessionDate) {
      let shiftText = String(shiftMap[normName] || '').trim();
      const departmentRaw = String(deptMap[normName] || '').trim();
      const department = departmentRaw.toUpperCase();
      const restDay = restDayMap[normName] ? String(restDayMap[normName]).trim().toLowerCase() : '';

      const txt = shiftText.toLowerCase().replace(/\s/g, '');
      const isDedicatedNight = txt.includes("6pm-8") || txt.includes("6pm-7am") || txt.includes("6pm-7:30") || txt.includes("6:30pm-8");
      
      const NIGHT_EXEMPT_DEPT_LABELS = new Set([
        "NURSES", "NURSE", "NURSE/PA", "HOUSEKEEPING( PA)", "HOUSEKEEPING(JANITORS)",
        "LABORATORY", "THEATRE", "FRONT OFFICE", "Nurse/PA", "Laboratory"
      ]);
      const isExemptDept = NIGHT_EXEMPT_DEPT_LABELS.has(department) || NIGHT_EXEMPT_DEPT_LABELS.has(departmentRaw);

      const isStandard = (restDay === 'sunday' || restDay === 'saturday') && !isDedicatedNight && !isExemptDept;

      if (isStandard && sessionDate) {
        const daysOfWeek = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
        const dayName = daysOfWeek[sessionDate.getDay()];
        if (dayName === restDay) {
          return []; // Rest day
        }
        if (dayName === 'saturday' || dayName === 'sunday') {
          return [{ Shift_ID: "Weekend_Half", Start_Time: "08:00", End_Time: "14:00", Is_Night_Shift: false }];
        }
      }

      // Shift text overrides
      if (txt.includes("6:30pm") || txt.includes("630pm")) {
        return [{ Shift_ID: "FO_Night_630", Start_Time: "18:30", End_Time: "08:00", Is_Night_Shift: true }];
      }
      if (txt.includes("6pm-8") || txt.includes("6pm- 8")) {
        return [{ Shift_ID: "Night_6to8", Start_Time: "18:00", End_Time: "08:00", Is_Night_Shift: true }];
      }
      if (txt.includes("11am-8") || txt.includes("11am8pm")) {
        return [{ Shift_ID: "Day_11to8", Start_Time: "11:00", End_Time: "20:00", Is_Night_Shift: false }];
      }
      if (txt.includes("10am-7") || txt.includes("10am7pm") || txt.includes("10:00-19:00")) {
        return [{ Shift_ID: "Day_10to7", Start_Time: "10:00", End_Time: "19:00", Is_Night_Shift: false }];
      }
      if (txt.includes("9am-6") || txt.includes("9am6pm") || txt.includes("09:00-18:00") || txt.includes("9:00-18:00")) {
        return [{ Shift_ID: "Day_9to6", Start_Time: "09:00", End_Time: "18:00", Is_Night_Shift: false }];
      }

      if (shiftText === "ORTHO_10_ROTATING") {
        return [cfg("Default_Day", "08:00", "17:00", false), cfg("Day_10to7", "10:00", "19:00", false)];
      }
      if (shiftText === "ORTHO_11_ROTATING") {
        return [cfg("Default_Day", "08:00", "17:00", false), cfg("Day_11to8", "11:00", "20:00", false)];
      }
      if (shiftText === "THEATRE_ROTATING") {
        return [
          cfg("PA_Nurse_Day", "07:30", "18:00", false),
          cfg("Day_10to7", "10:00", "19:00", false),
          cfg("Day_11to8", "11:00", "20:00", false)
        ];
      }
      if (department === "OTHER" || department === "DEFAULT" || department === "") {
        return [
          cfg("Default_Day", "08:00", "17:00", false),
          cfg("Day_10to7", "10:00", "19:00", false),
          cfg("Day_11to8", "11:00", "20:00", false)
        ];
      }
      if (shiftText === "FO_ROTATING") {
        return [
          cfg("FO_Day_8", "08:00", "17:00", false),
          cfg("FO_Day_9", "09:00", "18:00", false),
          cfg("FO_Day_10", "10:00", "19:00", false),
          cfg("FO_Day_11", "11:00", "20:00", false),
          cfg("FO_Night", "18:30", "08:00", true)
        ];
      }

      // Default configs for pool mapping
      const SHIFT_ID_DEFAULTS = {
        "Default_Day":    ["08:00", "17:00", false],
        "PA_Nurse_Day":   ["07:30", "18:00", false],
        "PA_Nurse_Night": ["18:00", "07:30", true],
        "Lab_Day_7":      ["07:00", "16:00", false],
        "Lab_Day_8":      ["08:00", "17:00", false],
        "Lab_Day_9":      ["09:00", "18:00", false],
        "Lab_Night":      ["18:00", "07:00", true],
        "Janitor_Day":    ["07:00", "16:00", false],
        "Janitor_Night":  ["18:00", "07:00", true],
        "FO_Day_8":       ["08:00", "17:00", false],
        "FO_Day_9":       ["09:00", "18:00", false],
        "FO_Day_10":      ["10:00", "19:00", false],
        "FO_Day_11":      ["11:00", "20:00", false],
        "FO_Night":       ["18:30", "08:00", true],
        "Night_6to8":     ["18:00", "08:00", true],
        "FO_Night_630":   ["18:30", "08:00", true],
        "Day_11to8":      ["11:00", "20:00", false],
        "Day_10to7":      ["10:00", "19:00", false],
        "Day_9to6":       ["09:00", "18:00", false],
        "Weekend_Half":   ["08:00", "14:00", false],
      };

      const BUILTIN_DEPT_RULES = {
        "NURSES":                  ["PA_Nurse_Day", "PA_Nurse_Night"],
        "NURSE":                   ["PA_Nurse_Day", "PA_Nurse_Night"],
        "NURSE/PA":                ["PA_Nurse_Day", "PA_Nurse_Night"],
        "HOUSEKEEPING( PA)":       ["PA_Nurse_Day", "PA_Nurse_Night"],
        "HOUSEKEEPING(JANITORS)":  ["Janitor_Day",  "Janitor_Night"],
        "HOUSEKEEPING(JANITORS) ": ["Janitor_Day",  "Janitor_Night"],
        "LABORATORY":              ["Lab_Day_7", "Lab_Day_8", "Lab_Day_9", "Lab_Night"],
        "THEATRE":                 ["PA_Nurse_Day", "PA_Nurse_Night"],
        "CLINICAL":                ["Default_Day", "FO_Day_9", "Day_11to8", "Night_6to8"],
        "FRONT OFFICE":            ["FO_Day_8", "FO_Day_9", "FO_Day_10", "FO_Day_11", "FO_Night"],
        "ORTHOPEDIC":              ["Default_Day", "Day_10to7", "Day_11to8"],
        "ORTHOPAEDIC":             ["Default_Day", "Day_10to7", "Day_11to8"],
      };

      function buildPool(sids) {
        return sids.map(sid => {
          if (sid in SHIFT_ID_DEFAULTS) {
            const [s, e, n] = SHIFT_ID_DEFAULTS[sid];
            return cfg(sid, s, e, n);
          }
          return cfg(sid, "08:00", "17:00", false);
        });
      }

      // 1. Try Department Rules sheet
      if (departmentRaw && departmentRaw in deptRules) {
        const pool = buildPool(deptRules[departmentRaw]);
        if (pool.length > 0) return pool;
      }

      // 2. Try Builtin Department rules
      if (department in BUILTIN_DEPT_RULES) {
        return buildPool(BUILTIN_DEPT_RULES[department]);
      }

      // Try normalized comparison
      const deptNorm = department.replace(/\s/g, '').replace(/\//g, '/');
      for (let key in BUILTIN_DEPT_RULES) {
        if (key.replace(/\s/g, '').replace(/\//g, '/') === deptNorm) {
          return buildPool(BUILTIN_DEPT_RULES[key]);
        }
      }

      // 3. Fallback
      return [cfg("Default_Day", "08:00", "17:00", false)];
    }

    function cfg(sid, start, end, night) {
      if (sid in shiftsConfigLookup) return shiftsConfigLookup[sid];
      return { Shift_ID: sid, Start_Time: start, End_Time: end, Is_Night_Shift: night };
    }

    function getClosestShift(refTime, allowedConfigs, isCheckoutOnly) {
      let bestConfig = null;
      let bestDiff = null;

      allowedConfigs.forEach(sh => {
        const targetTimeStr = isCheckoutOnly ? sh.End_Time : sh.Start_Time;
        const [sh_h, sh_m] = targetTimeStr.split(':').map(Number);
        
        // Same day
        const dt = new Date(refTime.getFullYear(), refTime.getMonth(), refTime.getDate(), sh_h, sh_m, 0);
        let diff = Math.abs(refTime.getTime() - dt.getTime());

        // Overnight shifts checks
        if (sh.Is_Night_Shift) {
          const dtPrev = new Date(refTime.getFullYear(), refTime.getMonth(), refTime.getDate() - 1, sh_h, sh_m, 0);
          const diffPrev = Math.abs(refTime.getTime() - dtPrev.getTime());
          if (diffPrev < diff) diff = diffPrev;

          const dtNext = new Date(refTime.getFullYear(), refTime.getMonth(), refTime.getDate() + 1, sh_h, sh_m, 0);
          const diffNext = Math.abs(refTime.getTime() - dtNext.getTime());
          if (diffNext < diff) diff = diffNext;
        }

        // Tiebreaker
        if (bestDiff !== null && Math.abs(diff - bestDiff) <= 1000) {
          const refHour = refTime.getHours();
          let preferNight = false;
          if (!isCheckoutOnly) {
            preferNight = refHour >= 17 && refHour <= 23;
          } else {
            preferNight = refHour >= 0 && refHour <= 10;
          }
          if (preferNight && sh.Is_Night_Shift && !bestConfig.Is_Night_Shift) {
            bestDiff = diff;
            bestConfig = sh;
            return;
          }
        }

        if (bestDiff === null || diff < bestDiff) {
          bestDiff = diff;
          bestConfig = sh;
        }
      });

      return bestConfig;
    }

    function calculateLatenessMins(cinTime, shiftConfig) {
      const [sh_h, sh_m] = shiftConfig.Start_Time.split(':').map(Number);
      const shiftStartDt = new Date(cinTime.getFullYear(), cinTime.getMonth(), cinTime.getDate(), sh_h, sh_m, 0);

      if (shiftConfig.Is_Night_Shift) {
        if (cinTime.getHours() < 12) {
          shiftStartDt.setDate(shiftStartDt.getDate() - 1);
        }
      }

      const diffMins = Math.floor((cinTime.getTime() - shiftStartDt.getTime()) / 60000);
      const lateMins = Math.max(0, diffMins - 6); // 6 min grace

      let status = "On Time";
      if (lateMins === 0) {
        status = "On Time";
      } else if (lateMins <= 15) {
        status = "Slightly Late";
      } else if (lateMins <= 60) {
        status = "Late";
      } else {
        status = "Very Late";
      }

      return { lateMins, status };
    }

    // 4. Generate Summaries
    log("Generating summaries & analytics sheets...");
    
    // Sort sessions
    sessions.sort((a, b) => {
      const nameComp = a.Employee.localeCompare(b.Employee);
      if (nameComp !== 0) return nameComp;
      return a.Date.getTime() - b.Date.getTime();
    });

    const scheduledSessions = sessions.filter(s => s.DayAttendance !== 'Rest Day');
    
    // Employee Detailed Summary
    const empSummaryMap = {};
    scheduledSessions.forEach(s => {
      if (!empSummaryMap[s.Employee]) {
        empSummaryMap[s.Employee] = {
          Employee: s.Employee,
          Total_Sessions: 0,
          Days_Attended: 0,
          Full_Days: 0,
          Early_Departures: 0,
          Incomplete_Days: 0,
          Total_Late_Minutes: 0,
          Hours_Worked_Sum: 0,
          Hours_Worked_Count: 0,
          Missing_Checkin_Days: 0,
          Missing_Checkout_Days: 0,
          Flagged_Records: 0
        };
      }

      const row = empSummaryMap[s.Employee];
      row.Total_Sessions++;
      row.Days_Attended++;
      
      if (s.DayAttendance === 'Full Day') row.Full_Days++;
      else if (s.DayAttendance === 'Early Departure') row.Early_Departures++;
      else if (s.DayAttendance === 'Incomplete') row.Incomplete_Days++;
      
      row.Total_Late_Minutes += s.LateMinutes;
      
      if (s.HoursWorked !== null) {
        row.Hours_Worked_Sum += s.HoursWorked;
        row.Hours_Worked_Count++;
      }

      if (s.MissingCheckin === 'Yes') row.Missing_Checkin_Days++;
      if (s.MissingCheckout === 'Yes') row.Missing_Checkout_Days++;
      if (s.DataQualityFlag && s.DataQualityFlag !== '') row.Flagged_Records++;
    });

    const employeeSummary = Object.keys(empSummaryMap).sort().map(k => {
      const row = empSummaryMap[k];
      const avgLate = row.Total_Sessions > 0 ? Math.round((row.Total_Late_Minutes / row.Total_Sessions) * 10) / 10 : 0;
      const avgHours = row.Hours_Worked_Count > 0 ? Math.round((row.Hours_Worked_Sum / row.Hours_Worked_Count) * 100) / 100 : 0;
      return {
        Employee: row.Employee,
        Total_Sessions: row.Total_Sessions,
        Days_Attended: row.Days_Attended,
        Full_Days: row.Full_Days,
        Early_Departures: row.Early_Departures,
        Incomplete_Days: row.Incomplete_Days,
        Total_Late_Minutes: row.Total_Late_Minutes,
        Average_Late_Minutes: avgLate,
        Average_Hours_Worked: avgHours,
        Total_Hours_Worked: Math.round(row.Hours_Worked_Sum * 100) / 100,
        Missing_Checkin_Days: row.Missing_Checkin_Days,
        Missing_Checkout_Days: row.Missing_Checkout_Days,
        Flagged_Records: row.Flagged_Records
      };
    });

    // Executive Status Summary
    const statusSummaryMap = {};
    scheduledSessions.forEach(s => {
      if (!statusSummaryMap[s.Employee]) {
        statusSummaryMap[s.Employee] = {
          Employee: s.Employee,
          Days_Attended: 0,
          Full_Days: 0,
          Early_Departures: 0,
          OnTime: 0,
          SlightlyLate: 0,
          Late: 0,
          VeryLate: 0,
          MissingCin: 0,
          MissingCout: 0,
          TotalSessions: 0
        };
      }

      const row = statusSummaryMap[s.Employee];
      row.Days_Attended++;
      row.TotalSessions++;
      
      if (s.DayAttendance === 'Full Day') row.Full_Days++;
      else if (s.DayAttendance === 'Early Departure') row.Early_Departures++;
      
      if (s.MissingCheckout === 'Yes') row.MissingCout++;

      if (s.MissingCheckin === 'Yes') {
        row.MissingCin++;
      } else {
        const status = s.AttendanceStatus;
        if (status.startsWith('On Time')) row.OnTime++;
        else if (status.startsWith('Slightly Late')) row.SlightlyLate++;
        else if (status.startsWith('Very Late')) row.VeryLate++;
        else if (status.startsWith('Late')) row.Late++;
      }
    });

    const statusSummary = Object.keys(statusSummaryMap).sort().map(k => {
      const row = statusSummaryMap[k];
      const validCinDays = row.OnTime + row.SlightlyLate + row.Late + row.VeryLate;
      const lateDays = row.SlightlyLate + row.Late + row.VeryLate;
      const latePct = validCinDays > 0 ? Math.round((lateDays / validCinDays) * 100 * 10) / 10 : 0.0;
      
      return {
        Employee: row.Employee,
        "Days Attended": row.Days_Attended,
        "Full Days": row.Full_Days,
        "Early Departures": row.Early_Departures,
        "On Time": row.OnTime,
        "Slightly Late": row.SlightlyLate,
        "Late": row.Late,
        "Very Late": row.VeryLate,
        "Missing C/In (Unknown Lateness)": row.MissingCin,
        "Missing C/Out Days": row.MissingCout,
        Total_Sessions: row.TotalSessions,
        "Valid C/In Days": validCinDays,
        "Late Days": lateDays,
        "Late Percentage": latePct
      };
    });

    const summaryMetrics = computeSummaryMetrics(sessions, statusSummary, unmatchedEmployees, duplicates);

    return {
      sessions,
      employeeSummary,
      statusSummary,
      unmatchedEmployees,
      duplicates,
      summaryMetrics
    };
  }

  async function downloadExcel(results, log) {
    const { sessions, employeeSummary, statusSummary, summaryMetrics = {} } = results;

    function toExcelLocalSerial(date) {
      if (!(date instanceof Date)) return date;
      return 25569 + Date.UTC(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        date.getHours(),
        date.getMinutes(),
        date.getSeconds(),
        date.getMilliseconds()
      ) / 86400000;
    }

    function toExcelLocalDateSerial(date) {
      if (!(date instanceof Date)) return date;
      return 25569 + Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000;
    }

    log("Building formatted Excel workbook...");
    const wb = new ExcelJS.Workbook();
    
    // Font & Styles configurations
    const FONT_FAMILY = 'Calibri';
    const f_title = { name: FONT_FAMILY, size: 16, bold: true, color: { argb: 'FF1B4F72' } };
    const f_meta = { name: FONT_FAMILY, size: 9, italic: true, color: { argb: 'FF5D6D7E' } };
    const f_header = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    const f_data = { name: FONT_FAMILY, size: 10, color: { argb: 'FF2C3E50' } };
    const f_data_bold = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF2C3E50' } };
    const f_kpi_lbl = { name: FONT_FAMILY, size: 9, color: { argb: 'FF566573' } };
    const f_kpi_val = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: 'FF1B4F72' } };

    const fill_header = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };
    const fill_zebra = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9F9' } };
    const fill_white = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    const fill_kpi = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAECEE' } };
    
    const border_thin = {
      top: { style: 'thin', color: { argb: 'FFBDC3C7' } },
      left: { style: 'thin', color: { argb: 'FFBDC3C7' } },
      bottom: { style: 'thin', color: { argb: 'FFBDC3C7' } },
      right: { style: 'thin', color: { argb: 'FFBDC3C7' } }
    };
    const border_double_bottom = {
      top: { style: 'thin', color: { argb: 'FFBDC3C7' } },
      bottom: { style: 'double', color: { argb: 'FF2C3E50' } }
    };

    const align_left = { horizontal: 'left', vertical: 'middle' };
    const align_center = { horizontal: 'center', vertical: 'middle' };
    const align_right = { horizontal: 'right', vertical: 'middle' };

    // Date strings for metadata
    let minD = new Date();
    let maxD = new Date(0);
    sessions.forEach(s => {
      if (s.Date.getTime() < minD.getTime()) minD = s.Date;
      if (s.Date.getTime() > maxD.getTime()) maxD = s.Date;
    });
    
    const formatDateStr = (d) => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    const minDateStr = formatDateStr(minD);
    const maxDateStr = formatDateStr(maxD);
    const genTimeStr = `${formatDateStr(new Date())} ${String(new Date().getHours()).padStart(2,'0')}:${String(new Date().getMinutes()).padStart(2,'0')}`;

    // -----------------------------------------
    // SHEET: Status Summary
    // -----------------------------------------
    log("Styling sheet: Status Summary...");
    const ws_status = wb.addWorksheet('Status Summary');
    
    // Title Blocks
    ws_status.getCell('A1').value = 'OASIS SPECIALIST HOSPITAL';
    ws_status.getCell('A1').font = f_title;
    ws_status.getCell('A2').value = `Reporting Period: ${minDateStr} to ${maxDateStr}   |   Generated: ${genTimeStr}`;
    ws_status.getCell('A2').font = f_meta;
    
    // KPI card calculation
    const totalEmployees = summaryMetrics.totalEmployees ?? statusSummary.length;
    const totalShiftsLogged = sessions.length;
    const avgLatePct = summaryMetrics.latenessRate !== undefined
      ? Math.round(summaryMetrics.latenessRate * 10) / 10
      : 0;
    
    const totalDaysAttended = summaryMetrics.totalDaysAttended ?? statusSummary.reduce((sum, s) => sum + s["Days Attended"], 0);
    const totalMissingCin = summaryMetrics.totalMissingCin ?? statusSummary.reduce((sum, s) => sum + s["Missing C/In (Unknown Lateness)"], 0);
    const incompletePct = totalDaysAttended > 0 ? Math.round((totalMissingCin / totalDaysAttended) * 100 * 10) / 10 : 0;

    // Render KPIs
    renderKPICard(ws_status, 1, 'EMPLOYEES AUDITED', totalEmployees);
    renderKPICard(ws_status, 3, 'TOTAL SHIFTS LOGGED', totalShiftsLogged);
    renderKPICard(ws_status, 5, 'OVERALL LATE PERCENTAGE', `${avgLatePct}%`);
    renderKPICard(ws_status, 7, 'INCOMPLETE SHIFT RATE', `${incompletePct}%`);

    function renderKPICard(ws, colStart, label, val) {
      ws.mergeCells(4, colStart, 4, colStart + 1);
      ws.mergeCells(5, colStart, 5, colStart + 1);
      
      const c_lbl = ws.getCell(4, colStart);
      c_lbl.value = label;
      c_lbl.font = f_kpi_lbl;
      c_lbl.fill = fill_kpi;
      c_lbl.alignment = align_center;
      
      const c_val = ws.getCell(5, colStart);
      c_val.value = val;
      c_val.font = f_kpi_val;
      c_val.fill = fill_kpi;
      c_val.alignment = align_center;

      for (let r = 4; r <= 5; r++) {
        for (let c = colStart; c <= colStart + 1; c++) {
          ws.getCell(r, c).border = border_thin;
        }
      }
    }

    // Table Headers
    const statusHeaders = Object.keys(statusSummary[0] || {});
    ws_status.getRow(7).values = statusHeaders;
    ws_status.getRow(7).height = 28;
    
    for (let c = 1; c <= statusHeaders.length; c++) {
      const cell = ws_status.getCell(7, c);
      cell.font = f_header;
      cell.fill = fill_header;
      cell.alignment = align_center;
      cell.border = border_thin;
    }

    // Populate Data rows
    const fill_p_green = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2F0D9' } };
    const font_p_green = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF385723' } };
    const fill_p_amber = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    const font_p_amber = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF7F6000' } };
    const fill_p_red   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
    const font_p_red   = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FFC65911' } };
    const fill_p_gray  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAECEE' } };
    const font_p_gray  = { name: FONT_FAMILY, size: 10, color: { argb: 'FF5D6D7E' } };

    statusSummary.forEach((row, rIdx) => {
      const rowNum = 8 + rIdx;
      ws_status.getRow(rowNum).height = 20;
      const isZebra = rowNum % 2 === 0;
      const fill = isZebra ? fill_zebra : fill_white;

      statusHeaders.forEach((key, cIdx) => {
        const colNum = 1 + cIdx;
        const cell = ws_status.getCell(rowNum, colNum);
        cell.value = row[key];
        cell.font = f_data;
        cell.fill = fill;
        cell.border = border_thin;
        cell.alignment = colNum === 1 ? align_left : align_center;

        // Conditional formatting on Late Percentage (Column 14)
        if (key === 'Late Percentage') {
          cell.font = f_data_bold;
          cell.alignment = align_right;
          const val = Number(row[key]);
          if (!isNaN(val)) {
            if (val <= 10.0) {
              cell.fill = fill_p_green;
              cell.font = font_p_green;
            } else if (val <= 30.0) {
              cell.fill = fill_p_amber;
              cell.font = font_p_amber;
            } else {
              cell.fill = fill_p_red;
              cell.font = font_p_red;
            }
          }
        }

        // Highlight Missing C/Out Days (Column 10)
        if (key === 'Missing C/Out Days') {
          const val = parseInt(row[key], 10);
          if (val > 0) {
            cell.fill = fill_p_gray;
            cell.font = font_p_gray;
          }
        }
      });
    });

    // Summary Row at the bottom
    const summaryRowLoc = 8 + statusSummary.length;
    ws_status.getCell(summaryRowLoc, 1).value = 'Total / Average';
    ws_status.getCell(summaryRowLoc, 1).font = f_data_bold;
    ws_status.getCell(summaryRowLoc, 1).alignment = align_left;
    ws_status.getCell(summaryRowLoc, 1).border = border_double_bottom;

    for (let c = 2; c <= statusHeaders.length; c++) {
      const colLetter = ws_status.getColumn(c).letter;
      const cell = ws_status.getCell(summaryRowLoc, c);
      cell.font = f_data_bold;
      cell.border = border_double_bottom;
      
      if (statusHeaders[c - 1] === 'Late Percentage') {
        cell.value = { formula: `AVERAGE(${colLetter}8:${colLetter}${summaryRowLoc - 1})` };
        cell.alignment = align_right;
        cell.numFmt = '0.0';
      } else {
        cell.value = { formula: `SUM(${colLetter}8:${colLetter}${summaryRowLoc - 1})` };
        cell.alignment = align_center;
      }
    }

    // Freeze top header and employee column
    ws_status.views = [{ state: 'frozen', xSplit: 1, ySplit: 7 }];

    // -----------------------------------------
    // SHEET: Employee Summary
    // -----------------------------------------
    log("Styling sheet: Employee Summary...");
    const ws_emp = wb.addWorksheet('Employee Summary');
    
    const empHeaders = Object.keys(employeeSummary[0] || {});
    ws_emp.getRow(1).values = empHeaders;
    ws_emp.getRow(1).height = 26;
    
    for (let c = 1; c <= empHeaders.length; c++) {
      const cell = ws_emp.getCell(1, c);
      cell.font = f_header;
      cell.fill = fill_header;
      cell.alignment = align_center;
      cell.border = border_thin;
    }

    employeeSummary.forEach((row, rIdx) => {
      const rowNum = 2 + rIdx;
      ws_emp.getRow(rowNum).height = 19;
      const isZebra = rowNum % 2 === 0;
      const fill = isZebra ? fill_zebra : fill_white;

      empHeaders.forEach((key, cIdx) => {
        const colNum = 1 + cIdx;
        const cell = ws_emp.getCell(rowNum, colNum);
        cell.value = row[key];
        cell.font = f_data;
        cell.fill = fill;
        cell.border = border_thin;

        if (colNum === 1) {
          cell.alignment = align_left;
        } else if (key === 'Average_Late_Minutes' || key === 'Average_Hours_Worked' || key === 'Total_Hours_Worked') {
          cell.alignment = align_right;
          cell.numFmt = '0.00';
        } else {
          cell.alignment = align_center;
        }
      });
    });

    ws_emp.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

    // -----------------------------------------
    // SHEET: Sessions
    // -----------------------------------------
    log("Styling sheet: Sessions...");
    const ws_sess = wb.addWorksheet('Sessions');
    
    const sessHeaders = [
      "Employee", "Date", "Shift", "Check-In", "Check-Out", 
      "Late Minutes", "Attendance Status", "Hours Worked", "Day Attendance", 
      "Missing Checkin", "Missing Checkout", "Night Shift", "Data Quality Flag"
    ];
    
    ws_sess.getRow(1).values = sessHeaders;
    ws_sess.getRow(1).height = 26;
    
    for (let c = 1; c <= sessHeaders.length; c++) {
      const cell = ws_sess.getCell(1, c);
      cell.font = f_header;
      cell.fill = fill_header;
      cell.alignment = align_center;
      cell.border = border_thin;
    }

    const fill_p_orange = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    const font_p_orange = { name: FONT_FAMILY, size: 10, color: { argb: 'FF7F6000' } };

    sessions.forEach((s, rIdx) => {
      const rowNum = 2 + rIdx;
      ws_sess.getRow(rowNum).height = 19;
      const isZebra = rowNum % 2 === 0;
      const fill = isZebra ? fill_zebra : fill_white;

      ws_sess.getCell(rowNum, 1).value = s.Employee;
      ws_sess.getCell(rowNum, 2).value = toExcelLocalDateSerial(s.Date);
      ws_sess.getCell(rowNum, 3).value = s.Shift;
      ws_sess.getCell(rowNum, 4).value = s.CheckIn ? toExcelLocalSerial(s.CheckIn) : null;
      ws_sess.getCell(rowNum, 5).value = s.CheckOut ? toExcelLocalSerial(s.CheckOut) : null;
      ws_sess.getCell(rowNum, 6).value = s.LateMinutes;
      ws_sess.getCell(rowNum, 7).value = s.AttendanceStatus;
      ws_sess.getCell(rowNum, 8).value = s.HoursWorked;
      ws_sess.getCell(rowNum, 9).value = s.DayAttendance;
      ws_sess.getCell(rowNum, 10).value = s.MissingCheckin;
      ws_sess.getCell(rowNum, 11).value = s.MissingCheckout;
      ws_sess.getCell(rowNum, 12).value = s.NightShift;
      ws_sess.getCell(rowNum, 13).value = s.DataQualityFlag;

      for (let c = 1; c <= sessHeaders.length; c++) {
        const cell = ws_sess.getCell(rowNum, c);
        cell.font = f_data;
        cell.fill = fill;
        cell.border = border_thin;

        if (c === 1) {
          cell.alignment = align_left;
        } else if (c === 2) {
          cell.alignment = align_center;
          cell.numFmt = 'yyyy-mm-dd';
        } else if (c === 4 || c === 5) {
          cell.alignment = align_center;
          if (cell.value) cell.numFmt = 'yyyy-mm-dd hh:mm';
        } else if (c === 8) {
          cell.alignment = align_right;
          cell.numFmt = '0.00';
        } else {
          cell.alignment = align_center;
        }
      }

      // Conditional styling on Attendance Status (Column 7)
      const statusCell = ws_sess.getCell(rowNum, 7);
      const val = String(s.AttendanceStatus);
      if (val.startsWith("On Time")) {
        statusCell.fill = fill_p_green;
        statusCell.font = font_p_green;
      } else if (val.startsWith("Slightly Late")) {
        statusCell.fill = fill_p_amber;
        statusCell.font = font_p_amber;
      } else if (val.startsWith("Late") || val.startsWith("Very Late")) {
        statusCell.fill = fill_p_red;
        statusCell.font = font_p_red;
      } else {
        statusCell.fill = fill_p_gray;
        statusCell.font = font_p_gray;
      }

      // Highlight Data Quality Warnings (Column 13)
      const flagCell = ws_sess.getCell(rowNum, 13);
      if (s.DataQualityFlag && String(s.DataQualityFlag).trim() !== '') {
        flagCell.fill = fill_p_orange;
        flagCell.font = font_p_orange;
      }
    });

    ws_sess.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

    // -----------------------------------------
    // AUTO-FIT COLUMN WIDTHS FOR ALL SHEETS
    // -----------------------------------------
    log("Optimizing column widths...");
    [ws_status, ws_emp, ws_sess].forEach(ws => {
      ws.columns.forEach(col => {
        let maxWidth = 10;
        
        col.eachCell({ includeEmpty: false }, cell => {
          if (ws.name === 'Status Summary' && cell.row < 7) return; // ignore title and KPIs
          
          let cellValStr = '';
          if (cell.value instanceof Date) {
            cellValStr = 'YYYY-MM-DD HH:MM'; // estimate length
          } else if (cell.value && cell.value.formula) {
            cellValStr = '123456.78'; // estimate result length
          } else {
            cellValStr = String(cell.value || '');
          }
          
          if (cellValStr.length > maxWidth) {
            maxWidth = cellValStr.length;
          }
        });
        
        col.width = Math.min(Math.max(maxWidth + 4, 12), 45);
      });
    });

    // -----------------------------------------
    // SAVE / DOWNLOAD FILE
    // -----------------------------------------
    log("Compiling file binary and launching download...");
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Attendance_Report.xlsx';
    document.body.appendChild(a);
    a.click();
    
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  }

  // Expose engine to global window object
  window.AttendanceEngine = {
    normalizeName,
    parseExcelDate,
    calculate,
    downloadExcel
  };
})();
