/**
 * TimetableManager.gs
 * Manages timetable data, class schedules, and date/time calculations
 * Handles manual entry, bulk import, and validation
 */

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================

const TIMETABLE_SHEET_NAME = 'SequencerData';
const TIMETABLE_SHEET_ID = '1w8OSfgjmlmAYgjDnmq_1QxOBv6UpFNN9ow9XBUakC8U';
const TIMETABLE_TAB = 'Timetable';
const CLASS_MAPPING_TAB = 'ClassMapping';

// Timetable column indices (0-based)
const TT_COL = {
  CLASS_CODE: 0,
  DAY: 1,
  PERIOD: 2,
  START_TIME: 3,
  END_TIME: 4,
  ROOM: 5,
  SLOT_TYPE: 6,
  ACTIVITY: 7,
  PERIOD_LABEL: 8
};

const TT_SLOT_TYPES = {
  CLASS: 'CLASS',
  FREE: 'FREE',
  DUTY: 'DUTY',
  RECESS: 'RECESS',
  LUNCH: 'LUNCH',
  BEFORE_SCHOOL: 'BEFORE_SCHOOL',
  AFTER_SCHOOL: 'AFTER_SCHOOL',
  OTHER: 'OTHER'
};

const TT_HEADERS = ['classCode', 'day', 'period', 'startTime', 'endTime', 'room', 'slotType', 'activity', 'periodLabel'];

// Class mapping column indices
const MAP_COL = {
  CLASS_CODE: 0,
  COURSE_ID: 1,
  COURSE_NAME: 2
};

function normalizeSlotType_(v) {
  var t = String(v || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/\//g, '_');
  if (!t) return TT_SLOT_TYPES.CLASS;
  if (t === 'OFF') return TT_SLOT_TYPES.FREE;
  if (t === 'BEFORE' || t === 'BEFORESCHOOL') return TT_SLOT_TYPES.BEFORE_SCHOOL;
  if (t === 'AFTER' || t === 'AFTERSCHOOL') return TT_SLOT_TYPES.AFTER_SCHOOL;
  if (
    t === TT_SLOT_TYPES.CLASS ||
    t === TT_SLOT_TYPES.FREE ||
    t === TT_SLOT_TYPES.DUTY ||
    t === TT_SLOT_TYPES.RECESS ||
    t === TT_SLOT_TYPES.LUNCH ||
    t === TT_SLOT_TYPES.BEFORE_SCHOOL ||
    t === TT_SLOT_TYPES.AFTER_SCHOOL ||
    t === TT_SLOT_TYPES.OTHER
  ) return t;
  return TT_SLOT_TYPES.OTHER;
}

function isSchedulableClassSlot_(slot) {
  var t = normalizeSlotType_(slot && slot.slotType);
  return t === TT_SLOT_TYPES.CLASS;
}

function slotToRow_(classCode, slot) {
  return [
    String(classCode || '').trim(),
    Number(slot.day),
    Number(slot.period),
    slot.startTime || '',
    slot.endTime || '',
    slot.room || '',
    normalizeSlotType_(slot.slotType),
    slot.activity || '',
    slot.periodLabel || ''
  ];
}

function ensureTimetableSchema_(sheet) {
  if (!sheet) return;
  var needCols = TT_HEADERS.length;
  if (sheet.getMaxColumns() < needCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  }

  var header = sheet.getRange(1, 1, 1, needCols).getValues()[0];
  var changed = false;
  for (var i = 0; i < needCols; i++) {
    if (String(header[i] || '').trim() !== TT_HEADERS[i]) {
      header[i] = TT_HEADERS[i];
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, needCols).setValues([header]);
    sheet.getRange(1, 1, 1, needCols).setFontWeight('bold');
  }
}

// ============================================================================
// MAIN API FUNCTIONS (Called from HTML)
// ============================================================================

/**
 * Get complete timetable for all classes or specific class
 * @param {string} classCode - Optional: specific class code
 * @returns {Object} Timetable data organized by class
 */
function apiGetTimetable(classCode = null) {
  try {
    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(TIMETABLE_TAB);

    if (!sheet) {
      return {
        success: false,
        error: 'Timetable not found. Please create it first.'
      };
    }

    ensureTimetableSchema_(sheet);

    const lastRow = sheet.getLastRow();
    const timetable = {};
    if (lastRow < 2) {
      return {
        success: true,
        timetable: classCode ? [] : {}
      };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, TT_HEADERS.length).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[TT_COL.CLASS_CODE]) continue;

      const code = String(row[TT_COL.CLASS_CODE] || '').trim();
      if (classCode && code !== String(classCode || '').trim()) continue;

      if (!timetable[code]) timetable[code] = [];

      timetable[code].push({
        day: Number(row[TT_COL.DAY]),
        period: Number(row[TT_COL.PERIOD]),
        startTime: formatTimeForUi_(row[TT_COL.START_TIME]),
        endTime: formatTimeForUi_(row[TT_COL.END_TIME]),
        room: row[TT_COL.ROOM] || '',
        slotType: normalizeSlotType_(row[TT_COL.SLOT_TYPE]),
        activity: row[TT_COL.ACTIVITY] || '',
        periodLabel: row[TT_COL.PERIOD_LABEL] || ''
      });
    }

    Object.keys(timetable).forEach(code => {
      timetable[code].sort((a, b) => {
        if (Number(a.day) !== Number(b.day)) return Number(a.day) - Number(b.day);
        if (Number(a.period) !== Number(b.period)) return Number(a.period) - Number(b.period);
        return String(a.startTime || '').localeCompare(String(b.startTime || ''));
      });
    });

    return {
      success: true,
      timetable: classCode ? (timetable[classCode] || []) : timetable
    };

  } catch (e) {
    Logger.log('apiGetTimetable error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Save timetable slots for a class
 * Replaces existing slots for that class
 * @param {string} classCode - Class code
 * @param {Array} slots - Array of slot objects
 * @returns {Object} Result
 */
function apiSaveTimetable(classCode, slots) {
  try {
    if (!classCode || !slots || !Array.isArray(slots)) {
      throw new Error('Invalid parameters: classCode and slots array required');
    }

    for (let slot of slots) {
      validateTimetableSlot(slot);
    }

    const code = String(classCode || '').trim();
    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(TIMETABLE_TAB);
    if (!sheet) throw new Error('Timetable tab not found');

    ensureTimetableSchema_(sheet);

    const lastRow = sheet.getLastRow();
    const existing = [['classCode', 'day', 'period', 'startTime', 'endTime', 'room', 'slotType', 'activity', 'periodLabel']];
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, TT_HEADERS.length).getValues();
      rows.forEach(function (row) {
        if (String(row[TT_COL.CLASS_CODE] || '').trim() !== code) {
          existing.push([
            String(row[TT_COL.CLASS_CODE] || '').trim(),
            row[TT_COL.DAY],
            row[TT_COL.PERIOD],
            formatTimeForUi_(row[TT_COL.START_TIME]),
            formatTimeForUi_(row[TT_COL.END_TIME]),
            row[TT_COL.ROOM] || '',
            normalizeSlotType_(row[TT_COL.SLOT_TYPE]),
            row[TT_COL.ACTIVITY] || '',
            row[TT_COL.PERIOD_LABEL] || ''
          ]);
        }
      });
    }

    slots.forEach(function (slot) {
      existing.push(slotToRow_(code, slot));
    });

    sheet.clearContents();
    sheet.getRange(1, 1, existing.length, TT_HEADERS.length).setValues(existing);

    return {
      success: true,
      message: 'Timetable saved: ' + slots.length + ' slots for ' + code
    };

  } catch (e) {
    Logger.log('apiSaveTimetable error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Delete timetable for a class
 * @param {string} classCode - Class code to delete
 * @returns {Object} Result
 */
function apiDeleteTimetable(classCode) {
  try {
    const code = String(classCode || '').trim();
    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(TIMETABLE_TAB);

    if (!sheet) {
      throw new Error('Timetable tab not found');
    }

    ensureTimetableSchema_(sheet);

    const lastRow = sheet.getLastRow();
    const out = [TT_HEADERS.slice()];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, TT_HEADERS.length).getValues();
      data.forEach(function (row) {
        if (String(row[TT_COL.CLASS_CODE] || '').trim() !== code) {
          out.push([
            String(row[TT_COL.CLASS_CODE] || '').trim(),
            row[TT_COL.DAY],
            row[TT_COL.PERIOD],
            formatTimeForUi_(row[TT_COL.START_TIME]),
            formatTimeForUi_(row[TT_COL.END_TIME]),
            row[TT_COL.ROOM] || '',
            normalizeSlotType_(row[TT_COL.SLOT_TYPE]),
            row[TT_COL.ACTIVITY] || '',
            row[TT_COL.PERIOD_LABEL] || ''
          ]);
        }
      });
    }

    sheet.clearContents();
    sheet.getRange(1, 1, out.length, TT_HEADERS.length).setValues(out);

    return {
      success: true,
      message: 'Timetable deleted for ' + code
    };

  } catch (e) {
    Logger.log('apiDeleteTimetable error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Bulk import timetable from CSV data
 * Format: classCode,day,period,startTime,endTime,room
 * @param {string} csvData - CSV formatted timetable data
 * @param {boolean} replaceAll - If true, clear existing data first
 * @returns {Object} Result with import summary
 */
function apiImportTimetable(csvData, replaceAll = false) {
  try {
    const lines = String(csvData || '').trim().split(/\r?\n/);
    const imported = [];
    const errors = [];

    lines.forEach((line, idx) => {
      if (!line || !String(line).trim()) return;

      const raw = String(line).trim();
      const lower = raw.toLowerCase();
      if (idx === 0 && (lower.includes('classcode') || lower.includes('class code'))) return;

      const parts = raw.split(',').map(p => String(p || '').trim());
      if (parts.length < 6) {
        errors.push('Line ' + (idx + 1) + ': Not enough columns (need classCode,day,period,startTime,endTime,slotType[,activity])');
        return;
      }

      // Preferred format:
      // classCode,day,period,startTime,endTime,slotType,activity
      // Legacy format still accepted:
      // classCode,day,period,startTime,endTime,room,slotType,activity,periodLabel
      var slotTypePart = parts[5] || 'CLASS';
      var roomPart = '';
      var activityPart = parts[6] || '';
      var periodLabelPart = parts[7] || '';

      var normalizedCol6 = normalizeSlotType_(parts[6] || '');
      var col6IsType = (
        normalizedCol6 === TT_SLOT_TYPES.CLASS ||
        normalizedCol6 === TT_SLOT_TYPES.FREE ||
        normalizedCol6 === TT_SLOT_TYPES.DUTY ||
        normalizedCol6 === TT_SLOT_TYPES.RECESS ||
        normalizedCol6 === TT_SLOT_TYPES.LUNCH ||
        normalizedCol6 === TT_SLOT_TYPES.BEFORE_SCHOOL ||
        normalizedCol6 === TT_SLOT_TYPES.AFTER_SCHOOL ||
        normalizedCol6 === TT_SLOT_TYPES.OTHER
      );

      if (parts.length >= 7 && col6IsType && normalizeSlotType_(parts[5] || '') === TT_SLOT_TYPES.OTHER) {
        roomPart = parts[5] || '';
        slotTypePart = parts[6] || 'CLASS';
        activityPart = parts[7] || '';
        periodLabelPart = parts[8] || '';
      }

      const slot = {
        classCode: parts[0],
        day: parseInt(parts[1], 10),
        period: parseInt(parts[2], 10),
        startTime: parts[3],
        endTime: parts[4],
        room: roomPart,
        slotType: slotTypePart,
        activity: activityPart,
        periodLabel: periodLabelPart || ('Period ' + parseInt(parts[2], 10))
      };

      try {
        validateTimetableSlot(slot);
        imported.push(slot);
      } catch (e) {
        errors.push('Line ' + (idx + 1) + ': ' + e.message);
      }
    });

    if (imported.length === 0) {
      throw new Error('No valid slots to import');
    }

    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(TIMETABLE_TAB);
    ensureTimetableSchema_(sheet);

    const out = [TT_HEADERS.slice()];

    if (!replaceAll && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, TT_HEADERS.length).getValues();
      data.forEach(row => {
        out.push([
          String(row[TT_COL.CLASS_CODE] || '').trim(),
          row[TT_COL.DAY],
          row[TT_COL.PERIOD],
          formatTimeForUi_(row[TT_COL.START_TIME]),
          formatTimeForUi_(row[TT_COL.END_TIME]),
          row[TT_COL.ROOM] || '',
          normalizeSlotType_(row[TT_COL.SLOT_TYPE]),
          row[TT_COL.ACTIVITY] || '',
          row[TT_COL.PERIOD_LABEL] || ''
        ]);
      });
    }

    imported.forEach(slot => {
      out.push(slotToRow_(slot.classCode, slot));
    });

    sheet.clearContents();
    sheet.getRange(1, 1, out.length, TT_HEADERS.length).setValues(out);

    return {
      success: true,
      imported: imported.length,
      errors: errors,
      message: 'Imported ' + imported.length + ' slots. ' + errors.length + ' errors.'
    };

  } catch (e) {
    Logger.log('apiImportTimetable error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Export timetable to CSV format
 * @param {string} classCode - Optional: export specific class only
 * @returns {Object} Result with CSV data
 */
function apiExportTimetable(classCode = null) {
  try {
    const result = apiGetTimetable(classCode);

    if (!result.success) {
      return result;
    }

    const timetable = result.timetable;
    const rows = ['classCode,day,period,startTime,endTime,room,slotType,activity,periodLabel'];

    if (classCode) {
      (timetable || []).forEach(slot => {
        rows.push(
          [
            classCode,
            slot.day,
            slot.period,
            slot.startTime,
            slot.endTime,
            slot.room || '',
            normalizeSlotType_(slot.slotType),
            slot.activity || '',
            slot.periodLabel || ''
          ].join(',')
        );
      });
    } else {
      Object.keys(timetable || {}).forEach(code => {
        (timetable[code] || []).forEach(slot => {
          rows.push(
            [
              code,
              slot.day,
              slot.period,
              slot.startTime,
              slot.endTime,
              slot.room || '',
              normalizeSlotType_(slot.slotType),
              slot.activity || '',
              slot.periodLabel || ''
            ].join(',')
          );
        });
      });
    }

    return {
      success: true,
      csv: rows.join('\n')
    };

  } catch (e) {
    Logger.log('apiExportTimetable error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================================================
// CLASS MAPPING FUNCTIONS
// ============================================================================

/**
 * Get all class mappings (classCode → Classroom courseId)
 * @returns {Object} Result with mappings array
 */
function apiGetClassMappings() {
  try {
    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(CLASS_MAPPING_TAB);
    
    if (!sheet) {
      return {
        success: false,
        error: 'ClassMapping tab not found'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const mappings = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[MAP_COL.CLASS_CODE]) continue;
      
      mappings.push({
        classCode: row[MAP_COL.CLASS_CODE],
        courseId: row[MAP_COL.COURSE_ID],
        courseName: row[MAP_COL.COURSE_NAME] || ''
      });
    }
    
    return {
      success: true,
      mappings: mappings
    };
    
  } catch (e) {
    Logger.log('apiGetClassMappings error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Save class mapping
 * @param {string} classCode - Class code
 * @param {string} courseId - Google Classroom course ID
 * @param {string} courseName - Course name (optional)
 * @returns {Object} Result
 */
function apiSaveClassMapping(classCode, courseId, courseName = '') {
  try {
    if (!classCode || !courseId) {
      throw new Error('classCode and courseId are required');
    }
    
    const ss = getTimetableSheet();
    const sheet = ss.getSheetByName(CLASS_MAPPING_TAB);
    
    if (!sheet) {
      throw new Error('ClassMapping tab not found');
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Check if mapping exists
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MAP_COL.CLASS_CODE] === classCode) {
        // Update existing
        data[i][MAP_COL.COURSE_ID] = courseId;
        data[i][MAP_COL.COURSE_NAME] = courseName;
        found = true;
        break;
      }
    }
    
    // Add new mapping if not found
    if (!found) {
      data.push([classCode, courseId, courseName]);
    }
    
    // Write back
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    if (typeof resetSequencerRuntimeCaches_ === 'function') {
      resetSequencerRuntimeCaches_();
    }
    
    return {
      success: true,
      message: `Mapping saved: ${classCode} → ${courseId}`
    };
    
  } catch (e) {
    Logger.log('apiSaveClassMapping error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Auto-populate class mappings from Google Classroom courses
 * Attempts to match class codes with course names
 * @returns {Object} Result with matched courses
 */
function apiAutoMapClasses() {
  try {
    // Get all Classroom courses
    const courses = Classroom.Courses.list({
      pageSize: 100,
      courseStates: ['ACTIVE']
    }).courses || [];

    // Get existing timetable class codes
    const timetableResult = apiGetTimetable();
    if (!timetableResult.success) {
      throw new Error('Could not load timetable');
    }

    const classCodes = Object.keys(timetableResult.timetable);
    const mappings = [];
    const unmatched = [];

    // Try to match each class code
    classCodes.forEach(classCode => {
      const matched = courses.find(course => {
        const name = course.name || '';
        // Try exact match or contains
        return name.includes(classCode) ||
               name.replace(/\s+/g, '').toUpperCase().includes(classCode.replace(/\s+/g, '').toUpperCase());
      });

      if (matched) {
        mappings.push({
          classCode: classCode,
          courseId: matched.id,
          courseName: matched.name
        });
      } else {
        unmatched.push(classCode);
      }
    });

    if (mappings.length) {
      const ss = getTimetableSheet();
      const sheet = ss.getSheetByName(CLASS_MAPPING_TAB);
      if (!sheet) throw new Error('ClassMapping tab not found');

      const data = sheet.getDataRange().getValues();
      const headers = data.length ? data[0] : ['classCode', 'courseId', 'courseName'];
      const byClass = {};

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const code = row[MAP_COL.CLASS_CODE];
        if (!code) continue;
        byClass[code] = [
          code,
          row[MAP_COL.COURSE_ID] || '',
          row[MAP_COL.COURSE_NAME] || ''
        ];
      }

      mappings.forEach(mapping => {
        byClass[mapping.classCode] = [mapping.classCode, mapping.courseId, mapping.courseName || ''];
      });

      const rows = [headers];
      Object.keys(byClass).sort().forEach(code => rows.push(byClass[code]));
      sheet.clearContents();
      sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

      if (typeof resetSequencerRuntimeCaches_ === 'function') {
        resetSequencerRuntimeCaches_();
      }
    }

    return {
      success: true,
      matched: mappings.length,
      unmatched: unmatched,
      mappings: mappings,
      message: 'Auto-mapped ' + mappings.length + ' classes. ' + unmatched.length + ' unmatched.'
    };

  } catch (e) {
    Logger.log('apiAutoMapClasses error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================================================
// DATE/TIME CALCULATION UTILITIES
// ============================================================================

/**
 * Calculate next class occurrence for a specific class
 * @param {string} classCode - Class code
 * @param {Date} afterDate - Calculate after this date
 * @returns {Object} Next occurrence details
 */
function calculateNextClassOccurrence(classCode, afterDate) {
  const timetableResult = apiGetTimetable(classCode);
  
  if (!timetableResult.success || !timetableResult.timetable || timetableResult.timetable.length === 0) {
    throw new Error(`No timetable found for ${classCode}`);
  }
  
  const slots = (timetableResult.timetable || []).filter(isSchedulableClassSlot_);
  
  // Get current school day (1-10)
  const currentDay = calculateSchoolDayFromDate(afterDate);
  
  // Find next slot in cycle
  let nextSlot = null;
  
  // Try to find slot on a later day in current cycle
  for (let slot of slots) {
    if (slot.day > currentDay) {
      nextSlot = slot;
      break;
    }
  }
  
  // If no slot found, wrap to first slot of next cycle
  if (!nextSlot) {
    nextSlot = slots[0];
  }
  
  // Convert school day to calendar date
  const calendarDate = convertSchoolDayToCalendarDate(nextSlot.day, afterDate);
  
  // Calculate assign time (5 min before)
  const assignTime = calculateAssignTime(nextSlot.startTime);
  
  return {
    date: formatDateToISO(calendarDate),
    assignTime: assignTime,
    day: nextSlot.day,
    period: nextSlot.period,
    startTime: nextSlot.startTime,
    endTime: nextSlot.endTime,
    room: nextSlot.room
  };
}

/**
 * Calculate school day number (1-10) from calendar date
 * Uses existing term settings from SettingsManager
 * @param {Date} date - Calendar date
 * @returns {number} School day (1-10)
 */
function calculateSchoolDayFromDate(date) {
  // Use existing calculateSchoolWeekDayTerm function from main app
  // This handles term dates, A/B weeks, etc.
  const result = calculateSchoolWeekDayTerm(date);
  return result.day; // Returns 1-10
}

/**
 * Convert school day number to calendar date
 * @param {number} targetDay - School day (1-10)
 * @param {Date} referenceDate - Reference date to calculate from
 * @returns {Date} Calendar date
 */
function convertSchoolDayToCalendarDate(targetDay, referenceDate) {
  const currentDay = calculateSchoolDayFromDate(referenceDate);
  
  let daysToAdd = targetDay - currentDay;
  
  // If target is in past of current cycle, add full cycle (10 days)
  if (daysToAdd <= 0) {
    daysToAdd += 10;
  }
  
  // Add school days (skip weekends)
  let date = new Date(referenceDate);
  let schoolDaysAdded = 0;
  
  while (schoolDaysAdded < daysToAdd) {
    date.setDate(date.getDate() + 1);
    const dayOfWeek = date.getDay();
    
    // Only count weekdays (Monday-Friday)
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      schoolDaysAdded++;
    }
  }
  
  return date;
}

/**
 * Calculate assign time (5 minutes before class start).
 * Accepts either a "HH:MM" string or a Date/time value from Sheets.
 * @param {string|Date} classStartTime - e.g. "08:55" or a Date object
 * @returns {string} Assign time "HH:MM"
 */
function calculateAssignTime(classStartTime) {
  // Normalise to "HH:MM" string
  var timeStr;

  if (classStartTime instanceof Date) {
    var h = classStartTime.getHours();
    var m = classStartTime.getMinutes();
    timeStr = String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  } else {
    timeStr = String(classStartTime || '');
  }

  var parts = timeStr.split(':');
  if (parts.length < 2) {
    // Fallback to a sensible default if time is malformed
    var now = new Date();
    var fh = now.getHours();
    var fm = now.getMinutes();
    parts = [String(fh), String(fm)];
  }

  var hours = Number(parts[0]);
  var minutes = Number(parts[1]);

  if (!isFinite(hours)) hours = 9;
  if (!isFinite(minutes)) minutes = 0;

  // Subtract 5 minutes
  minutes -= 5;
  if (minutes < 0) {
    minutes += 60;
    hours -= 1;
    if (hours < 0) {
      hours = 23; // Wrap to previous day
    }
  }

  return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
}


/**
 * Get all class occurrences in a date range
 * @param {string} classCode - Class code
 * @param {Date} startDate - Start of range
 * @param {Date} endDate - End of range
 * @returns {Array} Array of occurrence objects
 */
function getClassOccurrencesInRange(classCode, startDate, endDate) {
  const occurrences = [];
  let currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    try {
      const next = calculateNextClassOccurrence(classCode, new Date(currentDate.getTime() - 86400000)); // Day before
      const nextDate = toLocalDateOnlyTimetable_(next.date);
      
      if (nextDate > endDate) break;
      if (nextDate >= currentDate) {
        occurrences.push(next);
        currentDate = new Date(nextDate);
      }
      
      currentDate.setDate(currentDate.getDate() + 1);
    } catch (e) {
      break;
    }
  }
  
  return occurrences;
}

// ============================================================================
// TIMETABLE RUNDOWN (Settings UI)
// ============================================================================
function apiGetTimetableRundown(classCode) {
  if (!classCode) throw new Error('Missing classCode');

  var settings = getServerSettings_();
  var termDates = (settings && settings.termDates) ? settings.termDates : [];
  if (!termDates.length) throw new Error('No term settings found');

  var today = new Date();
  var term = findCurrentTerm_(termDates, today);
  if (!term) throw new Error('No current term found for today');

  var startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var endDate = toLocalDateOnlyTimetable_(term.end);
  if (isNaN(endDate.getTime())) throw new Error('Invalid term end date');

  var occurrences = getClassOccurrencesInRange(classCode, startDate, endDate);

  var courseId = getCourseIdFromClassCode(classCode);
  var cwItems = listCourseWorkAll_(courseId);
  var cwByDate = indexCourseWorkByDate_(cwItems);

  var tz = Session.getScriptTimeZone() || 'Australia/Sydney';
  var output = occurrences.map(function(occ) {
    var dateIso = occ.date;
    var weekDay = calculateSchoolWeekDayTerm(toLocalDateOnlyTimetable_(dateIso));
    var assignTime = occ.assignTime || '';
    var title = pickBestTitleForSlot_(dateIso, assignTime, cwByDate);
    return {
      date: dateIso,
      time: assignTime,
      day: occ.day,
      period: occ.period,
      week: weekDay && weekDay.week ? weekDay.week : null,
      term: term.term,
      title: title || ''
    };
  });

  var delivered = output.filter(function(o){ return o.title; }).length;
  var remaining = output.length - delivered;

  return {
    classCode: classCode,
    term: term.term,
    termStart: term.start,
    termEnd: term.end,
    totalSlots: output.length,
    deliveredCount: delivered,
    remainingCount: remaining,
    occurrences: output
  };
}

function toLocalDateOnlyTimetable_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  var raw = String(value || '').trim();
  if (!raw) return new Date();

  var iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }

  var parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return new Date();
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function findCurrentTerm_(termDates, today) {
  var d = toLocalDateOnlyTimetable_(today);
  for (var i = 0; i < termDates.length; i++) {
    var t = termDates[i];
    var start = toLocalDateOnlyTimetable_(t.start);
    var end = toLocalDateOnlyTimetable_(t.end);
    if (d >= start && d <= end) return t;
  }
  return null;
}

function listCourseWorkAll_(courseId) {
  var all = [];
  var pageToken = null;
  do {
    var res = Classroom.Courses.CourseWork.list(courseId, {
      pageSize: 100,
      pageToken: pageToken || undefined,
      courseWorkStates: ['PUBLISHED', 'DRAFT']
    });
    (res.courseWork || []).forEach(function(w){
      all.push({
        id: w.id,
        title: w.title || '',
        creationTime: w.creationTime || '',
        scheduledTime: w.scheduledTime || ''
      });
    });
    pageToken = res.nextPageToken || null;
  } while (pageToken);
  return all;
}

function indexCourseWorkByDate_(items) {
  var tz = Session.getScriptTimeZone() || 'Australia/Sydney';
  var byDate = {};
  items.forEach(function(it){
    var ts = it.scheduledTime || it.creationTime || '';
    if (!ts) return;
    var d = new Date(ts);
    if (isNaN(d.getTime())) return;
    var dateKey = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var timeKey = Utilities.formatDate(d, tz, 'HH:mm');
    if (!byDate[dateKey]) byDate[dateKey] = [];
    byDate[dateKey].push({ title: it.title, time: timeKey });
  });
  return byDate;
}

function pickBestTitleForSlot_(dateIso, assignTime, byDate) {
  var items = byDate[dateIso] || [];
  if (!items.length) return '';
  if (!assignTime) return items[0].title || '';
  var slotMin = timeToMinutes_(assignTime);
  var best = null;
  var bestDiff = 99999;
  items.forEach(function(it){
    var diff = Math.abs(timeToMinutes_(it.time) - slotMin);
    if (diff < bestDiff) { bestDiff = diff; best = it; }
  });
  return best ? best.title : '';
}

function timeToMinutes_(t) {
  var parts = String(t || '').split(':');
  var hh = Number(parts[0] || 0);
  var mm = Number(parts[1] || 0);
  return hh * 60 + mm;
}

// ============================================================================
// DASHBOARD SUMMARY API
// ============================================================================

function apiGetDashboardTimetableSummary() {
  try {
    var tz = Session.getScriptTimeZone() || 'Australia/Sydney';
    var now = new Date();
    var cycle = calculateSchoolWeekDayTerm(now);
    var cycleDay = Number((cycle && cycle.day) || 1);
    var currentWeek = cycleDay <= 5 ? 'A' : 'B';

    var timetableResult = apiGetTimetable();
    if (!timetableResult || !timetableResult.success) {
      return {
        success: false,
        currentWeek: currentWeek,
        today: Utilities.formatDate(now, tz, 'EEE, d MMM yyyy'),
        cycleDay: cycleDay,
        nextAvailable: { available: false },
        error: timetableResult && timetableResult.error ? timetableResult.error : 'Could not load timetable'
      };
    }

    var timetable = timetableResult.timetable || {};
    var slotIndex = {};

    Object.keys(timetable).forEach(function(classCode) {
      (timetable[classCode] || []).forEach(function(slot) {
        var slotType = normalizeSlotType_(slot && slot.slotType);
        var day = Number(slot && slot.day);
        if (!day || day < 1 || day > 10) return;

        var startTime = formatTimeForUi_(slot && slot.startTime);
        var endTime = formatTimeForUi_(slot && slot.endTime);
        var period = Number(slot && slot.period) || 0;
        var periodLabel = String((slot && slot.periodLabel) || '').trim();

        var key = [day, period, startTime, endTime, slotType, periodLabel].join('|');
        if (!slotIndex[key]) {
          slotIndex[key] = {
            day: day,
            period: period,
            startTime: startTime,
            endTime: endTime,
            slotType: slotType,
            periodLabel: periodLabel
          };
        }
      });
    });

    var freeSlots = Object.keys(slotIndex).map(function(k) { return slotIndex[k]; })
      .filter(function(slot) { return slot.slotType === TT_SLOT_TYPES.FREE; });

    if (!freeSlots.length) {
      return {
        success: true,
        currentWeek: currentWeek,
        today: Utilities.formatDate(now, tz, 'EEE, d MMM yyyy'),
        cycleDay: cycleDay,
        term: cycle && cycle.term ? cycle.term : null,
        nextAvailable: {
          available: false
        }
      };
    }

    freeSlots.sort(function(a, b) {
      var dayDelta = distanceInCycleDays_(cycleDay, a.day) - distanceInCycleDays_(cycleDay, b.day);
      if (dayDelta !== 0) return dayDelta;

      if (Number(a.period) !== Number(b.period)) return Number(a.period) - Number(b.period);

      return String(a.startTime || '').localeCompare(String(b.startTime || ''));
    });

    var next = freeSlots[0];
    var weekType = next.day <= 5 ? 'A' : 'B';

    return {
      success: true,
      currentWeek: currentWeek,
      today: Utilities.formatDate(now, tz, 'EEE, d MMM yyyy'),
      cycleDay: cycleDay,
      term: cycle && cycle.term ? cycle.term : null,
      nextAvailable: {
        available: true,
        day: cycleDayLabel_(next.day),
        cycleDay: next.day,
        weekType: weekType,
        slot: next.periodLabel || (next.period ? 'Period ' + next.period : ''),
        time: formatTimeRange12Hour_(next.startTime, next.endTime)
      }
    };

  } catch (e) {
    Logger.log('apiGetDashboardTimetableSummary error: ' + e.message);
    return {
      success: false,
      currentWeek: 'A',
      today: '',
      cycleDay: 1,
      nextAvailable: { available: false },
      error: e.message
    };
  }
}

function distanceInCycleDays_(fromDay, toDay) {
  var delta = Number(toDay) - Number(fromDay);
  if (delta < 0) delta += 10;
  return delta;
}

function cycleDayLabel_(dayNum) {
  var names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  var idx = (Number(dayNum) - 1) % 5;
  if (idx < 0) idx = 0;
  return names[idx] + ' (Day ' + Number(dayNum) + ')';
}

function formatTimeRange12Hour_(startTime, endTime) {
  var start = to12HourTime_(startTime);
  var end = to12HourTime_(endTime);
  if (!start && !end) return '';
  if (!end) return start;
  if (!start) return end;
  return start + ' - ' + end;
}

function to12HourTime_(timeValue) {
  var t = String(timeValue || '').trim();
  if (!t) return '';

  var parts = t.split(':');
  if (parts.length < 2) return t;

  var h = Number(parts[0]);
  var m = Number(parts[1]);
  if (!isFinite(h) || !isFinite(m)) return t;

  var ampm = h >= 12 ? 'pm' : 'am';
  var h12 = h % 12;
  if (h12 === 0) h12 = 12;

  return h12 + ':' + String(m).padStart(2, '0') + ' ' + ampm;
}

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

/**
 * Validate timetable slot data
 * @param {Object} slot - Slot object to validate
 * @throws {Error} If validation fails
 */
function validateTimetableSlot(slot) {
  var dayNum = Number(slot.day);
  if (!dayNum || dayNum < 1 || dayNum > 10) {
    throw new Error('Day must be between 1 and 10');
  }

  slot.slotType = normalizeSlotType_(slot.slotType);

  var periodNum = Number(slot.period);
  if (slot.slotType === TT_SLOT_TYPES.CLASS) {
    if (!periodNum || periodNum < 1 || periodNum > 10) {
      throw new Error('Class period must be between 1 and 10');
    }
  } else {
    if (!isFinite(periodNum)) periodNum = 0;
    if (periodNum < 0 || periodNum > 10) {
      throw new Error('Non-class period must be between 0 and 10');
    }
  }

  var hasStart = String(slot.startTime || '').trim() !== '';
  var hasEnd = String(slot.endTime || '').trim() !== '';

  if (slot.slotType === TT_SLOT_TYPES.CLASS) {
    if (!validateTimeFormat(String(slot.startTime || ''))) {
      throw new Error('Invalid start time format: ' + slot.startTime + '. Expected HH:MM');
    }
    if (!validateTimeFormat(String(slot.endTime || ''))) {
      throw new Error('Invalid end time format: ' + slot.endTime + '. Expected HH:MM');
    }
    const start = timeToMinutes(String(slot.startTime));
    const end = timeToMinutes(String(slot.endTime));
    if (end <= start) {
      throw new Error('End time must be after start time');
    }
  } else if (hasStart || hasEnd) {
    if (!validateTimeFormat(String(slot.startTime || '')) || !validateTimeFormat(String(slot.endTime || ''))) {
      throw new Error('Non-class slots with times must use HH:MM for start/end');
    }
    const start = timeToMinutes(String(slot.startTime));
    const end = timeToMinutes(String(slot.endTime));
    if (end <= start) {
      throw new Error('End time must be after start time');
    }
  }

  slot.day = dayNum;
  slot.period = periodNum;
}

/**
 * Validate time format (HH:MM)
 * @param {string} time - Time string
 * @returns {boolean} True if valid
 */
function validateTimeFormat(time) {
  if (!time || typeof time !== 'string') return false;
  
  const parts = time.split(':');
  if (parts.length !== 2) return false;
  
  const hours = parseInt(parts[0]);
  const minutes = parseInt(parts[1]);
  
  if (isNaN(hours) || isNaN(minutes)) return false;
  if (hours < 0 || hours > 23) return false;
  if (minutes < 0 || minutes > 59) return false;
  
  return true;
}

/**
 * Convert time string to minutes since midnight
 * @param {string} time - Format "HH:MM"
 * @returns {number} Minutes
 */
function timeToMinutes(time) {
  const parts = time.split(':');
  return parseInt(parts[0]) * 60 + parseInt(parts[1]);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get or create timetable spreadsheet
 * @returns {Spreadsheet} Timetable spreadsheet
 */
function getTimetableSheet() {
  // Prefer explicit sheet ID (reliable)
  if (TIMETABLE_SHEET_ID) {
    return SpreadsheetApp.openById(TIMETABLE_SHEET_ID);
  }
  // Try to find existing by name
  const files = DriveApp.getFilesByName(TIMETABLE_SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  
  // Create new
  const ss = SpreadsheetApp.create(TIMETABLE_SHEET_NAME);
  initializeTimetableSheets(ss);
  return ss;
}

/**
 * Initialize timetable sheet structure
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function initializeTimetableSheets(ss) {
  // Create Timetable tab
  const ttSheet = ss.getSheetByName('Sheet1') || ss.insertSheet(TIMETABLE_TAB);
  ttSheet.setName(TIMETABLE_TAB);
  
  const ttHeaders = TT_HEADERS.slice();
  ttSheet.getRange(1, 1, 1, ttHeaders.length).setValues([ttHeaders]);
  ttSheet.setFrozenRows(1);
  ttSheet.getRange(1, 1, 1, ttHeaders.length).setFontWeight('bold');
  
  // Create ClassMapping tab
  const mapSheet = ss.insertSheet(CLASS_MAPPING_TAB);
  const mapHeaders = ['classCode', 'courseId', 'courseName'];
  mapSheet.getRange(1, 1, 1, mapHeaders.length).setValues([mapHeaders]);
  mapSheet.setFrozenRows(1);
  mapSheet.getRange(1, 1, 1, mapHeaders.length).setFontWeight('bold');
  
  // Create Sequences tab (for SequenceManager)
  const seqSheet = ss.insertSheet('Sequences');
  const seqHeaders = [
    'classCode', 'lessonId', 'position', 'title', 'assignDate', 'assignTime',
    'dueDate', 'dueTime', 'status', 'classworkId', 'description', 'materials', 'points'
  ];
  seqSheet.getRange(1, 1, 1, seqHeaders.length).setValues([seqHeaders]);
  seqSheet.setFrozenRows(1);
  seqSheet.getRange(1, 1, 1, seqHeaders.length).setFontWeight('bold');
}

function formatTimeForUi_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(value || '').trim();
}

/**
 * Format date to ISO string (YYYY-MM-DD)
 * @param {Date} date - Date object
 * @returns {string} ISO formatted date
 */
function formatDateToISO(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Get list of all class codes in timetable
 * @returns {Array} Array of class codes
 */
function getAllClassCodes() {
  const result = apiGetTimetable();
  if (!result.success) return [];
  return Object.keys(result.timetable);
}

// ============================================================================
// TESTING FUNCTIONS (Remove in production)
// ============================================================================

/**
 * Test function - create sample timetable
 */
function testCreateSampleTimetable() {
  const slots7TECHA = [
    { day: 1, period: 1, startTime: '08:55', endTime: '09:55', room: 'T3' },
    { day: 1, period: 4, startTime: '12:20', endTime: '13:20', room: 'T3' },
    { day: 3, period: 1, startTime: '08:55', endTime: '09:55', room: 'T3' },
    { day: 5, period: 2, startTime: '09:55', endTime: '10:55', room: 'T3' }
  ];
  
  const result = apiSaveTimetable('7TECHA', slots7TECHA);
  Logger.log(result);
}

/**
 * Test function - calculate next occurrence
 */
function testNextOccurrence() {
  const today = new Date();
  const next = calculateNextClassOccurrence('7TECHA', today);
  Logger.log('Next occurrence: ' + JSON.stringify(next, null, 2));
}

/**
 * Test function - get occurrences in range
 */
function testOccurrencesInRange() {
  const start = new Date('2026-02-07');
  const end = new Date('2026-02-28');
  const occurrences = getClassOccurrencesInRange('7TECHA', start, end);
  Logger.log(`Found ${occurrences.length} occurrences`);
  occurrences.forEach(occ => {
    Logger.log(`  ${occ.date} @ ${occ.assignTime} (Day ${occ.day}, Period ${occ.period})`);
  });
}
