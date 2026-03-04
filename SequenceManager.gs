/**
 * SequenceManager.gs
 * Manages lesson sequences for timetable-aware scheduling
 * Handles CRUD operations, shuffle logic, and Google Sheets persistence
 */

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================

const SEQ_SHEET_NAME = 'SequencerData';
const SEQ_DATA_SHEET_ID = '1w8OSfgjmlmAYgjDnmq_1QxOBv6UpFNN9ow9XBUakC8U';
const TAB_SEQUENCES = 'Sequences';
const TAB_TIMETABLE = 'Timetable';
const TAB_CLASS_MAPPING = 'ClassMapping';

// Column indices for Sequences tab (0-based)
const SEQ_COL = {
  CLASS_CODE: 0,
  LESSON_ID: 1,
  POSITION: 2,
  TITLE: 3,
  ASSIGN_DATE: 4,
  ASSIGN_TIME: 5,
  DUE_DATE: 6,
  DUE_TIME: 7,
  STATUS: 8,
  CLASSWORK_ID: 9,
  DESCRIPTION: 10,
  MATERIALS: 11,
  POINTS: 12,
  TOPIC: 13
};

const SEQ_HEADERS = [
  'classCode', 'lessonId', 'position', 'title', 'assignDate', 'assignTime',
  'dueDate', 'dueTime', 'status', 'classworkId', 'description', 'materials', 'points', 'topic'
];

// Per-execution runtime cache for hot lookups.
var __seqRuntimeCache = {
  timetableByClass: null,
  classToCourse: null,
  courseToClass: null,
  settings: null,
  topicIndexByCourse: {}
};

function resetSequencerRuntimeCaches_() {
  __seqRuntimeCache.timetableByClass = null;
  __seqRuntimeCache.classToCourse = null;
  __seqRuntimeCache.courseToClass = null;
}

// ============================================================================
// MAIN API FUNCTIONS (Called from HTML)
// ============================================================================

/**
 * Get complete sequence for a class
 * @param {string} classCode - Class code (e.g., "7TECHA")
 * @returns {Object} { classCode, lessons: [...] }
 */
function apiGetSequence(classCode) {
  try {
    const sequence = loadSequence(classCode);
    return {
      success: true,
      classCode: classCode,
      lessons: sequence
    };
  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Create or update a lesson in sequence
 * Handles automatic shuffling if date conflict exists
 * @param {string} classCode - Class code
 * @param {Object} lessonData - Lesson details
 * @returns {Object} Result with success/error
 */
function apiSaveLesson(classCode, lessonData) {
  try {
    Logger.log('apiSaveLesson START classCode=' + classCode +
      ' title=' + (lessonData && lessonData.title || '') +
      ' assignDate=' + (lessonData && lessonData.assignDate || '') +
      ' useTimetableTime=' + (lessonData && lessonData.useTimetableTime) +
      ' skipCustomDateBump=' + (lessonData && lessonData.skipCustomDateBump));

    var skipCustomDateBump = !!(lessonData && lessonData.skipCustomDateBump);
    var preventEquivalentMatch = !!(lessonData && lessonData.preventEquivalentMatch);
    var preserveAssignDate = !!(lessonData && lessonData.preserveAssignDate);
    // Sheet-first: never pull Classroom state before applying bump/save logic.
    // Sequence sheet is source-of-truth for scheduling and numbering.
    var processLog = [];
    processLog.push('1) Validate lesson payload');
    // Validate required fields
    if (!lessonData.title || !lessonData.assignDate) {
      throw new Error('Missing required fields: title, assignDate');
    }

    // If requested (or missing), derive assignTime from timetable for that calendar date
    // (e.g. custom date in sequence mode).
    if (lessonData.useTimetableTime || !lessonData.assignTime) {
      var requestedDate = String(lessonData.assignDate || '').trim();
      processLog.push('2) Resolve timetable time for selected date');
      const derived = deriveAssignDateTimeFromTimetable_(classCode, lessonData.assignDate);
      lessonData.assignDate = preserveAssignDate && requestedDate ? requestedDate : derived.date;
      lessonData.assignTime = derived.assignTime;
    }

    if (!lessonData.assignTime) {
      throw new Error('Missing required field: assignTime');
    }
    // Normalize times to HH:MM for consistent slot keys
    lessonData.assignTime = normaliseTimeToHHMM_(lessonData.assignTime);
    if (lessonData.dueTime) lessonData.dueTime = normaliseTimeToHHMM_(lessonData.dueTime);

    var sequence = loadSequence(classCode);

    // Check if this is an update (existing lesson). Prefer id match, then
    // classworkId match so edited Classroom items move instead of duplicating.
    var existingIndex = -1;
    if (lessonData.id) {
      existingIndex = sequence.findIndex(function (l) { return l.id === lessonData.id; });
    }
    if (existingIndex < 0 && lessonData.classworkId) {
      existingIndex = sequence.findIndex(function (l) {
        return String(l && l.classworkId || '') === String(lessonData.classworkId || '');
      });
      if (existingIndex >= 0 && !lessonData.id) {
        lessonData.id = sequence[existingIndex].id;
      }
    }

    // Reuse/custom-date saves can arrive as "new" (no edit context) even when they
    // represent moving an already-scheduled lesson. Detect and convert to UPDATE to
    // avoid duplicates.
    if (!preventEquivalentMatch && existingIndex < 0 && !lessonData.id && !lessonData.classworkId && lessonData.useTimetableTime) {
      var moveMatchIdx = findSingleEquivalentScheduledLessonIndex_(sequence, lessonData);
      if (moveMatchIdx >= 0) {
        existingIndex = moveMatchIdx;
        lessonData.id = sequence[moveMatchIdx].id;
        if (sequence[moveMatchIdx].classworkId) {
          lessonData.classworkId = sequence[moveMatchIdx].classworkId;
        }
        processLog.push('3) Matched existing lesson and switched to move mode');
      }
    }

    // Generate lesson ID if truly new
    if (!lessonData.id) {
      lessonData.id = Utilities.getUuid();
    }
    if (!lessonData.status) lessonData.status = 'scheduled';

    var newSlotKey = slotKey_(lessonData.assignDate, lessonData.assignTime);
    var moved = [];
    var slotChanged = false;
    var occupied = new Set();
    sequence.forEach(function (l) {
      if (!isSchedulableForBump_(l)) return;
      occupied.add(slotKey_(l.assignDate, l.assignTime));
    });

    if (existingIndex >= 0) {
      // UPDATE existing lesson
      processLog.push('3) Update existing lesson');
      var oldSlotKey = slotKey_(sequence[existingIndex].assignDate, sequence[existingIndex].assignTime);
      slotChanged = oldSlotKey !== newSlotKey;

      // Update the lesson
      sequence[existingIndex] = lessonData;

      // If slot changed, check for conflicts and shuffle
      if (oldSlotKey !== newSlotKey) {
        var conflict = sequence.find(function (l) {
          return slotKey_(l.assignDate, l.assignTime) === newSlotKey &&
            l.id !== lessonData.id &&
            isSchedulableForBump_(l);
        });
        var sameDateConflictUpdate = lessonData.useTimetableTime && sequence.find(function (l) {
          return l.id !== lessonData.id && isSchedulableForBump_(l) && String(l.assignDate || '') === String(lessonData.assignDate || '');
        });
        if (lessonData.useTimetableTime && !skipCustomDateBump) {
          var beforeCountUpdate = countSchedulableOnDate_(sequence, lessonData.assignDate, lessonData.id);
          processLog.push('4) Bump existing lessons first');
          // Custom-date semantics: clear the chosen date before assignment.
          moved = forceClearDateBeforeInsert_(sequence, lessonData.assignDate, classCode, lessonData.id) || [];
          processLog.push('5) Bumped ' + moved.length + ' lesson(s)');
          if (beforeCountUpdate > 0 && moved.length <= 0) {
            throw new Error('Bump verification failed: expected lessons to move from ' + lessonData.assignDate);
          }
          verifyBumpClearedDate_(sequence, lessonData.assignDate, lessonData.id);
          processLog.push('6) Verified bump complete for ' + lessonData.assignDate);
        } else if (lessonData.useTimetableTime && skipCustomDateBump) {
          processLog.push('4) Bump already completed and confirmed in previous step');
        } else if (conflict || sameDateConflictUpdate) {
          processLog.push('4) Bump existing lessons first');
          // Exact slot semantics for explicit time edits.
          moved = shuffleDownFrom(sequence, lessonData.assignDate, lessonData.assignTime, classCode, lessonData.id) || [];
          processLog.push('5) Bumped ' + moved.length + ' lesson(s)');
        }
      }
    } else {
      // NEW lesson - check for conflicts
      processLog.push('3) Prepare new lesson insertion');
      var conflict = sequence.find(function (l) {
        return slotKey_(l.assignDate, l.assignTime) === newSlotKey && isSchedulableForBump_(l);
      });
      var sameDateConflict = lessonData.useTimetableTime && sequence.find(function (l) {
        return isSchedulableForBump_(l) && l.assignDate === lessonData.assignDate;
      });
      if (lessonData.useTimetableTime && !skipCustomDateBump) {
        var beforeCountNew = countSchedulableOnDate_(sequence, lessonData.assignDate, '');
        processLog.push('4) Bump existing lessons first');
        // For custom-date insertion, clear the chosen date before assignment.
        moved = forceClearDateBeforeInsert_(sequence, lessonData.assignDate, classCode) || [];
        processLog.push('5) Bumped ' + moved.length + ' lesson(s)');
        if (beforeCountNew > 0 && moved.length <= 0) {
          throw new Error('Bump verification failed: expected lessons to move from ' + lessonData.assignDate);
        }
        verifyBumpClearedDate_(sequence, lessonData.assignDate, '');
        processLog.push('6) Verified bump complete for ' + lessonData.assignDate);
      } else if (lessonData.useTimetableTime && skipCustomDateBump) {
        processLog.push('4) Bump already completed and confirmed in previous step');
      } else if (conflict || occupied.has(newSlotKey) || sameDateConflict) {
        processLog.push('4) Bump existing lessons first');
        // For exact time insertion, push from slot.
        moved = shuffleDownFrom(sequence, lessonData.assignDate, lessonData.assignTime, classCode) || [];
        processLog.push('5) Bumped ' + moved.length + ' lesson(s)');
      }
      // Add new lesson
      processLog.push('7) Assign new lesson to selected date');
      sequence.push(lessonData);
    }

    // Check term boundary warnings
    var warnings = checkTermBoundaries(sequence);

    // Re-sort chronologically, then renumber ALL lessons sequentially and
    // rebuild every lesson's topic from its current delivery date.
    // This is the single correct place to do this — covers both the inserted
    // lesson and every bumped lesson in one pass.
    processLog.push('8) Resequence: renumber all lessons and rebuild all topics from delivery dates');
    sequence.sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + a.assignTime);
      var dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });

    resequenceAll_(sequence, classCode, lessonData.courseId || null);

    sequence.forEach(function (lesson, idx) {
      lesson.position = idx + 1;
    });

    // SHEET-FIRST: Save to sheet before any Classroom calls.
    // This guarantees the lesson is persisted even if Classroom sync fails.
    processLog.push('9) Save sequence to Google Sheet (sheet-first)');
    Logger.log('apiSaveLesson: about to saveSequence for ' + classCode +
      ' with ' + sequence.length + ' lessons. New lesson id=' + lessonData.id);
    saveSequence(classCode, sequence);
    Logger.log('apiSaveLesson: saveSequence completed for ' + classCode);

    // Reload authoritative state from sheet so Classroom receives exact sheet values.
    var sheetSeq = loadSequence(classCode);
    var savedLesson = sheetSeq.find(function (l) { return l.id === lessonData.id; }) || lessonData;

    // Sync affected lessons to Classroom using sheet state.
    if (lessonData && lessonData.publishToClassroom) {
      processLog.push('10) Sync affected lessons to Classroom from sheet state');
      var affected = {};
      var classroomErrors = [];
      affected[lessonData.id] = true;
      moved.forEach(function (l) { affected[l.id] = true; });

      sheetSeq.forEach(function (l) {
        if (!affected[l.id]) return;
        if (!isSchedulableForBump_(l)) return;
        try {
          var cwId = createOrUpdateClassworkForLesson_(classCode, l, lessonData.courseId || null);
          if (cwId && cwId !== l.classworkId) {
            // Write new classworkId back to sheet immediately
            var seq2 = loadSequence(classCode);
            var idx2 = seq2.findIndex(function (x) { return x.id === l.id; });
            if (idx2 >= 0) { seq2[idx2].classworkId = cwId; saveSequence(classCode, seq2); }
          }
        } catch (e) {
          var errMsg = e && e.message ? e.message : String(e);
          Logger.log('Classroom sync warning for ' + l.id + ': ' + errMsg);
          // Never throw here — sheet is already saved, Classroom is best-effort.
          var isPermDenied = errMsg.toLowerCase().indexOf('@projectpermissiondenied') !== -1 ||
                             errMsg.toLowerCase().indexOf('project is not permitted') !== -1;
          if (isPermDenied) {
            warnings.push({
              lessonId: l.id,
              lessonTitle: l.title || '',
              date: l.assignDate || '',
              message: 'Saved to sheet. Classroom sync blocked by project permissions — will sync via 1-min trigger.'
            });
          } else {
            classroomErrors.push({ lessonId: l.id, message: errMsg });
          }
        }
      });

      if (classroomErrors.length) {
        classroomErrors.slice(0, 3).forEach(function (x) {
          warnings.push({ lessonId: x.lessonId, lessonTitle: '', date: '',
            message: 'Classroom sync warning: ' + x.message });
        });
      }
    }

    return { success: true, lesson: savedLesson, warnings: warnings, processLog: processLog };
  } catch (e) {
    Logger.log('apiSaveLesson error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function splitLessonPrefixServer_(title) {
  var raw = String(title || '').trim();
  var m = raw.match(/^Lesson\s*(\d+)\s*:?\s*(.*)$/i);
  if (!m) return { hasPrefix: false, number: 0, base: raw };
  return {
    hasPrefix: true,
    number: Number(m[1] || 0),
    base: String(m[2] || '').trim()
  };
}

function lessonBaseTitle_(title) {
  var split = splitLessonPrefixServer_(title);
  return String(split.base || title || '').trim();
}

function lessonTextKey_(lessonLike) {
  var base = lessonBaseTitle_(lessonLike && lessonLike.title || '').toLowerCase();
  var desc = String(lessonLike && lessonLike.description || '').replace(/\s+/g, ' ').trim().toLowerCase();
  var pts = Number(lessonLike && lessonLike.points || 0);
  return base + '||' + desc + '||' + pts;
}

function findSingleEquivalentScheduledLessonIndex_(sequence, lessonData) {
  if (!Array.isArray(sequence) || !lessonData) return -1;
  var key = lessonTextKey_(lessonData);
  var matches = [];
  for (var i = 0; i < sequence.length; i++) {
    var l = sequence[i];
    if (!isSchedulableForBump_(l)) continue;
    if (lessonTextKey_(l) === key) matches.push(i);
  }
  return matches.length === 1 ? matches[0] : -1;
}

function countSchedulableOnDate_(sequence, dateIso, excludeId) {
  var d = String(dateIso || '').trim();
  if (!d || !Array.isArray(sequence)) return 0;
  var n = 0;
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (excludeId && String(l.id || '') === String(excludeId || '')) return;
    if (String(l.assignDate || '') !== d) return;
    n++;
  });
  return n;
}

function parseLessonNumberFromTitleServer_(title) {
  return splitLessonPrefixServer_(title).number || 0;
}

function lessonSortTimeMs_(lesson) {
  var d = String(lesson && lesson.assignDate || '').trim();
  var t = normaliseTimeToHHMM_(String(lesson && lesson.assignTime || '').trim() || '00:00');
  if (!d) return 0;
  var dt = new Date(d + 'T' + t + ':00');
  return isNaN(dt.getTime()) ? 0 : dt.getTime();
}

function shiftLessonNumbersFromIndex_(sequence, startIndex, startNumber) {
  if (!Array.isArray(sequence) || startIndex < 0 || !(startNumber >= 1)) return;
  var next = startNumber;
  for (var i = startIndex; i < sequence.length; i++) {
    var lesson = sequence[i];
    if (!isSchedulableForBump_(lesson)) continue;
    var split = splitLessonPrefixServer_(lesson.title);
    var base = split.base || String(lesson.title || '').trim();
    lesson.title = base ? ('Lesson ' + next + ': ' + base) : ('Lesson ' + next);
    next++;
  }
}

/**
 * Master resequence: after any bump/insert/delete, call this to:
 *   1. Renumber every schedulable lesson sequentially (Lesson 1, 2, 3…)
 *      based purely on chronological order.
 *   2. Rebuild every lesson's topicText from its current assignDate.
 *
 * This is the single source of truth for numbering and topics.
 * It replaces the partial rebalanceLessonNumbersOnSave_ approach for
 * bump operations where multiple lessons change date simultaneously.
 *
 * @param {Array}  sequence  - Sorted (chronological) array of lesson objects (modified in place)
 * @param {string} classCode - Class code for topic text generation
 * @param {string} courseId  - Classroom course ID for topic text generation (may be null)
 */
function resequenceAll_(sequence, classCode, courseId) {
  if (!Array.isArray(sequence) || !sequence.length) return;
  var lessonNum = 1;
  sequence.forEach(function (lesson) {
    if (!isSchedulableForBump_(lesson)) return;
    // Renumber: preserve the base title (strip existing "Lesson N:" prefix)
    var split = splitLessonPrefixServer_(lesson.title);
    var base = split.base || String(lesson.title || '').trim();
    lesson.title = base ? ('Lesson ' + lessonNum + ': ' + base) : ('Lesson ' + lessonNum);
    // Rebuild topic from current delivery date
    if (lesson.assignDate) {
      lesson.topicText = buildTopicText_(lesson.assignDate, courseId || null, classCode);
    }
    lessonNum++;
  });
}

function maxLessonNumberBeforeMs_(sequence, targetMs) {
  if (!Array.isArray(sequence) || !(targetMs >= 0)) return 0;
  var max = 0;
  sequence.forEach(function (lesson) {
    if (!isSchedulableForBump_(lesson)) return;
    var ms = lessonSortTimeMs_(lesson);
    if (ms < targetMs) {
      var num = parseLessonNumberFromTitleServer_(lesson.title);
      if (num && num > max) max = num;
    }
  });
  return max;
}

function rebalanceLessonNumbersOnSave_(sequence, lessonData, moved) {
  if (!Array.isArray(sequence) || !sequence.length) return;
  var targetMs = lessonData ? lessonSortTimeMs_(lessonData) : 0;
  if (!targetMs) return;
  var startIndex = findFirstScheduledIndexAfterMs_(sequence, targetMs);
  if (startIndex < 0) return;
  var insertedId = String(lessonData && lessonData.id || '');
  var prevNumber = maxLessonNumberBeforeMs_(sequence, targetMs);
  var startNumber = 1;
  var foundAtOrAfter = false;
  if (prevNumber > 0) {
    startNumber = prevNumber + 1;
  } else {
    // If no earlier numbered lesson exists, preserve the number already present
    // in the sequence at/after this point (excluding the inserted lesson itself).
    for (var i = startIndex; i < sequence.length; i++) {
      var candidate = sequence[i];
      if (!isSchedulableForBump_(candidate)) continue;
      if (insertedId && String(candidate.id || '') === insertedId) continue;
      var num = parseLessonNumberFromTitleServer_(candidate.title);
      if (num && num > 0) { startNumber = num; foundAtOrAfter = true; break; }
    }
    if (!foundAtOrAfter) {
      // No existing numbered lesson found at/after insertion point.
      // Prefer inserted lesson number if valid, otherwise derive from prev.
      var insertedNum = parseLessonNumberFromTitleServer_(lessonData && lessonData.title || '');
      if (insertedNum > 0) {
        startNumber = insertedNum;
      } else {
        startNumber = prevNumber > 0 ? prevNumber + 1 : 1;
      }
    }
  }
  shiftLessonNumbersFromIndex_(sequence, startIndex, startNumber);
}

function findFirstScheduledIndexAfterMs_(sequence, targetMs) {
  if (!Array.isArray(sequence) || !(targetMs >= 0)) return -1;
  for (var i = 0; i < sequence.length; i++) {
    var lesson = sequence[i];
    if (!isSchedulableForBump_(lesson)) continue;
    if (lessonSortTimeMs_(lesson) >= targetMs) return i;
  }
  return -1;
}
/**
 * Return the lesson number that should be used when inserting at a target slot.
 * If a lesson already exists at/after the slot, insertion takes that number.
 * Otherwise it uses previous numbered lesson + 1.
 */
function apiGetInsertionLessonNumber(classCode, assignDate, assignTime) {
  try {
    var code = String(classCode || '').trim();
    var date = String(assignDate || '').trim();
    var time = normaliseTimeToHHMM_(String(assignTime || '').trim() || '00:00');
    if (!code || !date) throw new Error('Missing classCode/assignDate');

    var targetMs = lessonSortTimeMs_({ assignDate: date, assignTime: time });
    if (!targetMs) throw new Error('Invalid target date/time');

    var seq = loadSequence(code) || [];
    var numbered = seq
      .filter(function (l) { return l && isSchedulableForBump_(l) && splitLessonPrefixServer_(l.title).hasPrefix; })
      .sort(function (a, b) { return lessonSortTimeMs_(a) - lessonSortTimeMs_(b); });

    if (!numbered.length) return { success: true, lessonNumber: 1 };

    var prevNum = 0;
    for (var i = 0; i < numbered.length; i++) {
      var ms = lessonSortTimeMs_(numbered[i]);
      var num = parseLessonNumberFromTitleServer_(numbered[i].title);
      if (!num) continue;
      if (ms >= targetMs) {
        return { success: true, lessonNumber: num };
      }
      prevNum = num;
    }

    return { success: true, lessonNumber: (prevNum > 0 ? prevNum + 1 : 1) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Phase 1 for custom-date insert flow: bump all schedulable lessons off target date,
 * persist, and return what moved. Assignment happens later after user confirmation.
 */
function apiBumpDateOnly(classCode, assignDate, courseId) {
  try {
    var code = String(classCode || '').trim();
    var date = String(assignDate || '').trim();
    if (!code) throw new Error('Missing classCode');
    if (!date) throw new Error('Missing assignDate');

    var processLog = [];
    processLog.push('1) Load sequence');
    var sequence = loadSequence(code) || [];
    var targetSlot = deriveAssignDateTimeFromTimetable_(code, date);
    var targetDate = String(targetSlot && targetSlot.date || date || '').trim();
    var targetTime = normaliseTimeToHHMM_(String(targetSlot && targetSlot.assignTime || '').trim() || '00:00');
    processLog.push('2) Target insertion slot: ' + targetDate + ' ' + targetTime);

    var beforeCount = countSchedulableAtOrAfterSlot_(sequence, targetDate, targetTime);
    processLog.push('3) Lessons at/after insertion slot: ' + beforeCount);

    // If no lessons exist at/after the chosen date, the slot is already free —
    // no bump needed. Return success with movedCount=0 so the UI proceeds to insert.
    if (beforeCount <= 0) {
      saveSequence(code, sequence);
      return {
        success: true,
        movedCount: 0,
        movedLessons: [],
        processLog: processLog.concat(['3b) Slot is free — no bump required'])
      };
    }

    processLog.push('4) Cascading bump from insertion slot');
    var moved = bumpCascadeFromSlot_(sequence, targetDate, targetTime, code) || [];
    processLog.push('5) Bumped ' + moved.length + ' lesson(s)');

    // If no lessons moved despite beforeCount > 0, they may already be on later dates.
    // Treat as no-op rather than hard failure.
    if (moved.length <= 0) {
      saveSequence(code, sequence);
      return {
        success: true,
        movedCount: 0,
        movedLessons: [],
        processLog: processLog.concat(['5b) No lessons needed moving — slot effectively free'])
      };
    }

    sequence.sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + (a.assignTime || '00:00'));
      var dateB = new Date(b.assignDate + 'T' + (b.assignTime || '00:00'));
      return dateA - dateB;
    });

    // Renumber all lessons sequentially and rebuild topics from new delivery dates.
    // Note: the new lesson has not been inserted yet (this is phase 1 of 2).
    // We renumber now so bumped lessons get correct numbers immediately;
    // apiSaveLesson (phase 2) will resequence again after inserting the new lesson.
    var cid = String(courseId || '').trim() || String(getCourseIdFromClassCode(code) || '').trim();
    resequenceAll_(sequence, code, cid || null);
    sequence.forEach(function (lesson, idx) { lesson.position = idx + 1; });
    processLog.push('6) Resequenced: renumbered all lessons and rebuilt topics from delivery dates');

    saveSequence(code, sequence);
    processLog.push('7) Saved bumped sequence to Google Sheet (sheet is source of truth)');

    // Reload authoritative state from the sheet so Classroom receives exact sheet values.
    var sheetSeq = loadSequence(code);
    var movedIdSet = {};
    moved.forEach(function (l) { if (l.id) movedIdSet[l.id] = true; });

    // Sync moved scheduled items in Classroom to keep dates/topics aligned.
    if (moved.length) {
      sheetSeq.forEach(function (l) {
        if (!l || !movedIdSet[l.id]) return;
        if (!isSchedulableForBump_(l)) return;
        try {
          createOrUpdateClassworkForLesson_(code, l, cid || null);
        } catch (e) {
          Logger.log('apiBumpDateOnly sync warning for lesson ' + (l.id || '') + ': ' + (e && e.message ? e.message : String(e)));
        }
      });
    }

    return {
      success: true,
      movedCount: moved.length,
      movedLessons: moved.map(function (l) {
        return {
          id: l.id || '',
          title: l.title || '',
          assignDate: l.assignDate || '',
          assignTime: l.assignTime || ''
        };
      }),
      processLog: processLog
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * Delete a lesson from sequence
 * Handles shuffle up for scheduled lessons
 * Shows warning for published lessons
 * @param {string} classCode - Class code
 * @param {string} lessonId - Lesson ID to delete
 * @param {boolean} confirmed - True if user confirmed deletion of published lesson
 * @returns {Object} Result
 */
function apiDeleteLesson(classCode, lessonId, confirmed = false) {
  try {
    const sequence = loadSequence(classCode);
    const lessonIndex = sequence.findIndex(l => l.id === lessonId);
    
    if (lessonIndex < 0) {
      throw new Error('Lesson not found');
    }
    
    const lesson = sequence[lessonIndex];
    
    // If published and not confirmed, return warning
    if (lesson.status === 'published' && !confirmed) {
      return {
        success: false,
        needsConfirmation: true,
        lesson: lesson,
        message: 'This lesson is already published. Student submissions will be lost if deleted.'
      };
    }
    
    // Delete linked Classroom post when present.
    // Published lessons require confirmation above; scheduled drafts are deleted directly.
    if (lesson.classworkId) {
      try {
        const courseId = getCourseIdFromClassCode(classCode);
        Classroom.Courses.CourseWork.remove(courseId, lesson.classworkId);
      } catch (e) {
        Logger.log('Failed to delete from Classroom: ' + e.message);
        // Continue anyway - remove from sequence
      }
    }
    
    // Get the deleted date for shuffle calculations
    const deletedDate = lesson.assignDate;
    
    // Remove from sequence
    sequence.splice(lessonIndex, 1);
    
    // Shuffle up all scheduled lessons after deleted date
    const moved = shuffleUpFrom(sequence, deletedDate, classCode);
    
    // Re-sort, renumber all lessons sequentially, rebuild all topics from delivery dates.
    sequence.sort((a, b) => {
      const dateA = new Date(a.assignDate + 'T' + a.assignTime);
      const dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });

    var courseId = '';
    try { courseId = getCourseIdFromClassCode(classCode); } catch (e) {}
    resequenceAll_(sequence, classCode, courseId || null);

    sequence.forEach((l, idx) => { l.position = idx + 1; });

    // Sheet-first: save before pushing to Classroom.
    saveSequence(classCode, sequence);

    // Reload authoritative state from the sheet so Classroom gets sheet-exact values.
    var sheetSeqAfterDelete = loadSequence(classCode);
    var movedIdSetDelete = {};
    if (moved && moved.length) {
      moved.forEach(function (l) { if (l && l.id) movedIdSetDelete[l.id] = true; });
    }

    // Push updated lessons to Classroom using sheet state.
    if (moved && moved.length) {
      sheetSeqAfterDelete.forEach(function (l) {
        if (!movedIdSetDelete[l.id]) return;
        if (!l.classworkId || l.status !== 'scheduled') return;
        try {
          createOrUpdateClassworkForLesson_(classCode, l, courseId || null);
        } catch (e) {
          Logger.log('Classroom update failed for shuffled lesson ' + l.id + ': ' + e.message);
        }
      });
    }

    return {
      success: true,
      message: `Lesson deleted. ${sheetSeqAfterDelete.filter(l => l.status === 'scheduled').length} lessons rescheduled.`
    };
    
  } catch (e) {
    Logger.log('apiDeleteLesson error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Update lesson status (called by auto-publish trigger)
 * @param {string} classCode - Class code
 * @param {string} lessonId - Lesson ID
 * @param {string} status - New status ('published', 'scheduled', etc.)
 * @param {string} classworkId - Classroom courseWork ID
 * @returns {boolean} Success
 */
function apiUpdateLessonStatus(classCode, lessonId, status, classworkId = null) {
  try {
    const sequence = loadSequence(classCode);
    const lesson = sequence.find(l => l.id === lessonId);
    
    if (!lesson) {
      throw new Error('Lesson not found');
    }
    
    lesson.status = status;
    if (classworkId) {
      lesson.classworkId = classworkId;
    }
    
    saveSequence(classCode, sequence);
    return true;
    
  } catch (e) {
    Logger.log('apiUpdateLessonStatus error: ' + e.message);
    return false;
  }
}

/**
 * Reschedule/update a sequence lesson by matching Classroom classwork ID.
 * Used by edit flows where the UI starts from a Classroom item but the source
 * of truth for scheduling is the sequence.
 * @param {string} courseId
 * @param {string} classworkId
 * @param {Object} patch
 * @returns {Object}
 */
function apiUpdateSequenceLessonByClasswork(courseId, classworkId, patch) {
  try {
    var cid = String(courseId || '').trim();
    var cwid = String(classworkId || '').trim();
    if (!cid) throw new Error('Missing courseId');
    if (!cwid) throw new Error('Missing classworkId');

    var classCode = getClassCodeFromCourseId_(cid);
    if (!classCode) throw new Error('No class mapping found for courseId ' + cid);

    var seq = loadSequence(classCode) || [];
    var incoming = patch && typeof patch === 'object' ? patch : {};
    var idx = seq.findIndex(function (l) {
      return String(l && l.classworkId || '') === cwid;
    });
    if (idx < 0 && incoming && incoming.id) {
      var wantedId = String(incoming.id || '').trim();
      if (wantedId) {
        idx = seq.findIndex(function (l) { return String(l && l.id || '') === wantedId; });
      }
    }

    if (idx < 0) {
      var oldDate = String(incoming.existingAssignDate || '').trim();
      var oldTime = String(incoming.existingAssignTime || '').trim();
      var normOldTime = oldTime ? normaliseTimeToHHMM_(oldTime) : '';
      var wantedTitle = String(incoming.title || '').trim().toLowerCase();
      if (wantedTitle && oldDate) {
        idx = seq.findIndex(function (l) {
          var lt = String(l && l.title || '').trim().toLowerCase();
          var ld = String(l && l.assignDate || '').trim();
          var lm = String(l && l.assignTime || '').trim();
          if (lt !== wantedTitle) return false;
          if (ld !== oldDate) return false;
          if (!normOldTime) return true;
          return normaliseTimeToHHMM_(lm) === normOldTime;
        });
      }
    }

    if (idx < 0) {
      return { success: false, error: 'No matching sequence lesson for classworkId ' + cwid };
    }

    var current = seq[idx] || {};
    var merged = Object.assign({}, current);

    if (Object.prototype.hasOwnProperty.call(incoming, 'id') && incoming.id) merged.id = String(incoming.id || '').trim();
    if (Object.prototype.hasOwnProperty.call(incoming, 'title')) merged.title = String(incoming.title || merged.title || '').trim();
    if (Object.prototype.hasOwnProperty.call(incoming, 'description')) merged.description = String(incoming.description || '');
    if (Object.prototype.hasOwnProperty.call(incoming, 'assignDate')) merged.assignDate = String(incoming.assignDate || '').trim();
    if (Object.prototype.hasOwnProperty.call(incoming, 'assignTime')) merged.assignTime = normaliseTimeToHHMM_(String(incoming.assignTime || '').trim());
    if (Object.prototype.hasOwnProperty.call(incoming, 'dueDate')) merged.dueDate = String(incoming.dueDate || '').trim();
    if (Object.prototype.hasOwnProperty.call(incoming, 'dueTime')) merged.dueTime = String(incoming.dueTime || '').trim() ? normaliseTimeToHHMM_(String(incoming.dueTime || '').trim()) : '';
    if (Object.prototype.hasOwnProperty.call(incoming, 'points')) merged.points = Number(incoming.points || 0);
    if (Object.prototype.hasOwnProperty.call(incoming, 'materials')) merged.materials = Array.isArray(incoming.materials) ? incoming.materials : [];
    if (Object.prototype.hasOwnProperty.call(incoming, 'topicText')) merged.topicText = String(incoming.topicText || '').trim();
    if (Object.prototype.hasOwnProperty.call(incoming, 'useTimetableTime')) merged.useTimetableTime = !!incoming.useTimetableTime;

    merged.courseId = cid;
    merged.classworkId = cwid;
    merged.status = 'scheduled';
    merged.publishToClassroom = true;

    if (!merged.title) throw new Error('Missing lesson title');
    if (!merged.assignDate) throw new Error('Missing assignDate');

    return apiSaveLesson(classCode, merged);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Calculate next available datetime for a class.
 * INSTRUMENTED FOR DEBUGGING.
 * @param {string} classCode - Class code (e.g., "12TAS3").
 * @param {string} afterDate - ISO date string "YYYY-MM-DD"; optional.
 * @returns {{success:boolean, date?:string, time?:string, day?:number,
 *           period?:number, classStartTime?:any, error?:string}}
 */
function apiGetNextAvailable(classCode, afterDate, courseId) {
  Logger.log('apiGetNextAvailable v3 START: classCode=' + classCode + ', afterDate=' + afterDate);
  
  try {
    Logger.log('Step 1: Validating classCode');
    if (!classCode) {
      throw new Error('Missing classCode.');
    }

    Logger.log('Step 2: Normalising reference date');
    var referenceDate;
    if (afterDate) {
      if (typeof afterDate === 'number') {
        referenceDate = new Date(afterDate);
      } else if (typeof afterDate === 'string' && /^\d+$/.test(afterDate)) {
        referenceDate = new Date(Number(afterDate));
      } else {
        referenceDate = new Date(afterDate);
      }
      if (isNaN(referenceDate.getTime())) {
        throw new Error('Invalid afterDate: ' + afterDate);
      }
    } else {
      referenceDate = new Date();
    }
    Logger.log('Step 2 complete: referenceDate=' + referenceDate.toISOString());

    Logger.log('Step 3: Finding next FREE occurrence (timetable + existing sequence)...');
    var nextSlot = calculateNextFreeOccurrence_(classCode, referenceDate);
    Logger.log('Step 3 complete: nextSlot=' + JSON.stringify(nextSlot));

    Logger.log('Step 4: Validating nextSlot');
    if (!nextSlot) {
      throw new Error('No next occurrence found for ' + classCode);
    }

    Logger.log('Step 5: Extracting fields');
    var date = nextSlot.date || null;
    var assignTime = nextSlot.assignTime || null;
    var day = typeof nextSlot.day === 'number' ? nextSlot.day : null;
    var period = typeof nextSlot.period === 'number' ? nextSlot.period : null;

    Logger.log('Step 6: Checking completeness');
    if (!date || !assignTime || day === null || period === null) {
      throw new Error('Incomplete nextSlot data for ' + classCode +
                      ': date=' + date + ', assignTime=' + assignTime +
                      ', day=' + day + ', period=' + period);
    }

        Logger.log('Step 7: Building success response');
    
    // Ensure classStartTime is a string (not a Date object) so it can serialize
    var classStartTimeStr = null;
    if (nextSlot.classStartTime instanceof Date) {
      var h = nextSlot.classStartTime.getHours();
      var m = nextSlot.classStartTime.getMinutes();
      classStartTimeStr = String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
    } else if (typeof nextSlot.classStartTime === 'string') {
      classStartTimeStr = nextSlot.classStartTime;
    }
    
    var weekTerm = calculateSchoolWeekDayTerm(new Date(date + 'T00:00:00'));
    var topicText = buildTopicText_(date, courseId || null, classCode);

    var result = {
      success: true,
      date: date,
      time: assignTime,
      day: day,
      period: period,
      classStartTime: classStartTimeStr,
      week: weekTerm.week,
      term: weekTerm.term,
      topicText: topicText
    };
    Logger.log('apiGetNextAvailable v3 SUCCESS: ' + JSON.stringify(result));
    return result;


  } catch (e) {
    Logger.log('apiGetNextAvailable v3 CATCH: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Get timetable slot (assign time) for a specific calendar date.
 * Used for custom sequence date UI.
 * @param {string} classCode
 * @param {string} dateIso - YYYY-MM-DD
 */
function apiGetTimetableSlotForDate(classCode, dateIso, courseId) {
  try {
    if (!classCode) throw new Error('Missing classCode.');
    if (!dateIso) throw new Error('Missing date.');
    var derived = deriveAssignDateTimeFromTimetable_(classCode, dateIso);
    var weekTerm = calculateSchoolWeekDayTerm(new Date(derived.date + 'T00:00:00'));
    var topicText = buildTopicText_(derived.date, courseId || null, classCode);
    return {
      success: true,
      date: derived.date,
      time: derived.assignTime,
      day: derived.day,
      period: derived.period,
      classStartTime: derived.classStartTime,
      week: weekTerm.week,
      term: weekTerm.term,
      topicText: topicText
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}



// ============================================================================
// CORE SEQUENCE OPERATIONS
// ============================================================================

/**
 * Load sequence from Google Sheets
 * @param {string} classCode - Class code
 * @returns {Array} Array of lesson objects
 */
function loadSequence(classCode) {
  const ss = getSequencerSheet();
  const sheet = ss.getSheetByName(TAB_SEQUENCES);
  
  if (!sheet) {
    throw new Error('Sequences tab not found. Please create it first.');
  }
  
  ensureSequencesHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  
  const lessons = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip if not this class or empty row
    if (row[SEQ_COL.CLASS_CODE] !== classCode || !row[SEQ_COL.LESSON_ID]) {
      continue;
    }
    
    const lesson = {
      id: row[SEQ_COL.LESSON_ID],
      position: row[SEQ_COL.POSITION] || 0,
      title: row[SEQ_COL.TITLE] || '',
      assignDate: formatDateToISO(row[SEQ_COL.ASSIGN_DATE]),
      assignTime: normaliseTimeToHHMM_(row[SEQ_COL.ASSIGN_TIME]) || '',
      dueDate: formatDateToISO(row[SEQ_COL.DUE_DATE]),
      dueTime: normaliseTimeToHHMM_(row[SEQ_COL.DUE_TIME]) || '',
      status: row[SEQ_COL.STATUS] || 'scheduled',
      classworkId: row[SEQ_COL.CLASSWORK_ID] || null,
      description: row[SEQ_COL.DESCRIPTION] || '',
      materials: row[SEQ_COL.MATERIALS] ? JSON.parse(row[SEQ_COL.MATERIALS]) : [],
      points: row[SEQ_COL.POINTS] || 0,
      topicText: row[SEQ_COL.TOPIC] || ''
    };
    
    lessons.push(lesson);
  }
  
  // Sort by position
  lessons.sort((a, b) => a.position - b.position);
  
  return lessons;
}

/**
 * Save sequence to Google Sheets
 * Overwrites all lessons for this class code
 * @param {string} classCode - Class code
 * @param {Array} lessons - Array of lesson objects
 */
function saveSequence(classCode, lessons) {
  const ss = getSequencerSheet();
  let sheet = ss.getSheetByName(TAB_SEQUENCES);
  
  if (!sheet) {
    // Create sheet if doesn't exist
    sheet = ss.insertSheet(TAB_SEQUENCES);
    sheet.getRange(1, 1, 1, SEQ_HEADERS.length).setValues([SEQ_HEADERS]);
    sheet.setFrozenRows(1);
  }
  ensureSequencesHeaders_(sheet);
  
  // Load all data
  const data = sheet.getDataRange().getValues();
  
  // Remove existing lessons for this class
  const filteredData = data
    .map(function (row) {
      var copy = Array.isArray(row) ? row.slice(0, SEQ_HEADERS.length) : [];
      while (copy.length < SEQ_HEADERS.length) copy.push('');
      return copy;
    })
    .filter((row, idx) => {
    if (idx === 0) return true; // Keep headers
    return row[SEQ_COL.CLASS_CODE] !== classCode;
  });
  
  // Add updated lessons
  lessons.forEach(lesson => {
    filteredData.push([
      classCode,
      lesson.id,
      lesson.position,
      lesson.title,
      lesson.assignDate,
      lesson.assignTime,
      lesson.dueDate || '',
      lesson.dueTime || '',
      lesson.status,
      lesson.classworkId || '',
      lesson.description || '',
      JSON.stringify(lesson.materials || []),
      lesson.points || 0,
      lesson.topicText || ''
    ]);
  });
  
  // Clear and write
  sheet.clearContents();
  sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  resetSequencerRuntimeCaches_();
}

function ensureSequencesHeaders_(sheet) {
  if (!sheet) return;
  var width = SEQ_HEADERS.length;
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, width).setValues([SEQ_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  var current = sheet.getRange(1, 1, 1, Math.max(width, 1)).getValues()[0] || [];
  var needsRewrite = false;
  for (var i = 0; i < width; i++) {
    if (String(current[i] || '') !== SEQ_HEADERS[i]) { needsRewrite = true; break; }
  }
  if (needsRewrite) {
    sheet.getRange(1, 1, 1, width).setValues([SEQ_HEADERS]);
    sheet.setFrozenRows(1);
  }
}

// ============================================================================
// SHUFFLE ALGORITHMS
// ============================================================================

/**
 * Shuffle lessons DOWN (push forward) starting from a specific slot
 * Used when inserting a lesson at an occupied slot (date+time)
 * @param {Array} sequence - Lesson array (modified in place)
 * @param {string} fromDate - Start shuffling from this date (ISO format)
 * @param {string} fromTime - Start shuffling from this time ("HH:MM")
 * @param {string} classCode - Class code for timetable lookups
 * @param {string} excludeId - Don't shuffle this lesson (optional)
 */
function shuffleDownFrom(sequence, fromDate, fromTime, classCode, excludeId) {
  if (sequence.length === 0) return;
  if (!classCode) {
    Logger.log('shuffleDownFrom: classCode is missing, cannot shuffle');
    return;
  }
  
  var fromKey = slotKey_(fromDate, fromTime);

  // Find all scheduled lessons at or after from slot (exclude published and excludeId)
  var toShift = sequence
    .filter(function (l) {
      if (!isSchedulableForBump_(l)) return false;
      if (l.id === excludeId) return false;
      return slotKey_(l.assignDate, l.assignTime) >= fromKey;
    })
    .sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + a.assignTime);
      var dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });
  
  if (toShift.length === 0) return;

  // Build occupied set for slots that must remain (all scheduled lessons NOT being shifted)
  var shiftingIds = {};
  toShift.forEach(function (l) { shiftingIds[l.id] = true; });
  var occupied = new Set();
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (shiftingIds[l.id]) return; // will move
    if (l.id === excludeId) return; // inserted/updated lesson
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });

  // Cursor starts at the occupied slot we're inserting into; first shifted lesson goes AFTER it
  var cursor = new Date(fromDate + 'T' + (fromTime || '00:00'));

  var moved = [];
  toShift.forEach(function (lesson) {
    var next = findNextFreeOccurrenceAfter_(classCode, cursor, occupied);
    lesson.assignDate = next.date;
    lesson.assignTime = next.assignTime;
    occupied.add(slotKey_(lesson.assignDate, lesson.assignTime));
    cursor = new Date(next.date + 'T' + next.assignTime);
    moved.push(lesson);
  });
  return moved;
}

function shuffleDownFromDate(sequence, fromDate, classCode, excludeId) {
  if (sequence.length === 0) return;
  if (!classCode) {
    Logger.log('shuffleDownFromDate: classCode is missing, cannot shuffle');
    return;
  }

  var toShift = sequence
    .filter(function (l) {
      if (!isSchedulableForBump_(l)) return false;
      if (l.id === excludeId) return false;
      return String(l.assignDate || '') >= String(fromDate || '');
    })
    .sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + a.assignTime);
      var dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });

  if (toShift.length === 0) return;

  var shiftingIds = {};
  toShift.forEach(function (l) { shiftingIds[l.id] = true; });
  var occupied = new Set();
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (shiftingIds[l.id]) return;
    if (l.id === excludeId) return;
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });

  // Date-based push-first: first shifted lesson goes after the insertion day.
  var cursor = new Date(String(fromDate || '') + 'T23:59:59');
  var moved = [];
  toShift.forEach(function (lesson) {
    var next = findNextFreeOccurrenceAfter_(classCode, cursor, occupied);
    lesson.assignDate = next.date;
    lesson.assignTime = next.assignTime;
    occupied.add(slotKey_(lesson.assignDate, lesson.assignTime));
    cursor = new Date(next.date + 'T' + next.assignTime);
    moved.push(lesson);
  });
  return moved;
}

function isSchedulableForBump_(lesson) {
  var s = String(lesson && lesson.status || '').toLowerCase();
  // Some legacy rows are marked published even when still sequence-managed.
  return s === 'scheduled' || s === 'draft' || s === 'published';
}

/**
 * Force-clear all schedulable lessons on the target calendar date before insertion.
 * This guarantees "bump first, assign last" behavior for custom date mode.
 */
function forceClearDateBeforeInsert_(sequence, targetDate, classCode, excludeId) {
  if (!Array.isArray(sequence) || !sequence.length) return [];
  if (!classCode) throw new Error('Missing classCode for date bump');
  var target = String(targetDate || '').trim();
  if (!target) throw new Error('Missing target date for bump');

  var toShift = sequence
    .filter(function (l) {
      if (!isSchedulableForBump_(l)) return false;
      if (excludeId && String(l.id || '') === String(excludeId || '')) return false;
      return String(l.assignDate || '') === target;
    })
    .sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + a.assignTime);
      var dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });

  if (!toShift.length) return [];

  var shiftingIds = {};
  toShift.forEach(function (l) { shiftingIds[String(l.id || '')] = true; });
  var occupied = new Set();
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (shiftingIds[String(l.id || '')]) return;
    if (excludeId && String(l.id || '') === String(excludeId || '')) return;
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });

  // Start after end-of-day so bumped lessons move to following available date.
  var cursor = new Date(target + 'T23:59:59');
  var moved = [];
  toShift.forEach(function (lesson) {
    var next = findNextFreeOccurrenceAfter_(classCode, cursor, occupied);
    lesson.assignDate = next.date;
    lesson.assignTime = next.assignTime;
    occupied.add(slotKey_(lesson.assignDate, lesson.assignTime));
    cursor = new Date(next.date + 'T' + next.assignTime);
    moved.push(lesson);
  });
  return moved;
}

function verifyBumpClearedDate_(sequence, targetDate, excludeId) {
  var target = String(targetDate || '').trim();
  if (!target) return;
  var remaining = (sequence || []).filter(function (l) {
    if (!isSchedulableForBump_(l)) return false;
    if (excludeId && String(l.id || '') === String(excludeId || '')) return false;
    return String(l.assignDate || '') === target;
  });
  if (remaining.length) {
    throw new Error('Bump verification failed: ' + remaining.length + ' lesson(s) still on ' + target);
  }
}

function countSchedulableAtOrAfterSlot_(sequence, dateIso, timeStr, excludeId) {
  if (!Array.isArray(sequence)) return 0;
  var d = String(dateIso || '').trim();
  var t = normaliseTimeToHHMM_(String(timeStr || '').trim() || '00:00');
  var targetMs = lessonSortTimeMs_({ assignDate: d, assignTime: t });
  var n = 0;
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (excludeId && String(l.id || '') === String(excludeId || '')) return;
    if (lessonSortTimeMs_(l) >= targetMs) n++;
  });
  return n;
}

function bumpCascadeFromSlot_(sequence, fromDate, fromTime, classCode, excludeId) {
  if (!Array.isArray(sequence) || !sequence.length) return [];
  if (!classCode) throw new Error('Missing classCode for bump cascade');

  var slotDate = String(fromDate || '').trim();
  var slotTime = normaliseTimeToHHMM_(String(fromTime || '').trim() || '00:00');
  var targetMs = lessonSortTimeMs_({ assignDate: slotDate, assignTime: slotTime });
  if (!targetMs) return [];

  // Only shift lessons that are AT the target slot or later.
  var toShift = sequence
    .filter(function (l) {
      if (!isSchedulableForBump_(l)) return false;
      if (excludeId && String(l.id || '') === String(excludeId || '')) return false;
      return lessonSortTimeMs_(l) >= targetMs;
    })
    .sort(function (a, b) {
      return lessonSortTimeMs_(a) - lessonSortTimeMs_(b);
    });

  if (!toShift.length) return [];

  var shiftingIds = {};
  toShift.forEach(function (l) { shiftingIds[String(l.id || '')] = true; });

  // Occupied set: slots held by lessons NOT being shifted (anchored lessons).
  var occupied = new Set();
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    var id = String(l.id || '');
    if (shiftingIds[id]) return;
    if (excludeId && id === String(excludeId || '')) return;
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });

  // Cascade: each lesson moves to the next free slot AFTER the previous
  // lesson's new position. Cursor starts just before the target slot so
  // the first bumped lesson lands on the first free slot at/after target.
  // We use end-of-day on the day BEFORE fromDate so the cascade starts fresh
  // from fromDate itself (the new lesson will occupy fromDate).
  var startCursor = new Date(slotDate + 'T23:59:59');
  // Advance one full day so the first free slot found is the day AFTER fromDate.
  startCursor = new Date(startCursor.getTime() + 1000); // push past 23:59:59

  var cursor = startCursor;
  var moved = [];
  toShift.forEach(function (lesson) {
    var next = findNextFreeOccurrenceAfter_(classCode, cursor, occupied);
    lesson.assignDate = next.date;
    lesson.assignTime = next.assignTime;
    occupied.add(slotKey_(lesson.assignDate, lesson.assignTime));
    // Advance cursor to end-of-day of the slot just assigned so the next
    // lesson cascades to the following available date.
    cursor = new Date(next.date + 'T23:59:59');
    moved.push(lesson);
  });
  return moved;
}

function getCourseWorkIdsByTopicName_(courseId, topicName) {
  var cid = String(courseId || '').trim();
  var wanted = String(topicName || '').trim().toLowerCase();
  if (!cid || !wanted) return [];

  var topicRes = Classroom.Courses.Topics.list(cid, { pageSize: 100 });
  var topics = topicRes && topicRes.topic ? topicRes.topic : [];
  var topicId = '';
  for (var i = 0; i < topics.length; i++) {
    var name = String(topics[i].name || '').trim().toLowerCase();
    if (name === wanted) { topicId = String(topics[i].topicId || '').trim(); break; }
  }
  if (!topicId) return [];

  var ids = [];
  var pageToken = null;
  do {
    var res = Classroom.Courses.CourseWork.list(cid, {
      pageSize: 100,
      pageToken: pageToken || undefined,
      courseWorkStates: ['PUBLISHED', 'DRAFT']
    });
    var arr = res && res.courseWork ? res.courseWork : [];
    arr.forEach(function (cw) {
      if (String(cw.topicId || '') === topicId && cw.id) ids.push(String(cw.id));
    });
    pageToken = res && res.nextPageToken ? res.nextPageToken : null;
  } while (pageToken);

  return ids;
}

function countSchedulableByClassworkIds_(sequence, classworkIds) {
  if (!Array.isArray(sequence) || !Array.isArray(classworkIds) || !classworkIds.length) return 0;
  var lookup = {};
  classworkIds.forEach(function (id) { lookup[String(id || '').trim()] = true; });
  var n = 0;
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    var cwid = String(l && l.classworkId || '').trim();
    if (cwid && lookup[cwid]) n++;
  });
  return n;
}

function forceClearLessonsByClassworkIds_(sequence, classworkIds, classCode) {
  if (!Array.isArray(sequence) || !sequence.length) return [];
  if (!classCode) throw new Error('Missing classCode for bump');
  if (!Array.isArray(classworkIds) || !classworkIds.length) return [];

  var lookup = {};
  classworkIds.forEach(function (id) { lookup[String(id || '').trim()] = true; });
  var toShift = sequence
    .filter(function (l) {
      if (!isSchedulableForBump_(l)) return false;
      var cwid = String(l && l.classworkId || '').trim();
      return !!(cwid && lookup[cwid]);
    })
    .sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + a.assignTime);
      var dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });

  if (!toShift.length) return [];

  var shiftingIds = {};
  toShift.forEach(function (l) { shiftingIds[String(l.id || '')] = true; });
  var occupied = new Set();
  sequence.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (shiftingIds[String(l.id || '')]) return;
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });

  // Cursor starts at end-of-day of the first shifted lesson date so bumped
  // lessons move to following available timetable dates, not from "now".
  var firstLesson = toShift[0];
  var cursorBase = firstLesson && firstLesson.assignDate ? (firstLesson.assignDate + 'T23:59:59') : '';
  var cursor = cursorBase ? new Date(cursorBase) : new Date();
  if (isNaN(cursor.getTime())) cursor = new Date();
  cursor.setSeconds(59, 0);
  var moved = [];
  toShift.forEach(function (lesson) {
    var base = lesson.assignDate ? (lesson.assignDate + 'T' + (lesson.assignTime || '00:00')) : '';
    var startFrom = base ? new Date(base) : cursor;
    if (isNaN(startFrom.getTime())) startFrom = cursor;
    var next = findNextFreeOccurrenceAfter_(classCode, startFrom, occupied);
    lesson.assignDate = next.date;
    lesson.assignTime = next.assignTime;
    occupied.add(slotKey_(lesson.assignDate, lesson.assignTime));
    cursor = new Date(next.date + 'T' + next.assignTime);
    moved.push(lesson);
  });
  return moved;
}

// ============================================================================
// TIMETABLE-AWARE SLOT HELPERS (time-aware + occupancy-aware)
// ============================================================================

function slotKey_(dateStr, timeStr) {
  var d = String(dateStr || '');
  var t = String(timeStr || '');
  // Normalise HH:MM to 5 chars if possible
  if (t && t.length >= 4) {
    var parts = t.split(':');
    if (parts.length >= 2) {
      t = String(parts[0]).padStart(2, '0') + ':' + String(parts[1]).padStart(2, '0');
    }
  }
  return d + 'T' + t;
}

function normaliseTimeToHHMM_(t) {
  if (t instanceof Date) {
    var h = t.getHours();
    var m = t.getMinutes();
    var s = t.getSeconds();
    if (s >= 30) {
      m += 1;
      if (m >= 60) { m = 0; h = (h + 1) % 24; }
    }
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  var s = String(t || '').trim();
  if (!s) return '';
  var parts = s.split(':');
  if (parts.length < 2) return s;
  var hh = Number(parts[0]);
  var mm = Number(parts[1]);
  var ss = Number(parts[2] || 0);
  if (ss >= 30) {
    mm += 1;
    if (mm >= 60) { mm = 0; hh = (hh + 1) % 24; }
  }
  return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
}

function combineDateAndTime_(dateObj, hhmm) {
  var parts = String(hhmm || '').split(':');
  var h = Number(parts[0] || 0);
  var m = Number(parts[1] || 0);
  var d = new Date(dateObj);
  d.setHours(h, m, 0, 0);
  return d;
}

function getSlotsForSchoolDay_(classCode, schoolDayNum) {
  var tt = getTimetableForClass(classCode) || [];
  return tt
    .filter(function (s) {
      var t = String((s && s.slotType) || 'CLASS').toUpperCase();
      var schedulable = !t || t === 'CLASS';
      return Number(s.day) === Number(schoolDayNum) && schedulable && !!normaliseTimeToHHMM_(s.startTime);
    })
    .sort(function (a, b) {
      var pa = Number(a.period || 0);
      var pb = Number(b.period || 0);
      if (pa !== pb) return pa - pb;
      var ta = normaliseTimeToHHMM_(a.startTime);
      var tb = normaliseTimeToHHMM_(b.startTime);
      return ta.localeCompare(tb);
    });
}

/**
 * Return next timetable occurrence strictly AFTER the given datetime (time-aware).
 * Scans forward over upcoming weekdays until it finds the first class period after afterDateTime.
 */
function calculateNextOccurrence(classCode, afterDateTime) {
  var timetable = getTimetableForClass(classCode);
  if (!timetable || timetable.length === 0) {
    throw new Error('No timetable found for ' + classCode);
  }

  var start = new Date(afterDateTime);
  if (isNaN(start.getTime())) throw new Error('Invalid afterDateTime');

  // Look ahead up to ~12 school weeks (plenty for sparse timetables)
  var lookaheadDays = 120;
  var cursor = new Date(start);
  cursor.setHours(0, 0, 0, 0);
  var settings = getServerSettings_();
  var terms = (settings && settings.termDates) ? settings.termDates.slice() : [];
  terms.sort(function (a, b) { return String(a.start).localeCompare(String(b.start)); });

  if (!isDateInAnyTerm_(cursor, terms)) {
    var nextTerm = getNextTermAfter_(cursor, terms);
    if (nextTerm) {
      cursor = new Date(nextTerm.start + 'T00:00:00');
    }
  }

  for (var i = 0; i <= lookaheadDays; i++) {
    if (i > 0) cursor.setDate(cursor.getDate() + 1);

    var dow = cursor.getDay();
    if (dow === 0 || dow === 6) continue; // weekends
    if (!isDateInAnyTerm_(cursor, terms)) continue;

    var schoolDayNum = calculateSchoolDayNumber(cursor);
    var slotsToday = getSlotsForSchoolDay_(classCode, schoolDayNum);
    if (!slotsToday.length) continue;

    for (var j = 0; j < slotsToday.length; j++) {
      var slot = slotsToday[j];
      var assignHHMM = calculateAssignTime(slot.startTime);
      var assignDt = combineDateAndTime_(cursor, assignHHMM);
      if (assignDt > start) {
        return {
          date: formatDateToISO(cursor),
          assignTime: assignHHMM,
          day: Number(slot.day),
          period: Number(slot.period),
          classStartTime: slot.startTime
        };
      }
    }
  }

  throw new Error('No next occurrence found for ' + classCode);
}

function isDateInAnyTerm_(date, terms) {
  if (!terms || !terms.length) return true;
  for (var i = 0; i < terms.length; i++) {
    var t = terms[i];
    var start = new Date(t.start + 'T00:00:00');
    var end = new Date(t.end + 'T23:59:59');
    if (date >= start && date <= end) return true;
  }
  return false;
}

function getNextTermAfter_(date, terms) {
  if (!terms || !terms.length) return null;
  for (var i = 0; i < terms.length; i++) {
    var t = terms[i];
    var start = new Date(t.start + 'T00:00:00');
    if (date < start) return t;
  }
  return null;
}

function buildOccupiedSlotSet_(classCode) {
  var seq = loadSequence(classCode) || [];
  var occupied = new Set();
  seq.forEach(function (l) {
    if (!isSchedulableForBump_(l)) return;
    if (!l.assignDate || !l.assignTime) return;
    occupied.add(slotKey_(l.assignDate, l.assignTime));
  });
  return occupied;
}

function findNextFreeOccurrenceAfter_(classCode, afterDateTime, occupiedSet) {
  var cursor = new Date(afterDateTime);
  // Ensure "strictly after" by nudging 1 second forward
  cursor = new Date(cursor.getTime() + 1000);

  for (var guard = 0; guard < 500; guard++) {
    var occ = calculateNextOccurrence(classCode, cursor);
    var key = slotKey_(occ.date, occ.assignTime);
    if (!occupiedSet || !occupiedSet.has(key)) {
      return occ;
    }
    if (guard === 0) {
      Logger.log('findNextFreeOccurrenceAfter_: occupied hit for ' + key + ' (cursor=' + cursor.toISOString() + ')');
    }
    // Move cursor to this occurrence so we can look for the next
    cursor = new Date(occ.date + 'T' + occ.assignTime);
  }
  throw new Error('Could not find a free slot for ' + classCode);
}

function calculateNextFreeOccurrence_(classCode, referenceDateTime) {
  var occupied = buildOccupiedSlotSet_(classCode);
  return findNextFreeOccurrenceAfter_(classCode, referenceDateTime, occupied);
}

/**
 * For a chosen calendar date, derive the timetable start/assign time for that day (earliest period).
 * Throws if the class does not occur on that date.
 */
function deriveAssignDateTimeFromTimetable_(classCode, assignDateIso) {
  if (!classCode) throw new Error('Missing classCode (needed to derive timetable time).');
  var d = new Date(assignDateIso + 'T00:00:00');
  if (isNaN(d.getTime())) throw new Error('Invalid assignDate: ' + assignDateIso);

  var termInfo = calculateSchoolWeekDayTerm(d);
  if (!termInfo.inTerm) {
    throw new Error('Selected date is outside term dates: ' + assignDateIso);
  }
  var schoolDayNum = termInfo.day;
  var slots = getSlotsForSchoolDay_(classCode, schoolDayNum);
  if (!slots.length) {
    throw new Error('No class period for ' + classCode + ' on ' + assignDateIso + ' (Day ' + schoolDayNum + ')');
  }

  var first = slots[0];
  return {
    date: formatDateToISO(d),
    assignTime: calculateAssignTime(first.startTime),
    day: Number(first.day),
    period: Number(first.period),
    classStartTime: first.startTime
  };
}


/**
 * Shuffle lessons UP (pull backward) starting from a date
 * Used when deleting a lesson
 * @param {Array} sequence - Lesson array (modified in place)
 * @param {string} fromDate - Start shuffling from this date
 * @param {string} classCode - Class code for date calculations
 */
function shuffleUpFrom(sequence, fromDate, classCode) {
  // Find all scheduled lessons after fromDate
  const toShift = sequence
    .filter(l => 
      l.assignDate > fromDate && 
      l.status === 'scheduled'
    )
    .sort((a, b) => {
      const dateA = new Date(a.assignDate + 'T' + a.assignTime);
      const dateB = new Date(b.assignDate + 'T' + b.assignTime);
      return dateA - dateB;
    });
  
  if (toShift.length === 0) return [];
  
  // Start from the deleted date
  let currentDate = new Date(fromDate);
  const moved = [];
  
  // Pull each lesson to current slot, then advance
  toShift.forEach(lesson => {
    const slot = calculateNextOccurrence(classCode, new Date(currentDate.getTime() - 86400000)); // Get slot for current date
    lesson.assignDate = slot.date;
    lesson.assignTime = slot.assignTime;
    currentDate = toLocalDateOnly_(calculateNextOccurrence(classCode, toLocalDateOnly_(slot.date)).date);
    moved.push(lesson);
  });
  return moved;
}
/**
 * Server-side helper to load the same APP_SETTINGS used by the UI.
 * Reads from PropertiesService.getUserProperties() / 'APP_SETTINGS'.
 */
function getServerSettings_() {
  if (__seqRuntimeCache.settings) return __seqRuntimeCache.settings;

  try {
    var parsed = (typeof apiGetSettings === 'function') ? (apiGetSettings() || {}) : {};
    if (!parsed || typeof parsed !== 'object') parsed = {};

    var year = new Date().getFullYear();
    if (!Array.isArray(parsed.termDates) || !parsed.termDates.length) {
      parsed.termDates = buildDefaultTermDates_(year);
    }
    if (!parsed.subjectMappings || typeof parsed.subjectMappings !== 'object') {
      parsed.subjectMappings = {};
    }

    __seqRuntimeCache.settings = parsed;
    return __seqRuntimeCache.settings;
  } catch (e) {
    Logger.log('getServerSettings_ parse error: ' + e.message);
    var y = new Date().getFullYear();
    __seqRuntimeCache.settings = {
      termDates: buildDefaultTermDates_(y),
      subjectMappings: {}
    };
    return __seqRuntimeCache.settings;
  }
}

function lockTermDatesToYear_(termDates, year) {
  if (!termDates || !termDates.length) return buildDefaultTermDates_(year);
  var allSameYear = termDates.every(function (t) {
    return t.start && t.end && String(t.start).indexOf(year + '-') === 0 && String(t.end).indexOf(year + '-') === 0;
  });
  if (allSameYear) return termDates;
  return buildDefaultTermDates_(year);
}

function buildDefaultTermDates_(year) {
  // Placeholder NSW-style structure per current request:
  // Term 1 starts Jan 26, 11-week term. 2-week breaks.
  // Terms 2–4 are 10-week terms.
  var term1Start = new Date(year, 0, 26);
  var term1End = addSchoolDays_(term1Start, 11 * 5 - 1);
  var term2Start = nextMondayAfter_(addDays_(term1End, 14));
  var term2End = addSchoolDays_(term2Start, 10 * 5 - 1);
  var term3Start = nextMondayAfter_(addDays_(term2End, 14));
  var term3End = addSchoolDays_(term3Start, 10 * 5 - 1);
  var term4Start = nextMondayAfter_(addDays_(term3End, 14));
  var term4End = addSchoolDays_(term4Start, 10 * 5 - 1);

  return [
    { term: 1, start: formatDateToISO(term1Start), end: formatDateToISO(term1End), startCycle: 'B' },
    { term: 2, start: formatDateToISO(term2Start), end: formatDateToISO(term2End), startCycle: 'A' },
    { term: 3, start: formatDateToISO(term3Start), end: formatDateToISO(term3End), startCycle: 'A' },
    { term: 4, start: formatDateToISO(term4Start), end: formatDateToISO(term4End), startCycle: 'A' }
  ];
}

function addDays_(date, days) {
  var d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function addSchoolDays_(startDate, schoolDaysToAdd) {
  var d = new Date(startDate);
  var added = 0;
  while (added < schoolDaysToAdd) {
    d.setDate(d.getDate() + 1);
    var dow = d.getDay();
    if (dow >= 1 && dow <= 5) {
      added++;
    }
  }
  return d;
}

function nextMondayAfter_(date) {
  var d = new Date(date);
  while (d.getDay() !== 1) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

/**
 * Convert incoming value to a local date-only Date (00:00 local).
 * Avoids timezone drift from parsing YYYY-MM-DD with UTC semantics.
 */
function toLocalDateOnly_(value) {
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

/**
 * Server-side version of calculateSchoolWeekDayTerm(date) used by
 * apiGetNextAvailable and timetable helpers.
 * Returns { week: number, day: 1-10, term: termNumber }.
 */
function calculateSchoolWeekDayTerm(date) {
  var settings = getServerSettings_();
  var terms = (settings.termDates || []).slice();
  terms.sort(function (a, b) { return String(a.start || '').localeCompare(String(b.start || '')); });
  if (!terms.length) {
    return { week: 1, day: 1, term: 1, inTerm: false };
  }

  var dateOnly = toLocalDateOnly_(date);
  var currentTerm = getTermForDate_(dateOnly, terms);
  if (!currentTerm) {
    return { week: 1, day: 1, term: 1, inTerm: false };
  }

  // Count school days (Mon-Fri) from start of term up to given date (inclusive)
  var schoolDays = 0;
  var cur = toLocalDateOnly_(currentTerm.start);
  while (cur <= dateOnly) {
    var dow = cur.getDay();
    if (dow >= 1 && dow <= 5) {
      schoolDays++;
    }
    cur.setDate(cur.getDate() + 1);
  }

  var startOffset = String(currentTerm.startCycle || 'A').toUpperCase() === 'B' ? 5 : 0;
  var adjusted = schoolDays + startOffset;

  var dayInCycle = ((adjusted - 1) % 10) + 1;
  var weekNumber = Math.floor((schoolDays - 1) / 5) + 1;

  return {
    week: weekNumber,
    day: dayInCycle,
    term: currentTerm.term,
    inTerm: true
  };
}

function getTermForDate_(date, terms) {
  var d = toLocalDateOnly_(date);
  for (var i = 0; i < terms.length; i++) {
    var t = terms[i];
    var start = toLocalDateOnly_(t.start);
    var end = toLocalDateOnly_(t.end);
    if (d >= start && d <= end) return t;
  }
  return null;
}

// ============================================================================
// DATE/TIME CALCULATIONS
// ============================================================================

// NOTE: calculateNextOccurrence is implemented in the timetable-aware helpers section
// below (time-aware, supports multiple periods per day, and used by sequencing).

/**
 * Calculate assign time (5 minutes before class start).
 * Accepts either a "HH:MM" string or a Date/time value from Sheets.
 * @param {string|Date} classStartTime - e.g. "08:55" or a Date object
 * @returns {string} Assign time in "HH:MM" format
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

  var assignMin = minutes - 5;
  var assignHour = hours;

  if (assignMin < 0) {
    assignMin = 60 + assignMin;
    assignHour = hours - 1;
    if (assignHour < 0) assignHour = 23;
  }

  return String(assignHour).padStart(2, '0') + ':' + String(assignMin).padStart(2, '0');
}


/**
 * Calculate school day number (1-10) for a given date
 * Reuses existing term/cycle logic from main app
 * @param {Date} date - Calendar date
 * @returns {number} Day number (1-10)
 */
function calculateSchoolDayNumber(date) {
  // This uses your existing calculateSchoolWeekDayTerm function
  const result = calculateSchoolWeekDayTerm(date);
  return result.day; // Returns 1-10
}

/**
 * Convert school day number to calendar date
 * @param {number} dayNumber - School day (1-10)
 * @param {Date} afterDate - Reference date
 * @returns {Date} Calendar date
 */
function schoolDayToCalendarDate(dayNumber, afterDate) {
  const currentDay = calculateSchoolDayNumber(afterDate);
  
  let daysToAdd = dayNumber - currentDay;
  
  // If target day is in past of current cycle, add full cycle
  if (daysToAdd <= 0) {
    daysToAdd += 10;
  }
  
  // Add days, skipping weekends
  let targetDate = new Date(afterDate);
  let schoolDaysAdded = 0;
  
  while (schoolDaysAdded < daysToAdd) {
    targetDate.setDate(targetDate.getDate() + 1);
    const dayOfWeek = targetDate.getDay();
    
    // Only count weekdays
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      schoolDaysAdded++;
    }
  }
  
  return targetDate;
}

// ============================================================================
// TIMETABLE OPERATIONS
// ============================================================================

/**
 * Get timetable slots for a class
 * @param {string} classCode - Class code
 * @returns {Array} Array of { day, period, startTime, endTime, room }
 */
function loadTimetableCache_() {
  if (__seqRuntimeCache.timetableByClass) return __seqRuntimeCache.timetableByClass;

  const ss = getSequencerSheet();
  const sheet = ss.getSheetByName(TAB_TIMETABLE);

  if (!sheet) {
    throw new Error('Timetable tab not found');
  }

  const data = sheet.getDataRange().getValues();
  const byClass = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = String(row[0] || '').trim();
    if (!code) continue;

    if (!byClass[code]) byClass[code] = [];
    byClass[code].push({
      day: row[1],
      period: row[2],
      startTime: normaliseTimeToHHMM_(row[3]),
      endTime: normaliseTimeToHHMM_(row[4]),
      room: row[5] || '',
      slotType: String(row[6] || 'CLASS').trim().toUpperCase() || 'CLASS',
      activity: row[7] || '',
      periodLabel: row[8] || ''
    });
  }

  Object.keys(byClass).forEach(code => {
    byClass[code].sort((a, b) => {
      if (a.day !== b.day) return a.day - b.day;
      return a.period - b.period;
    });
  });

  __seqRuntimeCache.timetableByClass = byClass;
  return byClass;
}

/**
 * Get timetable slots for a class
 * @param {string} classCode - Class code
 * @returns {Array} Array of { day, period, startTime, endTime, room }
 */

function canonicalClassCodeKey_(code) {
  var raw = String(code || '').trim().toUpperCase();
  if (!raw) return '';
  var head = raw.split('-')[0].trim();
  head = head.split('(')[0].trim();
  head = head.replace(/\s+/g, '');
  if (!head) head = raw.replace(/\s+/g, '');
  return head;
}

function getTimetableForClass(classCode) {
  const byClass = loadTimetableCache_();
  const code = String(classCode || '').trim();
  var slots = byClass[code] || [];

  if (!slots.length) {
    var wanted = canonicalClassCodeKey_(code);
    var merged = [];
    Object.keys(byClass).forEach(function (k) {
      var ck = canonicalClassCodeKey_(k);
      if (!ck || !wanted) return;
      if (ck === wanted || ck.indexOf(wanted) === 0 || wanted.indexOf(ck) === 0) {
        merged = merged.concat(byClass[k] || []);
      }
    });
    slots = merged;
  }

  return slots.map(slot => ({
    day: slot.day,
    period: slot.period,
    startTime: slot.startTime,
    endTime: slot.endTime,
    room: slot.room,
    slotType: String(slot.slotType || 'CLASS').toUpperCase(),
    activity: slot.activity || '',
    periodLabel: slot.periodLabel || ''
  }));
}

function loadClassMappingCache_() {
  if (__seqRuntimeCache.classToCourse && __seqRuntimeCache.courseToClass) {
    return;
  }

  const ss = getSequencerSheet();
  const sheet = ss.getSheetByName(TAB_CLASS_MAPPING);

  if (!sheet) {
    throw new Error('ClassMapping tab not found');
  }

  const data = sheet.getDataRange().getValues();
  const classToCourse = {};
  const courseToClass = {};

  for (let i = 1; i < data.length; i++) {
    const classCode = String(data[i][0] || '').trim();
    const courseId = String(data[i][1] || '').trim();
    if (!classCode || !courseId) continue;
    classToCourse[classCode] = courseId;
    courseToClass[courseId] = classCode;
  }

  __seqRuntimeCache.classToCourse = classToCourse;
  __seqRuntimeCache.courseToClass = courseToClass;
}

/**
 * Get course ID from class code
 * @param {string} classCode - Class code
 * @returns {string} Google Classroom course ID
 */
function getCourseIdFromClassCode(classCode) {
  loadClassMappingCache_();
  const code = String(classCode || '').trim();
  const courseId = __seqRuntimeCache.classToCourse[code];
  if (courseId) {
    return courseId;
  }

  throw new Error('No course mapping found for ' + classCode);
}

function getClassCodeFromCourseId_(courseId) {
  loadClassMappingCache_();
  const id = String(courseId || '').trim();
  return __seqRuntimeCache.courseToClass[id] || null;
}

function apiSyncSequenceWithClassroom(courseId) {
  if (!courseId) return { success: false, message: 'Missing courseId' };
  const classCode = getClassCodeFromCourseId_(courseId);
  if (!classCode) return { success: false, message: 'No classCode mapping for courseId' };

  const allCw = [];
  let pageToken = null;
  do {
    const res = Classroom.Courses.CourseWork.list(courseId, {
      pageSize: 100,
      pageToken: pageToken || undefined,
      courseWorkStates: ['PUBLISHED', 'DRAFT']
    });
    (res.courseWork || []).forEach(cw => allCw.push(cw));
    pageToken = res.nextPageToken || null;
  } while (pageToken);
  var topicById = getTopicNameByIdForCourse_(courseId);
  const cwById = {};
  allCw.forEach(cw => { cwById[cw.id] = cw; });

  const seq = loadSequence(classCode) || [];
  const seqByClassworkId = {};
  seq.forEach(function (l) {
    var cwid = String(l && l.classworkId || '').trim();
    if (cwid) seqByClassworkId[cwid] = l;
  });
  const kept = [];
  const removed = [];
  let updated = 0;
  let added = 0;
  seq.forEach(l => {
    if (l.classworkId && !cwById[l.classworkId]) {
      removed.push(l);
      return;
    }
    if (l.classworkId && cwById[l.classworkId]) {
      const cw = cwById[l.classworkId];
      // Sheet-first: do not overwrite dates/titles from Classroom for linked rows.
      // Only mirror lifecycle status and optional topic label.
      var cwState = String(cw.state || '').toUpperCase();
      if (cwState === 'PUBLISHED') l.status = 'published';
      if (cwState === 'DRAFT' && String(l.status || '').toLowerCase() === 'published') l.status = 'scheduled';
      if (cw.topicId && topicById[cw.topicId]) {
        l.topicText = String(topicById[cw.topicId] || '');
      }
      updated++;
    }
    kept.push(l);
  });

  // Handle Classroom items not found in sheet by classworkId:
  // - DRAFT = orphan from a previous bump/edit cycle. Never import. Try to delete.
  // - PUBLISHED = may be manually created in Classroom. Import only if no
  //   sheet lesson already has the same title+date (avoids classworkId-mismatch dupes).
  var deleted = 0;
  allCw.forEach(function (cw) {
    var cwid = String(cw && cw.id || '').trim();
    if (!cwid) return;
    if (seqByClassworkId[cwid]) return; // already linked in sheet

    var cwState = String(cw.state || '').toUpperCase();

    if (cwState === 'DRAFT') {
      // Orphaned draft — attempt to delete. If blocked, skip silently (do NOT import).
      try {
        Classroom.Courses.CourseWork.remove(courseId, cwid);
        Logger.log('apiSyncSequenceWithClassroom: deleted orphaned draft ' + cwid +
          ' ("' + (cw.title || '') + '")');
        deleted++;
      } catch (delErr) {
        Logger.log('apiSyncSequenceWithClassroom: could not delete orphaned draft ' + cwid +
          ' (skipped, not imported): ' + String(delErr && delErr.message ? delErr.message : delErr));
      }
      return; // never import orphaned drafts
    }

    if (cwState === 'PUBLISHED') {
      // Check if sheet already has a lesson with the same title + date.
      // If so, just update its classworkId rather than creating a duplicate row.
      var cwTitle = String(cw.title || '').trim().toLowerCase();
      var cwDate = '';
      try {
        var ts = String(cw.scheduledTime || cw.creationTime || '');
        if (ts) {
          var d2 = new Date(ts);
          if (!isNaN(d2.getTime())) {
            cwDate = Utilities.formatDate(d2, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        }
      } catch (e2) {}

      var matchIdx = kept.findIndex(function (l) {
        return String(l.title || '').trim().toLowerCase() === cwTitle &&
               String(l.assignDate || '') === cwDate;
      });

      if (matchIdx >= 0) {
        // Repair classworkId mismatch instead of duplicating
        kept[matchIdx].classworkId = cwid;
        Logger.log('apiSyncSequenceWithClassroom: repaired classworkId for "' +
          kept[matchIdx].title + '" on ' + cwDate + ' -> ' + cwid);
        return;
      }

      // Truly new published item created in Classroom — import it
      var lesson = classroomCourseworkToSequenceLesson_(classCode, cw, topicById);
      if (!lesson) return;
      kept.push(lesson);
      added++;
    }
  });

  if (removed.length || updated || added || deleted) {
    kept.sort(function (a, b) {
      var dateA = new Date(a.assignDate + 'T' + (a.assignTime || '00:00'));
      var dateB = new Date(b.assignDate + 'T' + (b.assignTime || '00:00'));
      return dateA - dateB;
    });
    kept.forEach(function (lesson, idx) { lesson.position = idx + 1; });
    saveSequence(classCode, kept);
  }
  return { success: true, classCode: classCode, removed: removed.length, updated: updated, added: added, deleted: deleted };
}

function classroomCourseworkToSequenceLesson_(classCode, cw, topicById) {
  if (!cw) return null;
  var state = String(cw.state || '').toUpperCase();
  if (state !== 'DRAFT' && state !== 'PUBLISHED') return null;

  // Include all Classroom-created lessons with best available timestamp.
  var ts = String(cw.scheduledTime || cw.creationTime || cw.updateTime || '').trim();
  if (!ts) return null;
  var d = new Date(ts);
  if (isNaN(d.getTime())) return null;

  var assignDate = formatDateToISO(d);
  var assignTime = normaliseTimeToHHMM_(Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm'));
  var dueDate = '';
  var dueTime = '';
  if (cw.dueDate) {
    var yy = Number(cw.dueDate.year || 0);
    var mm = Number(cw.dueDate.month || 0);
    var dd = Number(cw.dueDate.day || 0);
    if (yy && mm && dd) {
      dueDate = yy + '-' + String(mm).padStart(2, '0') + '-' + String(dd).padStart(2, '0');
    }
  }
  if (cw.dueTime) {
    var hh = String(cw.dueTime.hours || 0).padStart(2, '0');
    var mi = String(cw.dueTime.minutes || 0).padStart(2, '0');
    dueTime = normaliseTimeToHHMM_(hh + ':' + mi);
  }

  return {
    id: Utilities.getUuid(),
    position: 0,
    title: String(cw.title || '(Untitled)'),
    assignDate: assignDate,
    assignTime: assignTime,
    dueDate: dueDate,
    dueTime: dueTime,
    status: state === 'PUBLISHED' ? 'published' : 'scheduled',
    classworkId: String(cw.id || ''),
    description: String(cw.description || ''),
    materials: Array.isArray(cw.materials) ? cw.materials : [],
    points: (typeof cw.maxPoints === 'number') ? cw.maxPoints : 0,
    topicText: (cw.topicId && topicById && topicById[cw.topicId]) ? String(topicById[cw.topicId]) : ''
  };
}

function getTopicNameByIdForCourse_(courseId) {
  var map = {};
  try {
    var res = Classroom.Courses.Topics.list(courseId, { pageSize: 100 });
    var topics = res && res.topic ? res.topic : [];
    topics.forEach(function (t) {
      var id = String(t.topicId || '').trim();
      if (!id) return;
      map[id] = String(t.name || '');
    });
  } catch (e) {
    Logger.log('getTopicNameByIdForCourse_ warning: ' + (e && e.message ? e.message : String(e)));
  }
  return map;
}

// ============================================================================
// TERM BOUNDARY CHECKS
// ============================================================================

/**
 * Check if any lessons cross term boundaries
 * @param {Array} sequence - Lesson array
 * @returns {Array} Array of warning objects
 */
function checkTermBoundaries(sequence) {
  const warnings = [];

  for (let i = 0; i < sequence.length; i++) {
    const lesson = sequence[i];
    const term = getTermForDate(new Date(lesson.assignDate));

    if (i > 0) {
      const prevTerm = getTermForDate(new Date(sequence[i - 1].assignDate));

      if (term !== prevTerm) {
        warnings.push({
          lessonId: lesson.id,
          lessonTitle: lesson.title,
          date: lesson.assignDate,
          message: 'Lesson moved from ' + prevTerm + ' to ' + term + '. Timetable may differ.'
        });
      }
    }
  }

  return warnings;
}

/**
 * Get term name for a date (server-side)
 * @param {Date} date - Date to check
 * @returns {string} Term name (e.g., "Term 1", "Term 2")
 */
function getTermForDate(date) {
  var settings = getServerSettings_();
  var terms = settings.termDates || [];
  for (var i = 0; i < terms.length; i++) {
    var t = terms[i];
    var start = new Date(t.start + 'T00:00:00');
    var end = new Date(t.end + 'T23:59:59');
    if (date >= start && date <= end) {
      return 'Term ' + t.term;
    }
  }
  return 'Unknown Term';
}

function createOrUpdateClassworkForLesson_(classCode, lesson, explicitCourseId) {
  if (!lesson) return null;
  var courseId = explicitCourseId || getCourseIdFromClassCode(classCode);
  if (!courseId) throw new Error('No course mapping found for ' + classCode);

  if (lesson.topicText && !lesson.topicId) {
    lesson.topicId = ensureTopicId_(courseId, lesson.topicText);
  }

  var cw = buildCourseWorkFromLesson_(lesson);

  if (lesson.classworkId) {
    Logger.log('Updating Classroom work for lesson ' + lesson.id + ' classworkId=' + lesson.classworkId);

    var fields = {
      title: cw.title,
      description: cw.description,
      dueDate: cw.dueDate || null,
      dueTime: cw.dueTime || null,
      maxPoints: (typeof cw.maxPoints === 'number') ? cw.maxPoints : null,
      scheduledTime: cw.scheduledTime || null,
      topicId: lesson.topicId || null
    };

    function patchWithMask(mask) {
      var body = {};
      mask.forEach(function (k) { body[k] = fields[k]; });
      return Classroom.Courses.CourseWork.patch(body, courseId, lesson.classworkId, {
        updateMask: mask.join(',')
      });
    }

    var attempts = [
      ['title', 'description', 'dueDate', 'dueTime', 'maxPoints', 'scheduledTime', 'topicId'],
      ['title', 'description', 'dueDate', 'dueTime', 'maxPoints', 'topicId'],
      ['title', 'description', 'dueDate', 'dueTime', 'maxPoints'],
      ['title', 'description', 'dueDate', 'dueTime'],
      ['title', 'description']
    ];

    var done = false;
    var lastErr = null;
    var topicPatchedInline = false;
    for (var i = 0; i < attempts.length; i++) {
      var mask = attempts[i];
      try {
        patchWithMask(mask);
        if (mask.indexOf('topicId') !== -1) topicPatchedInline = true;
        done = true;
        break;
      } catch (e) {
        lastErr = e;
        var msg = String(e && e.message ? e.message : e);
        var lower = msg.toLowerCase();
        var scheduledMask = mask.indexOf('scheduledTime') !== -1;
        var isMaskError = msg.indexOf('Non-supported update mask fields specified') !== -1;
        var isScheduleFieldError = scheduledMask && (
          lower.indexOf('scheduledtime') !== -1 ||
          lower.indexOf('schedule') !== -1 ||
          lower.indexOf('draft') !== -1
        );
        if (!isMaskError && !isScheduleFieldError) {
          throw e;
        }
        Logger.log('createOrUpdateClassworkForLesson_ retry mask=' + mask.join(',') + ' error=' + msg);
      }
    }

    if (!done && lastErr) throw lastErr;
    // Keep topic assignment in sync for existing coursework items.
    if (lesson.topicId && !topicPatchedInline) {
      try {
        updateClassworkTopic_(courseId, lesson.classworkId, lesson.topicId);
      } catch (topicErr) {
        throw topicErr;
      }
    }
    return lesson.classworkId;
  }

  Logger.log('Creating Classroom scheduled work for lesson ' + lesson.id + ' at ' + cw.scheduledTime);
  var created = Classroom.Courses.CourseWork.create(cw, courseId);
  return created && created.id ? created.id : null;
}


function normalizeClassroomMaterialPayload_(m) {
  if (!m || typeof m !== 'object') return null;
  if (m.youtubeVideo || m.youTubeVideo) {
    var y = m.youtubeVideo || m.youTubeVideo;
    return { youtubeVideo: y };
  }
  if (m.driveFile) return { driveFile: m.driveFile };
  if (m.link) return { link: m.link };
  if (m.form) return { form: m.form };
  return null;
}

function normalizeClassroomMaterialsPayload_(materials) {
  return (materials || []).map(normalizeClassroomMaterialPayload_).filter(function (x) { return !!x; });
}

function buildCourseWorkFromLesson_(lesson) {
  var cw = {
    title: lesson.title || '(Untitled)',
    description: lesson.description || '',
    workType: 'ASSIGNMENT',
    state: 'DRAFT',
    scheduledTime: buildScheduledTimeString_(lesson.assignDate, lesson.assignTime)
  };

  if (lesson.topicId) {
    cw.topicId = lesson.topicId;
  }

  if (lesson.points) {
    cw.maxPoints = lesson.points;
  }

  if (lesson.materials && lesson.materials.length) {
    cw.materials = normalizeClassroomMaterialsPayload_(lesson.materials);
  }

  if (lesson.dueDate) {
    var d = lesson.dueDate.split('-');
    if (d.length === 3) {
      cw.dueDate = {
        year: Number(d[0]),
        month: Number(d[1]),
        day: Number(d[2])
      };
    }
    if (lesson.dueTime) {
      var t = lesson.dueTime.split(':');
      cw.dueTime = {
        hours: Number(t[0] || 0),
        minutes: Number(t[1] || 0)
      };
    }
  }

  return cw;
}

function getTopicIndexForCourse_(courseId) {
  if (!__seqRuntimeCache.topicIndexByCourse[courseId]) {
    var res = Classroom.Courses.Topics.list(courseId, { pageSize: 100 });
    var topics = res.topic || [];
    var byName = {};
    topics.forEach(function (t) {
      var key = String(t.name || '').trim().toLowerCase();
      if (key && t.topicId) byName[key] = t.topicId;
    });
    __seqRuntimeCache.topicIndexByCourse[courseId] = byName;
  }
  return __seqRuntimeCache.topicIndexByCourse[courseId];
}

function ensureTopicId_(courseId, topicText) {
  var name = String(topicText || '').trim();
  if (!name) return null;
  var key = name.toLowerCase();
  var byName = getTopicIndexForCourse_(courseId);
  if (byName[key]) return byName[key];

  var created = Classroom.Courses.Topics.create({ name: name }, courseId);
  var topicId = created && created.topicId ? created.topicId : null;
  if (topicId) byName[key] = topicId;
  return topicId;
}

function updateClassworkTopic_(courseId, classworkId, topicId) {
  if (!courseId || !classworkId || !topicId) return;
  Classroom.Courses.CourseWork.patch({ topicId: topicId }, courseId, classworkId, {
    updateMask: 'topicId'
  });
}

function onEdit(e) {
  try {
    // Simple triggers run with restricted auth and cannot call services like
    // SpreadsheetApp.openById/Classroom. Only proceed when FULL auth is present.
    if (e && e.authMode && e.authMode !== ScriptApp.AuthMode.FULL) return;
    syncEditedSequenceRowToClassroom_(e);
  } catch (err) {
    Logger.log('onEdit sync error: ' + (err && err.message ? err.message : String(err)));
  }
}

function onSequenceSheetEditInstalled(e) {
  try {
    syncEditedSequenceRowToClassroom_(e);
  } catch (err) {
    Logger.log('onSequenceSheetEditInstalled sync error: ' + (err && err.message ? err.message : String(err)));
  }
}

function syncEditedSequenceRowToClassroom_(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();
  if (!sheet || sheet.getName() !== TAB_SEQUENCES) return;
  var startRow = Math.max(2, range.getRow());
  var endRow = range.getLastRow();
  if (endRow < startRow) return;

  ensureSequencesHeaders_(sheet);
  for (var r = startRow; r <= endRow; r++) {
    syncEditedSequenceRowByNumber_(sheet, r);
  }
}

function syncEditedSequenceRowByNumber_(sheet, rowNumber) {
  var row = sheet.getRange(rowNumber, 1, 1, SEQ_HEADERS.length).getValues()[0];
  if (!row || !row[SEQ_COL.CLASS_CODE] || !row[SEQ_COL.LESSON_ID]) return;

  var classCode = String(row[SEQ_COL.CLASS_CODE] || '').trim();
  var materialsRaw = row[SEQ_COL.MATERIALS];
  var materials = [];
  if (materialsRaw) {
    try {
      materials = JSON.parse(materialsRaw);
      if (!Array.isArray(materials)) materials = [];
    } catch (err) {
      Logger.log('syncEditedSequenceRowByNumber_ invalid materials JSON row ' + rowNumber + ': ' + String(err && err.message ? err.message : err));
      materials = [];
    }
  }

  var lesson = {
    id: String(row[SEQ_COL.LESSON_ID] || '').trim(),
    position: Number(row[SEQ_COL.POSITION] || 0),
    title: String(row[SEQ_COL.TITLE] || '').trim(),
    assignDate: formatDateToISO(row[SEQ_COL.ASSIGN_DATE]),
    assignTime: normaliseTimeToHHMM_(row[SEQ_COL.ASSIGN_TIME]) || '',
    dueDate: formatDateToISO(row[SEQ_COL.DUE_DATE]),
    dueTime: normaliseTimeToHHMM_(row[SEQ_COL.DUE_TIME]) || '',
    status: String(row[SEQ_COL.STATUS] || 'scheduled'),
    classworkId: String(row[SEQ_COL.CLASSWORK_ID] || '').trim(),
    description: String(row[SEQ_COL.DESCRIPTION] || ''),
    materials: materials,
    points: Number(row[SEQ_COL.POINTS] || 0),
    topicText: String(row[SEQ_COL.TOPIC] || '').trim()
  };

  if (!classCode || !lesson.title || !lesson.assignDate) return;
  if (!lesson.assignTime) lesson.assignTime = '00:00';
  if (!isSchedulableForBump_(lesson)) return;

  var courseId = getCourseIdFromClassCode(classCode);

  // If classworkId is missing, try to resolve from existing Classroom items first.
  // This avoids duplicate creates while still allowing sheet-driven updates.
  if (!lesson.classworkId) {
    var resolvedId = findExistingClassworkIdForLesson_(courseId, lesson);
    if (resolvedId) {
      lesson.classworkId = resolvedId;
      sheet.getRange(rowNumber, SEQ_COL.CLASSWORK_ID + 1).setValue(resolvedId);
    }
  }

  var classworkId = '';
  try {
    classworkId = createOrUpdateClassworkForLesson_(classCode, lesson, courseId);
  } catch (err) {
    var msg = String(err && err.message ? err.message : err || '');
    var blocked = /@ProjectPermissionDenied|not permitted to make this request|developer console project is not permitted/i.test(msg);
    // Workaround for environments where patch is blocked:
    // create a new Classroom item and relink this sequence row.
    if (blocked && lesson.classworkId) {
      Logger.log('syncEditedSequenceRowByNumber_ patch blocked at row ' + rowNumber + ', creating replacement classwork.');
      lesson.classworkId = '';
      classworkId = createOrUpdateClassworkForLesson_(classCode, lesson, courseId);
    } else {
      throw err;
    }
  }
  if (classworkId && String(classworkId) !== String(lesson.classworkId || '')) {
    sheet.getRange(rowNumber, SEQ_COL.CLASSWORK_ID + 1).setValue(classworkId);
  }
  if (!lesson.status) {
    sheet.getRange(rowNumber, SEQ_COL.STATUS + 1).setValue('scheduled');
  }
}

function findExistingClassworkIdForLesson_(courseId, lesson) {
  if (!courseId || !lesson) return '';
  var targetTitle = String(lesson.title || '').trim().toLowerCase();
  if (!targetTitle) return '';
  var targetDate = String(lesson.assignDate || '').trim();
  var targetTime = normaliseTimeToHHMM_(lesson.assignTime || '') || '';

  var all = [];
  var pageToken = null;
  do {
    var res = Classroom.Courses.CourseWork.list(courseId, {
      pageSize: 100,
      pageToken: pageToken || undefined,
      courseWorkStates: ['PUBLISHED', 'DRAFT']
    });
    (res.courseWork || []).forEach(function (cw) { all.push(cw); });
    pageToken = res.nextPageToken || null;
  } while (pageToken);

  var candidates = all.filter(function (cw) {
    return String(cw && cw.title || '').trim().toLowerCase() === targetTitle;
  });
  if (!candidates.length) return '';

  var exact = candidates.find(function (cw) {
    var dt = getClassworkLocalDateTime_(cw);
    if (!dt) return false;
    if (String(dt.date) !== targetDate) return false;
    if (!targetTime) return true;
    return String(dt.time) === targetTime;
  });
  if (exact && exact.id) return String(exact.id);

  var dateOnly = candidates.find(function (cw) {
    var dt = getClassworkLocalDateTime_(cw);
    return !!dt && String(dt.date) === targetDate;
  });
  if (dateOnly && dateOnly.id) return String(dateOnly.id);

  return '';
}

function getClassworkLocalDateTime_(cw) {
  var ts = String((cw && (cw.scheduledTime || cw.creationTime || cw.updateTime)) || '').trim();
  if (!ts) return null;
  var d = new Date(ts);
  if (isNaN(d.getTime())) return null;
  var tz = Session.getScriptTimeZone();
  return {
    date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
    time: Utilities.formatDate(d, tz, 'HH:mm')
  };
}

function buildTopicText_(assignDate, courseId, classCode) {
  if (!assignDate) return '';
  var settings = getServerSettings_();
  var subject = '';
  if (settings && settings.subjectMappings && courseId && settings.subjectMappings[courseId]) {
    subject = settings.subjectMappings[courseId];
  } else {
    subject = classCode || 'Subject';
  }
  var d = new Date(assignDate + 'T00:00:00');
  var info = calculateSchoolWeekDayTerm(d);
  return 'Week ' + info.week + ', Day ' + info.day + ' (T' + info.term + '): ' + subject;
}

function buildScheduledTimeString_(dateIso, timeStr) {
  var dParts = String(dateIso || '').split('-');
  var tParts = String(timeStr || '00:00').split(':');
  var year = Number(dParts[0]);
  var month = Number(dParts[1]) - 1;
  var day = Number(dParts[2]);
  var hour = Number(tParts[0] || 0);
  var min = Number(tParts[1] || 0);
  var d = new Date(year, month, day, hour, min, 0, 0);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}


// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get or create SequencerData spreadsheet
 * @returns {Spreadsheet} The sequencer data spreadsheet
 */
function getSequencerSheet() {
  // Prefer explicit sheet ID (reliable)
  if (SEQ_DATA_SHEET_ID) {
    return SpreadsheetApp.openById(SEQ_DATA_SHEET_ID);
  }
  // Fallback: try to find by name
  const files = DriveApp.getFilesByName(SEQ_SHEET_NAME);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }
  
  // Create new sheet
  const ss = SpreadsheetApp.create(SEQ_SHEET_NAME);
  
  // Create tabs
  createSequencesTab(ss);
  createTimetableTab(ss);
  createClassMappingTab(ss);
  
  return ss;
}

/**
 * Create Sequences tab with headers
 */
function createSequencesTab(ss) {
  const sheet = ss.getSheetByName('Sheet1') || ss.insertSheet(TAB_SEQUENCES);
  sheet.setName(TAB_SEQUENCES);
  
  const headers = [
    'classCode', 'lessonId', 'position', 'title', 'assignDate', 'assignTime',
    'dueDate', 'dueTime', 'status', 'classworkId', 'description', 'materials', 'points'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

/**
 * Create Timetable tab with headers
 */
function createTimetableTab(ss) {
  const sheet = ss.insertSheet(TAB_TIMETABLE);
  
  const headers = ['classCode', 'day', 'period', 'startTime', 'endTime', 'room'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

/**
 * Create ClassMapping tab with headers
 */
function createClassMappingTab(ss) {
  const sheet = ss.insertSheet(TAB_CLASS_MAPPING);
  
  const headers = ['classCode', 'courseId', 'courseName'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

/**
 * Format date to ISO string (YYYY-MM-DD)
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
 * Get class code from sequence (helper)
 */
function getClassCodeFromSequence(sequence) {
  // This is a bit of a hack - we'd need to pass classCode properly
  // For now, return empty and handle in calling function
  return '';
}

// ============================================================================
// TESTING FUNCTIONS (Remove in production)
// ============================================================================

/**
 * Test function - create sample timetable data
 */
function testCreateSampleData() {
  const ss = getSequencerSheet();
  
  // Add timetable
  const ttSheet = ss.getSheetByName(TAB_TIMETABLE);
  const ttData = [
    ['7TECHA', 1, 1, '08:55', '09:55', 'T3'],
    ['7TECHA', 3, 1, '08:55', '09:55', 'T3'],
    ['7TECHA', 5, 2, '09:55', '10:55', 'T3'],
    ['8TECHB', 2, 2, '09:55', '10:55', 'T4']
  ];
  ttSheet.getRange(2, 1, ttData.length, ttData[0].length).setValues(ttData);
  
  // Add class mapping (use your actual course IDs)
  const mapSheet = ss.getSheetByName(TAB_CLASS_MAPPING);
  const mapData = [
    ['7TECHA', 'YOUR_COURSE_ID_HERE', 'Year 7 Tech A - Materials'],
    ['8TECHB', 'YOUR_COURSE_ID_HERE', 'Year 8 Tech B - Engineering']
  ];
  mapSheet.getRange(2, 1, mapData.length, mapData[0].length).setValues(mapData);
  
  Logger.log('Sample data created!');
}