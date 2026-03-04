/**
 * PublishTrigger.gs
 * Reliable trigger management + due publishing for sequence lessons.
 */

const PUBLISH_HANDLER = 'checkAndPublishLessons';
const SYNC_HANDLER = 'syncSequencesWithClassroom';
const SHEET_EDIT_HANDLER = 'onSequenceSheetEditInstalled';
const PUBLISH_INTERVAL_MIN = 1;

/**
 * Run once: recreate publish trigger and remove standalone sync trigger
 * (sync is executed inside publish run to avoid lock contention).
 */
function setupPublishTrigger() {
  recreateTrigger_(PUBLISH_HANDLER, function () {
    return ScriptApp.newTrigger(PUBLISH_HANDLER)
      .timeBased()
      .everyMinutes(PUBLISH_INTERVAL_MIN)
      .create();
  });

  removeHandlerTriggers_(SYNC_HANDLER);

  Logger.log('Auto trigger recreated: publish every ' + PUBLISH_INTERVAL_MIN + ' min (standalone sync trigger removed)');
}

function installAllAutomationTriggers() {
  setupPublishTrigger();
  installSequenceSheetEditTrigger();
  Logger.log('All automation triggers ensured.');
}

function checkAndPublishLessons() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // Process any sheet edits queued by doPost (runs before publish checks)
    try { processSheetEditQueue_(); } catch (qErr) {
      Logger.log('processSheetEditQueue_ failed: ' + String(qErr && qErr.message ? qErr.message : qErr));
    }

    var now = new Date();
    var classRows = getAllClassRowsForAutomation_();
    var publishedCount = 0;
    var checkedCount = 0;
    var permissionDeniedSeen = false;

    classRows.forEach(function (row) {
      var classCode = row.classCode;
      var courseId = row.courseId;

      if (courseId) {
        try {
          apiSyncSequenceWithClassroom(courseId);
        } catch (syncErr) {
          Logger.log('checkAndPublishLessons sync failed for ' + classCode + ': ' + String(syncErr && syncErr.message ? syncErr.message : syncErr));
        }
      }

      var result;
      try {
        result = apiGetSequence(classCode);
      } catch (e) {
        Logger.log('checkAndPublishLessons: failed loading sequence for ' + classCode + ' - ' + String(e && e.message ? e.message : e));
        return;
      }

      var lessons = (result && result.success && Array.isArray(result.lessons)) ? result.lessons : [];
      lessons.forEach(function (lesson) {
        if (permissionDeniedSeen) return;
        checkedCount++;
        if (!shouldPublishNow_(lesson, now)) return;

        var outcome = publishDueLesson_(classCode, lesson);
        if (outcome && outcome.success) publishedCount++;
        if (outcome && outcome.permissionDenied) {
          permissionDeniedSeen = true;
          Logger.log('checkAndPublishLessons halted early: Classroom API permission denied. Fix Cloud project permissions, then rerun.');
        }
      });
    });

    Logger.log('checkAndPublishLessons complete. classes=' + classRows.length + ' checked=' + checkedCount + ' published=' + publishedCount);
  } catch (e) {
    Logger.log('checkAndPublishLessons error: ' + String(e && e.message ? e.message : e));
  } finally {
    lock.releaseLock();
  }
}

function syncSequencesWithClassroom() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var classRows = getAllClassRowsForAutomation_();
    classRows.forEach(function (row) {
      if (!row.courseId) return;
      try {
        apiSyncSequenceWithClassroom(row.courseId);
      } catch (e) {
        Logger.log('syncSequencesWithClassroom error for ' + row.classCode + ': ' + String(e && e.message ? e.message : e));
      }
    });
  } catch (e) {
    Logger.log('syncSequencesWithClassroom fatal: ' + String(e && e.message ? e.message : e));
  } finally {
    lock.releaseLock();
  }
}

function installSyncSequencesTrigger() {
  // Deprecated in normal flow; publish trigger already performs sync before publish checks.
  removeHandlerTriggers_(SYNC_HANDLER);
  Logger.log('Standalone sync trigger removed (sync runs inside publish trigger).');
}

function installSequenceSheetEditTrigger() {
  // Remove legacy handler triggers to avoid simple-trigger style failures.
  removeHandlerTriggers_('onEdit');
  ensureTriggerSingleton_(SHEET_EDIT_HANDLER, function () {
    return ScriptApp.newTrigger(SHEET_EDIT_HANDLER)
      .forSpreadsheet(getSequencerSheet())
      .onEdit()
      .create();
  });
}

function publishDueLesson_(classCode, lesson) {
  try {
    if (lesson.classworkId) {
      return publishExistingDraftClasswork_(classCode, lesson);
    }
    return publishLessonToClassroom(classCode, lesson);
  } catch (e) {
    Logger.log('publishDueLesson_ failed for class=' + classCode + ' lesson=' + (lesson && lesson.id ? lesson.id : '?') + ': ' + String(e && e.message ? e.message : e));
    return { success: false, permissionDenied: isProjectPermissionDeniedError_(e) };
  }
}

function publishExistingDraftClasswork_(classCode, lesson) {
  var courseId;
  try {
    courseId = getCourseIdFromClassCode(classCode);
  } catch (mapErr) {
    Logger.log('publishExistingDraftClasswork_: no course mapping for ' + classCode + ': ' + String(mapErr && mapErr.message ? mapErr.message : mapErr));
    return { success: false, permissionDenied: false };
  }

  var cw;
  try {
    cw = Classroom.Courses.CourseWork.get(courseId, String(lesson.classworkId));
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    Logger.log('publishExistingDraftClasswork_: could not fetch classwork ' + lesson.classworkId + ' for ' + classCode + ': ' + msg);
    return publishLessonToClassroom(classCode, lesson);
  }

  if (cw && cw.state === 'PUBLISHED') {
    apiUpdateLessonStatus(classCode, lesson.id, 'published', lesson.classworkId);
    return { success: true, permissionDenied: false };
  }

  try {
    Classroom.Courses.CourseWork.patch({ state: 'PUBLISHED' }, courseId, String(lesson.classworkId), {
      updateMask: 'state'
    });
    apiUpdateLessonStatus(classCode, lesson.id, 'published', lesson.classworkId);
    Logger.log('Published existing draft classwork ' + lesson.classworkId + ' for ' + classCode);
    return { success: true, permissionDenied: false };
  } catch (e) {
    var errMsg = String(e && e.message ? e.message : e);
    Logger.log('Failed to publish existing draft classwork ' + lesson.classworkId + ' for ' + classCode + ': ' + errMsg);

    // Workaround when patch is blocked by project permissions:
    // create a fresh published coursework item instead of mutating the draft.
    if (isProjectPermissionDeniedError_(e)) {
      Logger.log('Patch permission denied; falling back to create-and-publish for lesson ' + (lesson.id || '?'));
      return publishLessonToClassroom(classCode, lesson);
    }
    return { success: false, permissionDenied: false };
  }
}

function publishLessonToClassroom(classCode, lesson) {
  try {
    var courseId = getCourseIdFromClassCode(classCode);
    if (!courseId) {
      Logger.log('publishLessonToClassroom: no course mapping found for ' + classCode);
      return { success: false, permissionDenied: false };
    }

    var courseWork = {
      title: lesson.title,
      description: lesson.description || '',
      workType: 'ASSIGNMENT',
      state: 'PUBLISHED',
      maxPoints: lesson.points || 0
    };

    if (lesson.topicText) {
      try {
        var ensured = apiEnsureTopic(courseId, lesson.topicText);
        if (ensured && ensured.topicId) courseWork.topicId = ensured.topicId;
      } catch (topicErr) {
        Logger.log('publishLessonToClassroom topic ensure failed: ' + String(topicErr && topicErr.message ? topicErr.message : topicErr));
      }
    }

    if (lesson.dueDate) {
      var duePayload = buildFutureDuePayload_(lesson.dueDate, lesson.dueTime);
      if (duePayload) {
        courseWork.dueDate = duePayload.dueDate;
        if (duePayload.dueTime) courseWork.dueTime = duePayload.dueTime;
      } else {
        Logger.log('publishLessonToClassroom: skipped past dueDate for lesson ' + (lesson.id || '?'));
      }
    }

    if (lesson.materials && lesson.materials.length > 0) {
      courseWork.materials = lesson.materials;
    }

    var created = Classroom.Courses.CourseWork.create(courseWork, courseId);
    apiUpdateLessonStatus(classCode, lesson.id, 'published', created && created.id ? created.id : lesson.classworkId);
    Logger.log('Published new classwork for lesson ' + lesson.id + ' in ' + classCode + ' -> ' + (created && created.id ? created.id : 'no-id'));
    return { success: true, permissionDenied: false };
  } catch (e) {
    Logger.log('publishLessonToClassroom failed for lesson ' + (lesson && lesson.id ? lesson.id : '?') + ': ' + String(e && e.message ? e.message : e));
    return { success: false, permissionDenied: isProjectPermissionDeniedError_(e) };
  }
}

function testPublishLesson() {
  var classCode = '7TECHA';
  var lessonId = 'sample-001';

  var result = apiGetSequence(classCode);
  var lesson = (result && result.lessons || []).find(function (l) { return l.id === lessonId; });

  if (!lesson) {
    Logger.log('Lesson not found');
    return;
  }

  var outcome = publishDueLesson_(classCode, lesson);
  Logger.log(outcome && outcome.success ? 'Test publish successful' : 'Test publish failed');
}

function isProjectPermissionDeniedError_(err) {
  var msg = String(err && err.message ? err.message : err || '');
  return /@ProjectPermissionDenied|not permitted to make this request|developer console project is not permitted/i.test(msg);
}

function buildFutureDuePayload_(dueDateStr, dueTimeStr) {
  var dueDate = new Date(String(dueDateStr || '').trim() + 'T00:00:00');
  if (isNaN(dueDate.getTime())) return null;

  var hh = 23;
  var mm = 59;
  if (dueTimeStr) {
    var parts = String(dueTimeStr).split(':');
    hh = Number(parts[0] || 0);
    mm = Number(parts[1] || 0);
    if (!isFinite(hh)) hh = 23;
    if (!isFinite(mm)) mm = 59;
    hh = Math.max(0, Math.min(23, hh));
    mm = Math.max(0, Math.min(59, mm));
  }
  var dueAt = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate(), hh, mm, 0, 0);
  if (dueAt.getTime() <= Date.now()) return null;

  return {
    dueDate: {
      year: dueDate.getFullYear(),
      month: dueDate.getMonth() + 1,
      day: dueDate.getDate()
    },
    dueTime: dueTimeStr ? { hours: hh, minutes: mm } : null
  };
}

function removePublishTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === PUBLISH_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' publish trigger(s)');
}

function listActiveTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Active Triggers:');
  triggers.forEach(function (trigger) {
    Logger.log(' - ' + trigger.getHandlerFunction() + ' (' + trigger.getEventType() + ')');
  });
  if (!triggers.length) Logger.log(' (No triggers set up)');
}

function removeHandlerTriggers_(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function recreateTrigger_(handlerName, factoryFn) {
  removeHandlerTriggers_(handlerName);
  factoryFn();
}

function ensureTriggerSingleton_(handlerName, factoryFn) {
  var triggers = ScriptApp.getProjectTriggers();
  var matched = triggers.filter(function (t) { return t.getHandlerFunction() === handlerName; });

  if (matched.length > 1) {
    for (var i = 1; i < matched.length; i++) {
      ScriptApp.deleteTrigger(matched[i]);
    }
  }

  if (!matched.length) {
    factoryFn();
  }
}

function getMappedClassRows_() {
  var ss = getSequencerSheet();
  var mapSheet = ss.getSheetByName(TAB_CLASS_MAPPING);
  if (!mapSheet) return [];

  var data = mapSheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var classCode = String(data[i][0] || '').trim();
    var courseId = String(data[i][1] || '').trim();
    if (!classCode) continue;
    out.push({ classCode: classCode, courseId: courseId || '' });
  }
  return out;
}

function getAllClassRowsForAutomation_() {
  var out = [];
  var seen = {};

  getMappedClassRows_().forEach(function (r) {
    var code = String(r.classCode || '').trim();
    if (!code || seen[code]) return;
    seen[code] = true;
    out.push({ classCode: code, courseId: String(r.courseId || '').trim() });
  });

  var fromTimetable = [];
  try {
    fromTimetable = getAllClassCodes() || [];
  } catch (e) {
    fromTimetable = [];
  }

  fromTimetable.forEach(function (codeRaw) {
    var code = String(codeRaw || '').trim();
    if (!code || seen[code]) return;
    seen[code] = true;
    var courseId = '';
    try {
      courseId = String(getCourseIdFromClassCode(code) || '').trim();
    } catch (e) {
      courseId = '';
    }
    out.push({ classCode: code, courseId: courseId });
  });

  return out;
}

function shouldPublishNow_(lesson, now) {
  if (!lesson) return false;
  if (String(lesson.status || '').toLowerCase() !== 'scheduled') return false;

  var assignDate = String(lesson.assignDate || '').trim();
  if (!assignDate) return false;

  var assignTime = normalizeHHMM_(String(lesson.assignTime || '').trim());
  var dueAt = new Date(assignDate + 'T' + assignTime + ':00');
  if (isNaN(dueAt.getTime())) return false;

  return dueAt.getTime() <= now.getTime();
}

function normalizeHHMM_(value) {
  var v = String(value || '').trim();
  if (!v) return '00:00';

  var m = v.match(/^(\d{1,2}):(\d{1,2})/);
  if (!m) return '00:00';

  var hh = Math.max(0, Math.min(23, Number(m[1] || 0)));
  var mm = Math.max(0, Math.min(59, Number(m[2] || 0)));
  return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
}

// ============================================================================
// SHEET EDIT QUEUE PROCESSOR
// Drains items enqueued by doPost (which can't call Classroom directly).
// Called at the top of checkAndPublishLessons so edits sync within ~1 minute.
// ============================================================================

function processSheetEditQueue_() {
  var props    = PropertiesService.getScriptProperties();
  var all      = props.getProperties();
  var keys     = Object.keys(all).filter(function(k) { return k.indexOf('EDITQ_') === 0; });
  if (!keys.length) return;

  Logger.log('processSheetEditQueue_: processing ' + keys.length + ' queued edit(s)');

  var ss       = SpreadsheetApp.openById(SEQ_DATA_SHEET_ID);
  var seqSheet = ss.getSheetByName(TAB_SEQUENCES);

  keys.forEach(function(key) {
    try {
      var item      = JSON.parse(all[key]);
      var classCode = String(item.classCode || '').trim();
      var lesson    = item.lesson;
      if (!classCode || !lesson) { props.deleteProperty(key); return; }

      var courseId = getCourseIdFromClassCode(classCode);
      if (!courseId) {
        Logger.log('processSheetEditQueue_: no courseId for ' + classCode);
        props.deleteProperty(key);
        return;
      }

      // Ensure topic exists
      if (lesson.topicText && !lesson.topicId) {
        try { lesson.topicId = ensureTopicId_(courseId, lesson.topicText); } catch (e) {}
      }

      // patch is blocked in this environment (@ProjectPermissionDenied).
      // Strategy: delete the old draft then create a fresh one.
      // Published items are left alone (don't delete student work).
      var newClassworkId = lesson.classworkId;

      if (lesson.classworkId) {
        var existingState = '';
        try {
          var existing = Classroom.Courses.CourseWork.get(courseId, lesson.classworkId);
          existingState = String(existing.state || '').toUpperCase();
        } catch (getErr) {
          Logger.log('processSheetEditQueue_: could not fetch existing classwork ' +
            lesson.classworkId + ': ' + getErr.message);
        }

        if (existingState === 'PUBLISHED') {
          // Don't touch published items — student work may be attached.
          Logger.log('processSheetEditQueue_: skipping published classwork ' + lesson.classworkId);
          props.deleteProperty(key);
          return;
        }

        if (existingState === 'DRAFT') {
          // patch and delete are both blocked (@ProjectPermissionDenied).
          // Clear the classworkId so a fresh draft gets created below.
          // The orphaned old draft will be reconciled by apiSyncSequenceWithClassroom.
          lesson.classworkId = '';
          newClassworkId = '';
        }
      }

      // Create fresh draft with updated values
      var cw = buildCourseWorkFromLesson_(lesson);
      if (lesson.topicId) cw.topicId = lesson.topicId;
      var created = Classroom.Courses.CourseWork.create(cw, courseId);
      newClassworkId = created && created.id ? String(created.id) : '';
      Logger.log('processSheetEditQueue_: created classwork ' + newClassworkId +
        ' for ' + (lesson.title || '?'));

      // Write classworkId back to the sheet
      if (newClassworkId && seqSheet) {
        var data = seqSheet.getDataRange().getValues();
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][SEQ_COL.LESSON_ID] || '') === String(lesson.id || '')) {
            seqSheet.getRange(r + 1, SEQ_COL.CLASSWORK_ID + 1).setValue(newClassworkId);
            break;
          }
        }
      }

      props.deleteProperty(key);

    } catch (err) {
      Logger.log('processSheetEditQueue_ error for ' + key + ': ' +
        (err && err.message ? err.message : String(err)));
      // Retry next minute; discard after 10 min to avoid pile-up.
      try {
        var staleCheck = JSON.parse(all[key]);
        if (staleCheck.queuedAt && (Date.now() - staleCheck.queuedAt) > 600000) {
          Logger.log('processSheetEditQueue_: discarding stale item ' + key);
          props.deleteProperty(key);
        }
      } catch (e2) { props.deleteProperty(key); }
    }
  });
}