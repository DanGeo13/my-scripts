const APP_CONFIG = {
  APP_NAME: 'Classroom Sequence Architect',
  SHEETS: {
    COURSES: 'Courses',
    TIMETABLE: 'Timetable',
    HOLIDAYS: 'Holidays',
    TERM_DATES: 'TermDates',
    SEQUENCES: 'Sequences'
  },
  HEADERS: {
    COURSES: ['CourseId', 'CourseName', 'SubjectCode', 'SubjectGroup', 'SortOrder', 'DriveFolderId', 'Active'],
    TIMETABLE: ['CourseId', 'DayOfWeek', 'PeriodCode', 'StartTime', 'EndTime'],
    HOLIDAYS: ['Date', 'Label'],
    TERM_DATES: ['TermCode', 'StartDate', 'EndDate'],
    SEQUENCES: [
      'CourseId',
      'LessonNumber',
      'LessonDate',
      'DayOfWeek',
      'PeriodCode',
      'Topic',
      'Title',
      'Description',
      'ClassroomCourseWorkId',
      'DriveFileId',
      'Status',
      'UpdatedAt'
    ]
  },
  DATE_FORMAT: 'yyyy-MM-dd',
  SUBJECT_DEFAULT: 'General',
  DAYS: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  DAY_INDEX: {
    Sunday: 0,
    Monday: 1,
    Tuesday: 2,
    Wednesday: 3,
    Thursday: 4,
    Friday: 5,
    Saturday: 6
  }
};

function doGet() {
  ensureSchema_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function bootstrapApp() {
  return runSafely_(function () {
    ensureSchema_();
    const courses = syncCoursesFromClassroom_();
    const home = buildHomeData_(courses);
    return {
      appName: APP_CONFIG.APP_NAME,
      timeZone: Session.getScriptTimeZone(),
      courses: courses,
      home: home
    };
  });
}

function getHomeData() {
  return runSafely_(function () {
    ensureSchema_();
    const courses = getActiveCourses_();
    return buildHomeData_(courses);
  });
}

function getClasswork(payload) {
  return runSafely_(function () {
    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const items = [];
    let token = null;
    do {
      const resp = Classroom.Courses.CourseWork.list(courseId, {
        pageSize: 100,
        pageToken: token
      });
      const list = (resp && resp.courseWork) || [];
      list.forEach(function (cw) {
        items.push({
          id: cw.id,
          title: cw.title || '',
          description: cw.description || '',
          state: cw.state || '',
          updateTime: cw.updateTime || cw.creationTime || ''
        });
      });
      token = resp && resp.nextPageToken;
    } while (token);

    items.sort(function (a, b) {
      return (b.updateTime || '').localeCompare(a.updateTime || '');
    });
    return { items: items };
  });
}

function getAnnouncements(payload) {
  return runSafely_(function () {
    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const items = [];
    let token = null;
    do {
      const resp = Classroom.Courses.Announcements.list(courseId, {
        pageSize: 50,
        pageToken: token
      });
      const list = (resp && resp.announcements) || [];
      list.forEach(function (a) {
        items.push({
          id: a.id,
          text: a.text || '',
          state: a.state || '',
          updateTime: a.updateTime || a.creationTime || ''
        });
      });
      token = resp && resp.nextPageToken;
    } while (token);

    items.sort(function (a, b) {
      return (b.updateTime || '').localeCompare(a.updateTime || '');
    });
    return { items: items };
  });
}

function getSequence(payload) {
  return runSafely_(function () {
    ensureSchema_();
    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const all = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    const lessons = all
      .filter(function (r) {
        return String(r.CourseId) === courseId;
      })
      .map(normalizeSequenceRow_)
      .sort(compareLessons_);
    return { lessons: lessons };
  });
}

function getTimetable(payload) {
  return runSafely_(function () {
    ensureSchema_();
    const filterCourseId = payload && payload.courseId ? String(payload.courseId) : null;
    const rows = getSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE)
      .map(function (r) {
        return {
          courseId: String(r.CourseId || ''),
          dayOfWeek: String(r.DayOfWeek || ''),
          periodCode: String(r.PeriodCode || ''),
          startTime: String(r.StartTime || ''),
          endTime: String(r.EndTime || '')
        };
      })
      .filter(function (r) {
        return filterCourseId ? r.courseId === filterCourseId : true;
      })
      .sort(function (a, b) {
        if (a.dayOfWeek === b.dayOfWeek) return a.startTime.localeCompare(b.startTime);
        return dayNameToIndex_(a.dayOfWeek) - dayNameToIndex_(b.dayOfWeek);
      });
    return { rows: rows };
  });
}

function moveTile(payload) {
  return runSafely_(function () {
    ensureSchema_();
    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const direction = requireString_(payload && payload.direction, 'direction is required');

    const rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES)
      .filter(function (r) {
        return toBoolean_(r.Active, true);
      })
      .sort(function (a, b) {
        return Number(a.SortOrder || 0) - Number(b.SortOrder || 0);
      });

    const idx = rows.findIndex(function (r) {
      return String(r.CourseId) === courseId;
    });
    if (idx < 0) throw new Error('Course not found.');

    const swapWith = direction === 'up' ? idx - 1 : idx + 1;
    if (swapWith < 0 || swapWith >= rows.length) {
      return { changed: false, courses: getActiveCourses_() };
    }

    const temp = rows[idx];
    rows[idx] = rows[swapWith];
    rows[swapWith] = temp;

    rows.forEach(function (row, i) {
      row.SortOrder = i + 1;
    });

    applyCourseSortOrder_(rows);
    return { changed: true, courses: getActiveCourses_() };
  });
}

function saveTileOrder(payload) {
  return runSafely_(function () {
    ensureSchema_();
    const order = (payload && payload.courseIds) || [];
    if (!Array.isArray(order) || !order.length) {
      throw new Error('courseIds must be a non-empty array.');
    }

    const rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
    const rank = {};
    order.forEach(function (id, i) {
      rank[String(id)] = i + 1;
    });

    rows.forEach(function (r) {
      const cid = String(r.CourseId || '');
      if (rank[cid]) {
        r.SortOrder = rank[cid];
      }
    });

    setSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES, rows);
    return { courses: getActiveCourses_() };
  });
}

function addToSequence(payload) {
  return runSafely_(function () {
    ensureSchema_();

    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const customDate = requireString_(payload && payload.customDate, 'customDate is required');
    const title = String((payload && payload.title) || 'Lesson');
    const description = String((payload && payload.description) || '');

    const course = getCourseMap_()[courseId];
    if (!course) throw new Error('Course metadata not found for this course.');

    const timetableSlots = getCourseTimetableSlots_(courseId);
    if (!timetableSlots.length) {
      throw new Error('No timetable rows found for this course. Add timetable rows first.');
    }

    const holidaySet = getHolidaySet_();
    const targetSlot = findSlotOnOrAfterDate_(new Date(customDate + 'T00:00:00'), timetableSlots, holidaySet);
    if (!targetSlot) {
      throw new Error('No valid timetable slot available for this date.');
    }

    let allRows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    const courseRows = allRows
      .filter(function (r) {
        return String(r.CourseId) === courseId;
      })
      .map(normalizeSequenceRow_)
      .sort(compareLessons_);

    const map = normalizeLessonsIntoUniqueSlots_(courseRows, timetableSlots, holidaySet);

    const insertedLesson = {
      lessonNumber: 0,
      lessonDate: formatDate_(targetSlot.date),
      dayOfWeek: APP_CONFIG.DAYS[targetSlot.date.getDay()],
      periodCode: targetSlot.periodCode,
      topic: '',
      title: title,
      description: description,
      classroomCourseWorkId: '',
      driveFileId: '',
      status: 'Draft',
      courseId: courseId
    };

    applyCascadeInsert_(map, insertedLesson, targetSlot, timetableSlots, holidaySet);

    const finalLessons = Object.keys(map)
      .map(function (k) {
        return map[k];
      })
      .sort(compareLessons_);

    const terms = getTermDateRanges_();
    const courseMeta = getCourseMap_()[courseId] || {};
    const subjectCode = String(courseMeta.subjectCode || APP_CONFIG.SUBJECT_DEFAULT);

    finalLessons.forEach(function (lesson, idx) {
      lesson.lessonNumber = idx + 1;
      lesson.title = normalizeLessonTitleForNumber_(lesson.title, lesson.lessonNumber);
      lesson.topic = buildTopic_(lesson.lessonDate, lesson.periodCode, subjectCode, terms);
      lesson.dayOfWeek = APP_CONFIG.DAYS[new Date(lesson.lessonDate + 'T00:00:00').getDay()];
      lesson.status = 'Draft';
    });

    ensureDriveFilesForLessons_(courseMeta, finalLessons);
    const synced = syncLessonsToClassroom_(courseId, finalLessons);

    allRows = allRows.filter(function (r) {
      return String(r.CourseId) !== courseId;
    });

    finalLessons.forEach(function (lesson) {
      allRows.push(toSequenceSheetRow_(lesson));
    });

    setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, allRows);

    return {
      insertedSlot: {
        date: insertedLesson.lessonDate,
        periodCode: insertedLesson.periodCode
      },
      classroomUpdatedCount: synced,
      lessons: finalLessons
    };
  });
}

function normalizeLessonsIntoUniqueSlots_(rows, timetableSlots, holidaySet) {
  var map = {};
  (rows || []).slice().sort(compareLessons_).forEach(function (row) {
    var slot = resolvePreferredSlotForLesson_(row, timetableSlots, holidaySet);
    if (!slot) return;
    var clone = JSON.parse(JSON.stringify(row || {}));
    applyCascadeInsert_(map, clone, slot, timetableSlots, holidaySet);
  });
  return map;
}

function resolvePreferredSlotForLesson_(lesson, timetableSlots, holidaySet) {
  var d = new Date(String(lesson && lesson.lessonDate || '') + 'T00:00:00');
  if (isNaN(d.getTime())) d = new Date();

  var daySlots = buildSlotsForDate_(d, timetableSlots, holidaySet);
  if (daySlots.length) {
    var match = daySlots.find(function (s) {
      return String(s.periodCode || '') === String(lesson && lesson.periodCode || '');
    });
    if (match) return match;
    return daySlots[0];
  }

  return findSlotOnOrAfterDate_(d, timetableSlots, holidaySet);
}

function syncSequenceToClassroom(payload) {
  return runSafely_(function () {
    ensureSchema_();
    const courseId = requireString_(payload && payload.courseId, 'courseId is required');
    const courseMap = getCourseMap_();
    const courseMeta = courseMap[courseId];
    if (!courseMeta) throw new Error('Course metadata not found.');

    const allRows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    const lessons = allRows
      .filter(function (r) {
        return String(r.CourseId) === courseId;
      })
      .map(normalizeSequenceRow_)
      .sort(compareLessons_);

    ensureDriveFilesForLessons_(courseMeta, lessons);
    const synced = syncLessonsToClassroom_(courseId, lessons);

    const remaining = allRows.filter(function (r) {
      return String(r.CourseId) !== courseId;
    });
    lessons.forEach(function (lesson) {
      remaining.push(toSequenceSheetRow_(lesson));
    });
    setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, remaining);

    return {
      classroomUpdatedCount: synced,
      lessons: lessons
    };
  });
}

function runSafely_(fn) {
  try {
    return { success: true, data: fn(), error: null };
  } catch (err) {
    return { success: false, data: null, error: err && err.message ? err.message : String(err) };
  }
}

function ensureSchema_() {
  ensureSheet_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  ensureSheet_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE);
  ensureSheet_(APP_CONFIG.SHEETS.HOLIDAYS, APP_CONFIG.HEADERS.HOLIDAYS);
  ensureSheet_(APP_CONFIG.SHEETS.TERM_DATES, APP_CONFIG.HEADERS.TERM_DATES);
  ensureSheet_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
  migrateSequenceSheetToCanonicalSchema_();
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRange.getValues()[0];
  const needsHeader = headers.some(function (h, i) {
    return String(existingHeaders[i] || '') !== h;
  });

  if (needsHeader) {
    headerRange.setValues([headers]);
  }

  sheet.setFrozenRows(1);
}

function getSheetObjects_(sheetName, headers) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) {
      obj[h] = row[i];
    });
    return obj;
  });
}

function setSheetObjects_(sheetName, headers, rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  const maxRows = sheet.getMaxRows();
  const requiredRows = Math.max(2, rows.length + 1);
  if (requiredRows > maxRows) {
    sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }

  if (!rows.length) return;

  const values = rows.map(function (row) {
    return headers.map(function (h) {
      return row[h] !== undefined ? row[h] : '';
    });
  });

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function syncCoursesFromClassroom_() {
  const classroomCourses = [];
  let token = null;
  do {
    const resp = Classroom.Courses.list({
      teacherId: 'me',
      courseStates: ['ACTIVE'],
      pageSize: 100,
      pageToken: token
    });

    const list = (resp && resp.courses) || [];
    list.forEach(function (course) {
      classroomCourses.push(course);
    });

    token = resp && resp.nextPageToken;
  } while (token);

  const rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  const byId = {};
  rows.forEach(function (r) {
    byId[String(r.CourseId)] = r;
  });

  classroomCourses.forEach(function (course, i) {
    const courseId = String(course.id);
    const existing = byId[courseId];
    if (!existing) {
      byId[courseId] = {
        CourseId: courseId,
        CourseName: course.name || 'Untitled Course',
        SubjectCode: deriveSubjectCode_(course.name || ''),
        SubjectGroup: deriveSubjectGroup_(course.name || ''),
        SortOrder: i + 1,
        DriveFolderId: '',
        Active: true
      };
    } else {
      existing.CourseName = course.name || existing.CourseName;
      existing.SubjectGroup = existing.SubjectGroup || deriveSubjectGroup_(existing.CourseName);
      existing.SubjectCode = existing.SubjectCode || deriveSubjectCode_(existing.CourseName);
      existing.Active = true;
    }
  });

  const allowed = {};
  classroomCourses.forEach(function (c) {
    allowed[String(c.id)] = true;
  });

  const merged = Object.keys(byId).map(function (id) {
    const row = byId[id];
    row.Active = !!allowed[id];
    return row;
  });

  normalizeSortOrder_(merged);
  setSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES, merged);

  return merged
    .filter(function (r) {
      return toBoolean_(r.Active, true);
    })
    .sort(function (a, b) {
      return Number(a.SortOrder || 0) - Number(b.SortOrder || 0);
    })
    .map(function (r) {
      return {
        id: String(r.CourseId),
        name: String(r.CourseName || ''),
        subjectCode: String(r.SubjectCode || APP_CONFIG.SUBJECT_DEFAULT),
        subjectGroup: String(r.SubjectGroup || deriveSubjectGroup_(r.CourseName || '')),
        sortOrder: Number(r.SortOrder || 0),
        driveFolderId: String(r.DriveFolderId || '')
      };
    });
}

function getActiveCourses_() {
  return getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES)
    .filter(function (r) {
      return toBoolean_(r.Active, true);
    })
    .sort(function (a, b) {
      return Number(a.SortOrder || 0) - Number(b.SortOrder || 0);
    })
    .map(function (r) {
      return {
        id: String(r.CourseId),
        name: String(r.CourseName || ''),
        subjectCode: String(r.SubjectCode || APP_CONFIG.SUBJECT_DEFAULT),
        subjectGroup: String(r.SubjectGroup || deriveSubjectGroup_(r.CourseName || '')),
        sortOrder: Number(r.SortOrder || 0),
        driveFolderId: String(r.DriveFolderId || '')
      };
    });
}

function getCourseMap_() {
  const map = {};
  getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES).forEach(function (r) {
    map[String(r.CourseId)] = {
      id: String(r.CourseId),
      name: String(r.CourseName || ''),
      subjectCode: String(r.SubjectCode || APP_CONFIG.SUBJECT_DEFAULT),
      subjectGroup: String(r.SubjectGroup || deriveSubjectGroup_(r.CourseName || '')),
      sortOrder: Number(r.SortOrder || 0),
      driveFolderId: String(r.DriveFolderId || ''),
      active: toBoolean_(r.Active, true)
    };
  });
  return map;
}

function buildHomeData_(courses) {
  const timetableRows = getSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE);
  const today = new Date();
  const todayDate = formatDate_(today);
  const dayName = APP_CONFIG.DAYS[today.getDay()];

  const byCourse = {};
  courses.forEach(function (course) {
    byCourse[course.id] = [];
  });

  timetableRows.forEach(function (row) {
    const cid = String(row.CourseId || '');
    if (!byCourse[cid]) return;
    if (String(row.DayOfWeek || '') !== dayName) return;
    byCourse[cid].push({
      periodCode: String(row.PeriodCode || ''),
      startTime: String(row.StartTime || ''),
      endTime: String(row.EndTime || '')
    });
  });

  Object.keys(byCourse).forEach(function (cid) {
    byCourse[cid].sort(function (a, b) {
      return a.startTime.localeCompare(b.startTime);
    });
  });

  const sequenceRows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES)
    .map(normalizeSequenceRow_)
    .filter(function (row) {
      return row.lessonDate === todayDate;
    });

  const topicLookup = {};
  sequenceRows.forEach(function (row) {
    topicLookup[slotKey_(row.lessonDate, row.periodCode)] = row.topic;
  });

  const overview = courses.map(function (course) {
    const blocks = (byCourse[course.id] || []).map(function (p) {
      return {
        periodCode: p.periodCode,
        startTime: p.startTime,
        endTime: p.endTime,
        topic: topicLookup[slotKey_(todayDate, p.periodCode)] || ''
      };
    });

    return {
      courseId: course.id,
      courseName: course.name,
      blocks: blocks
    };
  });

  return {
    todayDate: todayDate,
    dayName: dayName,
    scheduleByCourse: overview
  };
}

function getCourseTimetableSlots_(courseId) {
  return getSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE)
    .filter(function (r) {
      return String(r.CourseId) === courseId;
    })
    .map(function (r) {
      return {
        dayOfWeek: String(r.DayOfWeek || ''),
        dayIndex: dayNameToIndex_(String(r.DayOfWeek || '')),
        periodCode: String(r.PeriodCode || ''),
        startTime: String(r.StartTime || ''),
        endTime: String(r.EndTime || '')
      };
    })
    .filter(function (slot) {
      return slot.dayIndex >= 1 && slot.dayIndex <= 5;
    })
    .sort(function (a, b) {
      if (a.dayIndex === b.dayIndex) {
        return a.startTime.localeCompare(b.startTime);
      }
      return a.dayIndex - b.dayIndex;
    });
}

function getHolidaySet_() {
  const rows = getSheetObjects_(APP_CONFIG.SHEETS.HOLIDAYS, APP_CONFIG.HEADERS.HOLIDAYS);
  const set = {};
  rows.forEach(function (r) {
    const d = asDate_(r.Date);
    if (!d) return;
    set[formatDate_(d)] = true;
  });
  return set;
}

function findSlotOnOrAfterDate_(date, timetableSlots, holidaySet) {
  let cursor = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  for (let i = 0; i < 730; i++) {
    const dateKey = formatDate_(cursor);
    const day = cursor.getDay();
    const isWeekend = day === 0 || day === 6;
    if (!isWeekend && !holidaySet[dateKey]) {
      const daySlots = timetableSlots
        .filter(function (slot) {
          return slot.dayIndex === day;
        })
        .sort(function (a, b) {
          return a.startTime.localeCompare(b.startTime);
        });
      if (daySlots.length) {
        return {
          date: new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate()),
          periodCode: daySlots[0].periodCode
        };
      }
    }
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
  }
  return null;
}

function findNextSlotAfter_(dateStr, periodCode, timetableSlots, holidaySet) {
  const currentDate = new Date(dateStr + 'T00:00:00');
  const dateSlots = buildSlotsForDate_(currentDate, timetableSlots, holidaySet);
  const idx = dateSlots.findIndex(function (slot) {
    return slot.periodCode === periodCode;
  });

  if (idx >= 0 && idx + 1 < dateSlots.length) {
    return dateSlots[idx + 1];
  }

  let cursor = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + 1);
  for (let i = 0; i < 730; i++) {
    const daySlots = buildSlotsForDate_(cursor, timetableSlots, holidaySet);
    if (daySlots.length) {
      return daySlots[0];
    }
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
  }
  throw new Error('Unable to find next timetable slot for cascade.');
}

function buildSlotsForDate_(date, timetableSlots, holidaySet) {
  const key = formatDate_(date);
  const day = date.getDay();
  const isWeekend = day === 0 || day === 6;
  if (isWeekend || holidaySet[key]) return [];

  return timetableSlots
    .filter(function (slot) {
      return slot.dayIndex === day;
    })
    .sort(function (a, b) {
      return a.startTime.localeCompare(b.startTime);
    })
    .map(function (slot) {
      return {
        date: new Date(date.getFullYear(), date.getMonth(), date.getDate()),
        periodCode: slot.periodCode
      };
    });
}

function applyCascadeInsert_(map, lesson, slot, timetableSlots, holidaySet) {
  let displaced = lesson;
  let curSlot = slot;

  for (let guard = 0; guard < 2000; guard++) {
    displaced.lessonDate = formatDate_(curSlot.date);
    displaced.periodCode = curSlot.periodCode;
    const key = slotKey_(displaced.lessonDate, displaced.periodCode);
    const previous = map[key];
    map[key] = displaced;

    if (!previous) {
      return;
    }

    displaced = previous;
    curSlot = findNextSlotAfter_(displaced.lessonDate, displaced.periodCode, timetableSlots, holidaySet);
  }

  throw new Error('Cascade bump exceeded safety limit. Check timetable/holiday setup.');
}

function getTermDateRanges_() {
  return getSheetObjects_(APP_CONFIG.SHEETS.TERM_DATES, APP_CONFIG.HEADERS.TERM_DATES)
    .map(function (row) {
      const start = asDate_(row.StartDate);
      const end = asDate_(row.EndDate);
      if (!start || !end) return null;
      return {
        code: String(row.TermCode || ''),
        start: new Date(start.getFullYear(), start.getMonth(), start.getDate()),
        end: new Date(end.getFullYear(), end.getMonth(), end.getDate())
      };
    })
    .filter(function (x) {
      return !!x;
    })
    .sort(function (a, b) {
      return a.start.getTime() - b.start.getTime();
    });
}

function buildTopic_(dateStr, periodCode, subjectCode, termRanges) {
  const date = new Date(dateStr + 'T00:00:00');
  const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const term = termRanges.find(function (t) {
    return dateOnly.getTime() >= t.start.getTime() && dateOnly.getTime() <= t.end.getTime();
  });

  let weekNum;
  if (term) {
    const days = Math.floor((dateOnly.getTime() - term.start.getTime()) / 86400000);
    weekNum = Math.floor(days / 7) + 1;
  } else {
    weekNum = isoWeek_(dateOnly);
  }

  const dayMap = { 1: 1, 2: 2, 3: 3, 4: 4, 5: 5 };
  const dayNum = dayMap[dateOnly.getDay()] || 1;
  return 'Week ' + weekNum + ', Day ' + dayNum + ' (' + periodCode + '): ' + subjectCode;
}

function ensureDriveFilesForLessons_(courseMeta, lessons) {
  if (!lessons.length) return;
  let folderId = String(courseMeta.driveFolderId || '');
  if (!folderId) {
    const folder = DriveApp.createFolder(safeFileName_(courseMeta.name + ' Lesson Assets'));
    folderId = folder.getId();
    updateCourseFolder_(courseMeta.id, folderId);
    courseMeta.driveFolderId = folderId;
  }

  lessons.forEach(function (lesson) {
    if (lesson.driveFileId) return;
    const folder = DriveApp.getFolderById(folderId);
    const contents = [
      lesson.topic,
      'Lesson Number: ' + lesson.lessonNumber,
      lesson.title ? 'Lesson Title: ' + lesson.title : '',
      lesson.description ? 'Notes: ' + lesson.description : ''
    ].filter(Boolean).join('\n\n');
    const file = folder.createFile(safeFileName_(lesson.topic) + '.txt', contents, MimeType.PLAIN_TEXT);
    lesson.driveFileId = file.getId();
  });
}

function syncLessonsToClassroom_(courseId, lessons) {
  let synced = 0;

  lessons.forEach(function (lesson) {
    if (lesson.classroomCourseWorkId) {
      try {
        Classroom.Courses.CourseWork.remove(courseId, lesson.classroomCourseWorkId);
      } catch (err) {
        // Keep going; recreation still enforces latest topic/title state.
      }
      lesson.classroomCourseWorkId = '';
    }

    // Enforce sequential lesson naming for all created coursework.
    lesson.title = normalizeLessonTitleForNumber_(lesson.title, lesson.lessonNumber);

    const cwBody = {
      title: lesson.title,
      description: buildCourseworkDescription_(lesson),
      state: 'DRAFT',
      workType: 'ASSIGNMENT'
    };

    if (lesson.driveFileId) {
      cwBody.materials = [
        {
          driveFile: {
            driveFile: {
              id: lesson.driveFileId
            },
            shareMode: 'VIEW'
          }
        }
      ];
    }

    const created = Classroom.Courses.CourseWork.create(cwBody, courseId);
    lesson.classroomCourseWorkId = created.id;
    lesson.status = created.state || 'DRAFT';
    synced++;
  });

  return synced;
}

function buildCourseworkDescription_(lesson) {
  const parts = [];
  if (lesson.title) parts.push('Lesson: ' + lesson.title);
  if (lesson.description) parts.push(lesson.description);
  parts.push('Scheduled date: ' + lesson.lessonDate + ' (' + lesson.periodCode + ')');
  return parts.join('\n\n');
}

function normalizeLessonTitleForNumber_(title, lessonNumber) {
  var raw = String(title || '').trim();
  var num = Number(lessonNumber || 0);
  if (!num) return raw || 'Lesson';
  if (!raw) return 'Lesson ' + num;
  var m = raw.match(/^Lesson\s*\d+[a-zA-Z]?\s*:?\s*(.*)$/i);
  var rest = m ? String(m[1] || '').trim() : raw;
  return rest ? ('Lesson ' + num + ': ' + rest) : ('Lesson ' + num);
}

function renumberSequenceRowsForCourse_(rows, courseId, canonicalClass) {
  var target = [];
  (rows || []).forEach(function (r, idx) {
    var rowCourse = String(r.CourseId || '');
    var sameCourse = rowCourse === String(courseId || '')
      || (!!canonicalClass && canonicalClassCodeServer_(rowCourse) === canonicalClass);
    if (!sameCourse) return;
    if (!String(r.LessonDate || '').trim()) return;
    target.push({ idx: idx, row: r });
  });

  if (!target.length) return false;

  target.sort(function (a, b) {
    return compareLessons_(normalizeSequenceRow_(a.row), normalizeSequenceRow_(b.row));
  });

  var changed = false;
  target.forEach(function (entry, pos) {
    var r = entry.row;
    var rowChanged = false;
    var expectedNum = pos + 1;
    var nextTitle = normalizeLessonTitleForNumber_(String(r.Title || ''), expectedNum);
    if (Number(r.LessonNumber || 0) !== expectedNum) {
      r.LessonNumber = expectedNum;
      changed = true;
      rowChanged = true;
    }
    if (String(r.Title || '') !== nextTitle) {
      r.Title = nextTitle;
      changed = true;
      rowChanged = true;
    }
    if (r.LessonDate) {
      var d = new Date(String(r.LessonDate) + 'T00:00:00');
      var day = isNaN(d.getTime()) ? '' : APP_CONFIG.DAYS[d.getDay()];
      if (day && String(r.DayOfWeek || '') !== day) {
        r.DayOfWeek = day;
        changed = true;
        rowChanged = true;
      }
    }
    if (rowChanged) r.UpdatedAt = new Date();
  });

  return changed;
}

function patchClassroomTitlesFromSequenceRows_(rows, courseId, canonicalClass) {
  var cid = String(courseId || '').trim();
  if (!cid) return;
  (rows || []).forEach(function (r) {
    var rowCourse = String(r.CourseId || '');
    var sameCourse = rowCourse === cid
      || (!!canonicalClass && canonicalClassCodeServer_(rowCourse) === canonicalClass);
    if (!sameCourse) return;
    var cwId = String(r.ClassroomCourseWorkId || '').trim();
    if (!cwId) return;
    try {
      Classroom.Courses.CourseWork.patch({
        title: String(r.Title || '(Untitled)')
      }, cid, cwId, { updateMask: 'title' });
    } catch (e) {
      // Non-fatal; sheet remains source-of-truth and next sync can retry.
    }
  });
}

function toSequenceSheetRow_(lesson) {
  var normalized = canonicalizeSequenceSheetRow_({
    CourseId: lesson.courseId,
    LessonNumber: lesson.lessonNumber,
    LessonDate: lesson.lessonDate,
    DayOfWeek: lesson.dayOfWeek,
    PeriodCode: lesson.periodCode,
    Topic: lesson.topic,
    Title: lesson.title,
    Description: lesson.description,
    ClassroomCourseWorkId: lesson.classroomCourseWorkId || '',
    DriveFileId: lesson.driveFileId || '',
    Status: lesson.status || 'Draft',
    UpdatedAt: new Date()
  });
  return {
    CourseId: normalized.CourseId,
    LessonNumber: normalized.LessonNumber,
    LessonDate: normalized.LessonDate,
    DayOfWeek: normalized.DayOfWeek,
    PeriodCode: normalized.PeriodCode,
    Topic: normalized.Topic,
    Title: normalized.Title,
    Description: normalized.Description,
    ClassroomCourseWorkId: normalized.ClassroomCourseWorkId,
    DriveFileId: normalized.DriveFileId,
    Status: normalized.Status,
    UpdatedAt: normalized.UpdatedAt || new Date()
  };
}

function normalizeSequenceRow_(row) {
  var canon = canonicalizeSequenceSheetRow_(row);
  return {
    courseId: String(canon.CourseId || ''),
    lessonNumber: Number(canon.LessonNumber || 0),
    lessonDate: asDateString_(canon.LessonDate),
    dayOfWeek: String(canon.DayOfWeek || ''),
    periodCode: String(canon.PeriodCode || ''),
    topic: String(canon.Topic || ''),
    title: String(canon.Title || ''),
    description: String(canon.Description || ''),
    classroomCourseWorkId: String(canon.ClassroomCourseWorkId || ''),
    driveFileId: String(canon.DriveFileId || ''),
    status: String(canon.Status || 'Draft')
  };
}

function migrateSequenceSheetToCanonicalSchema_() {
  var rows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
  if (!rows.length) return;

  var changed = false;
  var migrated = rows.map(function (r) {
    var before = serializeSequenceRowForCompare_(r);
    var canon = canonicalizeSequenceSheetRow_(r);
    var after = serializeSequenceRowForCompare_(canon);
    if (before !== after) changed = true;
    return canon;
  });

  var deduped = dedupeCanonicalSequenceRows_(migrated);
  if (deduped.length !== migrated.length) changed = true;

  if (changed) {
    setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, deduped);
  }
}

function serializeSequenceRowForCompare_(row) {
  return JSON.stringify([
    String(row.CourseId || '').trim(),
    String(row.LessonNumber || '').trim(),
    String(row.LessonDate || '').trim(),
    String(row.DayOfWeek || '').trim(),
    String(row.PeriodCode || '').trim(),
    String(row.Topic || '').trim(),
    String(row.Title || '').trim(),
    String(row.Description || '').trim(),
    String(row.ClassroomCourseWorkId || '').trim(),
    String(row.DriveFileId || '').trim(),
    String(row.Status || '').trim(),
    String(row.UpdatedAt || '').trim()
  ]);
}

function dedupeCanonicalSequenceRows_(rows) {
  var bestByKey = {};
  var passthrough = [];

  rows.forEach(function (r) {
    var cw = String(r.ClassroomCourseWorkId || '').trim();
    if (!cw) {
      passthrough.push(r);
      return;
    }
    var key = String(r.CourseId || '').trim() + '|' + cw;
    if (!bestByKey[key]) {
      bestByKey[key] = r;
      return;
    }
    var prev = bestByKey[key];
    var prevScore = scoreCanonicalSequenceRow_(prev);
    var nextScore = scoreCanonicalSequenceRow_(r);
    if (nextScore >= prevScore) bestByKey[key] = r;
  });

  var merged = passthrough.concat(Object.keys(bestByKey).map(function (k) { return bestByKey[k]; }));
  merged.sort(function (a, b) {
    var ca = String(a.CourseId || '');
    var cb = String(b.CourseId || '');
    if (ca !== cb) return ca.localeCompare(cb);
    var aa = normalizeSequenceRow_(a);
    var bb = normalizeSequenceRow_(b);
    return compareLessons_(aa, bb);
  });
  return merged;
}

function scoreCanonicalSequenceRow_(r) {
  var score = 0;
  if (String(r.LessonDate || '').trim()) score += 4;
  if (String(r.Title || '').trim()) score += 3;
  if (Number(r.LessonNumber || 0) > 0) score += 2;
  if (String(r.PeriodCode || '').trim()) score += 1;
  if (String(r.Description || '').trim()) score += 1;
  if (String(r.Topic || '').trim()) score += 1;
  return score;
}

function canonicalizeSequenceSheetRow_(row) {
  var r = row || {};
  var legacy = looksLegacySequenceShape_(r);
  var mappedCourseId = coerceSequenceCourseIdToClassroom_(legacy ? r.CourseId : r.CourseId);

  var result = {
    CourseId: String(mappedCourseId || r.CourseId || '').trim(),
    LessonNumber: 0,
    LessonDate: '',
    DayOfWeek: '',
    PeriodCode: '',
    Topic: '',
    Title: '',
    Description: '',
    ClassroomCourseWorkId: '',
    DriveFileId: '',
    Status: 'Draft',
    UpdatedAt: new Date()
  };

  if (legacy) {
    var oldTitle = String(r.DayOfWeek || '').trim();
    var oldAssignDate = asDateString_(r.PeriodCode);
    var oldAssignTime = normalizeHHMMFromAnyServer_(r.Topic);
    var oldStatus = String(r.ClassroomCourseWorkId || '').trim();
    var oldCwId = String(r.DriveFileId || '').trim();
    var oldDescription = String(r.Status || '');
    var oldPosition = Number(r.LessonDate || 0);

    var parsedNum = parseLessonNumberFromTitleServer_(oldTitle);
    result.LessonNumber = oldPosition > 0 ? oldPosition : parsedNum;
    result.LessonDate = oldAssignDate || '';
    result.DayOfWeek = result.LessonDate ? dayNameFromIsoDateServer_(result.LessonDate) : '';
    result.PeriodCode = inferPeriodCodeForCourseDateTimeServer_(result.CourseId, result.LessonDate, oldAssignTime) || 'P1';
    result.Topic = '';
    result.Title = oldTitle || String(r.Title || '').trim();
    result.Description = oldDescription || String(r.Description || '');
    result.ClassroomCourseWorkId = looksLikeClassworkIdServer_(oldCwId) ? oldCwId : '';
    result.DriveFileId = '';
    result.Status = normalizeSequenceStatusServer_(oldStatus);
    result.UpdatedAt = asDate_(r.UpdatedAt) || new Date();
  } else {
    var directTitle = String(r.Title || '').trim();
    var fallbackTitle = String(r.DayOfWeek || '').trim();
    var lessonDate = asDateString_(r.LessonDate);
    var period = normalizePeriodCodeServer_(r.PeriodCode);
    var topic = String(r.Topic || '').trim();
    var status = normalizeSequenceStatusServer_(r.Status);
    var cwId = String(r.ClassroomCourseWorkId || '').trim();
    var drive = String(r.DriveFileId || '').trim();

    if (!lessonDate && isIsoDateServer_(period)) {
      lessonDate = period;
      period = '';
    }

    if (!cwId && looksLikeClassworkIdServer_(drive) && looksLikeStatusTokenServer_(String(r.ClassroomCourseWorkId || ''))) {
      cwId = drive;
      drive = '';
      status = normalizeSequenceStatusServer_(String(r.ClassroomCourseWorkId || ''));
    }

    var lessonNum = Number(r.LessonNumber || 0);
    if (!(lessonNum > 0)) {
      lessonNum = parseLessonNumberFromTitleServer_(directTitle || fallbackTitle);
    }

    result.LessonNumber = lessonNum > 0 ? lessonNum : 0;
    result.LessonDate = lessonDate || '';
    result.DayOfWeek = result.LessonDate ? dayNameFromIsoDateServer_(result.LessonDate) : String(r.DayOfWeek || '').trim();
    result.PeriodCode = period || inferPeriodCodeForCourseDateTimeServer_(result.CourseId, result.LessonDate, normalizeHHMMFromAnyServer_(topic)) || 'P1';
    result.Topic = topic;
    result.Title = directTitle || fallbackTitle;
    result.Description = String(r.Description || '');
    result.ClassroomCourseWorkId = looksLikeClassworkIdServer_(cwId) ? cwId : '';
    result.DriveFileId = looksLikeClassworkIdServer_(drive) && !result.ClassroomCourseWorkId ? '' : drive;
    result.Status = status;
    result.UpdatedAt = asDate_(r.UpdatedAt) || new Date();
  }

  if (!result.CourseId) result.CourseId = String(r.CourseId || '').trim();
  if (!result.Title) result.Title = '(Untitled)';
  if (!result.PeriodCode) result.PeriodCode = 'P1';
  if (!result.Status) result.Status = 'Draft';
  return result;
}

function looksLegacySequenceShape_(row) {
  var lessonNumber = String(row && row.LessonNumber || '').trim();
  var lessonDate = String(row && row.LessonDate || '').trim();
  var periodCode = String(row && row.PeriodCode || '').trim();
  var dayOfWeek = String(row && row.DayOfWeek || '').trim();
  var cwColumn = String(row && row.ClassroomCourseWorkId || '').trim();
  var driveColumn = String(row && row.DriveFileId || '').trim();
  var statusColumn = String(row && row.Status || '').trim();

  var idLike = looksLikeUuidServer_(lessonNumber) || !/^\d+$/.test(lessonNumber);
  var positionLike = /^\d+$/.test(lessonDate);
  var assignDateLike = isIsoDateServer_(periodCode);
  var titleLike = /lesson/i.test(dayOfWeek);
  var statusLike = looksLikeStatusTokenServer_(cwColumn);
  var cwIdLike = looksLikeClassworkIdServer_(driveColumn);
  var longDescriptionLike = statusColumn.length > 40 || statusColumn.indexOf('\n') >= 0;

  return !!(idLike && positionLike && assignDateLike && titleLike && statusLike && cwIdLike && longDescriptionLike);
}

function looksLikeUuidServer_(value) {
  var s = String(value || '').trim();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
}

function looksLikeClassworkIdServer_(value) {
  var s = String(value || '').trim();
  return /^\d{6,}$/.test(s);
}

function looksLikeStatusTokenServer_(value) {
  var s = String(value || '').trim().toUpperCase();
  return s === 'PUBLISHED' || s === 'SCHEDULED' || s === 'DRAFT' || s === 'FAILED' || s === 'ARCHIVED' || s === 'DELETED' || s === 'PENDING' || s === 'PUBLISHED';
}

function normalizeSequenceStatusServer_(value) {
  var s = String(value || '').trim().toUpperCase();
  if (!s) return 'Draft';
  if (s === 'PUBLISHED') return 'PUBLISHED';
  if (s === 'SCHEDULED') return 'SCHEDULED';
  if (s === 'DRAFT') return 'Draft';
  if (s === 'FAILED') return 'FAILED';
  if (s === 'ARCHIVED') return 'ARCHIVED';
  if (s === 'DELETED') return 'DELETED';
  if (s === 'PENDING') return 'PENDING';
  return String(value || 'Draft');
}

function coerceSequenceCourseIdToClassroom_(rawCourse) {
  var raw = String(rawCourse || '').trim();
  if (!raw) return '';
  if (/^\d{6,}$/.test(raw)) return raw;
  var classCode = canonicalClassCodeServer_(extractClassCodeServer_(raw));
  if (!classCode) return raw;
  var mapped = courseIdFromClassCodeServer_(classCode);
  return String(mapped || raw);
}

function isIsoDateServer_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || '').trim());
}

function dayNameFromIsoDateServer_(isoDate) {
  var d = new Date(String(isoDate || '') + 'T00:00:00');
  if (isNaN(d.getTime())) return '';
  return APP_CONFIG.DAYS[d.getDay()];
}

function normalizePeriodCodeServer_(value) {
  var s = String(value || '').trim().toUpperCase();
  if (!s) return '';
  if (/^P\d+$/.test(s)) return s;
  var m = s.match(/(\d+)/);
  if (m) return 'P' + String(Number(m[1] || 0));
  return s;
}

function normalizeHHMMFromAnyServer_(value) {
  var s = String(value || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!m) return '';
  return normalizeHHMMServer_(s);
}

function parseLessonNumberFromTitleServer_(title) {
  var m = String(title || '').match(/Lesson\s*(\d+)/i);
  return m ? Number(m[1] || 0) : 0;
}

function inferPeriodCodeForCourseDateTimeServer_(courseId, lessonDate, hhmm) {
  var cid = String(courseId || '').trim();
  var d = String(lessonDate || '').trim();
  var t = String(hhmm || '').trim();
  if (!cid || !d || !t) return '';
  var classCode = classCodeForCourseIdServer_(cid);
  if (!classCode) {
    classCode = extractClassCodeServer_(cid);
  }
  var byClass = buildTimetableByClass_();
  var slots = byClass[canonicalClassCodeServer_(classCode)] || [];
  if (!slots.length) return '';
  var day = new Date(d + 'T00:00:00');
  if (isNaN(day.getTime())) return '';
  var weekday = day.getDay() === 0 ? 7 : day.getDay();
  var daySlots = slots.filter(function (s) {
    return dayToWeekdayServer_(s.day) === weekday;
  });
  if (!daySlots.length) return '';
  var exact = daySlots.find(function (s) {
    return normalizeHHMMServer_(s.startTime) === normalizeHHMMServer_(t);
  });
  if (exact && Number(exact.period || 0) > 0) return 'P' + String(Number(exact.period));
  var beforeOrEqual = daySlots
    .filter(function (s) { return normalizeHHMMServer_(s.startTime) <= normalizeHHMMServer_(t); })
    .sort(function (a, b) { return String(b.startTime || '').localeCompare(String(a.startTime || '')); });
  var candidate = beforeOrEqual[0] || daySlots.sort(function (a, b) { return String(a.startTime || '').localeCompare(String(b.startTime || '')); })[0];
  return (candidate && Number(candidate.period || 0) > 0) ? ('P' + String(Number(candidate.period))) : 'P1';
}

function compareLessons_(a, b) {
  const keyA = a.lessonDate + '|' + normalizeTimeCode_(a.periodCode);
  const keyB = b.lessonDate + '|' + normalizeTimeCode_(b.periodCode);
  return keyA.localeCompare(keyB);
}

function slotKey_(dateStr, periodCode) {
  return dateStr + '|' + periodCode;
}

function normalizeSortOrder_(rows) {
  rows.sort(function (a, b) {
    return Number(a.SortOrder || 0) - Number(b.SortOrder || 0);
  });

  rows.forEach(function (row, idx) {
    row.SortOrder = idx + 1;
  });
}

function applyCourseSortOrder_(activeRows) {
  const all = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  const lookup = {};
  activeRows.forEach(function (r) {
    lookup[String(r.CourseId)] = Number(r.SortOrder || 0);
  });

  all.forEach(function (row) {
    const cid = String(row.CourseId || '');
    if (lookup[cid]) {
      row.SortOrder = lookup[cid];
    }
  });

  setSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES, all);
}

function updateCourseFolder_(courseId, folderId) {
  const rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  rows.forEach(function (r) {
    if (String(r.CourseId) === String(courseId)) {
      r.DriveFolderId = folderId;
    }
  });
  setSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES, rows);
}

function deriveSubjectCode_(name) {
  const cleaned = String(name || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .trim();
  if (!cleaned) return APP_CONFIG.SUBJECT_DEFAULT;
  const parts = cleaned.split(/\s+/).filter(Boolean);
  return (parts[0] || APP_CONFIG.SUBJECT_DEFAULT).slice(0, 8);
}

function deriveSubjectGroup_(name) {
  const n = String(name || '').toLowerCase();
  if (n.indexOf('engineer') >= 0 || n.indexOf('tech') >= 0) return 'Engineering';
  if (n.indexOf('design') >= 0) return 'Design';
  if (n.indexOf('food') >= 0 || n.indexOf('cook') >= 0) return 'Food';
  if (n.indexOf('textile') >= 0 || n.indexOf('fabric') >= 0) return 'Textiles';
  if (n.indexOf('agri') >= 0 || n.indexOf('farm') >= 0) return 'Agriculture';
  if (n.indexOf('graphic') >= 0 || n.indexOf('media') >= 0) return 'Graphics';
  return 'Engineering';
}

function asDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  const d = new Date(value);
  if (isNaN(d.getTime())) return null;
  return d;
}

function asDateString_(value) {
  const d = asDate_(value);
  if (!d) return '';
  return formatDate_(d);
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), APP_CONFIG.DATE_FORMAT);
}

function isoWeek_(date) {
  const tmp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = tmp.getUTCDay() || 7;
  tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
  return Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7);
}

function dayNameToIndex_(dayName) {
  return APP_CONFIG.DAY_INDEX[String(dayName || '').trim()] !== undefined
    ? APP_CONFIG.DAY_INDEX[String(dayName || '').trim()]
    : -1;
}

function requireString_(val, msg) {
  if (val === null || val === undefined || String(val).trim() === '') {
    throw new Error(msg);
  }
  return String(val).trim();
}

function toBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  if (value === 'TRUE' || value === 'true' || value === 1 || value === '1') return true;
  if (value === 'FALSE' || value === 'false' || value === 0 || value === '0') return false;
  return defaultValue;
}

function safeFileName_(name) {
  return String(name || 'Lesson')
    .replace(/[\\/:*?"<>|#\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120);
}

function normalizeTimeCode_(periodCode) {
  const m = String(periodCode || '').match(/(\d+)/);
  if (!m) return periodCode;
  return ('000' + m[1]).slice(-3) + periodCode;
}

/* =========================
 * Compatibility API Layer
 * For legacy index.html (api* functions)
 * ========================= */

function apiListCourses() {
  try {
    ensureSchema_();
    var items = [];
    ['ACTIVE', 'ARCHIVED'].forEach(function (state) {
      var token = null;
      do {
        var resp = Classroom.Courses.list({
          teacherId: 'me',
          courseStates: [state],
          pageSize: 100,
          pageToken: token
        });
        (resp.courses || []).forEach(function (c) {
          items.push({
            id: String(c.id || ''),
            name: String(c.name || ''),
            section: String(c.section || ''),
            room: String(c.room || ''),
            courseState: String(c.courseState || state || 'ACTIVE'),
            creationTime: String(c.creationTime || ''),
            updateTime: String(c.updateTime || '')
          });
        });
        token = resp.nextPageToken;
      } while (token);
    });
    return items;
  } catch (e) {
    return [];
  }
}

function apiGetSettings() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('APP_SETTINGS_JSON');
    if (!raw) return {
      termDates: [],
      subjectMappings: {},
      courseOrder: [],
      courseColors: {},
      hiddenCourseIds: [],
      defaultReuseByCourse: {},
      reuseFiltersBySource: {}
    };
    var parsed = JSON.parse(raw);
    return {
      termDates: parsed.termDates || [],
      subjectMappings: parsed.subjectMappings || {},
      courseOrder: parsed.courseOrder || [],
      courseColors: parsed.courseColors || {},
      hiddenCourseIds: parsed.hiddenCourseIds || [],
      defaultReuseByCourse: parsed.defaultReuseByCourse || {},
      reuseFiltersBySource: parsed.reuseFiltersBySource || {}
    };
  } catch (e) {
    return {
      termDates: [],
      subjectMappings: {},
      courseOrder: [],
      courseColors: {},
      hiddenCourseIds: [],
      defaultReuseByCourse: {},
      reuseFiltersBySource: {}
    };
  }
}

function apiSaveSettings(payload) {
  var safe = payload || {};
  PropertiesService.getDocumentProperties().setProperty('APP_SETTINGS_JSON', JSON.stringify({
    termDates: safe.termDates || [],
    subjectMappings: safe.subjectMappings || {},
    courseOrder: safe.courseOrder || [],
    courseColors: safe.courseColors || {},
    hiddenCourseIds: safe.hiddenCourseIds || [],
    defaultReuseByCourse: safe.defaultReuseByCourse || {},
    reuseFiltersBySource: safe.reuseFiltersBySource || {}
  }));
  return { success: true };
}

function apiSaveCourseOrder(order) {
  return runSafely_(function () {
    return saveTileOrder({ courseIds: order || [] });
  });
}

function apiGetClassMappings() {
  try {
    ensureSchema_();
    var rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
    var mappings = rows
      .filter(function (r) { return toBoolean_(r.Active, true); })
      .map(function (r) {
        var name = String(r.CourseName || '');
        return {
          classCode: extractClassCodeServer_(name),
          courseId: String(r.CourseId || ''),
          courseName: name
        };
      })
      .filter(function (m) { return !!m.classCode; });
    return { success: true, mappings: mappings };
  } catch (e) {
    return { success: false, mappings: [], error: String(e && e.message || e) };
  }
}

function apiListTopics(courseId) {
  try {
    var out = [];
    var token = null;
    do {
      var resp = Classroom.Courses.Topics.list(String(courseId), { pageSize: 100, pageToken: token });
      (resp.topic || []).forEach(function (t) {
        out.push({ topicId: String(t.topicId || ''), name: String(t.name || '') });
      });
      token = resp.nextPageToken;
    } while (token);
    return out;
  } catch (e) {
    return [];
  }
}

function apiEnsureTopic(courseId, topicName) {
  var name = String(topicName || '').trim();
  if (!name) return { topicId: null };

  var existing = apiListTopics(courseId).find(function (t) {
    return String(t.name || '').toLowerCase() === name.toLowerCase();
  });
  if (existing) return { topicId: existing.topicId };

  var created = Classroom.Courses.Topics.create({ name: name }, String(courseId));
  return { topicId: String(created.topicId || '') };
}

function apiListCourseWork(courseId, opts) {
  ensureSchema_();
  // Keep Sequences sheet close to Classroom state whenever classwork is fetched.
  try {
    var classCodeKey = classCodeForCourseIdServer_(String(courseId || ''));
    syncSequenceSheetFromClassroom_(String(courseId || ''), classCodeKey);
  } catch (e) {
    // Non-fatal: continue serving classwork even if sync step fails.
  }

  var options = opts || {};
  var pageSize = Math.max(1, Math.min(100, Number(options.maxItems || 50)));
  var token = options.pageToken ? String(options.pageToken) : null;

  var resp = Classroom.Courses.CourseWork.list(String(courseId), {
    pageSize: pageSize,
    pageToken: token
  });

  var items = (resp.courseWork || []).map(courseWorkToClient_);
  if (options.returnMeta) {
    return {
      items: items,
      nextPageToken: resp.nextPageToken || null
    };
  }
  return items;
}

function classCodeForCourseIdServer_(courseId) {
  var wanted = String(courseId || '').trim();
  if (!wanted) return '';
  var rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].CourseId || '') === wanted) {
      return extractClassCodeServer_(rows[i].CourseName || '');
    }
  }
  return '';
}

function apiListAnnouncements(courseId) {
  var out = [];
  var token = null;
  do {
    var resp = Classroom.Courses.Announcements.list(String(courseId), { pageSize: 100, pageToken: token });
    (resp.announcements || []).forEach(function (a) {
      out.push({
        id: String(a.id || ''),
        text: String(a.text || ''),
        creationTime: String(a.creationTime || ''),
        updateTime: String(a.updateTime || ''),
        state: String(a.state || ''),
        attachments: materialsToClient_(a.materials || [])
      });
    });
    token = resp.nextPageToken;
  } while (token);
  return out;
}

function apiListCourseWorkMaterials(courseId) {
  return apiListCourseWork(courseId, { maxItems: 100 })
    .filter(function (it) {
      return String(it.workType || '').toUpperCase() === 'MATERIAL' || ((it.attachments || []).length > 0);
    });
}

function apiGetCourseWorkForReuse(courseId, courseWorkId) {
  var cw = Classroom.Courses.CourseWork.get(String(courseId), String(courseWorkId));
  return {
    id: String(cw.id || ''),
    title: String(cw.title || ''),
    description: String(cw.description || ''),
    maxPoints: cw.maxPoints,
    materials: cw.materials || []
  };
}

function apiGetDriveFileMeta(fileId) {
  var f = DriveApp.getFileById(String(fileId));
  return {
    id: String(f.getId()),
    name: String(f.getName() || ''),
    title: String(f.getName() || ''),
    thumbnailLink: '',
    alternateLink: String(f.getUrl() || '')
  };
}

function apiUploadFile(fileName, base64, mimeType) {
  var bytes = Utilities.base64Decode(String(base64 || ''));
  var blob = Utilities.newBlob(bytes, String(mimeType || MimeType.PLAIN_TEXT), String(fileName || 'upload.bin'));
  var f = DriveApp.createFile(blob);
  return String(f.getId());
}

function apiCreateAssignmentImmediate(courseId, draft) {
  var cw = buildCourseworkBodyFromDraft_(draft || {}, false);
  var created = Classroom.Courses.CourseWork.create(cw, String(courseId));
  return { id: String(created.id || ''), state: String(created.state || '') };
}

function apiCreateAssignmentScheduled(courseId, draft) {
  var cw = buildCourseworkBodyFromDraft_(draft || {}, true);
  var created = Classroom.Courses.CourseWork.create(cw, String(courseId));
  return { id: String(created.id || ''), state: String(created.state || ''), scheduledTime: String(created.scheduledTime || '') };
}

function apiUpdateCourseWork(courseId, courseWorkId, draft) {
  var body = buildCourseworkBodyFromDraft_(draft || {}, true);
  var mask = [];
  ['title', 'description', 'topicId', 'maxPoints', 'dueDate', 'dueTime', 'scheduledTime', 'materials'].forEach(function (k) {
    if (body[k] !== undefined) mask.push(k);
  });
  var updated = Classroom.Courses.CourseWork.patch(body, String(courseId), String(courseWorkId), {
    updateMask: mask.join(',')
  });
  return { success: true, id: String(updated.id || ''), state: String(updated.state || '') };
}

function apiDeleteCourseWork(courseId, courseWorkId) {
  Classroom.Courses.CourseWork.remove(String(courseId), String(courseWorkId));
  return { success: true };
}

function apiListStudentSubmissions(courseId, courseWorkId, userId) {
  var out = [];
  try {
    var token = null;
    do {
      var resp = Classroom.Courses.CourseWork.StudentSubmissions.list(String(courseId), String(courseWorkId), {
        pageSize: 100,
        pageToken: token,
        userId: userId ? String(userId) : undefined
      });
      (resp.studentSubmissions || []).forEach(function (s) {
        out.push({
          id: String(s.id || ''),
          userId: String(s.userId || ''),
          state: String(s.state || ''),
          late: !!s.late,
          updateTime: String(s.updateTime || ''),
          attachments: materialsToClient_((s.assignmentSubmission && s.assignmentSubmission.attachments) || [])
        });
      });
      token = resp.nextPageToken;
    } while (token);
  } catch (e) {
    // Draft/material items or inaccessible coursework can throw precondition errors.
    return [];
  }
  return out;
}

function apiGetSubmissionStats(courseId, courseWorkId) {
  var subs = apiListStudentSubmissions(courseId, courseWorkId);
  var assigned = subs.length;
  var submitted = subs.filter(function (s) {
    var st = String(s.state || '').toUpperCase();
    return st === 'TURNED_IN' || st === 'RETURNED';
  }).length;
  var late = subs.filter(function (s) { return !!s.late; }).length;
  var missing = Math.max(0, assigned - submitted);
  return {
    assigned: assigned,
    submitted: submitted,
    late: late,
    missing: missing
  };
}

function apiGetSubmissionDetails(courseId, courseWorkId) {
  var subs = apiListStudentSubmissions(courseId, courseWorkId);
  var missing = [];
  var late = [];
  subs.forEach(function (s) {
    var st = String(s.state || '').toUpperCase();
    var name = 'Student ' + String(s.userId || '');
    try {
      var profile = Classroom.UserProfiles.get(String(s.userId));
      if (profile && profile.name) name = String(profile.name.fullName || name);
    } catch (e) {}

    if (st !== 'TURNED_IN' && st !== 'RETURNED') {
      missing.push({ userId: s.userId, name: name });
    }
    if (s.late) {
      late.push({ userId: s.userId, name: name, lateText: 'Late' });
    }
  });

  return {
    missingStudents: missing,
    lateStudents: late,
    missing: missing.length,
    late: late.length,
    submitted: subs.length - missing.length,
    assigned: subs.length
  };
}

function apiGetTimetable(classCode) {
  ensureSchema_();
  var byClass = buildTimetableByClass_();
  if (classCode) {
    var key = canonicalClassCodeServer_(classCode);
    var single = {};
    single[key] = byClass[key] || [];
    return { success: true, timetable: single };
  }
  return { success: true, timetable: byClass };
}

function apiImportTimetable(csvText, replaceAll) {
  ensureSchema_();
  var lines = String(csvText || '').split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  var parsed = [];
  lines.forEach(function (line) {
    var cols = line.split(',').map(function (x) { return String(x || '').trim(); });
    if (cols.length < 7) return;
    if (String(cols[0]).toLowerCase() === 'classcode') return;
    parsed.push({
      classCode: canonicalClassCodeServer_(cols[0]),
      day: Number(cols[1] || 0),
      period: Number(cols[2] || 0),
      startTime: normalizeHHMMServer_(cols[3]),
      endTime: normalizeHHMMServer_(cols[4]),
      slotType: String(cols[5] || 'CLASS').toUpperCase(),
      activity: String(cols[6] || ''),
      periodLabel: String(cols[7] || '')
    });
  });

  var rows = replaceAll
    ? []
    : getSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE);

  parsed.forEach(function (p) {
    if (!p.classCode || !p.day || !p.startTime || !p.endTime) return;
    rows.push({
      CourseId: p.classCode,
      DayOfWeek: String(p.day),
      PeriodCode: p.period ? ('P' + p.period) : p.slotType,
      StartTime: p.startTime,
      EndTime: p.endTime
    });
  });

  setSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE, rows);
  return { success: true, imported: parsed.length, errors: [] };
}

function apiGetTimetableSlotForDate(classCode, isoDate) {
  var key = canonicalClassCodeServer_(classCode);
  var byClass = buildTimetableByClass_();
  var slots = byClass[key] || [];
  if (!slots.length) return { success: false, error: 'No timetable configured for class.' };

  var d = new Date(String(isoDate) + 'T00:00:00');
  if (isNaN(d.getTime())) return { success: false, error: 'Invalid date.' };
  var weekday = d.getDay() === 0 ? 7 : d.getDay();

  var daySlots = slots.filter(function (s) {
    return dayToWeekdayServer_(s.day) === weekday;
  }).sort(function (a, b) {
    return String(a.startTime || '').localeCompare(String(b.startTime || ''));
  });
  if (!daySlots.length) return { success: false, error: 'No slot for selected date.' };

  var first = daySlots[0];
  return {
    success: true,
    date: formatDate_(d),
    time: first.startTime,
    period: Number(first.period || 0),
    topicText: ''
  };
}

function apiGetNextAvailable(classCode, nowMs) {
  var key = canonicalClassCodeServer_(classCode);
  var byClass = buildTimetableByClass_();
  var slots = byClass[key] || [];
  if (!slots.length) return { success: false, error: 'No timetable configured.' };

  var cur = nowMs ? new Date(Number(nowMs)) : new Date();
  if (isNaN(cur.getTime())) cur = new Date();
  cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate());

  for (var i = 0; i < 365; i++) {
    var d = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + i);
    var weekday = d.getDay() === 0 ? 7 : d.getDay();
    var daySlots = slots.filter(function (s) {
      return dayToWeekdayServer_(s.day) === weekday;
    }).sort(function (a, b) {
      return String(a.startTime || '').localeCompare(String(b.startTime || ''));
    });
    if (!daySlots.length) continue;

    var first = daySlots[0];
    return {
      success: true,
      date: formatDate_(d),
      time: first.startTime,
      period: Number(first.period || 0),
      topicText: ''
    };
  }

  return { success: false, error: 'No available slot found.' };
}

function apiGetSequence(classCode) {
  ensureSchema_();
  var key = canonicalClassCodeServer_(classCode);
  var mappedCourseId = courseIdFromClassCodeServer_(key);
  syncSequenceSheetFromClassroom_(mappedCourseId, key);

  var all = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
  var rows = all
    .filter(function (r) {
      var cid = String(r.CourseId || '');
      return cid === key || (mappedCourseId && cid === mappedCourseId);
    })
    .map(normalizeSequenceRow_)
    .sort(compareLessons_);

  var timetable = buildTimetableByClass_();
  var classSlots = timetable[key] || timetable[canonicalClassCodeServer_(mappedCourseId)] || [];

  var lessons = rows.map(function (r, i) {
    var t = timeForLessonServer_(classSlots, r.lessonDate, r.periodCode);
    return {
      id: String(r.lessonDate + '_' + r.periodCode + '_' + (i + 1)),
      title: String(r.title || r.topic || ''),
      description: String(r.description || ''),
      assignDate: String(r.lessonDate || ''),
      assignTime: String(t || ''),
      dueDate: '',
      dueTime: '',
      classworkId: String(r.classroomCourseWorkId || ''),
      status: String(r.status || 'Draft'),
      position: i + 1,
      materials: r.driveFileId ? [{ driveFile: { driveFile: { id: r.driveFileId }, shareMode: 'VIEW' } }] : []
    };
  });

  return { success: true, lessons: lessons };
}

function apiSaveLesson(classCode, payload) {
  try {
    var data = payload || {};
    var courseId = String(data.courseId || '') || courseIdFromClassCodeServer_(classCode);
    if (!courseId) throw new Error('No mapped courseId for class: ' + classCode);

    var date = String(data.assignDate || '').trim();
    if (!date) throw new Error('assignDate is required');

    var res = addToSequence({
      courseId: courseId,
      customDate: date,
      title: String(data.title || 'Lesson'),
      description: String(data.description || '')
    });

    return {
      success: !!(res && res.success),
      lesson: res && res.data && res.data.insertedSlot ? {
        assignDate: res.data.insertedSlot.date,
        assignTime: String(data.assignTime || ''),
        classworkId: ''
      } : null,
      warnings: [],
      processLog: ['1) Payload received', '2) Sequence cascade applied', '3) Classroom sync complete']
    };
  } catch (e) {
    return { success: false, error: String(e && e.message || e), processLog: [] };
  }
}

function apiDeleteLesson(classCode, lessonId) {
  try {
    var seq = apiGetSequence(classCode);
    if (!seq || !seq.success) return { success: false, error: 'Sequence not available.' };
    var target = (seq.lessons || []).find(function (l) { return String(l.id || '') === String(lessonId || ''); });
    if (!target) return { success: false, error: 'Lesson not found.' };

    var mappedCourseId = courseIdFromClassCodeServer_(classCode);
    var all = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    var kept = all.filter(function (r) {
      var cid = String(r.CourseId || '');
      var matchCourse = (cid === mappedCourseId || cid === canonicalClassCodeServer_(classCode));
      if (!matchCourse) return true;
      return !(String(r.LessonDate || '') === String(target.assignDate || '') && String(r.ClassroomCourseWorkId || '') === String(target.classworkId || ''));
    });
    setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, kept);

    if (target.classworkId && mappedCourseId) {
      try { Classroom.Courses.CourseWork.remove(String(mappedCourseId), String(target.classworkId)); } catch (e) {}
    }

    return { success: true, message: 'Lesson deleted.' };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function apiUpdateSequenceLessonByClasswork(courseId, classworkId, payload) {
  try {
    ensureSchema_();
    var data = payload || {};
    var mappedCourseId = String(courseId || data.courseId || '').trim();
    var cwId = String(classworkId || data.classworkId || '').trim();
    if (!mappedCourseId) throw new Error('Missing courseId for sequence update.');
    var existingAssignDate = String(data.existingAssignDate || '').trim();
    var existingAssignTime = normalizeHHMMFromAnyServer_(String(data.existingAssignTime || '').trim()) || '';
    var previousTitle = String(data.previousTitle || '').trim();
    var parsedSeqId = parseSequenceUiIdServer_(String(data.id || '').trim());
    var existingPeriodCode = normalizePeriodCodeServer_(String(data.existingPeriodCode || parsedSeqId.periodCode || '').trim());

    if (cwId) {
      var body = {};
      if (data.title !== undefined) body.title = String(data.title || '').trim() || '(Untitled)';
      if (data.description !== undefined) body.description = String(data.description || '');
      if (data.points !== undefined) body.maxPoints = Number(data.points || 0);
      if (data.assignDate) {
        var hhmm = normalizeHHMMServer_(String(data.assignTime || '00:00'));
        var scheduled = new Date(String(data.assignDate) + 'T' + hhmm + ':00');
        if (!isNaN(scheduled.getTime())) body.scheduledTime = scheduled.toISOString();
      }
      if (data.dueDate) {
        var dparts = String(data.dueDate).split('-');
        body.dueDate = { year: Number(dparts[0] || 0), month: Number(dparts[1] || 0), day: Number(dparts[2] || 0) };
        if (data.dueTime) {
          var tparts = String(data.dueTime).split(':');
          body.dueTime = { hours: Number(tparts[0] || 0), minutes: Number(tparts[1] || 0) };
        } else {
          body.dueTime = null;
        }
      } else {
        body.dueDate = null;
        body.dueTime = null;
      }
      if (Array.isArray(data.materials)) {
        body.materials = data.materials.map(function (m) {
          var x = JSON.parse(JSON.stringify(m || {}));
          if (x.youTubeVideo && !x.youtubeVideo) {
            x.youtubeVideo = x.youTubeVideo;
            delete x.youTubeVideo;
          }
          return x;
        });
      }
      if (data.topicText) {
        var topicResp = apiEnsureTopic(mappedCourseId, data.topicText);
        body.topicId = topicResp && topicResp.topicId ? topicResp.topicId : null;
      }

      var mask = [];
      ['title', 'description', 'maxPoints', 'scheduledTime', 'dueDate', 'dueTime', 'materials', 'topicId'].forEach(function (k) {
        if (body[k] !== undefined) mask.push(k);
      });
      if (mask.length) {
        Classroom.Courses.CourseWork.patch(body, mappedCourseId, cwId, { updateMask: mask.join(',') });
      }
    }

    var all = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    var classCode = canonicalClassCodeServer_(extractClassCodeServer_(String(data.classCode || ''))) || canonicalClassCodeServer_(extractClassCodeServer_(mappedCourseId));
    var targetDate = String(data.assignDate || '').trim();
    var targetTitle = String(data.title || '').trim();
    var found = false;
    var fallbackMatchIndex = -1;
    var fallbackScore = -1;

    all.forEach(function (r, idx) {
      var cid = String(r.CourseId || '');
      var sameCourse = cid === mappedCourseId || (classCode && canonicalClassCodeServer_(cid) === classCode);
      if (!sameCourse) return;
      var rowCwId = String(r.ClassroomCourseWorkId || '').trim();
      var rowDate = String(asDateString_(r.LessonDate) || '').trim();
      var rowPeriod = normalizePeriodCodeServer_(String(r.PeriodCode || '').trim());

      if (cwId && rowCwId === cwId) {
        if (targetDate) r.LessonDate = targetDate;
        if (targetTitle) r.Title = targetTitle;
        if (data.description !== undefined) r.Description = String(data.description || '');
        if (data.topicText) r.Topic = String(data.topicText || '');
        if (data.status) r.Status = String(data.status || 'scheduled');
        if (!r.PeriodCode) r.PeriodCode = 'P1';
        if (r.LessonDate) {
          var d = new Date(String(r.LessonDate) + 'T00:00:00');
          r.DayOfWeek = isNaN(d.getTime()) ? String(r.DayOfWeek || '') : APP_CONFIG.DAYS[d.getDay()];
        }
        if (!r.LessonNumber) {
          var mm = String(r.Title || '').match(/Lesson\s*(\d+)/i);
          r.LessonNumber = mm ? Number(mm[1] || 0) : Number(r.LessonNumber || 0);
        }
        r.UpdatedAt = new Date();
        r.CourseId = mappedCourseId;
        found = true;
        return;
      }

      if (cwId) return;

      var score = 0;
      var prevTitleNorm = previousTitle.toLowerCase();
      var rowTitleNorm = String(r.Title || '').trim().toLowerCase();
      if (existingAssignDate && rowDate === existingAssignDate) score += 4;
      if (parsedSeqId.lessonDate && rowDate === parsedSeqId.lessonDate) score += 3;
      if (existingPeriodCode && rowPeriod && rowPeriod === existingPeriodCode) score += 2;
      if (prevTitleNorm && rowTitleNorm === prevTitleNorm) score += 2;
      if (!rowCwId) score += 1;

      if (score > fallbackScore) {
        fallbackScore = score;
        fallbackMatchIndex = idx;
      }
    });

    if (!found && !cwId && fallbackMatchIndex >= 0 && fallbackScore >= 4) {
      var rr = all[fallbackMatchIndex];
      if (targetDate) rr.LessonDate = targetDate;
      if (targetTitle) rr.Title = targetTitle;
      if (data.description !== undefined) rr.Description = String(data.description || '');
      if (data.topicText) rr.Topic = String(data.topicText || '');
      if (data.status) rr.Status = String(data.status || 'scheduled');
      if (!rr.PeriodCode) {
        rr.PeriodCode = existingPeriodCode || inferPeriodCodeForCourseDateTimeServer_(mappedCourseId, rr.LessonDate, existingAssignTime || '00:00') || 'P1';
      }
      if (rr.LessonDate) {
        var dd2 = new Date(String(rr.LessonDate) + 'T00:00:00');
        rr.DayOfWeek = isNaN(dd2.getTime()) ? String(rr.DayOfWeek || '') : APP_CONFIG.DAYS[dd2.getDay()];
      }
      if (!rr.LessonNumber) {
        var mm2 = String(rr.Title || '').match(/Lesson\s*(\d+)/i);
        rr.LessonNumber = mm2 ? Number(mm2[1] || 0) : Number(rr.LessonNumber || 0);
      }
      rr.CourseId = mappedCourseId;
      rr.UpdatedAt = new Date();
      found = true;
    }

    if (!found) {
      var lessonNum = 0;
      var mNum = targetTitle.match(/Lesson\s*(\d+)/i);
      if (mNum) lessonNum = Number(mNum[1] || 0);
      var dateForRow = targetDate || formatDate_(new Date());
      var dd = new Date(String(dateForRow) + 'T00:00:00');
      all.push({
        CourseId: mappedCourseId,
        LessonNumber: lessonNum,
        LessonDate: dateForRow,
        DayOfWeek: isNaN(dd.getTime()) ? '' : APP_CONFIG.DAYS[dd.getDay()],
        PeriodCode: 'P1',
        Topic: String(data.topicText || ''),
        Title: targetTitle || '(Untitled)',
        Description: String(data.description || ''),
        ClassroomCourseWorkId: cwId,
        DriveFileId: '',
        Status: String(data.status || 'scheduled'),
        UpdatedAt: new Date()
      });
    }

    setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, all);
    try {
      syncSequenceSheetFromClassroom_(mappedCourseId, classCode, {});
    } catch (syncErr) {}
    return { success: true, classworkId: cwId };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function parseSequenceUiIdServer_(id) {
  var raw = String(id || '').trim();
  if (!raw) return { lessonDate: '', periodCode: '', index: 0 };
  var parts = raw.split('_');
  if (parts.length < 3) return { lessonDate: '', periodCode: '', index: 0 };
  var date = isIsoDateServer_(parts[0]) ? parts[0] : '';
  var period = normalizePeriodCodeServer_(parts[1] || '');
  var idx = Number(parts[2] || 0);
  return { lessonDate: date, periodCode: period, index: isNaN(idx) ? 0 : idx };
}

function syncSequenceSheetFromClassroom_(courseId, classCodeKey, opts) {
  try {
    var options = opts || {};
    var strictClassroom = !!options.strictClassroom;
    var cid = String(courseId || '').trim();
    if (!cid) return;

    var allCourseWork = [];
    var token = null;
    do {
      var resp = Classroom.Courses.CourseWork.list(cid, { pageSize: 100, pageToken: token });
      (resp.courseWork || []).forEach(function (cw) {
        allCourseWork.push(cw);
      });
      token = resp.nextPageToken;
    } while (token);

    if (!allCourseWork.length) return;
    allCourseWork.sort(function (a, b) {
      var at = new Date(String(a.scheduledTime || a.creationTime || a.updateTime || ''));
      var bt = new Date(String(b.scheduledTime || b.creationTime || b.updateTime || ''));
      return at.getTime() - bt.getTime();
    });

    var rows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES);
    var cwIdSet = {};
    allCourseWork.forEach(function (cw) {
      var id = String(cw.id || '').trim();
      if (id) cwIdSet[id] = true;
    });

    // Remove stale rows for this course/class that no longer exist in Classroom.
    var canonicalClass = canonicalClassCodeServer_(classCodeKey || '');
    var keptRows = [];
    rows.forEach(function (r) {
      var rowCourse = String(r.CourseId || '');
      var sameCourse = rowCourse === cid || (canonicalClass && canonicalClassCodeServer_(rowCourse) === canonicalClass);
      if (!sameCourse) {
        keptRows.push(r);
        return;
      }
      var rowCwId = String(r.ClassroomCourseWorkId || '').trim();
      if (rowCwId && cwIdSet[rowCwId]) {
        keptRows.push(r);
      } else if (!rowCwId && !strictClassroom) {
        // Non-strict mode keeps manual/unlinked rows with metadata.
        if (String(r.LessonDate || '').trim() && String(r.Title || '').trim()) keptRows.push(r);
      }
    });
    rows = keptRows;
    var indexByCwId = {};
    rows.forEach(function (r, i) {
      var x = String(r.ClassroomCourseWorkId || '').trim();
      if (x && indexByCwId[x] === undefined) indexByCwId[x] = i;
    });

    var dayCounters = {};
    var changed = false;
    allCourseWork.forEach(function (cw) {
      var cwId = String(cw.id || '').trim();
      if (!cwId) return;
      var ts = String(cw.scheduledTime || cw.creationTime || cw.updateTime || '').trim();
      var dt = ts ? new Date(ts) : new Date();
      if (isNaN(dt.getTime())) dt = new Date();
      var dateStr = formatDate_(dt);
      var dayKey = dateStr;
      dayCounters[dayKey] = (dayCounters[dayKey] || 0) + 1;

      var idx = indexByCwId[cwId];
      var row = idx !== undefined ? rows[idx] : null;
      if (!row) {
        row = {
          CourseId: cid,
          LessonNumber: 0,
          LessonDate: dateStr,
          DayOfWeek: APP_CONFIG.DAYS[dt.getDay()],
          PeriodCode: 'P' + String(dayCounters[dayKey]),
          Topic: '',
          Title: '',
          Description: '',
          ClassroomCourseWorkId: cwId,
          DriveFileId: '',
          Status: '',
          UpdatedAt: new Date()
        };
        rows.push(row);
        indexByCwId[cwId] = rows.length - 1;
        changed = true;
      }

      var newTitle = String(cw.title || row.Title || '');
      var newDescription = String(cw.description || row.Description || '');
      var newStatus = String(cw.state || row.Status || '');
      var newDate = String(dateStr || row.LessonDate || '');
      var oldTuple = [row.Title, row.Description, row.Status, row.LessonDate].join('|');
      row.CourseId = cid;
      row.ClassroomCourseWorkId = cwId;
      row.Title = newTitle;
      row.Description = newDescription;
      row.Status = newStatus;
      row.LessonDate = newDate;
      row.DayOfWeek = APP_CONFIG.DAYS[new Date(newDate + 'T00:00:00').getDay()];
      if (!row.PeriodCode) row.PeriodCode = 'P' + String(dayCounters[dayKey]);
      var m = newTitle.match(/Lesson\s*(\d+)/i);
      if (m) row.LessonNumber = Number(m[1] || 0);
      row.UpdatedAt = new Date();
      var newTuple = [row.Title, row.Description, row.Status, row.LessonDate].join('|');
      if (oldTuple !== newTuple) changed = true;
    });

    if (renumberSequenceRowsForCourse_(rows, cid, canonicalClass)) changed = true;
    if (changed) {
      setSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES, rows);
      patchClassroomTitlesFromSequenceRows_(rows, cid, canonicalClass);
    }
  } catch (e) {
    // Non-fatal sync: keep serving sequence data from sheet even if Classroom sync fails.
  }
}

function apiForceSynchronizeAllSequences() {
  try {
    ensureSchema_();
    syncCoursesFromClassroom_();

    var courseRows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES)
      .filter(function (r) { return toBoolean_(r.Active, true); });

    var report = [];
    courseRows.forEach(function (r) {
      var courseId = String(r.CourseId || '').trim();
      if (!courseId) return;
      var classCode = extractClassCodeServer_(String(r.CourseName || ''));
      syncSequenceSheetFromClassroom_(courseId, classCode, { strictClassroom: true });
      var integrity = apiDiagnoseSequenceIntegrity(classCode);
      report.push({
        courseId: courseId,
        classCode: classCode,
        warnings: (integrity && integrity.warnings) ? integrity.warnings.length : 0
      });
    });

    return {
      success: true,
      processed: report.length,
      report: report
    };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function apiBumpDateOnly() {
  return { success: true, movedCount: 0, processLog: ['1) Compatibility mode: no-op bump'], warnings: [] };
}

function apiAutoMapClasses() {
  return { success: true, matched: 0, message: 'Using derived class code mapping from course names.' };
}

function apiGetTimetableRundown(classCode) {
  var seq = apiGetSequence(classCode);
  var lessons = (seq && seq.success && Array.isArray(seq.lessons)) ? seq.lessons : [];
  return {
    term: 1,
    termEnd: '',
    totalSlots: lessons.length,
    remainingCount: lessons.length,
    occurrences: lessons.slice(0, 100).map(function (l) {
      return {
        date: l.assignDate || '',
        time: l.assignTime || '',
        title: l.title || '',
        day: 1,
        week: 1,
        period: 1
      };
    })
  };
}

function apiDiagnoseSequenceIntegrity(classCode) {
  try {
    ensureSchema_();
    var key = canonicalClassCodeServer_(classCode);
    var mappedCourseId = courseIdFromClassCodeServer_(key);
    if (!mappedCourseId) {
      return {
        success: true,
        classCode: key,
        mappedCourseId: '',
        blockers: ['No course mapping found for class code.'],
        warnings: [],
        stats: { classroomItems: 0, sequenceRows: 0 }
      };
    }

    var byClass = buildTimetableByClass_();
    var slots = byClass[key] || byClass[canonicalClassCodeServer_(mappedCourseId)] || [];
    var blockers = [];
    var warnings = [];
    if (!slots.length) blockers.push('No timetable slots found for this class.');

    var allCourseWork = [];
    var token = null;
    do {
      var resp = Classroom.Courses.CourseWork.list(String(mappedCourseId), { pageSize: 100, pageToken: token });
      (resp.courseWork || []).forEach(function (cw) { allCourseWork.push(cw); });
      token = resp.nextPageToken;
    } while (token);

    var rows = getSheetObjects_(APP_CONFIG.SHEETS.SEQUENCES, APP_CONFIG.HEADERS.SEQUENCES).filter(function (r) {
      var cid = String(r.CourseId || '');
      return cid === String(mappedCourseId) || cid === key;
    });

    var rowByCw = {};
    rows.forEach(function (r) {
      var id = String(r.ClassroomCourseWorkId || '').trim();
      if (!id) return;
      if (rowByCw[id]) warnings.push('Duplicate sequence rows for Classroom item ' + id + '.');
      rowByCw[id] = r;
    });

    var cwIdSet = {};
    allCourseWork.forEach(function (cw) {
      var id = String(cw.id || '').trim();
      if (id) cwIdSet[id] = true;
      if (!rowByCw[id]) warnings.push('Classroom item missing from Sequences sheet: ' + id);
    });

    rows.forEach(function (r) {
      var id = String(r.ClassroomCourseWorkId || '').trim();
      if (id && !cwIdSet[id]) warnings.push('Sequence row points to deleted/missing Classroom item: ' + id);
      if (!String(r.LessonDate || '').trim()) warnings.push('Sequence row missing LessonDate for Classroom item: ' + (id || '(none)'));
      if (!String(r.PeriodCode || '').trim()) warnings.push('Sequence row missing PeriodCode for Classroom item: ' + (id || '(none)'));
    });

    // Check sequential lesson numbering in Classroom titles.
    var sortedCw = allCourseWork.slice().sort(function (a, b) {
      var at = new Date(String(a.scheduledTime || a.creationTime || a.updateTime || ''));
      var bt = new Date(String(b.scheduledTime || b.creationTime || b.updateTime || ''));
      return at.getTime() - bt.getTime();
    });
    var expected = 1;
    sortedCw.forEach(function (cw) {
      var title = String(cw.title || '');
      var m = title.match(/Lesson\s*(\d+)/i);
      var got = m ? Number(m[1] || 0) : 0;
      if (got && got !== expected) warnings.push('Non-sequential lesson title in Classroom: expected Lesson ' + expected + ', found "' + title + '".');
      if (!got) warnings.push('Classroom item missing "Lesson N" prefix: "' + title + '".');
      expected++;
    });

    return {
      success: true,
      classCode: key,
      mappedCourseId: mappedCourseId,
      blockers: blockers,
      warnings: warnings,
      stats: {
        classroomItems: allCourseWork.length,
        sequenceRows: rows.length,
        timetableSlots: slots.length
      }
    };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function apiForceSynchronizeSequence(classCode) {
  try {
    ensureSchema_();
    var key = canonicalClassCodeServer_(classCode);
    var mappedCourseId = courseIdFromClassCodeServer_(key);
    if (!mappedCourseId) return { success: false, error: 'No mapped course for class code: ' + key };

    syncSequenceSheetFromClassroom_(mappedCourseId, key);
    var report = apiDiagnoseSequenceIntegrity(key);
    return {
      success: true,
      classCode: key,
      mappedCourseId: mappedCourseId,
      report: report
    };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

/* -------- Helpers for compatibility layer -------- */

function courseWorkToClient_(cw) {
  return {
    id: String(cw.id || ''),
    title: String(cw.title || ''),
    description: String(cw.description || ''),
    text: String(cw.description || ''),
    state: String(cw.state || ''),
    workType: String(cw.workType || ''),
    maxPoints: cw.maxPoints,
    dueDate: cw.dueDate || null,
    dueTime: cw.dueTime || null,
    topicId: cw.topicId || null,
    creationTime: String(cw.creationTime || ''),
    updateTime: String(cw.updateTime || ''),
    scheduledTime: String(cw.scheduledTime || ''),
    attachments: materialsToClient_(cw.materials || [])
  };
}

function materialsToClient_(materials) {
  return (materials || []).map(function (m) {
    if (m.driveFile && m.driveFile.driveFile) {
      var df = m.driveFile.driveFile;
      return {
        type: 'DRIVE_FILE',
        title: String(df.title || 'Drive file'),
        url: String(df.alternateLink || ''),
        thumbnailUrl: String(df.thumbnailUrl || '')
      };
    }
    if (m.youTubeVideo || m.youtubeVideo) {
      var y = m.youTubeVideo || m.youtubeVideo;
      var id = String(y.id || '');
      return {
        type: 'YOUTUBE',
        title: String(y.title || 'YouTube'),
        url: id ? ('https://www.youtube.com/watch?v=' + id) : String(y.alternateLink || ''),
        thumbnailUrl: String(y.thumbnailUrl || '')
      };
    }
    if (m.link) {
      return {
        type: 'LINK',
        title: String(m.link.title || m.link.url || 'Link'),
        url: String(m.link.url || ''),
        thumbnailUrl: String(m.link.thumbnailUrl || '')
      };
    }
    if (m.form) {
      return {
        type: 'FORM',
        title: String(m.form.title || 'Form'),
        url: String(m.form.formUrl || ''),
        thumbnailUrl: String(m.form.thumbnailUrl || '')
      };
    }
    return null;
  }).filter(function (x) { return !!x; });
}

function buildCourseworkBodyFromDraft_(draft, includeScheduled) {
  var body = {
    title: String(draft.title || '(Untitled)'),
    description: String(draft.descriptionUnicode || draft.description || ''),
    workType: 'ASSIGNMENT',
    state: String(draft.state || 'DRAFT')
  };
  if (draft.topicId !== undefined) body.topicId = draft.topicId || null;
  if (draft.maxPoints !== undefined && draft.maxPoints !== null && draft.maxPoints !== '') body.maxPoints = Number(draft.maxPoints);
  if (draft.dueDate !== undefined) body.dueDate = draft.dueDate || null;
  if (draft.dueTime !== undefined) body.dueTime = draft.dueTime || null;
  if (includeScheduled && draft.scheduledTime) body.scheduledTime = String(draft.scheduledTime);
  if (Array.isArray(draft.materials) && draft.materials.length) {
    body.materials = draft.materials.map(function (m) {
      var x = JSON.parse(JSON.stringify(m || {}));
      if (x.youTubeVideo && !x.youtubeVideo) {
        x.youtubeVideo = x.youTubeVideo;
        delete x.youTubeVideo;
      }
      return x;
    });
  }
  return body;
}

function extractClassCodeServer_(name) {
  var text = String(name || '').replace(/[\u2014\u2013-]/g, ' ');
  var tokens = text.split(/\s+/).map(function (t) { return t.trim(); }).filter(Boolean);
  var cleaned = tokens.map(function (t) { return t.replace(/[^A-Za-z0-9._-]/g, ''); }).filter(Boolean);
  var strong = cleaned.find(function (t) {
    return /[A-Za-z]/.test(t) && /\d/.test(t);
  });
  return canonicalClassCodeServer_(strong || cleaned[0] || '');
}

function canonicalClassCodeServer_(code) {
  return String(code || '').trim().toUpperCase().replace(/\s+/g, '');
}

function courseIdFromClassCodeServer_(classCode) {
  var wanted = canonicalClassCodeServer_(classCode);
  if (!wanted) return '';
  var rows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  for (var i = 0; i < rows.length; i++) {
    var code = extractClassCodeServer_(rows[i].CourseName || '');
    if (code === wanted) return String(rows[i].CourseId || '');
  }
  return '';
}

function buildTimetableByClass_() {
  var rows = getSheetObjects_(APP_CONFIG.SHEETS.TIMETABLE, APP_CONFIG.HEADERS.TIMETABLE);
  var courseRows = getSheetObjects_(APP_CONFIG.SHEETS.COURSES, APP_CONFIG.HEADERS.COURSES);
  var courseIdToClass = {};
  courseRows.forEach(function (r) {
    courseIdToClass[String(r.CourseId || '')] = extractClassCodeServer_(r.CourseName || '');
  });

  var out = {};
  rows.forEach(function (r) {
    var rawCourse = String(r.CourseId || '');
    var classCode = canonicalClassCodeServer_(courseIdToClass[rawCourse] || rawCourse);
    if (!classCode) return;

    var dayRaw = String(r.DayOfWeek || '').trim();
    var dayNum = Number(dayRaw);
    if (!dayNum) {
      var d = dayNameToIndex_(dayRaw);
      dayNum = (d >= 1 && d <= 5) ? d : 1;
    }

    var periodMatch = String(r.PeriodCode || '').match(/(\d+)/);
    var period = periodMatch ? Number(periodMatch[1]) : 0;

    if (!out[classCode]) out[classCode] = [];
    out[classCode].push({
      classCode: classCode,
      day: dayNum,
      period: period,
      startTime: normalizeHHMMServer_(String(r.StartTime || '00:00')),
      endTime: normalizeHHMMServer_(String(r.EndTime || '00:00')),
      slotType: period > 0 ? 'CLASS' : String(r.PeriodCode || 'OTHER').toUpperCase(),
      activity: '',
      periodLabel: period > 0 ? ('Period ' + period) : String(r.PeriodCode || 'Slot')
    });
  });

  Object.keys(out).forEach(function (k) {
    out[k].sort(function (a, b) {
      if (Number(a.day) !== Number(b.day)) return Number(a.day) - Number(b.day);
      return String(a.startTime || '').localeCompare(String(b.startTime || ''));
    });
  });
  return out;
}

function dayToWeekdayServer_(day) {
  var d = Number(day || 0);
  if (d >= 1 && d <= 5) return d;
  if (d >= 6 && d <= 10) return d - 5;
  return 0;
}

function normalizeHHMMServer_(value) {
  var s = String(value || '').trim();
  var m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!m) return '00:00';
  var hh = Math.max(0, Math.min(23, Number(m[1] || 0)));
  var mm = Math.max(0, Math.min(59, Number(m[2] || 0)));
  return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
}

function timeForLessonServer_(classSlots, lessonDate, periodCode) {
  var d = new Date(String(lessonDate || '') + 'T00:00:00');
  if (isNaN(d.getTime())) return '00:00';
  var weekday = d.getDay() === 0 ? 7 : d.getDay();
  var match = String(periodCode || '').match(/(\d+)/);
  var p = match ? Number(match[1]) : 0;
  var slot = (classSlots || []).find(function (s) {
    return dayToWeekdayServer_(s.day) === weekday && (p ? Number(s.period || 0) === p : true);
  });
  return slot ? String(slot.startTime || '00:00') : '00:00';
}
