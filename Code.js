// Copyright (c) 2025 faetschi

function buildCalendarFromOccurrences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName("Calendar");
  const occSheet = ss.getSheetByName("TaskOccurrences");
  const personsSheet = ss.getSheetByName("Persons");

  // --- Read persons and build a map ---
  const personsData = personsSheet.getRange(2,1,personsSheet.getLastRow()-1,2).getValues();
  const personColors = {};
  personsData.forEach(row => {
    const name = row[0];
    const color = row[1] || "#000000";
    personColors[name] = color;
  });

  const year = cal.getRange("B1").getValue();
  const month = cal.getRange("B2").getValue();

  // Clear previous content
  const lastRow = cal.getMaxRows();
  const lastCol = cal.getMaxColumns();
  cal.getRange(4,1,lastRow-3,lastCol).clearContent().clearFormat();

  const startRow = 4;
  const startCol = 1;

  // Weekday headers
  const weekdays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const weekdayRange = cal.getRange(startRow,startCol,1,7);
  weekdayRange.setValues([weekdays]);
  weekdayRange.setFontWeight("bold").setFontSize(13);

  const firstDay = new Date(year, month-1, 1);
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstWeekDay = firstDay.getDay();

  const occData = occSheet.getRange(2,1,Math.max(occSheet.getLastRow()-1,0),7).getValues();
  const today = new Date();
  today.setHours(0,0,0,0);

  const expandedTasks = occData
    .filter(r => r[6] === true) // active
    .map(r => ({
      taskID: r[0],
      name: r[1],
      due: new Date(r[2]),
      person: r[3] || "",
      done: r[4],
      color: r[5]
    }));

  let day = 1;
  let rowOffset = 1;

  while (day <= daysInMonth) {
    let weekDates = Array(7).fill("");
    let weekTasks = Array(7).fill([]);

    for (let col = 0; col < 7 && day <= daysInMonth; col++) {
      if (day === 1 && col < firstWeekDay) continue; // offset for first week

      const date = new Date(year, month-1, day);
      weekDates[col] = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
      weekTasks[col] = expandedTasks.filter(t => t.due.getTime() === date.getTime());

      day++;
    }

    // Ensure arrays are length 7
    for (let c = 0; c < 7; c++) {
      if (!weekDates[c]) weekDates[c] = "";
      if (!weekTasks[c]) weekTasks[c] = [];
    }

    const weekRange = cal.getRange(startRow + rowOffset, startCol, 1, 7);
    weekRange.setValues([weekDates]).setWrap(true).setNumberFormat('@STRING@');

    const dateRow = startRow + rowOffset;

    // --- Determine today columns ---
    const todayCols = [];
    for (let c = 0; c < 7; c++) {
      if (!weekDates[c]) continue;
      const parts = weekDates[c].split(".");
      const cellDate = new Date(parseInt(parts[2],10), parseInt(parts[1],10)-1, parseInt(parts[0],10));
      if (isSameDate(cellDate, today)) todayCols.push(c);
    }

    rowOffset++; // move to first task row

    const maxTasks = Math.max(...weekTasks.map(t => t.length));
    for (let t = 0; t < maxTasks; t++) {
      const taskRow = Array(7).fill("");
      for (let c = 0; c < 7; c++) {
        if (weekTasks[c].length > t) {
          const task = weekTasks[c][t];
          taskRow[c] = task.name + (task.person ? "\n" + task.person : "");
        }
      }

      const range = cal.getRange(startRow + rowOffset, startCol, 1, 7);
      range.setValues([taskRow]).setWrap(true);

      for (let c = 0; c < 7; c++) {
        if (weekTasks[c].length <= t) continue;
        const task = weekTasks[c][t];
        const cell = cal.getRange(startRow + rowOffset, startCol + c);

        // Background
        const isPast = task.due.getTime() < today.getTime();
        const isToday = isSameDate(task.due, today);
        if (isToday) cell.setBackground(task.done ? "#00FF00" : "#FF0000");
        else if (isPast) cell.setBackground(task.done ? "#FFFFFF" : "#FF0000");
        else cell.setBackground("#FFFFFF");

        // Rich text (always apply task color, optionally apply person)
        const parts = taskRow[c].split("\n");
        if (parts.length === 2) {
          const personColor = personColors[parts[1]] || "#000000";
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(parts[0] + "\n" + parts[1])
            .setTextStyle(0, parts[0].length, SpreadsheetApp.newTextStyle().setForegroundColor(task.color || "#000000").build())
            .setTextStyle(parts[0].length + 1, parts[0].length + 1 + parts[1].length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(personColor).build())
            .build();
          cell.setRichTextValue(richText);
        } else if (parts.length === 1) {
          // No person, just the task
          cell.setFontColor(task.color || "#000000");
        }

        // Border for today
        if (todayCols.includes(c)) {
          cell.setFontWeight("bold")
          cell.setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
      }

      rowOffset++;
    }

    // Add border to date cells for today
    for (const c of todayCols) {
      const dateCell = cal.getRange(dateRow, startCol + c);
      dateCell.setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  ss.toast("Calendar rebuilt", "Task Tools", 3);
}

function generateTaskOccurrences(monthsAhead = 12) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");
  let occSheet = ss.getSheetByName("TaskOccurrences");

  // Create sheet if missing
  if (!occSheet) occSheet = ss.insertSheet("TaskOccurrences");

  // Headers (only set if sheet is empty)
  if (occSheet.getLastRow() === 0) {
    const headers = ["TaskID","Task","DueDate","Person","Done","Color","Active"];
    occSheet.getRange(1,1,1,headers.length).setValues([headers]);
  }

  // --- Read tasks (6 columns now) ---
  const taskData = tasksSheet.getRange(2,1,tasksSheet.getLastRow()-1,6).getValues();
  const today = new Date();
  const endDate = new Date();
  endDate.setMonth(endDate.getMonth() + monthsAhead);

  // Read existing occurrences
  const occLastRow = occSheet.getLastRow();
  let existingOccurrences = {};
  if (occLastRow > 1) {
    const occData = occSheet.getRange(2,1,occLastRow-1,7).getValues();
    occData.forEach(r => {
      const key = r[0] + "_" + new Date(r[2]).toDateString();
      existingOccurrences[key] = {
        TaskID: r[0],
        Task: r[1],
        DueDate: new Date(r[2]),
        Person: r[3],
        Done: r[4],
        Color: r[5],
        Active: r[6]
      };
    });
  }

  const updatedOccurrences = [];

  taskData.forEach(row => {
    const taskID = row[0];
    const taskName = row[1];
    let startDate = row[2] instanceof Date ? row[2] : new Date(row[2]);
    const freq = row[3];
    const color = row[4];
    const active = row[5];

    if (!active) return; // skip inactive tasks

    let current = new Date(startDate);
    while (current <= endDate) {
      const key = taskID + "_" + current.toDateString();

      if (existingOccurrences[key]) {
        // Update columns from Tasks but keep Done and Person
        const occ = existingOccurrences[key];
        occ.Task = taskName;
        occ.Color = color;
        occ.Active = active;
        updatedOccurrences.push([
          occ.TaskID,
          occ.Task,
          new Date(occ.DueDate),
          occ.Person, // preserve
          occ.Done,   // preserve
          occ.Color,
          occ.Active
        ]);
      } else {
        // New occurrence
        updatedOccurrences.push([
          taskID,
          taskName,
          new Date(current),
          "",       // Person empty initially
          false,    // Done
          color,
          true      // Active
        ]);
      }

      current.setDate(current.getDate() + freq);
    }
  });

  // Clear existing sheet but keep headers
  if (occSheet.getLastRow() > 1) {
    occSheet.getRange(2,1,occSheet.getLastRow()-1,7).clearContent();
  }

  // Sort occurrences by date ascending
  updatedOccurrences.sort((a, b) => new Date(a[2]) - new Date(b[2]));

  if (updatedOccurrences.length > 0) {
    occSheet.getRange(2,1,updatedOccurrences.length,7).setValues(updatedOccurrences);
  }
}

function generateTodayTasksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const occSheet = ss.getSheetByName("TaskOccurrences");
  let todaySheet = ss.getSheetByName("TodayTasks");

  if (!todaySheet) todaySheet = ss.insertSheet("TodayTasks");

  // Only clear rows from row 4 down (keep headers)
  todaySheet.getRange(4, 1, todaySheet.getMaxRows()-3, 4).clearContent().clearFormat();

  // Ensure headers are set
  const headers = ["Date", "Task", "Person", "Done"];
  todaySheet.getRange(3, 1, 1, headers.length).setValues([headers]);

  const today = new Date();
  today.setHours(0,0,0,0);

  const occData = occSheet.getRange(2,1,occSheet.getLastRow()-1,7).getValues();

  // Filter for tasks due today and active
  let todayTasks = occData
    .filter(r => r[6] === true) // Active
    .map(r => ({
      taskID: r[0],
      task: r[1],
      due: new Date(r[2]),
      person: r[3],
      done: r[4],
      color: r[5]
    }))
    .filter(t => isSameDate(t.due, today))
    .sort((a,b) => a.task.localeCompare(b.task)); // optional: sort by name

  if (todayTasks.length === 0) return;

  const startRow = 4;
  const output = todayTasks.map(t => [
    Utilities.formatDate(t.due, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy"),
    t.task,
    t.person, // can be empty initially
    t.done
  ]);

  // todaySheet.getRange(startRow, 1, output.length, 1).setNumberFormat('@STRING@');
  todaySheet.getRange(startRow,1,output.length,output[0].length).setValues(output);

  // Set checkboxes
  todaySheet.getRange(startRow,4,output.length).insertCheckboxes();

  // Apply font color and conditional backgrounds
  for (let i = 0; i < todayTasks.length; i++) {
    const row = startRow + i;
    const taskCell = todaySheet.getRange(row, 2);
    taskCell.setFontColor(todayTasks[i].color || "#000000");

    if (todayTasks[i].done) taskCell.setBackground("#00FF00");  // Done today → green
    else taskCell.setBackground("#FF0000");                     // Not done → red

    // Add border for all today tasks
    const range = todaySheet.getRange(row, 1, 1, 4);
    range.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  ss.toast("TodayTasks sheet refreshed", "Task Tools", 3);
}

/**
 * Unified handler used by the simple and installable onEdit triggers.
 * Detects multi-cell edits and pastes that intersect Tasks columns 1..7 and rows >=2.
 */
function handleEdit(e) {
  if (!e || !e.range) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- 1) Tasks sheet changes (single or multi-cell, paste, etc.) ---
    if (sheetName === "Tasks") {
      const startRow = range.getRow();
      const endRow = startRow + range.getNumRows() - 1;
      const startCol = range.getColumn();
      const endCol = startCol + range.getNumColumns() - 1;

      // intersects rows >=2 and cols 1..7
      if (endRow >= 2 && startCol <= 7 && endCol >= 1) {
        Logger.log(`handleEdit: Tasks edited rows ${startRow}-${endRow} cols ${startCol}-${endCol}`);

        // 1️⃣ Regenerate TaskOccurrences
        generateTaskOccurrences();
        ss.toast("TaskOccurrences regenerated on edit", "Task Tools", 3);

        // 2️⃣ Refresh TodayTasks
        generateTodayTasksSheet();
        ss.toast("TodayTasks refreshed on edit", "Task Tools", 3);

        // 3️⃣ Regenerate Calendar
        buildCalendarFromOccurrences();
        ss.toast("TodayTasks regenerated on edit", "Task Tools", 3);

        SpreadsheetApp.flush();
      }
    }

    // --- 2) TodayTasks checkbox edits (single-cell only) ---
    if (sheetName === "TodayTasks") {
      const row = range.getRow();
      const col = range.getColumn();
      if (row >= 4 && (col === 3 || col === 4)) { // Person or Done
        const todaySheet = ss.getSheetByName("TodayTasks");
        const occSheet = ss.getSheetByName("TaskOccurrences");

        const dateStr = todaySheet.getRange(row, 1).getValue(); // "dd.MM.yyyy"
        const date = parseDDMMYYYY(dateStr);
        const taskName = todaySheet.getRange(row, 2).getValue().toString().trim();
        const person = todaySheet.getRange(row, 3).getValue().toString().trim();
        const doneValue = todaySheet.getRange(row, 4).getValue();

        const occLast = Math.max(occSheet.getLastRow() - 1, 0);
        if (occLast > 0) {
          const occData = occSheet.getRange(2, 1, occLast, 7).getValues();
          for (let i = 0; i < occData.length; i++) {
            const occ = occData[i];
            const occDate = occ[2] instanceof Date ? occ[2] : parseDDMMYYYY(occ[2]);
            if (occ[1].toString().trim() === taskName && isSameDate(occDate, date)) {
              // Update Person if changed
              if (occ[3] !== person) occSheet.getRange(i + 2, 4).setValue(person);
              // Update Done if changed
              if (occ[4] !== doneValue) occSheet.getRange(i + 2, 5).setValue(doneValue);
              SpreadsheetApp.flush();
              break;
            }
          }
        }

        generateTodayTasksSheet();
        buildCalendarFromOccurrences();
      }
    }

    // --- 3) Calendar Year/Month edits (B1/B2) ---
    if (sheetName === "Calendar") {
      const row = range.getRow();
      const col = range.getColumn();
      if ((row === 1 && col === 2) || (row === 2 && col === 2)) {
        Logger.log(`handleEdit: Calendar Year/Month changed at B${row}`);
        buildCalendarFromOccurrences();
        ss.toast("Calendar rebuilt (Year/Month change)", "Task Tools", 3);
      }
    }

  } catch (err) {
    Logger.log("handleEdit error: " + err);
  }
}

// Simple onEdit wrapper (still useful for manual UI edits)
// function onEdit(e) { handleEdit(e); }

/**
 * Create an installable onEdit trigger that calls handleEdit.
 * Run this once manually and accept the OAuth scopes.
 */
function createInstallableOnEditTrigger() {
  // remove any duplicate triggers for the same handler to avoid multiples
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "handleEdit" && t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("handleEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  Logger.log("Installable onEdit trigger created for handleEdit");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // --- Hide the "data" sheets by default ---
    const hiddenSheets = ["Persons", "TaskOccurrences"];
    const visibleSheets = ["Tasks", "TodayTasks", "Calendar"];

    // Refresh TodayTasks
    generateTodayTasksSheet();
    ss.toast("TodayTasks refreshed", "Task Tools", 3);

    // Rebuild Calendar
    buildCalendarFromOccurrences();
    ss.toast("Calendar rebuilt", "Task Tools", 3);

    ss.getSheets().forEach(sheet => {
      if (hiddenSheets.includes(sheet.getName())) {
        sheet.hideSheet();
      } else if (visibleSheets.includes(sheet.getName())) {
        sheet.showSheet();
      }
    });

  } catch (err) {
    Logger.log("Error refreshing on open: " + err);
  }

  // Build custom menu
  ui.createMenu('🧠 Task Tools')
    .addItem('🔄 Refresh Today\'s Tasks', 'generateTodayTasksSheet')
    .addItem('📅 Rebuild Calendar', 'buildCalendarFromOccurrences')
    .addItem('🔄 Generate Task Occurrences', 'generateTaskOccurrences')
    .addSeparator()
    .addItem('🗂 Manage Data Sheets', 'showHiddenSheets')
    .addItem('🙈 Hide Data Sheets', 'hideHiddenSheets')
    .addToUi();
}

// --- Show hidden sheets ---
function showHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hiddenSheets = ["Persons", "TaskOccurrences"];

  hiddenSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.showSheet();
  });

  SpreadsheetApp.getUi().alert(
    "Hidden sheets are now visible. Remember to hide them again when finished.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --- Hide hidden sheets ---
function hideHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hiddenSheets = ["Persons", "TaskOccurrences"];

  hiddenSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.hideSheet();
  });

  SpreadsheetApp.getUi().alert("Hidden sheets are now hidden again.");
}


// --- helper ---
function isSameDate(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

function parseDDMMYYYY(str) {
  const parts = str.split(".");
  return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
}
