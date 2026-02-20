// ==========================================
// PROJECT ACTIVITY TRACKER - PHASE 2
// Multi-Sheet Support + Tasks + Google Integration
// ==========================================

// Configuration
const CONFIG = {
  SHEETS: {
    TRACKED_PROJECTS: ['cobuild', 'enablement'], // Add more sheet names as needed
    AUDIT: 'Audit Log',
    TASKS: 'Tasks'
  },
  
  // Column positions (must match in all tracked sheets)
  UUID_COLUMN: 20,              // Column T (hidden)
  CALENDAR_SYNC_COLUMN: 21,     // Column U - Calendar Sync checkbox
  ACTIVITY_COLUMN: 15,          // Column O - Next Steps
  LAST_CHECKIN_COLUMN: 12,      // Column L - Last Check In
  NEXT_CHECKIN_COLUMN: 13,      // Column M - Next Check In
  COMPLETION_DATE_COLUMN: 9,    // Column I - Completion Date
  STATUS_COLUMN: 7,             // Column G - Status
  PROJECT_TITLE_COLUMN: 3,      // Column C - Project Title
  
  // Settings
  DEFAULT_NEXT_CHECKIN_DAYS: 7,
  DEFAULT_SYNC_ENABLED: false,
  
  // Calendar settings
  CHECKIN_DURATION: 30,
  DEFAULT_CHECKIN_TIME: '14:00',
  WORK_HOURS_START: '09:00',
  WORK_HOURS_END: '17:00',
  PREFERRED_TIME_SLOTS: ['14:00', '10:00', '11:00', '15:00', '16:00', '09:00'],
  
  // Google Tasks settings
  GOOGLE_TASKS_AUTO_SYNC: true,
  
  // UI preferences
  NAVIGATE_TO_TASK_AFTER_CREATE: false,
  
  HISTORY_LIMIT: 10,
  EMAIL_REMINDER_HOUR: 8
};

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function isTrackedSheet(sheetName) {
  return CONFIG.SHEETS.TRACKED_PROJECTS.includes(sheetName);
}

function getActiveProjectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  return isTrackedSheet(activeSheet.getName()) ? activeSheet : null;
}

function generateUUID(sheetName) {
  const prefix = sheetName === 'enablement' ? 'enbl_' : 'proj_';
  return prefix + Utilities.getUuid().substring(0, 8);
}

function generateTaskID() {
  return 'task_' + Utilities.getUuid().substring(0, 8);
}

function formatTime(date) {
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const ampm = hours >= 12 ? 'PM' : 'AM';
  const displayHours = hours % 12 || 12;
  const displayMinutes = minutes.toString().padStart(2, '0');
  return `${displayHours}:${displayMinutes} ${ampm}`;
}

// ==========================================
// INITIAL SETUP
// ==========================================

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Create Audit Log
  let auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(CONFIG.SHEETS.AUDIT);
    auditSheet.appendRow(['Timestamp', 'Project UUID', 'Project Title', 'Sheet', 'Row', 'Column', 'Field Name', 'Old Value', 'New Value', 'User Email']);
    auditSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
    auditSheet.setFrozenRows(1);
    auditSheet.hideSheet();
  }
  
  // Setup each tracked sheet
  const setupMessages = [];
  CONFIG.SHEETS.TRACKED_PROJECTS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      setupMessages.push(`⚠️ Sheet "${sheetName}" not found - skipped`);
      return;
    }
    
    // Add Calendar Sync header if missing
    const lastCol = sheet.getLastColumn();
    if (lastCol >= CONFIG.CALENDAR_SYNC_COLUMN) {
      const header = sheet.getRange(1, CONFIG.CALENDAR_SYNC_COLUMN).getValue();
      if (!header) {
        sheet.getRange(1, CONFIG.CALENDAR_SYNC_COLUMN).setValue('Calendar Sync');
      }
    }
    
    // Add UUIDs
    addUUIDsToSheet(sheet, sheetName);
    
    // Hide UUID column
    sheet.hideColumns(CONFIG.UUID_COLUMN);
    
    setupMessages.push(`✅ ${sheetName} configured`);
  });
  
  // Create Tasks sheet
  let tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet(CONFIG.SHEETS.TASKS);
    setupTasksSheet(tasksSheet);
    setupMessages.push('✅ Tasks sheet created');
  }
  
  ui.alert('Setup Complete!', setupMessages.join('\n') + '\n\nTracker ready for: ' + CONFIG.SHEETS.TRACKED_PROJECTS.join(', '), ui.ButtonSet.OK);
}

function addUUIDsToSheet(sheet, sheetName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const uuidRange = sheet.getRange(2, CONFIG.UUID_COLUMN, lastRow - 1, 1);
  const uuids = uuidRange.getValues();
  
  for (let i = 0; i < uuids.length; i++) {
    if (!uuids[i][0]) {
      uuids[i][0] = generateUUID(sheetName);
    }
  }
  
  uuidRange.setValues(uuids);
}

function setupTasksSheet(sheet) {
  const headers = ['Task ID', 'Parent Task ID', 'Project UUID', 'Project Name', 'Task Description', 'Task Type', 'Due Date', 'Due Time', 'Duration (min)', 'Status', 'Priority', 'Assigned To', 'Source', 'Notes', 'Calendar Sync', 'Google Task ID', 'Created Date', 'Completed Date', 'Last Modified'];
  
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // Set column widths
  const widths = [100, 120, 120, 150, 250, 120, 100, 100, 100, 120, 100, 120, 100, 200, 100, 150, 120, 120, 120];
  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
}

function applyTaskRowValidation(sheet, row) {
  const taskTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Follow-up', 'Milestone', 'Check-in', 'Deliverable', 'Subtask'], true).build();
  sheet.getRange(row, 6).setDataValidation(taskTypeRule);
  
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Not Started', 'In Progress', 'Complete', 'Blocked', 'Cancelled'], true).build();
  sheet.getRange(row, 10).setDataValidation(statusRule);
  
  const priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(['High', 'Medium', 'Low'], true).build();
  sheet.getRange(row, 11).setDataValidation(priorityRule);
  
  const calSyncRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(row, 15).setDataValidation(calSyncRule);
}

// ==========================================
// AUDIT LOGGING
// ==========================================

function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!isTrackedSheet(sheetName)) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (row === 1 || col === CONFIG.UUID_COLUMN) return;
  
  const oldValue = e.oldValue || '[Initial value]';
  const newValue = e.value || '[Cleared]';
  
  if (oldValue === newValue) return;
  
  // Auto-generate UUID if missing
  const currentUUID = sheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  if (!currentUUID) {
    sheet.getRange(row, CONFIG.UUID_COLUMN).setValue(generateUUID(sheetName));
  }
  
  const formattedOldValue = formatValueForAudit(oldValue, col);
  const formattedNewValue = formatValueForAudit(newValue, col);
  
  logToAudit(sheet, row, col, formattedOldValue, formattedNewValue);
  
  // Auto-update check-in dates when Next Steps changes
  if (col === CONFIG.ACTIVITY_COLUMN) {
    const now = new Date();
    const nextCheckIn = new Date(now);
    nextCheckIn.setDate(nextCheckIn.getDate() + CONFIG.DEFAULT_NEXT_CHECKIN_DAYS);
    
    sheet.getRange(row, CONFIG.LAST_CHECKIN_COLUMN).setValue(now);
    sheet.getRange(row, CONFIG.NEXT_CHECKIN_COLUMN).setValue(nextCheckIn);
  }
}

function logToAudit(sheet, row, col, oldValue, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!auditSheet) return;
  
  const projectUUID = sheet.getRange(row, CONFIG.UUID_COLUMN).getValue() || 'unknown';
  const projectTitle = sheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const fieldName = sheet.getRange(1, col).getValue();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  auditSheet.appendRow([new Date(), projectUUID, projectTitle, sheet.getName(), row, col, fieldName, oldValue, newValue, userEmail]);
}

function formatValueForAudit(value, column) {
  if (!value) return '[Empty]';
  
  const dateColumns = [8, 9, 10, 12, 13];
  if (dateColumns.includes(column)) {
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    } else if (typeof value === 'number' && value > 40000 && value < 60000) {
      const date = new Date((value - 25569) * 86400 * 1000);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }
  }
  
  return String(value);
}

// ==========================================
// CHANGE HISTORY
// ==========================================

function showChangeHistory() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Switch to a tracked project sheet first.');
    return;
  }
  
  const cell = activeSheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a data cell (not header).');
    return;
  }
  
  const projectUUID = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const fieldName = activeSheet.getRange(1, col).getValue();
  
  const history = getFieldHistory(projectUUID, fieldName);
  
  if (history.length === 0) {
    SpreadsheetApp.getUi().alert('No History', `No changes recorded for "${fieldName}"`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const htmlOutput = createHistoryModal(projectTitle, fieldName, history);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, `Change History: ${fieldName}`);
}

function getFieldHistory(projectUUID, fieldName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!auditSheet) return [];
  
  const data = auditSheet.getDataRange().getValues();
  const history = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    const [timestamp, uuid, projectTitle, sheetName, row, col, field, oldVal, newVal, user] = data[i];
    
    if (uuid === projectUUID && field === fieldName) {
      history.push({
        timestamp: new Date(timestamp),
        oldValue: oldVal,
        newValue: newVal,
        user: user
      });
      
      if (history.length >= CONFIG.HISTORY_LIMIT) break;
    }
  }
  
  return history;
}

function createHistoryModal(projectTitle, fieldName, history) {
  let html = `<style>body{font-family:Arial;padding:15px;background:#f8f9fa}.header{font-size:16px;font-weight:bold;margin-bottom:20px;color:#1a73e8;background:white;padding:15px;border-radius:8px}.change-item{padding:15px;margin-bottom:12px;border-left:4px solid #1a73e8;background:white;border-radius:6px}.timestamp{font-size:12px;color:#5f6368;margin-bottom:8px}.values{font-size:14px;margin:8px 0}.old-value{color:#d93025;background:#fce8e6;padding:2px 6px;border-radius:3px}.new-value{color:#188038;background:#e6f4ea;padding:2px 6px;border-radius:3px}.user{font-size:11px;color:#5f6368;margin-top:8px}</style><div class="header">${fieldName}<div style="color:#5f6368;font-size:14px;margin-top:5px">${projectTitle}</div></div>`;
  
  history.forEach(change => {
    const date = Utilities.formatDate(change.timestamp, Session.getScriptTimeZone(), 'MMM dd, yyyy \'at\' h:mm a');
    html += `<div class="change-item"><div class="timestamp">📅 ${date}</div><div class="values"><span class="old-value">${change.oldValue}</span> → <span class="new-value">${change.newValue}</span></div><div class="user">Changed by: ${change.user}</div></div>`;
  });
  
  return HtmlService.createHtmlOutput(html).setWidth(550).setHeight(450);
}

// ==========================================
// TASKS - ADD TASK
// ==========================================

function addTaskForProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Switch to cobuild or enablement sheet first.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const activeRow = activeSheet.getActiveCell().getRow();
  
  if (activeRow === 1) {
    ui.alert('Select a project row (not header).');
    return;
  }
  
  const projectUUID = activeSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  if (!projectUUID || !projectName) {
    ui.alert('Project missing UUID or title. Run Setup Tracker.');
    return;
  }
  
  // Step 1: Description
  const descResp = ui.prompt('Add Task - Step 1 of 4', `Sheet: ${sheetName}\nProject: ${projectName}\n\nEnter task description:`, ui.ButtonSet.OK_CANCEL);
  if (descResp.getSelectedButton() !== ui.Button.OK) return;
  const taskDesc = descResp.getResponseText().trim();
  if (!taskDesc) {
    ui.alert('Description required.');
    return;
  }
  
  // Step 2: Type (default: Follow-up)
  const typeResp = ui.prompt('Add Task - Step 2 of 4', 'Task type (default: Follow-up):\n\n1 = Follow-up\n2 = Milestone\n3 = Check-in\n4 = Deliverable\n\nEnter 1-4 or leave blank:', ui.ButtonSet.OK_CANCEL);
  if (typeResp.getSelectedButton() !== ui.Button.OK) return;
  const typeMap = {'1': 'Follow-up', '2': 'Milestone', '3': 'Check-in', '4': 'Deliverable', '': 'Follow-up'};
  const taskType = typeMap[typeResp.getResponseText().trim()] || 'Follow-up';
  
  // Step 3: Due date (default: none)
  const dateResp = ui.prompt('Add Task - Step 3 of 4', 'Due date (default: none):\n\n+N = N days from today (e.g. +5, +7)\nMM/DD/YYYY = specific date\nLeave blank = no due date', ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  
  let dueDate = null;
  const dateText = dateResp.getResponseText().trim();
  if (dateText.startsWith('+')) {
    const days = parseInt(dateText.substring(1));
    if (!isNaN(days) && days > 0) {
      dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + days);
    }
  } else if (dateText) {
    const custom = new Date(dateText);
    if (!isNaN(custom.getTime())) dueDate = custom;
  }
  
  // Step 4: Priority (default: Low)
  const priResp = ui.prompt('Add Task - Step 4 of 4', 'Priority (default: Low):\n\n1 = High\n2 = Medium\n3 = Low\n\nEnter 1-3 or leave blank:', ui.ButtonSet.OK_CANCEL);
  if (priResp.getSelectedButton() !== ui.Button.OK) return;
  const priMap = {'1': 'High', '2': 'Medium', '3': 'Low', '': 'Low'};
  const priority = priMap[priResp.getResponseText().trim()] || 'Low';
  
  createTaskInSheet(projectUUID, projectName, sheetName, taskDesc, taskType, dueDate, priority);
}

function createTaskInSheet(projectUUID, projectName, sheetSource, taskDesc, taskType, dueDate, priority) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!tasksSheet) {
    ui.alert('Tasks sheet not found.');
    return;
  }
  
  // Duplicate check
  const data = tasksSheet.getDataRange().getValues();
  const fiveSecondsAgo = new Date(Date.now() - 5000);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === taskDesc && data[i][2] === projectUUID && data[i][16] instanceof Date && data[i][16] > fiveSecondsAgo) {
      ui.alert('⚠️ Duplicate', 'This task was just created.', ui.ButtonSet.OK);
      ss.setActiveSheet(tasksSheet);
      tasksSheet.setActiveRange(tasksSheet.getRange(i + 1, 1, 1, 19));
      return;
    }
  }
  
  const taskID = generateTaskID();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  tasksSheet.appendRow([taskID, '', projectUUID, projectName, taskDesc, taskType, dueDate, '', 30, 'Not Started', priority, userEmail, sheetSource, '', false, '', now, '', now]);
  
  const lastRow = tasksSheet.getLastRow();
  applyTaskRowValidation(tasksSheet, lastRow);
  
  if (dueDate) tasksSheet.getRange(lastRow, 7).setNumberFormat('M/d/yyyy');
  tasksSheet.getRange(lastRow, 17).setNumberFormat('M/d/yyyy h:mm');
  tasksSheet.getRange(lastRow, 19).setNumberFormat('M/d/yyyy h:mm');
  
  // Auto-sync to Google Tasks if enabled
  if (CONFIG.GOOGLE_TASKS_AUTO_SYNC) {
    syncTaskToGoogleTasks(taskID);
  }
  
  if (CONFIG.NAVIGATE_TO_TASK_AFTER_CREATE) {
    ss.setActiveSheet(tasksSheet);
    tasksSheet.setActiveRange(tasksSheet.getRange(lastRow, 1, 1, 19));
  } else {
    const viewResp = ui.alert('✅ Task Created', `Task: ${taskDesc}\n\nView in Tasks sheet?`, ui.ButtonSet.YES_NO);
    if (viewResp === ui.Button.YES) {
      ss.setActiveSheet(tasksSheet);
      tasksSheet.setActiveRange(tasksSheet.getRange(lastRow, 1, 1, 19));
    }
  }
}

// ==========================================
// TASKS - SUBTASKS
// ==========================================

function addSubtask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (activeSheet.getName() !== CONFIG.SHEETS.TASKS) {
    ui.alert('Switch to Tasks sheet and select a parent task.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    ui.alert('Select a task row (not header).');
    return;
  }
  
  const parentTaskID = tasksSheet.getRange(row, 1).getValue();
  const existingParent = tasksSheet.getRange(row, 2).getValue();
  const projectUUID = tasksSheet.getRange(row, 3).getValue();
  const projectName = tasksSheet.getRange(row, 4).getValue();
  const parentDesc = tasksSheet.getRange(row, 5).getValue();
  const sheetSource = tasksSheet.getRange(row, 13).getValue();
  
  if (existingParent) {
    ui.alert('⚠️ Cannot Add Subtask', 'Selected task is already a subtask. Only 1 level deep allowed.', ui.ButtonSet.OK);
    return;
  }
  
  const descResp = ui.prompt('Add Subtask', `Parent: ${parentDesc}\n\nEnter subtask description:`, ui.ButtonSet.OK_CANCEL);
  if (descResp.getSelectedButton() !== ui.Button.OK) return;
  const subtaskDesc = descResp.getResponseText().trim();
  if (!subtaskDesc) {
    ui.alert('Description required.');
    return;
  }
  
  const subtaskID = generateTaskID();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  tasksSheet.appendRow([subtaskID, parentTaskID, projectUUID, projectName, '  └─ ' + subtaskDesc, 'Subtask', null, '', 15, 'Not Started', 'Low', userEmail, sheetSource, '', false, '', now, '', now]);
  
  const lastRow = tasksSheet.getLastRow();
  applyTaskRowValidation(tasksSheet, lastRow);
  tasksSheet.getRange(lastRow, 17).setNumberFormat('M/d/yyyy h:mm');
  tasksSheet.getRange(lastRow, 19).setNumberFormat('M/d/yyyy h:mm');
  
  if (CONFIG.GOOGLE_TASKS_AUTO_SYNC) {
    syncTaskToGoogleTasks(subtaskID);
  }
  
  ui.alert('✅ Subtask Created', `Subtask under: ${parentDesc}\n\n${subtaskDesc}`, ui.ButtonSet.OK);
}

// ==========================================
// TASKS - VIEW
// ==========================================

function viewProjectTasks() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Switch to a tracked project sheet first.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row.');
    return;
  }
  
  const projectUUID = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const data = tasksSheet.getDataRange().getValues();
  
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === projectUUID) {
      tasks.push({
        description: data[i][4],
        type: data[i][5],
        dueDate: data[i][6],
        status: data[i][9],
        priority: data[i][10]
      });
    }
  }
  
  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No Tasks', `No tasks for: ${projectName}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let message = `Tasks for: ${projectName}\n\nTotal: ${tasks.length}\n\n`;
  tasks.slice(0, 10).forEach((t, i) => {
    message += `${i + 1}. [${t.priority}] ${t.description}\n   ${t.type} | ${t.status}\n\n`;
  });
  
  if (tasks.length > 10) message += `... and ${tasks.length - 10} more\n`;
  
  SpreadsheetApp.getUi().alert('Project Tasks', message, SpreadsheetApp.getUi().ButtonSet.OK);
  ss.setActiveSheet(tasksSheet);
}

// ==========================================
// GOOGLE TASKS INTEGRATION
// ==========================================

function getOrCreateGoogleTaskList(projectUUID, projectName, sheetSource) {
  try {
    const taskLists = Tasks.Tasklists.list().items || [];
    const listTitle = `[${sheetSource}] ${projectName}`;
    
    for (let list of taskLists) {
      if (list.title && list.title === listTitle) {
        return list.id;
      }
    }
    
    const newList = Tasks.Tasklists.insert({title: listTitle});
    return newList.id;
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Google Tasks Error', 'Enable Google Tasks API in Extensions → Apps Script → Services', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

function syncTaskToGoogleTasks(taskID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  if (!tasksSheet) return null;
  
  const data = tasksSheet.getDataRange().getValues();
  let taskRow = -1;
  let taskData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskID) {
      taskRow = i + 1;
      taskData = data[i];
      break;
    }
  }
  
  if (!taskData) return null;
  
  const parentTaskID = taskData[1];
  const projectUUID = taskData[2];
  const projectName = taskData[3];
  const taskDesc = taskData[4];
  const dueDate = taskData[6];
  const status = taskData[9];
  const sheetSource = taskData[12];
  const googleTaskID = taskData[15];
  
  const taskListID = getOrCreateGoogleTaskList(projectUUID, projectName, sheetSource);
  if (!taskListID) return null;
  
  try {
    let googleTask;
    
    if (googleTaskID) {
      googleTask = Tasks.Tasks.get(taskListID, googleTaskID);
      googleTask.title = taskDesc;
      googleTask.status = (status === 'Complete') ? 'completed' : 'needsAction';
      if (dueDate) googleTask.due = new Date(dueDate).toISOString();
      googleTask = Tasks.Tasks.update(googleTask, taskListID, googleTaskID);
    } else {
      googleTask = {
        title: taskDesc,
        status: (status === 'Complete') ? 'completed' : 'needsAction',
        notes: `Task ID: ${taskID}\nProject: ${projectName}`
      };
      
      if (dueDate) googleTask.due = new Date(dueDate).toISOString();
      
      if (parentTaskID) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === parentTaskID) {
            const parentGoogleID = data[i][15];
            if (parentGoogleID) googleTask.parent = parentGoogleID;
            break;
          }
        }
      }
      
      googleTask = Tasks.Tasks.insert(googleTask, taskListID);
      tasksSheet.getRange(taskRow, 16).setValue(googleTask.id);
    }
    
    return googleTask.id;
  } catch (error) {
    Logger.log(`Error syncing to Google Tasks: ${error.message}`);
    return null;
  }
}

function syncProjectToGoogleTasks() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Switch to a tracked project sheet first.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row.');
    return;
  }
  
  const projectUUID = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const data = tasksSheet.getDataRange().getValues();
  
  let syncCount = 0;
  
  // Sync parent tasks first
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === projectUUID && !data[i][1]) {
      if (syncTaskToGoogleTasks(data[i][0])) syncCount++;
    }
  }
  
  // Then subtasks
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === projectUUID && data[i][1]) {
      if (syncTaskToGoogleTasks(data[i][0])) syncCount++;
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ Synced', `Synced ${syncCount} task(s) to Google Tasks\n\nCheck your Google Tasks app!`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// CALENDAR - SMART SCHEDULING
// ==========================================

function findAvailableTimeSlot(calendar, targetDate) {
  const checkDate = new Date(targetDate);
  checkDate.setHours(0, 0, 0, 0);
  
  const nextDay = new Date(checkDate);
  nextDay.setDate(nextDay.getDate() + 1);
  
  const existingEvents = calendar.getEvents(checkDate, nextDay);
  
  for (let timeSlot of CONFIG.PREFERRED_TIME_SLOTS) {
    const [hours, minutes] = timeSlot.split(':');
    const slotStart = new Date(checkDate);
    slotStart.setHours(parseInt(hours), parseInt(minutes), 0, 0);
    
    const slotEnd = new Date(slotStart);
    slotEnd.setMinutes(slotEnd.getMinutes() + CONFIG.CHECKIN_DURATION);
    
    let hasConflict = false;
    for (let event of existingEvents) {
      const eventStart = event.getStartTime();
      const eventEnd = event.getEndTime();
      
      if ((slotStart >= eventStart && slotStart < eventEnd) || (slotEnd > eventStart && slotEnd <= eventEnd) || (slotStart <= eventStart && slotEnd >= eventEnd)) {
        hasConflict = true;
        break;
      }
    }
    
    if (!hasConflict) return slotStart;
  }
  
  const [hours, minutes] = CONFIG.DEFAULT_CHECKIN_TIME.split(':');
  const defaultSlot = new Date(checkDate);
  defaultSlot.setHours(parseInt(hours), parseInt(minutes), 0, 0);
  return defaultSlot;
}

function createCalendarEventForRow(row, sourceSheet) {
  const activeSheet = sourceSheet || getActiveProjectSheet();
  if (!activeSheet || row < 2) {
    SpreadsheetApp.getUi().alert('Invalid project selection.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const calendar = CalendarApp.getDefaultCalendar();
  
  const uuid = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue() || '';
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue() || '';
  const completionDate = activeSheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getValue() || null;
  const nextCheckIn = activeSheet.getRange(row, CONFIG.NEXT_CHECKIN_COLUMN).getValue() || null;
  const status = activeSheet.getRange(row, CONFIG.STATUS_COLUMN).getValue() || '';
  const nextSteps = activeSheet.getRange(row, CONFIG.ACTIVITY_COLUMN).getValue() || '';
  
  Logger.log(`Creating events for: ${projectTitle} (${sheetName})`);
  
  if (!completionDate && !nextCheckIn) {
    SpreadsheetApp.getUi().alert('❌ No Dates', `Set Column I or M first.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const description = `Project: ${projectTitle}\nStatus: ${status}\nNext Steps: ${nextSteps}\n\nUUID: ${uuid}\nSheet: ${sheetName}\nType: {{TYPE}}\nLink: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${activeSheet.getSheetId()}&range=A${row}`;
  
  let eventsCreated = 0;
  const messages = [];
  
  try {
    // Check-in event (time-blocked)
    if (nextCheckIn && nextCheckIn instanceof Date && !isNaN(nextCheckIn)) {
      const checkInStart = findAvailableTimeSlot(calendar, nextCheckIn);
      const checkInEnd = new Date(checkInStart);
      checkInEnd.setMinutes(checkInEnd.getMinutes() + CONFIG.CHECKIN_DURATION);
      
      const event = calendar.createEvent(`Check-in [${sheetName}]: ${projectTitle}`, checkInStart, checkInEnd, {description: description.replace('{{TYPE}}', 'check-in')});
      event.setColor(CalendarApp.EventColor.BLUE);
      event.addPopupReminder(1440);
      
      eventsCreated++;
      messages.push(`✓ Check-in: ${Utilities.formatDate(nextCheckIn, Session.getScriptTimeZone(), 'MM/dd/yyyy')} at ${formatTime(checkInStart)}`);
    }
    
    // Deadline event (all-day)
    if (completionDate && completionDate instanceof Date && !isNaN(completionDate)) {
      const event = calendar.createAllDayEvent(`DUE [${sheetName}]: ${projectTitle}`, completionDate, {description: description.replace('{{TYPE}}', 'completion')});
      event.setColor(CalendarApp.EventColor.ORANGE);
      event.addPopupReminder(1440);
      event.addPopupReminder(10080);
      
      eventsCreated++;
      messages.push(`✓ Deadline: ${Utilities.formatDate(completionDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')} (all-day)`);
    }
    
    if (eventsCreated > 0) {
      SpreadsheetApp.getUi().alert('✅ Events Created', `${eventsCreated} event(s) for ${projectTitle}:\n\n${messages.join('\n')}\n\nCheck your calendar!`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    SpreadsheetApp.getUi().alert('❌ Error', `Could not create events: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// CALENDAR - SYNC FUNCTIONS
// ==========================================

function syncWithCalendar() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Switch to cobuild or enablement sheet first.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const lastRow = activeSheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No projects found.');
    return;
  }
  
  SpreadsheetApp.getUi().alert('✅ Calendar Sync', `Scanned ${sheetName} sheet.\n\nUse "Enable Sync (This Project)" to create calendar events for individual projects.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function enableCalendarSync() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project in cobuild or enablement sheet.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row.');
    return;
  }
  
  activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).setValue(true);
  
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const sheetName = activeSheet.getName();
  
  const resp = SpreadsheetApp.getUi().alert('✅ Sync Enabled', `Sheet: ${sheetName}\nProject: ${projectTitle}\n\nCreate calendar events now?`, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  
  if (resp === SpreadsheetApp.getUi().Button.YES) {
    createCalendarEventForRow(row, activeSheet);
  }
}

function disableCalendarSync() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project in cobuild or enablement sheet.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row.');
    return;
  }
  
  activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).setValue(false);
  SpreadsheetApp.getUi().alert('⏸️ Sync Disabled', 'Calendar sync disabled for this project.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// DEBUG FUNCTIONS
// ==========================================

function testCalendarAccess() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const testDate = new Date();
    testDate.setDate(testDate.getDate() + 7);
    
    const testEvent = calendar.createAllDayEvent('TEST EVENT - Project Tracker', testDate, {description: 'Test event. You can delete this.'});
    
    SpreadsheetApp.getUi().alert('✅ Success!', `Calendar: ${calendar.getName()}\nTest event created for ${Utilities.formatDate(testDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n\nCheck your calendar!`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', `Calendar access failed: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function searchForProjectEvents() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const oneYearAgo = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
    const oneYearAhead = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));
    
    const allEvents = calendar.getEvents(oneYearAgo, oneYearAhead);
    const projectEvents = [];
    
    allEvents.forEach(event => {
      const desc = event.getDescription();
      const title = event.getTitle();
      
      if ((desc && (desc.includes('UUID: proj_') || desc.includes('UUID: enbl_'))) || (title && (title.includes('[cobuild]') || title.includes('[enablement]')))) {
        projectEvents.push({title: event.getTitle(), date: event.getAllDayStartDate() || event.getStartTime()});
      }
    });
    
    let msg = `🔍 Search Results:\n\nTotal events: ${allEvents.length}\nProject events: ${projectEvents.length}\n\n`;
    
    if (projectEvents.length > 0) {
      msg += 'Project Events:\n\n';
      projectEvents.slice(0, 10).forEach((e, i) => {
        msg += `${i + 1}. ${e.title}\n   ${Utilities.formatDate(e.date, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n\n`;
      });
    } else {
      msg += 'No project events found.';
    }
    
    SpreadsheetApp.getUi().alert('Calendar Search', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', `Search failed: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function debugSelectedProject() {
  const activeSheet = getActiveProjectSheet();
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project in cobuild or enablement sheet.');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const uuid = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const completionDate = activeSheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getValue();
  const nextCheckIn = activeSheet.getRange(row, CONFIG.NEXT_CHECKIN_COLUMN).getValue();
  const syncEnabled = activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).getValue();
  
  let msg = `🔍 Debug Info:\n\nSheet: ${sheetName}\nRow: ${row}\nUUID: ${uuid || '❌ MISSING'}\nProject: ${projectTitle}\nCompletion: ${completionDate || '❌ NOT SET'}\nNext Check In: ${nextCheckIn || '❌ NOT SET'}\nCalendar Sync: ${syncEnabled ? '✅ Enabled' : '❌ Disabled'}\n\n`;
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = findCalendarEventsByUUID(calendar, uuid);
    msg += `📅 Calendar Events: ${events.length} found`;
  } catch (error) {
    msg += `📅 Calendar: Error - ${error.message}`;
  }
  
  SpreadsheetApp.getUi().alert('Debug Info', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function findCalendarEventsByUUID(calendar, uuid) {
  const now = new Date();
  const oneYearAgo = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  const oneYearAhead = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));
  
  const events = calendar.getEvents(oneYearAgo, oneYearAhead);
  const found = [];
  
  for (let event of events) {
    const desc = event.getDescription();
    if (desc && desc.includes(`UUID: ${uuid}`)) {
      found.push(event);
    }
  }
  
  return found;
}

// ==========================================
// MENU
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Project Tracker')
    .addItem('🔧 Setup Tracker', 'setupTracker')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Tasks')
      .addItem('➕ Add Task for This Project', 'addTaskForProject')
      .addItem('➕ Add Subtask to This Task', 'addSubtask')
      .addItem('📊 View Project Tasks', 'viewProjectTasks')
      .addSeparator()
      .addItem('🔄 Sync to Google Tasks', 'syncProjectToGoogleTasks'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Calendar')
      .addItem('🔄 Sync with Calendar', 'syncWithCalendar')
      .addItem('✅ Enable Sync (This Project)', 'enableCalendarSync')
      .addItem('⏸️ Disable Sync (This Project)', 'disableCalendarSync'))
    .addSeparator()
    .addItem('📜 View Change History', 'showChangeHistory')
    .addSeparator()
    .addSubMenu(ui.createMenu('🐛 Debug')
      .addItem('🔍 Test Calendar', 'testCalendarAccess')
      .addItem('🔍 Search Events', 'searchForProjectEvents')
      .addItem('🔍 Debug Project', 'debugSelectedProject'))
    .addToUi();
}