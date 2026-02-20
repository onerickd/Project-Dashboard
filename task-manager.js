// ==========================================
// PROJECT ACTIVITY TRACKER WITH AUDIT LOG & CALENDAR SYNC
// Optimized for cobuild sheet + Calendar sync with selective control (Column U)
// ==========================================

// Configuration
const CONFIG = {
  SHEETS: {
    TRACKED_PROJECTS: ['cobuild', 'enablement'], // Sheets to track (add more as needed)
    AUDIT: 'Audit Log',
    TASKS: 'Tasks'
  },
  UUID_COLUMN: 20,              // Column T (hidden)
  CALENDAR_SYNC_COLUMN: 21,     // Column U - Calendar Sync checkbox
  ACTIVITY_COLUMN: 15,          // Column O - Next Steps (main activity tracking)
  LAST_CHECKIN_COLUMN: 12,      // Column L - Last Check In (auto-update on Next Steps change)
  NEXT_CHECKIN_COLUMN: 13,      // Column M - Next Check In (auto-update to +7 days)
  COMPLETION_DATE_COLUMN: 9,    // Column I - Completion Date
  STATUS_COLUMN: 7,             // Column G - Status
  PROJECT_TITLE_COLUMN: 3,      // Column C - Project Title
  
  DEFAULT_NEXT_CHECKIN_DAYS: 7, // Days to add for next check-in
  DEFAULT_SYNC_ENABLED: false,  // New projects start with sync OFF (user decides)
  
  // Calendar settings
  CALENDAR_NAME: 'primary',     // Which calendar to use
  CHECKIN_DURATION: 30,         // Check-in duration in minutes
  DEFAULT_CHECKIN_TIME: '14:00', // Default 2:00 PM if no empty slots
  WORK_HOURS_START: '09:00',    // Work day starts at 9 AM
  WORK_HOURS_END: '17:00',      // Work day ends at 5 PM
  PREFERRED_TIME_SLOTS: ['14:00', '10:00', '11:00', '15:00', '16:00', '09:00'], // Priority order
  SYNC_PRIORITY: 'calendar',    // 'calendar' or 'sheet' - calendar wins by default
  
  HISTORY_LIMIT: 10,            // Number of changes to show in modal
  EMAIL_REMINDER_HOUR: 8        // 8 AM daily reminder
};

// ==========================================
// HELPER FUNCTIONS - Multi-sheet support
// ==========================================

// Check if a sheet is a tracked project sheet
function isTrackedSheet(sheetName) {
  return CONFIG.SHEETS.TRACKED_PROJECTS.includes(sheetName);
}

// Get current active project sheet or return null
function getActiveProjectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  if (isTrackedSheet(sheetName)) {
    return activeSheet;
  }
  return null;
}

// Get sheet source name for display (e.g., "cobuild", "enablement")
function getSheetSource(sheetName) {
  return sheetName;
}

// ==========================================
// INITIAL SETUP - Run this once
// ==========================================
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  
  // Create Audit Log sheet if it doesn't exist
  let auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(CONFIG.SHEETS.AUDIT);
    auditSheet.appendRow([
      'Timestamp', 'Project UUID', 'Project Title', 'Sheet', 'Row', 
      'Column', 'Field Name', 'Old Value', 'New Value', 'User Email'
    ]);
    auditSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
    auditSheet.setFrozenRows(1);
    auditSheet.hideSheet(); // Hide audit log by default
  }
  
  // Add Calendar Sync column header if missing
  const headerRow = projectSheet.getRange(1, 1, 1, projectSheet.getLastColumn()).getValues()[0];
  if (!headerRow[CONFIG.CALENDAR_SYNC_COLUMN - 1] || headerRow[CONFIG.CALENDAR_SYNC_COLUMN - 1] === '') {
    projectSheet.getRange(1, CONFIG.CALENDAR_SYNC_COLUMN).setValue('Calendar Sync');
  }
  
  // Add UUIDs to existing projects if missing
  addUUIDsToProjects();
  
  // Hide UUID column
  projectSheet.hideColumns(CONFIG.UUID_COLUMN);
  
  // Create Tasks sheet if it doesn't exist
  let tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet(CONFIG.SHEETS.TASKS);
    setupTasksSheet(tasksSheet);
  }
  
  SpreadsheetApp.getUi().alert('Setup Complete!', 
    '✅ Audit Log created and hidden\n✅ UUIDs generated\n✅ UUID column (T) hidden\n✅ Calendar Sync column (U) added\n\n' +
    'Your cobuild sheet is now being tracked!\n\n' +
    'Features:\n' +
    '• All changes automatically logged to Audit Log\n' +
    '• Column O (Next Steps) is your main activity field\n' +
    '• Column L (Last Check In) & M (Next Check In) auto-update\n' +
    '• Column U (Calendar Sync) controls which projects sync to calendar\n' +
    '• Select any cell → View Change History to see past values\n\n' +
    'Next Steps:\n' +
    '1. Check Column U checkbox for projects you want on calendar\n' +
    '2. Click "📅 Sync with Calendar" to create calendar events',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// UUID GENERATION
// ==========================================
function generateUUID() {
  return 'proj_' + Utilities.getUuid().substring(0, 8);
}

function addUUIDsToProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return; // No data rows
  
  const uuidRange = sheet.getRange(2, CONFIG.UUID_COLUMN, lastRow - 1, 1);
  const uuids = uuidRange.getValues();
  
  for (let i = 0; i < uuids.length; i++) {
    if (!uuids[i][0]) { // If UUID is empty
      uuids[i][0] = generateUUID();
    }
  }
  
  uuidRange.setValues(uuids);
}

// ==========================================
// AUDIT LOGGING - Triggers on every edit
// ==========================================
function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Only track changes in tracked project sheets
  if (!isTrackedSheet(sheetName)) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Skip header row
  if (row === 1) return;
  
  // Skip UUID column changes
  if (col === CONFIG.UUID_COLUMN) return;
  
  const oldValue = e.oldValue || '[Initial value]';
  const newValue = e.value || '[Cleared]';
  
  // Only log if value actually changed
  if (oldValue === newValue) return;
  
  // Check if this row needs a UUID (new project)
  const currentUUID = sheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  if (!currentUUID) {
    sheet.getRange(row, CONFIG.UUID_COLUMN).setValue(generateUUID(sheetName));
  }
  
  // Format values for better readability in audit log
  const formattedOldValue = formatValueForAudit(oldValue, col);
  const formattedNewValue = formatValueForAudit(newValue, col);
  
  // Log to audit FIRST (before any auto-updates)
  logToAudit(sheet, row, col, formattedOldValue, formattedNewValue);
  
  // If Next Steps (Column O) was edited, auto-update Last Check In (Column L) and Next Check In (Column M)
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
  
  if (!auditSheet) {
    // Create audit sheet if it doesn't exist yet
    const newAuditSheet = ss.insertSheet(CONFIG.SHEETS.AUDIT);
    newAuditSheet.appendRow([
      'Timestamp', 'Project UUID', 'Project Title', 'Sheet', 'Row', 
      'Column', 'Field Name', 'Old Value', 'New Value', 'User Email'
    ]);
    newAuditSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
    newAuditSheet.setFrozenRows(1);
    newAuditSheet.hideSheet();
    return logToAudit(sheet, row, col, oldValue, newValue); // Retry with new sheet
  }
  
  // Get project info
  const projectUUID = sheet.getRange(row, CONFIG.UUID_COLUMN).getValue() || 'unknown';
  const projectTitle = sheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue(); // Column C - Project Title
  const fieldName = sheet.getRange(1, col).getValue(); // Header name
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  // Append to audit log
  auditSheet.appendRow([
    new Date(),
    projectUUID,
    projectTitle,
    sheet.getName(),
    row,
    col,
    fieldName,
    oldValue,
    newValue,
    userEmail
  ]);
}

// Format values for better readability in audit log (especially dates)
function formatValueForAudit(value, column) {
  if (!value) return '[Empty]';
  
  // Date columns: H (Start Date), I (Completion Date), J (First Check In), L (Last Check In), M (Next Check In)
  const dateColumns = [8, 9, 10, 12, 13];
  
  if (dateColumns.includes(column)) {
    // Check if it's a date object or date serial number
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    } else if (typeof value === 'number' && value > 40000 && value < 60000) {
      // Excel date serial number (between ~2009-2064)
      const date = new Date((value - 25569) * 86400 * 1000);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }
  }
  
  return String(value);
}

// ==========================================
// SHOW CHANGE HISTORY MODAL
// ==========================================
function showChangeHistory() {
  const activeSheet = getActiveProjectSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Switch to a tracked project sheet (cobuild or enablement) first.');
    return;
  }
  
  const cell = activeSheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a data cell (not a header) to view its change history.');
    return;
  }
  
  const projectUUID = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const fieldName = activeSheet.getRange(1, col).getValue();
  
  const history = getFieldHistory(projectUUID, fieldName);
  
  if (history.length === 0) {
    SpreadsheetApp.getUi().alert('No History', 
      `No changes recorded yet for "${fieldName}" in project "${projectTitle}".\n\n` +
      `Make an edit to this field to start tracking history.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
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
  
  // Start from row 1 (skip header at 0), go backwards (newest first)
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
  // Highlight if this is the Next Steps column
  const isActivityColumn = fieldName === 'Next Steps';
  const emoji = isActivityColumn ? '📝' : '📋';
  
  let html = `
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; margin: 0; background: #f8f9fa; }
      .header { 
        font-size: 16px; 
        font-weight: bold; 
        margin-bottom: 20px; 
        color: #1a73e8;
        background: white;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      }
      .project-title { color: #5f6368; font-size: 14px; margin-top: 5px; }
      .change-item { 
        padding: 15px; 
        margin-bottom: 12px; 
        border-left: 4px solid #1a73e8; 
        background: white;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      }
      .timestamp { 
        font-size: 12px; 
        color: #5f6368; 
        margin-bottom: 8px;
        font-weight: 500;
      }
      .values { 
        font-size: 14px; 
        margin: 8px 0; 
        line-height: 1.5;
      }
      .old-value { 
        color: #d93025; 
        background: #fce8e6;
        padding: 2px 6px;
        border-radius: 3px;
      }
      .new-value { 
        color: #188038; 
        background: #e6f4ea;
        padding: 2px 6px;
        border-radius: 3px;
      }
      .arrow { color: #5f6368; margin: 0 8px; }
      .user { 
        font-size: 11px; 
        color: #5f6368; 
        margin-top: 8px;
        font-style: italic;
      }
      .summary { 
        font-size: 12px; 
        color: #5f6368; 
        margin-top: 20px; 
        text-align: center;
        padding: 10px;
        background: white;
        border-radius: 6px;
      }
      .activity-badge {
        display: inline-block;
        background: #1a73e8;
        color: white;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        margin-left: 8px;
      }
    </style>
    <div class="header">
      ${emoji} ${fieldName}
      ${isActivityColumn ? '<span class="activity-badge">Main Activity Field</span>' : ''}
      <div class="project-title">${projectTitle}</div>
    </div>
  `;
  
  history.forEach((change, index) => {
    const date = Utilities.formatDate(change.timestamp, Session.getScriptTimeZone(), 'MMM dd, yyyy \'at\' h:mm a');
    html += `
      <div class="change-item">
        <div class="timestamp">📅 ${date}</div>
        <div class="values">
          <span class="old-value">${escapeHtml(change.oldValue)}</span>
          <span class="arrow">→</span>
          <span class="new-value">${escapeHtml(change.newValue)}</span>
        </div>
        <div class="user">Changed by: ${change.user}</div>
      </div>
    `;
  });
  
  html += `<div class="summary">Showing last ${history.length} change${history.length > 1 ? 's' : ''}</div>`;
  
  return HtmlService.createHtmlOutput(html).setWidth(550).setHeight(450);
}

function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;')
                     .replace(/</g, '&lt;')
                     .replace(/>/g, '&gt;')
                     .replace(/"/g, '&quot;');
}

// ==========================================
// GOOGLE TASKS INTEGRATION
// ==========================================

// Get or create a task list for a project
function getOrCreateGoogleTaskList(projectUUID, projectName) {
  try {
    // Search for existing task list with project UUID in title
    const taskLists = Tasks.Tasklists.list().items || [];
    
    for (let list of taskLists) {
      if (list.title && list.title.includes(`(${projectUUID})`)) {
        Logger.log(`Found existing task list: ${list.title}`);
        return list.id;
      }
    }
    
    // Create new task list
    const newList = Tasks.Tasklists.insert({
      title: `${projectName} (${projectUUID})`
    });
    
    Logger.log(`Created new task list: ${newList.title}`);
    return newList.id;
    
  } catch (error) {
    Logger.log(`Error accessing Google Tasks: ${error.message}`);
    SpreadsheetApp.getUi().alert('❌ Google Tasks Error',
      'Could not access Google Tasks.\n\n' +
      'Make sure Google Tasks API is enabled:\n' +
      'Extensions → Apps Script → Services → Add Google Tasks API',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

// Sync a single task to Google Tasks
function syncTaskToGoogleTasks(taskID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  
  if (!tasksSheet) {
    Logger.log('Tasks sheet not found');
    return;
  }
  
  // Find the task by ID
  const data = tasksSheet.getDataRange().getValues();
  let taskRow = -1;
  let taskData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskID) { // Column A - Task ID
      taskRow = i + 1;
      taskData = data[i];
      break;
    }
  }
  
  if (!taskData) {
    Logger.log(`Task ${taskID} not found`);
    return;
  }
  
  const parentTaskID = taskData[1];      // B - Parent Task ID
  const projectUUID = taskData[2];       // C - Project UUID
  const projectName = taskData[3];       // D - Project Name
  const taskDescription = taskData[4];   // E - Task Description
  const dueDate = taskData[6];           // G - Due Date
  const status = taskData[9];            // J - Status
  const googleTaskID = taskData[15];     // P - Google Task ID
  
  Logger.log(`Syncing task: ${taskDescription}`);
  
  // Get or create Google Tasks list for this project
  const taskListID = getOrCreateGoogleTaskList(projectUUID, projectName);
  if (!taskListID) return;
  
  try {
    let googleTask;
    
    if (googleTaskID) {
      // Update existing Google Task
      googleTask = Tasks.Tasks.get(taskListID, googleTaskID);
      googleTask.title = taskDescription;
      googleTask.status = (status === 'Complete') ? 'completed' : 'needsAction';
      
      if (dueDate) {
        googleTask.due = new Date(dueDate).toISOString();
      }
      
      googleTask = Tasks.Tasks.update(googleTask, taskListID, googleTaskID);
      Logger.log(`Updated Google Task: ${googleTaskID}`);
      
    } else {
      // Create new Google Task
      googleTask = {
        title: taskDescription,
        status: (status === 'Complete') ? 'completed' : 'needsAction',
        notes: `Task ID: ${taskID}\nProject: ${projectName}`
      };
      
      if (dueDate) {
        googleTask.due = new Date(dueDate).toISOString();
      }
      
      // If this is a subtask, link it to parent
      if (parentTaskID) {
        // Find parent's Google Task ID
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === parentTaskID) {
            const parentGoogleTaskID = data[i][15]; // P - Google Task ID
            if (parentGoogleTaskID) {
              googleTask.parent = parentGoogleTaskID;
              Logger.log(`Linking subtask to parent: ${parentGoogleTaskID}`);
            }
            break;
          }
        }
      }
      
      googleTask = Tasks.Tasks.insert(googleTask, taskListID);
      Logger.log(`Created Google Task: ${googleTask.id}`);
      
      // Save Google Task ID back to sheet
      tasksSheet.getRange(taskRow, 16).setValue(googleTask.id); // Column P
    }
    
    return googleTask.id;
    
  } catch (error) {
    Logger.log(`Error syncing to Google Tasks: ${error.message}`);
    return null;
  }
}

// Sync all tasks for a project to Google Tasks
function syncProjectToGoogleTasks(projectUUID, showPreview = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!tasksSheet) {
    ui.alert('Tasks sheet not found.');
    return;
  }
  
  const data = tasksSheet.getDataRange().getValues();
  const tasksToSync = [];
  
  // Collect all tasks for this project
  for (let i = 1; i < data.length; i++) {
    const taskProjectUUID = data[i][2]; // C - Project UUID
    const taskID = data[i][0];          // A - Task ID
    const parentTaskID = data[i][1];    // B - Parent Task ID
    const description = data[i][4];     // E - Description
    const taskType = data[i][5];        // F - Task Type
    const googleTaskID = data[i][15];   // P - Google Task ID
    const status = data[i][9];          // J - Status
    
    if (taskProjectUUID === projectUUID && taskID) {
      tasksToSync.push({
        taskID: taskID,
        description: description,
        type: taskType,
        isSubtask: parentTaskID ? true : false,
        parentTaskID: parentTaskID,
        googleTaskID: googleTaskID,
        status: status,
        action: googleTaskID ? 'Update' : 'Create'
      });
    }
  }
  
  if (tasksToSync.length === 0) {
    ui.alert('No tasks found for this project.');
    return;
  }
  
  // Show preview if requested
  if (showPreview) {
    const projectName = tasksToSync[0].description ? data.find(row => row[2] === projectUUID)[3] : 'Unknown';
    const previewHtml = createSyncPreviewDialog(projectName, tasksToSync);
    const htmlOutput = HtmlService.createHtmlOutput(previewHtml).setWidth(600).setHeight(500);
    
    const response = ui.showModalDialog(htmlOutput, 'Google Tasks Sync Preview');
    
    // User closed without confirming - handled in HTML
    return;
  }
  
  // Perform actual sync
  performProjectSync(projectUUID, tasksToSync, ui);
}

// Create sync preview dialog
function createSyncPreviewDialog(projectName, tasksToSync) {
  const createCount = tasksToSync.filter(t => t.action === 'Create').length;
  const updateCount = tasksToSync.filter(t => t.action === 'Update').length;
  const taskCount = tasksToSync.filter(t => !t.isSubtask).length;
  const subtaskCount = tasksToSync.filter(t => t.isSubtask).length;
  
  let html = `
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; background: #f8f9fa; margin: 0; }
      .header { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #1a73e8; }
      .summary { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
      .task-item { 
        background: white; 
        padding: 12px; 
        margin-bottom: 8px; 
        border-radius: 6px;
        border-left: 4px solid #34a853;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      }
      .task-item.update { border-left-color: #fbbc04; }
      .task-item.subtask { margin-left: 20px; border-left-color: #4285f4; }
      .task-name { font-weight: bold; color: #202124; margin-bottom: 4px; }
      .task-meta { font-size: 12px; color: #5f6368; }
      .badge { 
        display: inline-block; 
        padding: 2px 8px; 
        border-radius: 12px; 
        font-size: 11px; 
        margin-left: 8px;
      }
      .badge-create { background: #e6f4ea; color: #137333; }
      .badge-update { background: #fef7e0; color: #b06000; }
      .badge-subtask { background: #e8f0fe; color: #1967d2; }
      .buttons { 
        position: sticky;
        bottom: 0;
        background: white; 
        padding: 15px; 
        text-align: right; 
        border-top: 1px solid #e0e0e0;
        box-shadow: 0 -2px 4px rgba(0,0,0,0.05);
      }
      button { 
        padding: 10px 24px; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer; 
        font-size: 14px;
        margin-left: 8px;
      }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-primary:hover { background: #1765cc; }
      .btn-secondary { background: #e8eaed; color: #3c4043; }
      .btn-secondary:hover { background: #d2d4d6; }
      .scrollable { max-height: 300px; overflow-y: auto; }
    </style>
    
    <div class="header">🔄 Sync to Google Tasks</div>
    
    <div class="summary">
      <strong>Project:</strong> ${projectName}<br>
      <strong>Total Tasks:</strong> ${tasksToSync.length} (${taskCount} tasks, ${subtaskCount} subtasks)<br>
      <strong>Actions:</strong> ${createCount} new, ${updateCount} updates
    </div>
    
    <div class="scrollable">
  `;
  
  // Show parent tasks first
  tasksToSync.filter(t => !t.isSubtask).forEach(task => {
    const cssClass = task.action === 'Update' ? 'update' : '';
    html += `
      <div class="task-item ${cssClass}">
        <div class="task-name">
          ${task.description}
          <span class="badge badge-${task.action.toLowerCase()}">${task.action}</span>
        </div>
        <div class="task-meta">
          ${task.type} • Status: ${task.status}
        </div>
      </div>
    `;
    
    // Show subtasks under parent
    const subtasks = tasksToSync.filter(t => t.parentTaskID === task.taskID);
    subtasks.forEach(subtask => {
      html += `
        <div class="task-item subtask">
          <div class="task-name">
            ${subtask.description}
            <span class="badge badge-subtask">Subtask</span>
            <span class="badge badge-${subtask.action.toLowerCase()}">${subtask.action}</span>
          </div>
        </div>
      `;
    });
  });
  
  html += `
    </div>
    
    <div class="buttons">
      <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
      <button class="btn-primary" onclick="confirmSync()">Sync to Google Tasks</button>
    </div>
    
    <script>
      function confirmSync() {
        const btn = event.target;
        btn.disabled = true;
        btn.textContent = 'Syncing...';
        
        google.script.run
          .withSuccessHandler(function() {
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error);
            btn.disabled = false;
            btn.textContent = 'Sync to Google Tasks';
          })
          .performProjectSyncFromDialog('${projectName.replace(/'/g, "\\'")}');
      }
    </script>
  `;
  
  return html;
}

// Perform the actual sync (called from dialog or directly)
function performProjectSync(projectUUID, tasksToSync, ui) {
  let syncCount = 0;
  let errorCount = 0;
  
  // First pass: sync parent tasks
  tasksToSync.filter(t => !t.isSubtask).forEach(task => {
    const result = syncTaskToGoogleTasks(task.taskID);
    if (result) {
      syncCount++;
    } else {
      errorCount++;
    }
  });
  
  // Second pass: sync subtasks (after parents have Google Task IDs)
  tasksToSync.filter(t => t.isSubtask).forEach(task => {
    const result = syncTaskToGoogleTasks(task.taskID);
    if (result) {
      syncCount++;
    } else {
      errorCount++;
    }
  });
  
  if (ui) {
    ui.alert('✅ Google Tasks Sync Complete',
      `Synced ${syncCount} task(s) to Google Tasks\n` +
      (errorCount > 0 ? `${errorCount} error(s)\n\n` : '\n') +
      `Check your Google Tasks app to see them!`,
      ui.ButtonSet.OK);
  }
  
  return { syncCount, errorCount };
}

// Called from dialog confirmation
function performProjectSyncFromDialog(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  let projectUUID;
  
  if (activeSheet.getName() === CONFIG.SHEETS.PROJECTS) {
    const activeRow = activeSheet.getActiveCell().getRow();
    const projectSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    projectUUID = projectSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  } else {
    // Find project UUID from Tasks sheet
    const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
    const data = tasksSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === projectName) { // D - Project Name
        projectUUID = data[i][2]; // C - Project UUID
        break;
      }
    }
  }
  
  if (projectUUID) {
    syncProjectToGoogleTasks(projectUUID, false); // false = don't show preview again
  }
}

// Manual sync selected project
function syncSelectedProjectToGoogleTasks() {
  const activeSheet = getActiveProjectSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) and select a project row.');
    return;
  }
  
  const activeRow = activeSheet.getActiveCell().getRow();
  if (activeRow === 1) {
    ui.alert('Select a project row (not the header).');
    return;
  }
  
  const projectUUID = activeSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  if (!projectUUID) {
    ui.alert('Project missing UUID. Run "Setup Tracker".');
    return;
  }
  
  Logger.log(`Manually syncing project: ${projectName} (${projectUUID})`);
  syncProjectToGoogleTasks(projectUUID);
}

// Toggle auto-sync feature
function toggleGoogleTasksAutoSync() {
  const ui = SpreadsheetApp.getUi();
  const currentState = CONFIG.GOOGLE_TASKS_AUTO_SYNC;
  
  const response = ui.alert(
    'Toggle Google Tasks Auto-Sync',
    `Current state: ${currentState ? 'ON' : 'OFF'}\n\n` +
    `Auto-sync ${currentState ? 'automatically' : 'will'} sync tasks to Google Tasks when created/updated.\n\n` +
    `Change to ${currentState ? 'OFF' : 'ON'}?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Note: This changes the CONFIG value for this session only
    // To persist, user needs to edit the CONFIG.GOOGLE_TASKS_AUTO_SYNC value in code
    CONFIG.GOOGLE_TASKS_AUTO_SYNC = !currentState;
    
    ui.alert('⚙️ Setting Changed',
      `Google Tasks Auto-Sync is now ${CONFIG.GOOGLE_TASKS_AUTO_SYNC ? 'ON' : 'OFF'}\n\n` +
      `Note: This change lasts for this session.\n` +
      `To make it permanent, edit CONFIG.GOOGLE_TASKS_AUTO_SYNC in the script.`,
      ui.ButtonSet.OK);
  }
}

// View Google Tasks sync status
function viewGoogleTasksSyncStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!tasksSheet) {
    ui.alert('Tasks sheet not found.');
    return;
  }
  
  const data = tasksSheet.getDataRange().getValues();
  
  let totalTasks = 0;
  let syncedTasks = 0;
  let notSyncedTasks = 0;
  let syncedSubtasks = 0;
  let notSyncedSubtasks = 0;
  let tasksByProject = {};
  
  for (let i = 1; i < data.length; i++) {
    const taskID = data[i][0];           // A - Task ID
    const parentTaskID = data[i][1];     // B - Parent Task ID
    const projectUUID = data[i][2];      // C - Project UUID
    const projectName = data[i][3];      // D - Project Name
    const description = data[i][4];      // E - Description
    const googleTaskID = data[i][15];    // P - Google Task ID
    
    if (!taskID) continue; // Skip empty rows
    
    const isSubtask = parentTaskID ? true : false;
    const isSynced = googleTaskID ? true : false;
    
    // Initialize project stats if not exists
    if (!tasksByProject[projectUUID]) {
      tasksByProject[projectUUID] = {
        name: projectName,
        synced: 0,
        notSynced: 0,
        tasks: []
      };
    }
    
    if (isSubtask) {
      if (isSynced) {
        syncedSubtasks++;
      } else {
        notSyncedSubtasks++;
      }
    } else {
      totalTasks++;
      if (isSynced) {
        syncedTasks++;
      } else {
        notSyncedTasks++;
      }
    }
    
    // Track by project
    if (isSynced) {
      tasksByProject[projectUUID].synced++;
    } else {
      tasksByProject[projectUUID].notSynced++;
      tasksByProject[projectUUID].tasks.push(description);
    }
  }
  
  // Build status message
  let message = '📊 Google Tasks Sync Status\n\n';
  message += `⚙️ Auto-Sync: ${CONFIG.GOOGLE_TASKS_AUTO_SYNC ? 'ON' : 'OFF'}\n\n`;
  
  message += '📋 Overall:\n';
  message += `  Total Tasks: ${totalTasks}\n`;
  message += `  ✅ Synced: ${syncedTasks}\n`;
  message += `  ❌ Not Synced: ${notSyncedTasks}\n`;
  message += `  Subtasks: ${syncedSubtasks + notSyncedSubtasks} (${syncedSubtasks} synced)\n\n`;
  
  if (notSyncedTasks + notSyncedSubtasks === 0) {
    message += '✅ All tasks synced to Google Tasks!\n\n';
  } else {
    message += '⚠️ Tasks Not Synced by Project:\n\n';
    
    const projectList = Object.keys(tasksByProject)
      .filter(uuid => tasksByProject[uuid].notSynced > 0)
      .slice(0, 5);
    
    projectList.forEach(uuid => {
      const proj = tasksByProject[uuid];
      message += `📁 ${proj.name}:\n`;
      message += `  ${proj.synced} synced, ${proj.notSynced} not synced\n`;
      
      if (proj.tasks.length > 0) {
        proj.tasks.slice(0, 3).forEach(task => {
          message += `    • ${task}\n`;
        });
        if (proj.tasks.length > 3) {
          message += `    ... and ${proj.tasks.length - 3} more\n`;
        }
      }
      message += '\n';
    });
    
    if (Object.keys(tasksByProject).filter(uuid => tasksByProject[uuid].notSynced > 0).length > 5) {
      message += `... and more projects\n\n`;
    }
    
    message += `💡 Tip: Use "Sync to Google Tasks" to sync unsynced tasks.`;
  }
  
  ui.alert('Google Tasks Sync Status', message, ui.ButtonSet.OK);
}

// Sync all tasks from all projects (bulk sync)
function syncAllTasksToGoogleTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!tasksSheet) {
    ui.alert('Tasks sheet not found.');
    return;
  }
  
  const response = ui.alert(
    'Sync All Tasks to Google Tasks?',
    'This will sync ALL tasks from ALL projects to Google Tasks.\n\n' +
    'This may take a while if you have many tasks.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const data = tasksSheet.getDataRange().getValues();
  const projects = new Set();
  
  // Get unique project UUIDs
  for (let i = 1; i < data.length; i++) {
    const projectUUID = data[i][2]; // C - Project UUID
    if (projectUUID) {
      projects.add(projectUUID);
    }
  }
  
  Logger.log(`Syncing ${projects.size} projects...`);
  
  let successCount = 0;
  let errorCount = 0;
  
  // Sync each project
  projects.forEach(projectUUID => {
    try {
      syncProjectToGoogleTasks(projectUUID);
      successCount++;
    } catch (error) {
      Logger.log(`Error syncing project ${projectUUID}: ${error.message}`);
      errorCount++;
    }
  });
  
  ui.alert('✅ Bulk Sync Complete',
    `Synced ${successCount} project(s) to Google Tasks\n` +
    (errorCount > 0 ? `${errorCount} error(s)\n\n` : '\n') +
    `Check your Google Tasks app!`,
    ui.ButtonSet.OK);
}

// ==========================================
// SUBTASKS
// ==========================================

// Add subtask to selected task
function addSubtaskToTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (activeSheet.getName() !== CONFIG.SHEETS.TASKS) {
    ui.alert('Please switch to the Tasks sheet and select a parent task row.');
    return;
  }
  
  const activeRow = activeSheet.getActiveCell().getRow();
  if (activeRow === 1) {
    ui.alert('Select a task row (not the header) to add a subtask.');
    return;
  }
  
  // Get parent task info
  const parentTaskID = tasksSheet.getRange(activeRow, 1).getValue();      // A - Task ID
  const existingParentID = tasksSheet.getRange(activeRow, 2).getValue();  // B - Parent Task ID
  const projectUUID = tasksSheet.getRange(activeRow, 3).getValue();       // C - Project UUID
  const projectName = tasksSheet.getRange(activeRow, 4).getValue();       // D - Project Name
  const parentDescription = tasksSheet.getRange(activeRow, 5).getValue(); // E - Description
  
  if (existingParentID) {
    ui.alert('⚠️ Cannot Add Subtask',
      'The selected task is already a subtask.\n\n' +
      'Subtasks can only be 1 level deep.\n' +
      'Select a parent task instead.',
      ui.ButtonSet.OK);
    return;
  }
  
  Logger.log(`Adding subtask to parent: ${parentTaskID}`);
  
  // Prompt for subtask description
  const descResponse = ui.prompt(
    'Add Subtask',
    `Parent Task: ${parentDescription}\n\nEnter subtask description:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;
  const subtaskDescription = descResponse.getResponseText().trim();
  
  if (!subtaskDescription) {
    ui.alert('Subtask description cannot be empty.');
    return;
  }
  
  // Create subtask
  const subtaskID = generateTaskID();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  tasksSheet.appendRow([
    subtaskID,                 // A - Task ID
    parentTaskID,              // B - Parent Task ID
    projectUUID,               // C - Project UUID
    projectName,               // D - Project Name
    '  └─ ' + subtaskDescription, // E - Task Description (with indent)
    'Subtask',                 // F - Task Type
    null,                      // G - Due Date (inherits from parent or set separately)
    '',                        // H - Due Time
    15,                        // I - Duration (subtasks default to 15 min)
    'Not Started',             // J - Status
    'Low',                     // K - Priority (subtasks default to Low)
    userEmail,                 // L - Assigned To
    'Tasks',                   // M - Source
    '',                        // N - Notes
    false,                     // O - Calendar Sync
    '',                        // P - Google Task ID
    now,                       // Q - Created Date
    '',                        // R - Completed Date
    now                        // S - Last Modified
  ]);
  
  const lastRow = tasksSheet.getLastRow();
  applyTaskRowValidation(tasksSheet, lastRow);
  tasksSheet.getRange(lastRow, 17).setNumberFormat('M/d/yyyy h:mm');
  tasksSheet.getRange(lastRow, 19).setNumberFormat('M/d/yyyy h:mm');
  
  Logger.log(`Subtask created: ${subtaskID}`);
  
  // Auto-sync if enabled
  if (CONFIG.GOOGLE_TASKS_AUTO_SYNC) {
    Logger.log('Auto-syncing subtask to Google Tasks...');
    syncTaskToGoogleTasks(subtaskID);
  }
  
  ui.alert('✅ Subtask Created',
    `Subtask added under: ${parentDescription}\n\n` +
    `Subtask: ${subtaskDescription}\n\n` +
    `Subtask ID: ${subtaskID}`,
    ui.ButtonSet.OK);
}

// ==========================================
// SMART CALENDAR SYNC
// ==========================================
function findAvailableTimeSlot(calendar, targetDate) {
  // Create date object for the target day
  const checkDate = new Date(targetDate);
  checkDate.setHours(0, 0, 0, 0);
  
  const nextDay = new Date(checkDate);
  nextDay.setDate(nextDay.getDate() + 1);
  
  // Get all events for that day
  const existingEvents = calendar.getEvents(checkDate, nextDay);
  
  Logger.log(`Checking ${CONFIG.PREFERRED_TIME_SLOTS.length} time slots for ${checkDate.toDateString()}`);
  Logger.log(`Found ${existingEvents.length} existing events on this day`);
  
  // Try each preferred time slot
  for (let timeSlot of CONFIG.PREFERRED_TIME_SLOTS) {
    const [hours, minutes] = timeSlot.split(':');
    const slotStart = new Date(checkDate);
    slotStart.setHours(parseInt(hours), parseInt(minutes), 0, 0);
    
    const slotEnd = new Date(slotStart);
    slotEnd.setMinutes(slotEnd.getMinutes() + CONFIG.CHECKIN_DURATION);
    
    // Check if this slot conflicts with any existing events
    let hasConflict = false;
    
    for (let event of existingEvents) {
      const eventStart = event.getStartTime();
      const eventEnd = event.getEndTime();
      
      // Check for overlap
      if ((slotStart >= eventStart && slotStart < eventEnd) ||
          (slotEnd > eventStart && slotEnd <= eventEnd) ||
          (slotStart <= eventStart && slotEnd >= eventEnd)) {
        hasConflict = true;
        Logger.log(`  ${timeSlot} conflicts with: ${event.getTitle()}`);
        break;
      }
    }
    
    if (!hasConflict) {
      Logger.log(`  ✓ ${timeSlot} is available!`);
      return slotStart;
    }
  }
  
  // No available slots found, use default time
  Logger.log(`  No available slots, using default: ${CONFIG.DEFAULT_CHECKIN_TIME}`);
  const [hours, minutes] = CONFIG.DEFAULT_CHECKIN_TIME.split(':');
  const defaultSlot = new Date(checkDate);
  defaultSlot.setHours(parseInt(hours), parseInt(minutes), 0, 0);
  return defaultSlot;
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
// CALENDAR SYNC - Core Functions
// ==========================================
function syncWithCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const calendar = CalendarApp.getDefaultCalendar();
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) first.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const lastRow = activeSheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert('No projects found in this sheet.');
    return;
  }
  
  const data = activeSheet.getRange(2, 1, lastRow - 1, activeSheet.getLastColumn()).getValues();
  const changes = [];
  let syncEnabledCount = 0;
  let syncDisabledCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const uuid = data[i][CONFIG.UUID_COLUMN - 1];
    const syncEnabled = data[i][CONFIG.CALENDAR_SYNC_COLUMN - 1];
    const projectTitle = data[i][CONFIG.PROJECT_TITLE_COLUMN - 1];
    const completionDate = data[i][CONFIG.COMPLETION_DATE_COLUMN - 1];
    const nextCheckIn = data[i][CONFIG.NEXT_CHECKIN_COLUMN - 1];
    const status = data[i][CONFIG.STATUS_COLUMN - 1];
    
    if (!uuid) continue;
    
    // Check if project has dates but sync is disabled
    if (!syncEnabled && (completionDate || nextCheckIn)) {
      changes.push({
        type: 'sync-disabled',
        row: row,
        uuid: uuid,
        projectTitle: projectTitle,
        sheetDate: completionDate || nextCheckIn,
        status: status
      });
      syncDisabledCount++;
      continue;
    }
    
    if (!syncEnabled) {
      syncDisabledCount++;
      continue;
    }
    
    syncEnabledCount++;
    
    // Find calendar events for this project (could be 0, 1, or 2)
    const calendarEvents = findCalendarEventByUUID(calendar, uuid);
    
    // Separate check-in and completion events
    let checkInEvent = null;
    let completionEvent = null;
    
    calendarEvents.forEach(event => {
      const description = event.getDescription();
      if (description && description.includes('Type: check-in')) {
        checkInEvent = event;
      } else if (description && description.includes('Type: completion')) {
        completionEvent = event;
      }
    });
    
    // Check for missing check-in event
    if (!checkInEvent && completionDate) {
      changes.push({
        type: 'missing-checkin',
        row: row,
        uuid: uuid,
        projectTitle: projectTitle,
        sheetDate: completionDate,
        status: status
      });
    }
    
    // Check for missing completion event
    if (!completionEvent && completionDate) {
      changes.push({
        type: 'missing-completion',
        row: row,
        uuid: uuid,
        projectTitle: projectTitle,
        sheetDate: completionDate,
        status: status
      });
    }
    
    // Compare check-in date if both exist
    if (checkInEvent && completionDate) {
      const calDate = checkInEvent.getAllDayStartDate();
      const sheetDate = new Date(completionDate);
      sheetDate.setHours(0, 0, 0, 0);
      calDate.setHours(0, 0, 0, 0);
      
      if (calDate.getTime() !== sheetDate.getTime()) {
        changes.push({
          type: 'different-checkin',
          row: row,
          uuid: uuid,
          projectTitle: projectTitle,
          sheetDate: completionDate,
          calendarDate: calDate,
          eventId: checkInEvent.getId(),
          status: status
        });
      }
    }
    
    // Compare completion date if both exist
    if (completionEvent && completionDate) {
      const sheetDate = new Date(completionDate);
      sheetDate.setHours(0, 0, 0, 0);
      calDate.setHours(0, 0, 0, 0);
      
      if (calDate.getTime() !== sheetDate.getTime()) {
        changes.push({
          type: 'different-completion',
          row: row,
          uuid: uuid,
          projectTitle: projectTitle,
          sheetDate: completionDate,
          calendarDate: calDate,
          eventId: completionEvent.getId(),
          status: status
        });
      }
    }
  }
  
  if (changes.length === 0) {
    SpreadsheetApp.getUi().alert('✅ All In Sync', 
      `Sheet: ${sheetName}\n\n` +
      `Syncing ${syncEnabledCount} of ${syncEnabledCount + syncDisabledCount} projects.\n\n` +
      `All calendar events are up to date!\n\n` +
      `${syncDisabledCount} projects have Calendar Sync disabled.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Show changes dialog
  showSyncReviewDialog(changes, syncEnabledCount, syncDisabledCount);
}

function findCalendarEventByUUID(calendar, uuid) {
  const now = new Date();
  const oneYearAgo = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  const oneYearAhead = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));
  
  const events = calendar.getEvents(oneYearAgo, oneYearAhead);
  const foundEvents = [];
  
  for (let event of events) {
    const description = event.getDescription();
    if (description && description.includes(`UUID: ${uuid}`)) {
      foundEvents.push(event);
    }
  }
  
  return foundEvents; // Returns array - could be 0, 1, or 2 events (check-in + completion)
}

  
function showSyncReviewDialog(changes, syncEnabledCount, syncDisabledCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const sheetName = activeSheet ? activeSheet.getName() : 'unknown';
  
  let html = `
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; background: #f8f9fa; }
      .header { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #1a73e8; }
      .summary { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
      .change-item { 
        background: white; 
        padding: 15px; 
        margin-bottom: 12px; 
        border-radius: 8px;
        border-left: 4px solid #fbbc04;
      }
      .project-name { font-weight: bold; color: #202124; margin-bottom: 8px; }
      .date-info { font-size: 13px; color: #5f6368; margin: 5px 0; }
      .actions { margin-top: 10px; }
      .btn { 
        padding: 6px 12px; 
        margin-right: 8px; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer;
        font-size: 12px;
      }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-primary:hover { background: #1765cc; }
      .btn-success { background: #34a853; color: white; }
      .btn-success:hover { background: #2d9348; }
      .btn-secondary { background: #e8eaed; color: #202124; }
      .btn-secondary:hover { background: #d2d4d6; }
      .missing { border-left-color: #ea4335; }
      .different { border-left-color: #fbbc04; }
      .disabled-sync { border-left-color: #9aa0a6; }
      .bottom-actions {
        position: sticky;
        bottom: 0;
        background: white;
        padding: 15px;
        border-top: 2px solid #e0e0e0;
        text-align: right;
        box-shadow: 0 -2px 4px rgba(0,0,0,0.05);
      }
      .scrollable { max-height: 400px; overflow-y: auto; padding: 10px; }
    </style>
    <div class="header">📅 Calendar Sync Review - ${sheetName}</div>
    <div class="summary">
      Syncing ${syncEnabledCount} of ${syncEnabledCount + syncDisabledCount} projects<br>
      ${syncDisabledCount} projects have Calendar Sync disabled<br><br>
      <strong>Found ${changes.length} item(s) needing attention</strong>
    </div>
    <div class="scrollable">
  `;
  
  changes.forEach((change, index) => {
    const dateFormat = 'MM/dd/yyyy';
    let cssClass = 'different';
    
    if (change.type.includes('missing')) {
      cssClass = 'missing';
    } else if (change.type === 'sync-disabled') {
      cssClass = 'disabled-sync';
    }
    
    html += `<div class="change-item ${cssClass}" data-row="${change.row}" data-uuid="${change.uuid}" data-type="${change.type}">`;
    html += `<div class="project-name">${change.projectTitle}</div>`;
    
    if (change.type === 'sync-disabled') {
      html += `<div class="date-info">⏸️ Calendar Sync: DISABLED</div>`;
      html += `<div class="date-info">Column U needs to be checked to sync this project</div>`;
      html += `<div class="actions">`;
      html += `<button class="btn btn-success" onclick="enableSyncAndCreate(${change.row}, '${change.uuid}')">Enable Sync & Create Events</button>`;
      html += `</div>`;
    } else if (change.type === 'missing-checkin') {
      html += `<div class="date-info">📋 Sheet Next Check In: ${Utilities.formatDate(new Date(change.sheetDate), Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="date-info">📅 Calendar Check-in: <strong>NOT FOUND</strong></div>`;
      html += `<div class="date-info">Status: ${change.status}</div>`;
      html += `<div class="actions">`;
      html += `<button class="btn btn-primary" onclick="createEventsForRow(${change.row})">Create Events</button>`;
      html += `</div>`;
    } else if (change.type === 'missing-completion') {
      html += `<div class="date-info">📋 Sheet Completion: ${Utilities.formatDate(new Date(change.sheetDate), Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="date-info">📅 Calendar Deadline: <strong>NOT FOUND</strong></div>`;
      html += `<div class="date-info">Status: ${change.status}</div>`;
      html += `<div class="actions">`;
      html += `<button class="btn btn-primary" onclick="createEventsForRow(${change.row})">Create Events</button>`;
      html += `</div>`;
    } else if (change.type === 'different-checkin') {
      html += `<div class="date-info">📋 Sheet Next Check In: ${Utilities.formatDate(new Date(change.sheetDate), Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="date-info">📅 Calendar Check-in: ${Utilities.formatDate(change.calendarDate, Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="actions">`;
      html += `<button class="btn btn-primary" onclick="updateSheetFromCalendar(${change.row}, '${change.calendarDate.toISOString()}', 'checkin')">Update Sheet</button>`;
      html += `<button class="btn btn-secondary" onclick="updateCalendarFromSheet(${change.row}, 'checkin')">Update Calendar</button>`;
      html += `</div>`;
    } else if (change.type === 'different-completion') {
      html += `<div class="date-info">📋 Sheet Completion: ${Utilities.formatDate(new Date(change.sheetDate), Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="date-info">📅 Calendar Deadline: ${Utilities.formatDate(change.calendarDate, Session.getScriptTimeZone(), dateFormat)}</div>`;
      html += `<div class="actions">`;
      html += `<button class="btn btn-primary" onclick="updateSheetFromCalendar(${change.row}, '${change.calendarDate.toISOString()}', 'completion')">Update Sheet</button>`;
      html += `<button class="btn btn-secondary" onclick="updateCalendarFromSheet(${change.row}, 'completion')">Update Calendar</button>`;
      html += `</div>`;
    }
    
    html += `</div>`;
  });
  
  html += `</div>`;
  
  html += `<div class="bottom-actions">`;
  html += `<button class="btn btn-success" onclick="createAllMissingEvents()">Create All Missing Events</button>`;
  html += `<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>`;
  html += `</div>`;
  
  html += `
    <script>
      function enableSyncAndCreate(row, uuid) {
        const btn = event.target;
        btn.disabled = true;
        btn.textContent = 'Enabling...';
        google.script.run
          .withSuccessHandler(function() {
            btn.textContent = 'Enabled!';
            btn.style.background = '#34a853';
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error);
            btn.disabled = false;
            btn.textContent = 'Enable Sync & Create Events';
          })
          .enableSyncAndCreateEvents(row);
      }
      
      function createEventsForRow(row) {
        const btn = event.target;
        btn.disabled = true;
        btn.textContent = 'Creating...';
        google.script.run
          .withSuccessHandler(function() {
            btn.textContent = 'Created!';
            btn.style.background = '#34a853';
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error);
            btn.disabled = false;
            btn.textContent = 'Create Events';
          })
          .createCalendarEventsFromDialog(row);
      }
      
      function createAllMissingEvents() {
        const btn = event.target;
        btn.disabled = true;
        btn.textContent = 'Creating All...';
        google.script.run
          .withSuccessHandler(function(result) {
            alert('Created ' + result + ' event(s)!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error);
            btn.disabled = false;
            btn.textContent = 'Create All Missing Events';
          })
          .createAllMissingEventsFromDialog();
      }
      
      function updateSheetFromCalendar(row, calendarDate, eventType) {
        alert('Update sheet from calendar - coming soon!');
      }
      
      function updateCalendarFromSheet(row, eventType) {
        alert('Update calendar from sheet - coming soon!');
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Calendar Sync Review');
}

// ==========================================
// CALENDAR SYNC - Individual Actions
// ==========================================
function enableCalendarSyncForProject() {
  const activeSheet = getActiveProjectSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project row in a tracked sheet (cobuild or enablement).');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row (not the header).');
    return;
  }
  
  activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).setValue(true);
  
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const sheetName = activeSheet.getName();
  const response = SpreadsheetApp.getUi().alert(
    '✅ Calendar Sync Enabled',
    `Sheet: ${sheetName}\nProject: ${projectTitle}\n\nCalendar sync enabled!\n\nDo you want to create a calendar event now?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    Logger.log(`Creating calendar events for row ${row} in sheet ${sheetName}`);
    createCalendarEventForRow(row, activeSheet);
  }
}

function disableCalendarSyncForProject() {
  const activeSheet = getActiveProjectSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project row in a tracked sheet (cobuild or enablement).');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row (not the header).');
    return;
  }
  
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const uuid = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  
  activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).setValue(false);
  
  const response = SpreadsheetApp.getUi().alert(
    '⏸️ Calendar Sync Disabled',
    `Calendar sync disabled for: ${projectTitle}\n\nDo you want to remove the calendar event?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    removeCalendarEventForUUID(uuid, projectTitle);
  }
}

function createCalendarEventForRow(row, sourceSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = sourceSheet || getActiveProjectSheet();
  
  if (!activeSheet) {
    Logger.log('ERROR: No active project sheet found');
    SpreadsheetApp.getUi().alert('❌ Error', 'Select a project in a tracked sheet', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!row || row < 2) {
    Logger.log('ERROR: Invalid row number');
    SpreadsheetApp.getUi().alert('❌ Error', 'Please select a valid project row', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const sheetName = activeSheet.getName();
  const calendar = CalendarApp.getDefaultCalendar();
  
  // Safely get values with null checks
  const uuid = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue() || '';
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue() || '';
  const completionDate = activeSheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getValue() || null;
  const nextCheckIn = activeSheet.getRange(row, CONFIG.NEXT_CHECKIN_COLUMN).getValue() || null;
  const status = activeSheet.getRange(row, CONFIG.STATUS_COLUMN).getValue() || '';
  const nextSteps = activeSheet.getRange(row, CONFIG.ACTIVITY_COLUMN).getValue() || '';
  
  // Debug logging
  Logger.log('=== Creating Calendar Events ===');
  Logger.log(`Sheet: ${sheetName}`);
  Logger.log(`Row: ${row}`);
  Logger.log(`UUID: ${uuid}`);
  Logger.log(`Project Title: ${projectTitle}`);
  Logger.log(`Completion Date: ${completionDate}`);
  Logger.log(`Completion Date type: ${typeof completionDate}`);
  Logger.log(`Next Check In: ${nextCheckIn}`);
  Logger.log(`Next Check In type: ${typeof nextCheckIn}`);
  Logger.log(`Status: ${status}`);
  Logger.log(`Calendar: ${calendar.getName()}`);
  
  // Validate we have at least one date
  if (!completionDate && !nextCheckIn) {
    const message = `Project "${projectTitle}" needs at least one date:\n\n` +
      `• Column I (Completion Date): Currently ${completionDate ? 'SET' : 'EMPTY'}\n` +
      `• Column M (Next Check In): Currently ${nextCheckIn ? 'SET' : 'EMPTY'}\n\n` +
      `Please set at least one date before creating calendar events.`;
    Logger.log(`ERROR: ${message}`);
    SpreadsheetApp.getUi().alert('❌ No Dates Set', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!projectTitle) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Project must have a title (Column C)', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const description = `Project: ${projectTitle}\nStatus: ${status}\nNext Steps: ${nextSteps}\n\nUUID: ${uuid}\nType: {{TYPE}}\nSpreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${activeSheet.getSheetId()}&range=A${row}`;
  
  let eventsCreated = 0;
  let messages = [];
  
  try {
    // Create Check-In Event (if Next Check In date exists) - TIME BLOCKED
    if (nextCheckIn && nextCheckIn instanceof Date && !isNaN(nextCheckIn)) {
      const checkInDate = new Date(nextCheckIn);
      
      // Find available time slot
      const checkInStartTime = findAvailableTimeSlot(calendar, checkInDate);
      const checkInEndTime = new Date(checkInStartTime);
      checkInEndTime.setMinutes(checkInEndTime.getMinutes() + CONFIG.CHECKIN_DURATION);
      
      const checkInTitle = `Check-in [${sheetName}]: ${projectTitle}`;
      const checkInDescription = description.replace('{{TYPE}}', 'check-in');
      
      Logger.log(`Creating check-in event: ${checkInTitle}`);
      Logger.log(`Time: ${checkInStartTime} to ${checkInEndTime}`);
      
      const checkInEvent = calendar.createEvent(checkInTitle, checkInStartTime, checkInEndTime, {
        description: checkInDescription
      });
      checkInEvent.setColor(CalendarApp.EventColor.BLUE); // Blue for check-ins
      checkInEvent.addPopupReminder(1440); // 1 day before
      
      Logger.log(`Check-in event created! Event ID: ${checkInEvent.getId()}`);
      eventsCreated++;
      messages.push(`✓ Check-in: ${Utilities.formatDate(checkInDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')} at ${formatTime(checkInStartTime)} (Blue, ${CONFIG.CHECKIN_DURATION} min)`);
    } else if (nextCheckIn) {
      Logger.log(`SKIPPED: Next Check In is not a valid date: ${nextCheckIn}`);
    }
    
    // Create Completion Deadline Event (if Completion Date exists) - ALL DAY
    if (completionDate && completionDate instanceof Date && !isNaN(completionDate)) {
      const completionEventDate = new Date(completionDate);
      const completionTitle = `DUE [${sheetName}]: ${projectTitle}`;
      const completionDescription = description.replace('{{TYPE}}', 'completion');
      
      Logger.log(`Creating completion event: ${completionTitle} on ${completionEventDate}`);
      
      const completionEvent = calendar.createAllDayEvent(completionTitle, completionEventDate, {
        description: completionDescription
      });
      completionEvent.setColor(CalendarApp.EventColor.ORANGE); // Orange for deadlines
      completionEvent.addPopupReminder(1440); // 1 day before
      completionEvent.addPopupReminder(10080); // 1 week before
      
      Logger.log(`Completion event created! Event ID: ${completionEvent.getId()}`);
      eventsCreated++;
      messages.push(`✓ Deadline: ${Utilities.formatDate(completionEventDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')} (Orange, all-day)`);
    } else if (completionDate) {
      Logger.log(`SKIPPED: Completion Date is not a valid date: ${completionDate}`);
    }
    
    if (eventsCreated === 0) {
      SpreadsheetApp.getUi().alert('❌ No Dates Set', 
        `Project "${projectTitle}" needs at least one date:\n\n` +
        `• Column I (Completion Date) for deadline event\n` +
        `• Column M (Next Check In) for check-in event\n\n` +
        `Please set dates before creating calendar events.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    SpreadsheetApp.getUi().alert('✅ Calendar Events Created',
      `${eventsCreated} event(s) created for: ${projectTitle}\n\n` +
      messages.join('\n') + '\n\n' +
      `Check your Google Calendar!\n\n` +
      `Legend:\n` +
      `• 🔵 Blue = Check-ins (30-min time blocks)\n` +
      `• 🟠 Orange = Deadlines (all-day events)\n\n` +
      `You can drag check-in events to different times.\n` +
      `Next sync will update the sheet from calendar.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    Logger.log(`ERROR Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert('❌ Error Creating Event', 
      `Could not create calendar event(s).\n\n` +
      `Error: ${error.message}\n\n` +
      `Events created: ${eventsCreated}\n\n` +
      `Check the logs (View → Executions) for more details.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function removeCalendarEventForUUID(uuid, projectTitle) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = findCalendarEventByUUID(calendar, uuid);
  
  if (events.length > 0) {
    events.forEach(event => event.deleteEvent());
    SpreadsheetApp.getUi().alert('✅ Calendar Events Removed',
      `${events.length} calendar event(s) removed for: ${projectTitle}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No calendar events found for this project.');
  }
}

function viewCalendarSyncStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) first.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  const lastRow = activeSheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert('No projects found.');
    return;
  }
  
  const data = activeSheet.getRange(2, 1, lastRow - 1, activeSheet.getLastColumn()).getValues();
  let enabledProjects = [];
  let disabledProjects = [];
  
  for (let i = 0; i < data.length; i++) {
    const syncEnabled = data[i][CONFIG.CALENDAR_SYNC_COLUMN - 1];
    const projectTitle = data[i][CONFIG.PROJECT_TITLE_COLUMN - 1];
    const status = data[i][CONFIG.STATUS_COLUMN - 1];
    
    if (syncEnabled) {
      enabledProjects.push(`${projectTitle} (${status})`);
    } else {
      disabledProjects.push(`${projectTitle} (${status})`);
    }
  }
  
  let message = `📊 Calendar Sync Status\n\nSheet: ${sheetName}\n\n`;
  message += `✅ Synced to Calendar (${enabledProjects.length}):\n`;
  message += enabledProjects.slice(0, 10).join('\n');
  if (enabledProjects.length > 10) message += `\n... and ${enabledProjects.length - 10} more`;
  
  message += `\n\n⏸️ Not Synced (${disabledProjects.length}):\n`;
  message += disabledProjects.slice(0, 10).join('\n');
  if (disabledProjects.length > 10) message += `\n... and ${disabledProjects.length - 10} more`;
  
  SpreadsheetApp.getUi().alert('Calendar Sync Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// TASKS SHEET SETUP & MANAGEMENT
// ==========================================
function setupTasksSheet(sheet) {
  // Set up headers
  const headers = [
    'Task ID',           // A - Auto-generated
    'Project UUID',      // B - Links to cobuild
    'Project Name',      // C - Readable name
    'Task Description',  // D - What needs to be done
    'Task Type',         // E - Follow-up, Milestone, Check-in, Deliverable
    'Due Date',          // F - When it's due
    'Due Time',          // G - Optional specific time
    'Duration (min)',    // H - 15, 30, 60
    'Status',            // I - Not Started, In Progress, Complete, Blocked
    'Priority',          // J - High, Medium, Low
    'Assigned To',       // K - Who's responsible
    'Source',            // L - Tasks, cobuild, Google Tasks, Calendar
    'Notes',             // M - Additional details
    'Calendar Sync',     // N - Checkbox
    'Created Date',      // O - When task was added
    'Completed Date',    // P - When marked done
    'Last Modified'      // Q - Last change timestamp
  ];
  
  sheet.appendRow(headers);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
             .setBackground('#4285F4')
             .setFontColor('white')
             .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 100);  // Task ID
  sheet.setColumnWidth(2, 120);  // Project UUID
  sheet.setColumnWidth(3, 150);  // Project Name
  sheet.setColumnWidth(4, 250);  // Task Description
  sheet.setColumnWidth(5, 120);  // Task Type
  sheet.setColumnWidth(6, 100);  // Due Date
  sheet.setColumnWidth(7, 100);  // Due Time
  sheet.setColumnWidth(8, 100);  // Duration
  sheet.setColumnWidth(9, 120);  // Status
  sheet.setColumnWidth(10, 100); // Priority
  sheet.setColumnWidth(11, 120); // Assigned To
  sheet.setColumnWidth(12, 100); // Source
  sheet.setColumnWidth(13, 200); // Notes
  sheet.setColumnWidth(14, 100); // Calendar Sync
  sheet.setColumnWidth(15, 120); // Created Date
  sheet.setColumnWidth(16, 120); // Completed Date
  sheet.setColumnWidth(17, 120); // Last Modified
  
  Logger.log('Tasks sheet setup complete - data validation will be applied as rows are added');
}

function generateTaskID() {
  return 'task_' + Utilities.getUuid().substring(0, 8);
}

// Apply data validation to a specific task row
function applyTaskRowValidation(sheet, row) {
  // Task Type dropdown
  const taskTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Follow-up', 'Milestone', 'Check-in', 'Deliverable'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, 5).setDataValidation(taskTypeRule);
  
  // Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'In Progress', 'Complete', 'Blocked', 'Cancelled'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, 9).setDataValidation(statusRule);
  
  // Priority dropdown
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, 10).setDataValidation(priorityRule);
  
  // Calendar Sync checkbox
  const calSyncRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheet.getRange(row, 14).setDataValidation(calSyncRule);
}

// ==========================================
// TASKS SHEET CLEANUP
// ==========================================
function cleanupTasksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  if (!tasksSheet) {
    ui.alert('Tasks sheet not found.');
    return;
  }
  
  // STEP 1: Analyze what will be removed
  const allData = tasksSheet.getDataRange().getValues();
  const headerRow = allData[0];
  
  Logger.log(`Total rows in sheet: ${allData.length}`);
  
  // Find rows with actual task data
  const dataRows = [];
  const emptyRows = [];
  
  for (let i = 1; i < allData.length; i++) {
    const taskID = allData[i][0]; // Column A
    const projectUUID = allData[i][1]; // Column B
    const taskDesc = allData[i][3]; // Column D
    
    // Row has data if Task ID exists OR Project UUID exists OR Description exists
    if (taskID || projectUUID || taskDesc) {
      dataRows.push({
        rowNum: i + 1,
        data: allData[i],
        taskID: taskID,
        projectName: allData[i][2],
        description: taskDesc
      });
      Logger.log(`Row ${i + 1}: HAS DATA - ${taskID} - ${taskDesc}`);
    } else {
      emptyRows.push(i + 1);
    }
  }
  
  Logger.log(`Tasks with data: ${dataRows.length}`);
  Logger.log(`Empty rows: ${emptyRows.length}`);
  
  // STEP 2: Show preview and confirm
  let previewMessage = '🔍 CLEANUP PREVIEW\n\n';
  previewMessage += `Total rows: ${allData.length}\n`;
  previewMessage += `Empty rows to remove: ${emptyRows.length}\n`;
  previewMessage += `Tasks to preserve: ${dataRows.length}\n\n`;
  
  if (dataRows.length > 0) {
    previewMessage += '✅ TASKS TO KEEP:\n';
    dataRows.slice(0, 5).forEach(task => {
      previewMessage += `  • Row ${task.rowNum}: ${task.description || task.taskID}\n`;
    });
    if (dataRows.length > 5) {
      previewMessage += `  ... and ${dataRows.length - 5} more\n`;
    }
  }
  
  previewMessage += '\n❌ EMPTY ROWS TO REMOVE:\n';
  if (emptyRows.length > 10) {
    previewMessage += `  Rows ${emptyRows[0]} - ${emptyRows[emptyRows.length - 1]}\n`;
  } else if (emptyRows.length > 0) {
    previewMessage += `  ${emptyRows.join(', ')}\n`;
  } else {
    previewMessage += '  None!\n';
  }
  
  previewMessage += '\n⚠️ A backup will be created first.\n\nContinue?';
  
  const response = ui.alert('Cleanup Preview', previewMessage, ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    ui.alert('Cleanup cancelled. No changes made.');
    return;
  }
  
  // STEP 3: Create backup first
  try {
    const backupSheet = tasksSheet.copyTo(ss);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    backupSheet.setName(`Tasks_Backup_${timestamp}`);
    ss.moveActiveSheet(ss.getNumSheets()); // Move to end
    Logger.log(`Backup created: Tasks_Backup_${timestamp}`);
  } catch (error) {
    ui.alert('❌ Backup Failed', 
      `Could not create backup: ${error.message}\n\nCleanup cancelled for safety.`,
      ui.ButtonSet.OK);
    return;
  }
  
  // STEP 4: Perform cleanup
  try {
    tasksSheet.clear();
    tasksSheet.clearFormats();
    
    tasksSheet.appendRow(headerRow);
    const headerRange = tasksSheet.getRange(1, 1, 1, headerRow.length);
    headerRange.setFontWeight('bold').setBackground('#4285F4').setFontColor('white').setHorizontalAlignment('center');
    tasksSheet.setFrozenRows(1);
    
    // Set column widths
    tasksSheet.setColumnWidth(1, 100);
    tasksSheet.setColumnWidth(2, 120);
    tasksSheet.setColumnWidth(3, 150);
    tasksSheet.setColumnWidth(4, 250);
    tasksSheet.setColumnWidth(5, 120);
    tasksSheet.setColumnWidth(6, 100);
    tasksSheet.setColumnWidth(7, 100);
    tasksSheet.setColumnWidth(8, 100);
    tasksSheet.setColumnWidth(9, 120);
    tasksSheet.setColumnWidth(10, 100);
    tasksSheet.setColumnWidth(11, 120);
    tasksSheet.setColumnWidth(12, 100);
    tasksSheet.setColumnWidth(13, 200);
    tasksSheet.setColumnWidth(14, 100);
    tasksSheet.setColumnWidth(15, 120);
    tasksSheet.setColumnWidth(16, 120);
    tasksSheet.setColumnWidth(17, 120);
    
    if (dataRows.length > 0) {
      for (let i = 0; i < dataRows.length; i++) {
        tasksSheet.appendRow(dataRows[i].data);
        const row = i + 2;
        applyTaskRowValidation(tasksSheet, row);
        tasksSheet.getRange(row, 6).setNumberFormat('M/d/yyyy');
        tasksSheet.getRange(row, 15).setNumberFormat('M/d/yyyy h:mm');
        tasksSheet.getRange(row, 17).setNumberFormat('M/d/yyyy h:mm');
      }
    }
    
    ui.alert('✅ Cleanup Complete!', 
      `Tasks preserved: ${dataRows.length}\n` +
      `Empty rows removed: ${emptyRows.length}\n\n` +
      `Backup created: Check the backup sheet at the end\n\n` +
      `Next task will go to row ${dataRows.length + 2}`,
      ui.ButtonSet.OK);
      
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ui.alert('❌ Cleanup Error',
      `Error during cleanup: ${error.message}\n\n` +
      `Your backup sheet is safe! You can restore from it.`,
      ui.ButtonSet.OK);
  }
}

// ==========================================
// ADD TASK FOR PROJECT (V2 - with defaults)
// ==========================================
function addTaskForProjectV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  // Check if we're on a tracked sheet
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) and select a project row.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  
  // Get selected project
  const activeRow = activeSheet.getActiveCell().getRow();
  
  if (activeRow === 1) {
    ui.alert('Select a project row (not the header) to add a task.');
    return;
  }
  
  const projectUUID = activeSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  Logger.log('=== Adding Task V2 ===');
  Logger.log(`Active Row: ${activeRow}`);
  Logger.log(`Project UUID: ${projectUUID}`);
  Logger.log(`Project Name: ${projectName}`);
  
  if (!projectUUID || !projectName) {
    ui.alert('❌ Error', 
      `Selected project is missing UUID or title.\n\n` +
      `UUID: ${projectUUID || 'MISSING'}\n` +
      `Name: ${projectName || 'MISSING'}\n\n` +
      `Run "Setup Tracker" to generate UUIDs.`,
      ui.ButtonSet.OK);
    return;
  }
  
  if (!tasksSheet) {
    ui.alert('❌ Error', 
      'Tasks sheet not found. Run "Setup Tracker" first.',
      ui.ButtonSet.OK);
    return;
  }
  
  // STEP 1: Description
  const descResponse = ui.prompt(
    'Add Task - Step 1 of 4',
    `Sheet: ${sheetName}\nProject: ${projectName}\n\nEnter task description:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;
  const taskDescription = descResponse.getResponseText().trim();
  if (!taskDescription) {
    ui.alert('Task description cannot be empty.');
    return;
  }
  
  // STEP 2: Task Type (default: Follow-up)
  const typeResponse = ui.prompt(
    'Add Task - Step 2 of 4',
    'Task type (default: Follow-up):\n\n1 = Follow-up\n2 = Milestone\n3 = Check-in\n4 = Deliverable\n\nEnter 1-4 or leave blank:',
    ui.ButtonSet.OK_CANCEL
  );
  if (typeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const typeInput = typeResponse.getResponseText().trim();
  const typeMap = {'1': 'Follow-up', '2': 'Milestone', '3': 'Check-in', '4': 'Deliverable', '': 'Follow-up'};
  const taskType = typeMap[typeInput] || 'Follow-up';
  
  Logger.log(`Type selected: ${taskType}`);
  
  // STEP 3: Due Date (default: none)
  const dateResponse = ui.prompt(
    'Add Task - Step 3 of 4',
    'Due date (default: none):\n\n+N = N days from today (e.g. +5, +7)\nMM/DD/YYYY = specific date\nLeave blank = no due date',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  
  let dueDate = null;
  const dateText = dateResponse.getResponseText().trim();
  
  if (dateText !== '') {
    if (dateText.startsWith('+')) {
      const days = parseInt(dateText.substring(1));
      if (!isNaN(days) && days > 0) {
        dueDate = new Date();
        dueDate.setDate(dueDate.getDate() + days);
      }
    } else {
      const customDate = new Date(dateText);
      if (!isNaN(customDate.getTime())) {
        dueDate = customDate;
      }
    }
  }
  
  Logger.log(`Due date: ${dueDate}`);
  
  // STEP 4: Priority (default: Low)
  const priorityResponse = ui.prompt(
    'Add Task - Step 4 of 4',
    'Priority (default: Low):\n\n1 = High\n2 = Medium\n3 = Low\n\nEnter 1-3 or leave blank:',
    ui.ButtonSet.OK_CANCEL
  );
  if (priorityResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const priorityInput = priorityResponse.getResponseText().trim();
  const priorityMap = {'1': 'High', '2': 'Medium', '3': 'Low', '': 'Low'};
  const priority = priorityMap[priorityInput] || 'Low';
  
  Logger.log(`Priority: ${priority}`);
  
  // Create the task
  createTaskInSheet(projectUUID, projectName, taskDescription, taskType, dueDate, priority);
}

// ==========================================
// CREATE TASK IN SHEET (shared function)
// ==========================================
function createTaskInSheet(projectUUID, projectName, taskDescription, taskType, dueDate, priority) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  Logger.log('=== createTaskInSheet called ===');
  Logger.log(`Task Description: ${taskDescription}`);
  Logger.log(`Priority received: ${priority}`);
  
  if (!tasksSheet) {
    Logger.log('ERROR: Tasks sheet not found');
    ui.alert('❌ Error', 'Tasks sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  // Ensure priority has a value
  if (!priority || priority === 'undefined') {
    priority = 'Low';
    Logger.log('Priority was undefined, defaulting to Low');
  }
  
  const taskID = generateTaskID();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  Logger.log('=== Creating Task ===');
  Logger.log(`Task ID: ${taskID}`);
  Logger.log(`Project UUID: ${projectUUID}`);
  Logger.log(`Project Name: ${projectName}`);
  Logger.log(`Description: ${taskDescription}`);
  Logger.log(`Type: ${taskType}`);
  Logger.log(`Due Date: ${dueDate}`);
  Logger.log(`Priority: ${priority}`);
  
  // DUPLICATE PREVENTION: Check if this exact task was just created (within last 5 seconds)
  const allTasks = tasksSheet.getDataRange().getValues();
  const fiveSecondsAgo = new Date(now.getTime() - 5000);
  
  for (let i = 1; i < allTasks.length; i++) {
    const existingDesc = allTasks[i][3]; // Column D - Description
    const existingProjectUUID = allTasks[i][1]; // Column B - Project UUID
    const existingCreated = allTasks[i][14]; // Column O - Created Date
    
    if (existingDesc === taskDescription && 
        existingProjectUUID === projectUUID &&
        existingCreated instanceof Date &&
        existingCreated > fiveSecondsAgo) {
      Logger.log('DUPLICATE DETECTED - Task already exists, skipping creation');
      ui.alert('⚠️ Duplicate Detected',
        `This task was just created:\n"${taskDescription}"\n\nSkipping duplicate creation.`,
        ui.ButtonSet.OK);
      
      // Navigate to the existing task
      ss.setActiveSheet(tasksSheet);
      tasksSheet.setActiveRange(tasksSheet.getRange(i + 1, 1, 1, 17));
      return;
    }
  }
  
  Logger.log('No duplicate found, proceeding with task creation');
  
  try {
    tasksSheet.appendRow([
      taskID,                    // A - Task ID
      '',                        // B - Parent Task ID (empty for main tasks)
      projectUUID,               // C - Project UUID
      projectName,               // D - Project Name
      taskDescription,           // E - Task Description
      taskType,                  // F - Task Type
      dueDate,                   // G - Due Date
      '',                        // H - Due Time
      30,                        // I - Duration (default 30 min)
      'Not Started',             // J - Status
      priority,                  // K - Priority
      userEmail,                 // L - Assigned To
      'Tasks',                   // M - Source
      '',                        // N - Notes
      false,                     // O - Calendar Sync (default off)
      '',                        // P - Google Task ID (empty initially)
      now,                       // Q - Created Date
      '',                        // R - Completed Date
      now                        // S - Last Modified
    ]);
    
    Logger.log('Task row appended successfully');
    
    // Get the actual last row with data
    const lastRow = tasksSheet.getLastRow();
    Logger.log(`Task added to row: ${lastRow}`);
    
    // Apply data validation to this specific row
    applyTaskRowValidation(tasksSheet, lastRow);
    
    // Format the new row
    if (dueDate) {
      tasksSheet.getRange(lastRow, 7).setNumberFormat('M/d/yyyy'); // Due Date (now column G)
    }
    tasksSheet.getRange(lastRow, 17).setNumberFormat('M/d/yyyy h:mm'); // Created Date (now column Q)
    tasksSheet.getRange(lastRow, 19).setNumberFormat('M/d/yyyy h:mm'); // Last Modified (now column S)
    
    Logger.log('Task formatting complete');
    
    // Auto-sync to Google Tasks if enabled
    if (CONFIG.GOOGLE_TASKS_AUTO_SYNC) {
      Logger.log('Auto-sync enabled, syncing to Google Tasks...');
      syncTaskToGoogleTasks(taskID);
    }
    
    ui.alert('✅ Task Created',
      `Task added to Tasks sheet (Row ${lastRow}):\n\n` +
      `${taskDescription}\n\n` +
      `Project: ${projectName}\n` +
      `Type: ${taskType}\n` +
      `Priority: ${priority}\n` +
      `Due: ${dueDate ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'Not set'}\n\n` +
      `Task ID: ${taskID}`,
      ui.ButtonSet.OK);
      
    // Navigate to Tasks sheet only if configured to do so
    if (CONFIG.NAVIGATE_TO_TASK_AFTER_CREATE) {
      ss.setActiveSheet(tasksSheet);
      tasksSheet.setActiveRange(tasksSheet.getRange(lastRow, 1, 1, 19));
    } else {
      // Stay on current sheet but show success message included option to view
      const viewResponse = ui.alert('✅ Task Created',
        `Task: ${taskDescription}\n\n` +
        `Would you like to view it in the Tasks sheet?`,
        ui.ButtonSet.YES_NO);
      
      if (viewResponse === ui.Button.YES) {
        ss.setActiveSheet(tasksSheet);
        tasksSheet.setActiveRange(tasksSheet.getRange(lastRow, 1, 1, 19));
      }
    }
    
  } catch (error) {
    Logger.log(`ERROR creating task: ${error.message}`);
    Logger.log(`ERROR stack: ${error.stack}`);
    ui.alert('❌ Error Creating Task',
      `Could not create task.\n\n` +
      `Error: ${error.message}\n\n` +
      `Check View → Executions for details.`,
      ui.ButtonSet.OK);
  }
}

// ==========================================
// ADD TASK FOR PROJECT (OLD VERSION - keep for compatibility)
// ==========================================
function addTaskForProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  // Check if we're on the cobuild sheet
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== CONFIG.SHEETS.PROJECTS) {
    ui.alert('Please switch to the cobuild sheet and select a project row.');
    return;
  }
  
  // Get selected project
  const activeRow = activeSheet.getActiveCell().getRow();
  
  if (activeRow === 1) {
    ui.alert('Select a project row (not the header) to add a task.');
    return;
  }
  
  const projectUUID = projectSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = projectSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  Logger.log('=== Adding Task ===');
  Logger.log(`Active Row: ${activeRow}`);
  Logger.log(`Project UUID: ${projectUUID}`);
  Logger.log(`Project Name: ${projectName}`);
  
  if (!projectUUID || !projectName) {
    ui.alert('❌ Error', 
      `Selected project is missing UUID or title.\n\n` +
      `UUID: ${projectUUID || 'MISSING'}\n` +
      `Name: ${projectName || 'MISSING'}\n\n` +
      `Run "Setup Tracker" to generate UUIDs.`,
      ui.ButtonSet.OK);
    return;
  }
  
  if (!tasksSheet) {
    ui.alert('❌ Error', 
      'Tasks sheet not found. Run "Setup Tracker" first.',
      ui.ButtonSet.OK);
    return;
  }
  
  // Prompt for task details
  const descResponse = ui.prompt(
    'Add Task - Step 1 of 4',
    `Project: ${projectName}\n\nEnter task description:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;
  const taskDescription = descResponse.getResponseText().trim();
  if (!taskDescription) {
    ui.alert('Task description cannot be empty.');
    return;
  }
  
  // Task type
  const typeResponse = ui.prompt(
    'Add Task - Step 2 of 4',
    'Select task type:\n1 = Follow-up\n2 = Milestone\n3 = Check-in\n4 = Deliverable\n\nEnter number (1-4):',
    ui.ButtonSet.OK_CANCEL
  );
  if (typeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const typeMap = {'1': 'Follow-up', '2': 'Milestone', '3': 'Check-in', '4': 'Deliverable'};
  const taskType = typeMap[typeResponse.getResponseText().trim()] || 'Follow-up';
  
  // Due date - use HTML date picker
  const dateHtml = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-bottom: 10px; font-weight: bold; }
      input[type="date"] { 
        width: 100%; 
        padding: 8px; 
        font-size: 14px; 
        border: 1px solid #ccc;
        border-radius: 4px;
        margin-bottom: 15px;
      }
      button { 
        background: #4285F4; 
        color: white; 
        padding: 10px 20px; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer;
        font-size: 14px;
        margin-right: 10px;
      }
      button:hover { background: #357ae8; }
      .cancel { background: #ccc; color: #333; }
      .cancel:hover { background: #bbb; }
      .info { color: #5f6368; font-size: 12px; margin-top: 10px; }
    </style>
    <div>
      <label>Add Task - Step 3 of 4</label>
      <label for="dueDate">Select due date (optional):</label>
      <input type="date" id="dueDate" />
      <div class="info">Leave blank if no specific due date</div>
      <button onclick="submitDate()">Next</button>
      <button class="cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    <script>
      function submitDate() {
        const date = document.getElementById('dueDate').value;
        google.script.run
          .withSuccessHandler(function() {
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error);
          })
          .continueAddTaskWithDate('${taskDescription}', '${taskType}', date);
      }
    </script>
  `).setWidth(400).setHeight(250);
  
  ui.showModalDialog(dateHtml, 'Add Task - Due Date');
}

// Continue task creation after date is selected
function continueAddTaskWithDate(taskDescription, taskType, dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  const activeSheet = ss.getActiveSheet();
  const activeRow = activeSheet.getActiveCell().getRow();
  
  const projectUUID = projectSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = projectSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  let dueDate = null;
  if (dateString && dateString.trim() !== '') {
    dueDate = new Date(dateString);
    if (isNaN(dueDate)) {
      dueDate = null;
    }
  }
  
  // Priority
  const priorityResponse = ui.prompt(
    'Add Task - Step 4 of 4',
    'Select priority:\n1 = High\n2 = Medium\n3 = Low\n\nEnter number (1-3):',
    ui.ButtonSet.OK_CANCEL
  );
  if (priorityResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const priorityMap = {'1': 'High', '2': 'Medium', '3': 'Low'};
  const priority = priorityMap[priorityResponse.getResponseText().trim()] || 'Medium';
  
  // Create task
  const taskID = generateTaskID();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  
  Logger.log('Creating task with:');
  Logger.log(`Task ID: ${taskID}`);
  Logger.log(`Description: ${taskDescription}`);
  Logger.log(`Type: ${taskType}`);
  Logger.log(`Due Date: ${dueDate}`);
  Logger.log(`Priority: ${priority}`);
  
  try {
    tasksSheet.appendRow([
      taskID,                    // A - Task ID
      projectUUID,               // B - Project UUID
      projectName,               // C - Project Name
      taskDescription,           // D - Task Description
      taskType,                  // E - Task Type
      dueDate,                   // F - Due Date
      '',                        // G - Due Time
      30,                        // H - Duration (default 30 min)
      'Not Started',             // I - Status
      priority,                  // J - Priority
      userEmail,                 // K - Assigned To
      'Tasks',                   // L - Source
      '',                        // M - Notes
      false,                     // N - Calendar Sync (default off)
      now,                       // O - Created Date
      '',                        // P - Completed Date
      now                        // Q - Last Modified
    ]);
    
    Logger.log('Task row appended successfully');
    
    // Get the actual last row with data
    const lastRow = tasksSheet.getLastRow();
    Logger.log(`Task added to row: ${lastRow}`);
    
    // Apply data validation to this specific row
    applyTaskRowValidation(tasksSheet, lastRow);
    
    // Format the new row
    if (dueDate) {
      tasksSheet.getRange(lastRow, 6).setNumberFormat('M/d/yyyy');
    }
    tasksSheet.getRange(lastRow, 15).setNumberFormat('M/d/yyyy h:mm');
    tasksSheet.getRange(lastRow, 17).setNumberFormat('M/d/yyyy h:mm');
    
    Logger.log('Task formatting complete');
    
    ui.alert('✅ Task Created',
      `Task added to Tasks sheet (Row ${lastRow}):\n\n` +
      `${taskDescription}\n\n` +
      `Project: ${projectName}\n` +
      `Type: ${taskType}\n` +
      `Priority: ${priority}\n` +
      `Due: ${dueDate ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'Not set'}\n\n` +
      `Task ID: ${taskID}\n\n` +
      `View Tasks sheet to see it, or enable Calendar Sync.`,
      ui.ButtonSet.OK);
      
    // Navigate to Tasks sheet to show the new task
    ss.setActiveSheet(tasksSheet);
    tasksSheet.setActiveRange(tasksSheet.getRange(lastRow, 1, 1, 17));
    
  } catch (error) {
    Logger.log(`ERROR creating task: ${error.message}`);
    Logger.log(`ERROR stack: ${error.stack}`);
    ui.alert('❌ Error Creating Task',
      `Could not create task.\n\n` +
      `Error: ${error.message}\n\n` +
      `Check View → Executions for details.`,
      ui.ButtonSet.OK);
  }
}

// ==========================================
// VIEW PROJECT TASKS
// ==========================================
function viewProjectTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  const ui = SpreadsheetApp.getUi();
  
  // Check if we're on a tracked sheet
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) and select a project row.');
    return;
  }
  
  // Get selected project
  const activeRow = activeSheet.getActiveCell().getRow();
  
  if (activeRow === 1) {
    ui.alert('Select a project row (not the header) to view its tasks.');
    return;
  }
  
  const projectUUID = activeSheet.getRange(activeRow, CONFIG.UUID_COLUMN).getValue();
  const projectName = activeSheet.getRange(activeRow, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  
  // Count tasks
  const data = tasksSheet.getDataRange().getValues();
  let tasks = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === projectUUID) { // Column B - Project UUID
      tasks.push({
        description: data[i][3],
        type: data[i][4],
        dueDate: data[i][5],
        status: data[i][8],
        priority: data[i][9]
      });
    }
  }
  
  if (tasks.length === 0) {
    ui.alert('No Tasks',
      `No tasks found for project: ${projectName}\n\nUse "Add Task for This Project" to create tasks.`,
      ui.ButtonSet.OK);
    return;
  }
  
  // Build task list
  let message = `Tasks for: ${projectName}\n\n`;
  message += `Total: ${tasks.length} task(s)\n\n`;
  
  const statusCounts = {};
  tasks.forEach(task => {
    const status = task.status || 'Not Started';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });
  
  message += 'Status Summary:\n';
  Object.keys(statusCounts).forEach(status => {
    message += `  ${status}: ${statusCounts[status]}\n`;
  });
  
  message += '\n---\n\n';
  
  tasks.slice(0, 10).forEach((task, index) => {
    message += `${index + 1}. [${task.priority}] ${task.description}\n`;
    message += `   Type: ${task.type} | Status: ${task.status}\n`;
    if (task.dueDate) {
      message += `   Due: ${Utilities.formatDate(new Date(task.dueDate), Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n`;
    }
    message += '\n';
  });
  
  if (tasks.length > 10) {
    message += `... and ${tasks.length - 10} more tasks\n`;
  }
  
  message += '\nOpen Tasks sheet to see full details.';
  
  ui.alert('Project Tasks', message, ui.ButtonSet.OK);
  
  // Navigate to Tasks sheet
  ss.setActiveSheet(tasksSheet);
}
function testCalendarAccess() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== Testing Calendar Access ===');
    
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log(`✓ Calendar access OK`);
    Logger.log(`Calendar name: ${calendar.getName()}`);
    Logger.log(`Calendar ID: ${calendar.getId()}`);
    
    // Try to create a test event
    const testDate = new Date();
    testDate.setDate(testDate.getDate() + 7); // 7 days from now
    
    Logger.log(`Creating test event for: ${testDate}`);
    const testEvent = calendar.createAllDayEvent('TEST EVENT - Project Tracker', testDate, {
      description: 'This is a test event from Project Tracker. You can delete this.'
    });
    
    Logger.log(`✓ Test event created successfully!`);
    Logger.log(`Event ID: ${testEvent.getId()}`);
    
    ui.alert('✅ Calendar Test Successful!',
      `Calendar access is working!\n\n` +
      `Calendar: ${calendar.getName()}\n` +
      `Test event created for: ${Utilities.formatDate(testDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n\n` +
      `Check your Google Calendar - you should see:\n` +
      `"TEST EVENT - Project Tracker"\n\n` +
      `You can delete this test event.`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`✗ ERROR: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    
    ui.alert('❌ Calendar Access Error',
      `Could not access Google Calendar.\n\n` +
      `Error: ${error.message}\n\n` +
      `This might mean:\n` +
      `1. Script needs calendar permissions (run Setup Tracker first)\n` +
      `2. Calendar API is not accessible\n` +
      `3. Account permissions issue\n\n` +
      `Check View → Executions for detailed logs.`,
      ui.ButtonSet.OK);
  }
}

// ==========================================
// DEBUG & TESTING FUNCTIONS
// ==========================================
function testCalendarAccess() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== Testing Calendar Access ===');
    
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log(`✓ Calendar access OK`);
    Logger.log(`Calendar name: ${calendar.getName()}`);
    Logger.log(`Calendar ID: ${calendar.getId()}`);
    
    // Try to create a test event
    const testDate = new Date();
    testDate.setDate(testDate.getDate() + 7); // 7 days from now
    
    Logger.log(`Creating test event for: ${testDate}`);
    const testEvent = calendar.createAllDayEvent('TEST EVENT - Project Tracker', testDate, {
      description: 'This is a test event from Project Tracker. You can delete this.'
    });
    
    Logger.log(`✓ Test event created successfully!`);
    Logger.log(`Event ID: ${testEvent.getId()}`);
    
    ui.alert('✅ Calendar Test Successful!',
      `Calendar access is working!\n\n` +
      `Calendar: ${calendar.getName()}\n` +
      `Test event created for: ${Utilities.formatDate(testDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n\n` +
      `Check your Google Calendar - you should see:\n` +
      `"TEST EVENT - Project Tracker"\n\n` +
      `You can delete this test event.`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`✗ ERROR: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    
    ui.alert('❌ Calendar Access Error',
      `Could not access Google Calendar.\n\n` +
      `Error: ${error.message}\n\n` +
      `This might mean:\n` +
      `1. Script needs calendar permissions (run Setup Tracker first)\n` +
      `2. Calendar API is not accessible\n` +
      `3. Account permissions issue\n\n` +
      `Check View → Executions for detailed logs.`,
      ui.ButtonSet.OK);
  }
}

function viewExecutionLogs() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('View Execution Logs',
    'To see detailed logs:\n\n' +
    '1. In the spreadsheet, click Extensions → Apps Script\n' +
    '2. In the Apps Script editor, click "Executions" (⏱️ icon on left)\n' +
    '3. Click on any recent execution to see logs\n\n' +
    'Or:\n' +
    '1. In Apps Script editor, click "View" → "Logs"\n' +
    '2. Run a function to see its output',
    ui.ButtonSet.OK);
}

function listMyCalendars() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const calendars = CalendarApp.getAllCalendars();
    
    let message = `📅 Your Calendars (${calendars.length}):\n\n`;
    
    calendars.forEach((cal, index) => {
      message += `${index + 1}. ${cal.getName()}\n`;
      message += `   ID: ${cal.getId()}\n`;
      message += `   Color: ${cal.getColor()}\n\n`;
      Logger.log(`Calendar ${index + 1}: ${cal.getName()} (${cal.getId()})`);
    });
    
    message += `\nCurrently using: ${CalendarApp.getDefaultCalendar().getName()}`;
    
    ui.alert('Your Google Calendars', message, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    ui.alert('❌ Error', `Could not list calendars: ${error.message}`, ui.ButtonSet.OK);
  }
}

function searchForProjectEvents() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const oneYearAgo = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
    const oneYearAhead = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));
    
    Logger.log('Searching for events from', oneYearAgo, 'to', oneYearAhead);
    
    const allEvents = calendar.getEvents(oneYearAgo, oneYearAhead);
    Logger.log(`Total events in range: ${allEvents.length}`);
    
    let projectEvents = [];
    allEvents.forEach(event => {
      const description = event.getDescription();
      if (description && description.includes('UUID: proj_')) {
        projectEvents.push({
          title: event.getTitle(),
          date: event.getAllDayStartDate() || event.getStartTime(),
          description: description
        });
      }
    });
    
    Logger.log(`Project events found: ${projectEvents.length}`);
    
    let message = `🔍 Search Results:\n\n`;
    message += `Total calendar events: ${allEvents.length}\n`;
    message += `Project tracker events: ${projectEvents.length}\n\n`;
    
    if (projectEvents.length > 0) {
      message += `Project Events:\n\n`;
      projectEvents.forEach((evt, index) => {
        message += `${index + 1}. ${evt.title}\n`;
        message += `   Date: ${Utilities.formatDate(evt.date, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n\n`;
      });
    } else {
      message += `No project tracker events found.\n\n`;
      message += `This means:\n`;
      message += `• No events created yet, OR\n`;
      message += `• Events don't have UUID tags in description`;
    }
    
    ui.alert('Calendar Event Search', message, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    ui.alert('❌ Error', `Could not search calendar: ${error.message}`, ui.ButtonSet.OK);
  }
}

function debugSelectedProject() {
  const activeSheet = getActiveProjectSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Select a project row in a tracked sheet (cobuild or enablement).');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    ui.alert('Select a project row (not the header) to debug.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  
  Logger.log('=== Debug Selected Project ===');
  Logger.log(`Sheet: ${sheetName}`);
  Logger.log(`Row: ${row}`);
  
  const uuid = activeSheet.getRange(row, CONFIG.UUID_COLUMN).getValue();
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const completionDate = activeSheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getValue();
  const syncEnabled = activeSheet.getRange(row, CONFIG.CALENDAR_SYNC_COLUMN).getValue();
  const status = activeSheet.getRange(row, CONFIG.STATUS_COLUMN).getValue();
  
  Logger.log(`UUID: ${uuid}`);
  Logger.log(`Project Title: ${projectTitle}`);
  Logger.log(`Completion Date: ${completionDate}`);
  Logger.log(`Calendar Sync: ${syncEnabled}`);
  Logger.log(`Status: ${status}`);
  
  let message = `🔍 Project Debug Info:\n\n`;
  message += `Sheet: ${sheetName}\n`;
  message += `Row: ${row}\n`;
  message += `UUID: ${uuid || '❌ MISSING'}\n`;
  message += `Project: ${projectTitle}\n`;
  message += `Completion Date: ${completionDate || '❌ NOT SET'}\n`;
  message += `Calendar Sync: ${syncEnabled ? '✅ Enabled' : '❌ Disabled'}\n`;
  message += `Status: ${status}\n\n`;
  
      // Check if calendar event exists
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = findCalendarEventByUUID(calendar, uuid);
    
    if (events.length > 0) {
      message += `📅 Calendar Events: ${events.length} FOUND\n`;
      events.forEach((event, index) => {
        const eventDate = event.getAllDayStartDate();
        const eventType = event.getDescription().includes('Type: check-in') ? 'Check-in (Blue)' : 'Deadline (Orange)';
        message += `   ${index + 1}. ${event.getTitle()} - ${eventType}\n`;
        message += `      Date: ${Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n`;
      });
      Logger.log(`Calendar events found: ${events.length}`);
    } else {
      message += `📅 Calendar Events: NOT FOUND\n`;
      Logger.log('No calendar events found for this project');
    }
  } catch (error) {
    message += `📅 Calendar Events: ERROR - ${error.message}\n`;
    Logger.log(`Error checking calendar: ${error.message}`);
  }
  
  ui.alert('Project Debug Info', message, ui.ButtonSet.OK);
}
function addNewProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = getActiveProjectSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (!activeSheet) {
    ui.alert('Please switch to a tracked project sheet (cobuild or enablement) first.');
    return;
  }
  
  const sheetName = activeSheet.getName();
  
  // Get project details via prompts
  const titleResponse = ui.prompt(
    'Add New Project - Step 1 of 4',
    'Enter Project Title:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (titleResponse.getSelectedButton() !== ui.Button.OK) return;
  const projectTitle = titleResponse.getResponseText().trim();
  if (!projectTitle) {
    ui.alert('Project title cannot be empty.');
    return;
  }
  
  const platformResponse = ui.prompt(
    'Add New Project - Step 2 of 4',
    'Enter Platform (e.g., AWS, Azure, GCP):',
    ui.ButtonSet.OK_CANCEL
  );
  if (platformResponse.getSelectedButton() !== ui.Button.OK) return;
  const platform = platformResponse.getResponseText().trim();
  
  const descriptionResponse = ui.prompt(
    'Add New Project - Step 3 of 4',
    'Enter Description:',
    ui.ButtonSet.OK_CANCEL
  );
  if (descriptionResponse.getSelectedButton() !== ui.Button.OK) return;
  const description = descriptionResponse.getResponseText().trim();
  
  const ownerResponse = ui.prompt(
    'Add New Project - Step 4 of 4',
    'Enter Tenable Owner:',
    ui.ButtonSet.OK_CANCEL
  );
  if (ownerResponse.getSelectedButton() !== ui.Button.OK) return;
  const owner = ownerResponse.getResponseText().trim();
  
  // Generate UUID with appropriate prefix
  const uuid = generateUUID(sheetName);
  const today = new Date();
  
  // Add new row (following your column structure)
  const newRow = [
    true,              // A - SPT
    platform,          // B - Platform
    projectTitle,      // C - Project Title
    '',                // D - Reason/Value to TENB
    description,       // E - Description
    '',                // F - Category
    'Not Started',     // G - Status
    today,             // H - Start Date
    '',                // I - Completion Date
    today,             // J - First Check In
    '',                // K - Potential Blockers
    today,             // L - Last Check In
    '',                // M - Next Check In
    '',                // N - Tenb Product(s)
    '',                // O - Next Steps
    owner,             // P - Tenable Owner
    '',                // Q - Team
    '',                // R - AWS Owner
    '',                // S - Notes
    uuid,              // T - UUID
    CONFIG.DEFAULT_SYNC_ENABLED // U - Calendar Sync
  ];
  
  activeSheet.appendRow(newRow);
  
  // Format the new row
  const lastRow = activeSheet.getLastRow();
  activeSheet.getRange(lastRow, 8).setNumberFormat('M/d/yyyy'); // Start Date
  activeSheet.getRange(lastRow, 10).setNumberFormat('M/d/yyyy'); // First Check In
  activeSheet.getRange(lastRow, 12).setNumberFormat('M/d/yyyy'); // Last Check In
  
  // Select the new row
  activeSheet.setActiveRange(activeSheet.getRange(lastRow, 1, 1, activeSheet.getLastColumn()));
  
  ui.alert('✅ Project Added Successfully!', 
    `Sheet: ${sheetName}\n` +
    `Project: ${projectTitle}\nUUID: ${uuid}\n\nRow ${lastRow} created.\nCalendar Sync: ${CONFIG.DEFAULT_SYNC_ENABLED ? 'Enabled' : 'Disabled'}\n\nYou can now edit additional fields directly in the sheet.`,
    ui.ButtonSet.OK);
}

// ==========================================
// MANUALLY GENERATE UUIDS FOR ROWS
// ==========================================
function generateMissingUUIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }
  
  let added = 0;
  
  for (let row = 2; row <= lastRow; row++) {
    const uuidCell = sheet.getRange(row, CONFIG.UUID_COLUMN);
    const currentUUID = uuidCell.getValue();
    
    if (!currentUUID) {
      uuidCell.setValue(generateUUID());
      added++;
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ UUIDs Generated', 
    `Added UUIDs to ${added} project(s) that were missing them.`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// QUICK UPDATE NEXT STEPS
// ==========================================
function quickUpdateNextSteps() {
  const activeSheet = getActiveProjectSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project row in a tracked sheet (cobuild or enablement).');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row to update Next Steps.');
    return;
  }
  
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const currentNextSteps = activeSheet.getRange(row, CONFIG.ACTIVITY_COLUMN).getValue();
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    `Update Next Steps: ${projectTitle}`,
    `Current: ${currentNextSteps}\n\nEnter new Next Steps:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newValue = response.getResponseText();
    if (newValue && newValue.trim() !== '') {
      activeSheet.getRange(row, CONFIG.ACTIVITY_COLUMN).setValue(newValue);
      ui.alert('✅ Next Steps Updated', 
        `Updated for: ${projectTitle}\n\nOld value has been saved to Audit Log.\nLast Check In and Next Check In have been updated.`,
        ui.ButtonSet.OK);
    }
  }
}

// ==========================================
// VIEW ALL RECENT ACTIVITY
// ==========================================
function viewRecentActivity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  const activeSheet = getActiveProjectSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (!auditSheet) {
    ui.alert('No activity logged yet. Make some edits first!');
    return;
  }
  
  const sheetName = activeSheet ? activeSheet.getName() : 'all';
  
  const data = auditSheet.getDataRange().getValues();
  const recentActivities = [];
  
  // Get last 20 changes (optionally filter by current sheet)
  for (let i = Math.max(1, data.length - 20); i < data.length; i++) {
    const [timestamp, uuid, projectTitle, activitySheetName, row, col, field, oldVal, newVal, user] = data[i];
    
    // If on a specific sheet, only show activities from that sheet
    if (activeSheet && activitySheetName !== sheetName) continue;
    recentActivities.push({
      timestamp: new Date(timestamp),
      projectTitle: projectTitle,
      field: field,
      oldValue: oldVal,
      newValue: newVal,
      user: user
    });
  }
  
  recentActivities.reverse(); // Show newest first
  
  const title = activeSheet ? `Recent Activity - ${sheetName} sheet` : 'Recent Activity - All Sheets';
  const htmlOutput = createActivitySummary(recentActivities);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}

function createActivitySummary(activities) {
  let html = `
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; margin: 0; background: #f8f9fa; }
      .activity-item { 
        padding: 12px; 
        margin-bottom: 10px; 
        background: white;
        border-radius: 6px;
        border-left: 3px solid #34a853;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      }
      .activity-header { 
        font-weight: bold; 
        color: #1a73e8; 
        margin-bottom: 5px;
      }
      .activity-field { 
        font-size: 12px; 
        color: #5f6368;
        margin-bottom: 3px;
      }
      .activity-time { 
        font-size: 11px; 
        color: #80868b;
      }
      .next-steps-badge {
        background: #1a73e8;
        color: white;
        padding: 2px 6px;
        border-radius: 10px;
        font-size: 10px;
        margin-left: 5px;
      }
    </style>
  `;
  
  activities.forEach(activity => {
    const date = Utilities.formatDate(activity.timestamp, Session.getScriptTimeZone(), 'MMM dd, h:mm a');
    const isNextSteps = activity.field === 'Next Steps';
    
    html += `
      <div class="activity-item">
        <div class="activity-header">
          ${activity.projectTitle}
          ${isNextSteps ? '<span class="next-steps-badge">ACTIVITY</span>' : ''}
        </div>
        <div class="activity-field">${activity.field}: ${activity.newValue}</div>
        <div class="activity-time">📅 ${date} • ${activity.user}</div>
      </div>
    `;
  });
  
  return HtmlService.createHtmlOutput(html).setWidth(500).setHeight(500);
}

// ==========================================
// GOOGLE TASKS INTEGRATION
// ==========================================
function createTaskFromProject() {
  const activeSheet = getActiveProjectSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Select a project row in a tracked sheet (cobuild or enablement).');
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a project row to create a task.');
    return;
  }
  
  const projectTitle = activeSheet.getRange(row, CONFIG.PROJECT_TITLE_COLUMN).getValue();
  const description = activeSheet.getRange(row, 5).getValue(); // Column E
  const completionDate = activeSheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getValue();
  const nextSteps = activeSheet.getRange(row, CONFIG.ACTIVITY_COLUMN).getValue();
  
  try {
    const task = {
      title: projectTitle,
      notes: `Description: ${description}\n\nNext Steps: ${nextSteps}`,
    };
    
    if (completionDate) {
      task.due = new Date(completionDate).toISOString();
    }
    
    Tasks.Tasks.insert(task, '@default');
    
    SpreadsheetApp.getUi().alert('✅ Task Created', 
      `Task "${projectTitle}" added to Google Tasks\n\nNext Steps included in task notes.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 
      'Could not create task. Make sure Google Tasks API is enabled in Extensions > Apps Script > Services', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// EMAIL REMINDERS
// ==========================================
function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let allActiveProjects = [];
  let allOverdueProjects = [];
  let allRecentActivity = [];
  
  // Loop through all tracked sheets
  CONFIG.SHEETS.TRACKED_PROJECTS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][CONFIG.STATUS_COLUMN - 1];
    const completionDate = data[i][CONFIG.COMPLETION_DATE_COLUMN - 1];
    const projectTitle = data[i][CONFIG.PROJECT_TITLE_COLUMN - 1];
    const nextSteps = data[i][CONFIG.ACTIVITY_COLUMN - 1];
    const lastCheckIn = data[i][CONFIG.LAST_CHECKIN_COLUMN - 1];
    
    if (status !== 'Complete' && status !== 'Cancelled') {
      allActiveProjects.push({title: `[${sheetName}] ${projectTitle}`, nextSteps: nextSteps});
      
      if (completionDate && new Date(completionDate) < today) {
        allOverdueProjects.push(`[${sheetName}] ${projectTitle} (Due: ${Utilities.formatDate(new Date(completionDate), Session.getScriptTimeZone(), 'MMM dd, yyyy')})`);
      }
      
      // Check if updated in last 24 hours
      if (lastCheckIn && (today - new Date(lastCheckIn)) < 86400000) {
        allRecentActivity.push({title: `[${sheetName}] ${projectTitle}`, nextSteps: nextSteps});
      }
    }
  }
});
  
  const email = Session.getActiveUser().getEmail();
  const subject = `📊 Daily Project Summary - ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMM dd, yyyy')}`;
  
  let body = `<h2>Daily Project Summary</h2>`;
  body += `<p><strong>Active Projects:</strong> ${allActiveProjects.length}</p>`;
  
  if (allRecentActivity.length > 0) {
    body += `<h3>🔥 Recent Activity (Last 24h)</h3><ul>`;
    allRecentActivity.forEach(p => body += `<li><strong>${p.title}</strong><br/>Next Steps: ${p.nextSteps}</li>`);
    body += `</ul>`;
  }
  
  if (allOverdueProjects.length > 0) {
    body += `<h3>⚠️ Overdue Projects (${allOverdueProjects.length})</h3><ul>`;
    allOverdueProjects.forEach(p => body += `<li>${p}</li>`);
    body += `</ul>`;
  }
  
  body += `<h3>All Active Projects</h3><ul>`;
  allActiveProjects.forEach(p => body += `<li><strong>${p.title}</strong><br/>Next Steps: ${p.nextSteps || '(none)'}</li>`);
  body += `</ul>`;
  
  body += `<p><a href="${ss.getUrl()}">Open Tracker</a></p>`;
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body
  });
}

// ==========================================
// CUSTOM MENU
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Main Tracker Menu
  ui.createMenu('📊 Project Tracker')
    .addItem('🔧 Setup Tracker', 'setupTracker')
    .addSeparator()
    .addItem('➕ Add New Project', 'addNewProject')
    .addItem('🆔 Generate Missing UUIDs', 'generateMissingUUIDs')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Tasks')
      .addItem('➕ Add Task for This Project', 'addTaskForProjectV2')
      .addItem('➕ Add Subtask to This Task', 'addSubtaskToTask')
      .addItem('📊 View Project Tasks', 'viewProjectTasks')
      .addSeparator()
      .addItem('🔄 Sync This Project to Google Tasks', 'syncSelectedProjectToGoogleTasks')
      .addItem('🔄 Sync All Tasks to Google Tasks', 'syncAllTasksToGoogleTasks')
      .addItem('📊 View Google Tasks Sync Status', 'viewGoogleTasksSyncStatus')
      .addSeparator()
      .addItem('⚙️ Toggle Auto-Sync', 'toggleGoogleTasksAutoSync')
      .addSeparator()
      .addItem('🧹 Clean Up Tasks Sheet', 'cleanupTasksSheet'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Calendar Sync')
      .addItem('🔄 Sync with Calendar', 'syncWithCalendar')
      .addItem('📊 View Sync Status', 'viewCalendarSyncStatus')
      .addSeparator()
      .addItem('✅ Enable Sync (This Project)', 'enableCalendarSyncForProject')
      .addItem('⏸️ Disable Sync (This Project)', 'disableCalendarSyncForProject'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 Next Steps (Activity)')
      .addItem('Quick Update Next Steps', 'quickUpdateNextSteps')
      .addItem('View All Recent Activity', 'viewRecentActivity'))
    .addSeparator()
    .addItem('📜 View Change History (Selected Cell)', 'showChangeHistory')
    .addItem('✅ Create Google Task', 'createTaskFromProject')
    .addSeparator()
    .addItem('📧 Send Daily Summary Now', 'sendDailySummary')
    .addItem('⚙️ Setup Daily Email Trigger', 'setupDailyTrigger')
    .addSeparator()
    .addSubMenu(ui.createMenu('🐛 Debug & Testing')
      .addItem('🔍 Test Calendar Access', 'testCalendarAccess')
      .addItem('🔍 Debug Selected Project', 'debugSelectedProject')
      .addItem('🔍 Search for Project Events', 'searchForProjectEvents')
      .addItem('📅 List My Calendars', 'listMyCalendars')
      .addItem('📋 View Execution Logs', 'viewExecutionLogs'))
    .addToUi();
}

// ==========================================
// TRIGGER SETUP
// ==========================================
function setupDailyTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailySummary') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(CONFIG.EMAIL_REMINDER_HOUR)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Daily Email Reminder Set', 
    `You will receive a daily summary email at ${CONFIG.EMAIL_REMINDER_HOUR}:00 AM\n\n` +
    `The email will include:\n` +
    `• Active project count from all tracked sheets\n` +
    `• Recent activity (last 24h)\n` +
    `• Overdue projects\n` +
    `• All active projects with Next Steps`, 
    SpreadsheetApp.getUi().ButtonSet.OK);
}
