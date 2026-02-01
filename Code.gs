/**
 * Gmail Cleaner
 * Helps users analyse and clean up their Gmail inbox using Google Sheets
 */

const DEFAULT_SCAN_LIMIT = 1000;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail Cleaner')
    .addItem('Launch Sidebar', 'showSidebar')
    .addItem('Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('Analyse Inbox Now', 'scanInbox')
    .addItem('Run Cleanup Now', 'runCleanup')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gmail Cleaner')
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Sets up the required sheets: Analysis, Rules, and Logs
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let analysisSheet = ss.getSheetByName('Analysis');
  if (!analysisSheet) {
    analysisSheet = ss.insertSheet('Analysis');
    analysisSheet.appendRow(['Sender', 'Email Count', 'Last Received', 'Sample Subject']);
    analysisSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    analysisSheet.setFrozenRows(1);
    analysisSheet.setColumnWidths(1, 4, 200);
  }
  
  let rulesSheet = ss.getSheetByName('Rules');
  if (!rulesSheet) {
    rulesSheet = ss.insertSheet('Rules');
    rulesSheet.appendRow(['Rule Type', 'Value', 'Action', 'Status']);
    rulesSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    rulesSheet.setFrozenRows(1);
    rulesSheet.setColumnWidths(1, 4, 200);
    
    const ruleTypeRange = rulesSheet.getRange('A2:A1000');
    const ruleTypeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sender', 'Subject', 'Content'], true)
      .setAllowInvalid(false)
      .build();
    ruleTypeRange.setDataValidation(ruleTypeValidation);
    
    const actionRange = rulesSheet.getRange('C2:C1000');
    const actionValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Trash', 'Archive', 'Mark Read'], true)
      .setAllowInvalid(false)
      .build();
    actionRange.setDataValidation(actionValidation);
    
    const statusRange = rulesSheet.getRange('D2:D1000');
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Paused'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusValidation);
    
    rulesSheet.appendRow(['Sender', 'example@newsletter.com', 'Trash', 'Paused']);
  }
  
  let logsSheet = ss.getSheetByName('Logs');
  if (!logsSheet) {
    logsSheet = ss.insertSheet('Logs');
    logsSheet.appendRow(['Timestamp', 'Action', 'Details', 'Count']);
    logsSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#fbbc04').setFontColor('#000000');
    logsSheet.setFrozenRows(1);
    logsSheet.setColumnWidths(1, 4, 200);
  }
  
  return {
    success: true,
    message: 'Sheets setup complete! Analysis, Rules, and Logs sheets are ready.'
  };
}

/**
 * Scans the inbox in small batches for progress updates
 * @param {number} startIndex - Where to start (0, 50, 100, etc.)
 * @param {boolean} clearSheet - Whether to clear existing data first
 * @param {number} maxEmails - Maximum emails to scan (default: 1000)
 */
function scanInbox(startIndex = 0, clearSheet = true, maxEmails = DEFAULT_SCAN_LIMIT) {
  // Ensure maxEmails is a valid number
  maxEmails = Number(maxEmails) || DEFAULT_SCAN_LIMIT;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analysisSheet = ss.getSheetByName('Analysis');
    
    if (!analysisSheet) {
      return { success: false, message: 'Analysis sheet not found. Please run Setup Sheets first.' };
    }
    
    if (startIndex >= maxEmails) {
      return {
        success: true,
        message: `Scan limit reached (${maxEmails.toLocaleString()} emails)`,
        emailsProcessed: 0,
        hasMore: false,
        nextIndex: startIndex,
        limitReached: true
      };
    }
    
    if (clearSheet && startIndex === 0) {
      const lastRow = analysisSheet.getLastRow();
      if (lastRow > 1) {
        analysisSheet.getRange(2, 1, lastRow - 1, 4).clear();
      }
    }
    
    const chunkSize = 50;
    const threads = GmailApp.getInboxThreads(startIndex, chunkSize);
    
    // Check if we found any threads
    if (!threads || threads.length === 0) {
      if (startIndex === 0) {
        return {
          success: true,
          message: 'Inbox is empty! No emails found to analyse.',
          emailsProcessed: 0,
          hasMore: false,
          nextIndex: startIndex
        };
      }
      
      return {
        success: true,
        message: 'All emails analysed!',
        emailsProcessed: 0,
        hasMore: false,
        nextIndex: startIndex
      };
    }
    
    const senderMap = {};
    
    if (!clearSheet || startIndex > 0) {
      const lastRow = analysisSheet.getLastRow();
      if (lastRow > 1) {
        const existingData = analysisSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        existingData.forEach(row => {
          if (row[0]) {
            senderMap[row[0]] = {
              count: row[1] || 0,
              lastReceived: new Date(row[2]) || new Date(),
              sampleSubject: row[3] || ''
            };
          }
        });
      }
    }
    
    let emailsInThisChunk = 0;
    threads.forEach(thread => {
      const messages = thread.getMessages();
      emailsInThisChunk += messages.length;
      
      messages.forEach(message => {
        const sender = message.getFrom();
        const subject = message.getSubject();
        const date = message.getDate();
        
        const emailMatch = sender.match(/<(.+?)>/) || [null, sender];
        const email = emailMatch[1] || sender;
        
        if (!senderMap[email]) {
          senderMap[email] = { count: 0, lastReceived: date, sampleSubject: subject };
        }
        
        senderMap[email].count++;
        if (date > senderMap[email].lastReceived) {
          senderMap[email].lastReceived = date;
          senderMap[email].sampleSubject = subject;
        }
      });
    });
    
    const senderArray = Object.keys(senderMap).map(email => ({
      email: email,
      count: senderMap[email].count,
      lastReceived: senderMap[email].lastReceived,
      sampleSubject: senderMap[email].sampleSubject
    }));
    
    senderArray.sort((a, b) => b.count - a.count);
    
    const lastRow = analysisSheet.getLastRow();
    if (lastRow > 1) {
      analysisSheet.getRange(2, 1, lastRow - 1, 4).clear();
    }
    
    const dataToWrite = senderArray.map(item => [
      item.email,
      item.count,
      Utilities.formatDate(item.lastReceived, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      item.sampleSubject
    ]);
    
    if (dataToWrite.length > 0) {
      analysisSheet.getRange(2, 1, dataToWrite.length, 4).setValues(dataToWrite);
      analysisSheet.getRange(2, 2, dataToWrite.length, 1).setNumberFormat('#,##0');
    }
    
    const nextIndex = startIndex + chunkSize;
    const hasMore = threads.length === chunkSize && nextIndex < maxEmails;
    
    return {
      success: true,
      message: nextIndex >= maxEmails ? `Scan limit reached (${maxEmails.toLocaleString()} emails)` : `Processing...`,
      emailsProcessed: emailsInThisChunk,
      totalEmailsSoFar: startIndex + emailsInThisChunk,
      senderCount: senderArray.length,
      hasMore: hasMore,
      nextIndex: nextIndex,
      limitReached: nextIndex >= maxEmails
    };
    
  } catch (error) {
    return {
      success: false,
      message: `Error at position ${startIndex}: ${error.message}`,
      canResume: true,
      resumeIndex: startIndex
    };
  }
}

function getAnalysisProgress() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastIndex = scriptProperties.getProperty('lastAnalysisIndex');
    return {
      hasIncomplete: lastIndex !== null,
      lastIndex: parseInt(lastIndex) || 0
    };
  } catch (error) {
    return { hasIncomplete: false, lastIndex: 0 };
  }
}

function saveAnalysisProgress(index) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('lastAnalysisIndex', index.toString());
  } catch (error) {
    console.error('Failed to save progress:', error);
  }
}

function clearAnalysisProgress() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('lastAnalysisIndex');
  } catch (error) {
    console.error('Failed to clear progress:', error);
  }
}

/**
 * Clears all analysis data and progress for a fresh start
 */
function clearAnalysisData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analysisSheet = ss.getSheetByName('Analysis');
    
    if (!analysisSheet) {
      return { 
        success: false, 
        message: 'Analysis sheet not found. Please run Setup Sheets first.' 
      };
    }
    
    const lastRow = analysisSheet.getLastRow();
    if (lastRow > 1) {
      analysisSheet.getRange(2, 1, lastRow - 1, 4).clear();
    }
    
    clearAnalysisProgress();
    
    logAction('Clear Analysis', 'Analysis data and progress cleared', 0);
    
    return {
      success: true,
      message: 'Analysis data cleared. Ready for a fresh scan!'
    };
    
  } catch (error) {
    return {
      success: false,
      message: `Error clearing analysis data: ${error.message}`
    };
  }
}

/**
 * Runs cleanup in batches to handle large numbers of emails
 * @param {number} ruleIndex - Which rule to process (0-based)
 * @param {number} batchStart - Start index for this batch
 */
function runCleanup(ruleIndex = 0, batchStart = 0) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName('Rules');
    
    if (!rulesSheet) {
      return { success: false, message: 'Rules sheet not found. Please run Setup Sheets first.' };
    }
    
    const lastRow = rulesSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: 'No rules found. Please add rules to the Rules sheet.' };
    }
    
    const allRules = rulesSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const activeRules = allRules
      .map((rule, idx) => ({ rule, originalIndex: idx }))
      .filter(({ rule }) => rule[3] === 'Active' && rule[1] && rule[2]);
    
    if (activeRules.length === 0) {
      return { success: false, message: 'No active rules found. Please add active rules to the Rules sheet.' };
    }
    
    if (ruleIndex >= activeRules.length) {
      return {
        success: true,
        message: 'All rules processed!',
        isComplete: true,
        ruleIndex: ruleIndex,
        batchStart: 0
      };
    }
    
    const { rule, originalIndex } = activeRules[ruleIndex];
    const [ruleType, value, action, status] = rule;
    
    try {
      let searchQuery = '';
      switch (ruleType) {
        case 'Sender':
          searchQuery = `from:${value}`;
          break;
        case 'Subject':
          searchQuery = `subject:${value}`;
          break;
        case 'Content':
          searchQuery = value;
          break;
        default:
          return runCleanup(ruleIndex + 1, 0);
      }
      
      const batchSize = 50;
      const threads = GmailApp.search(searchQuery, batchStart, batchSize);
      
      if (threads.length === 0) {
        if (batchStart === 0) {
          logAction(`Cleanup Rule ${originalIndex + 1}`, `${action} - ${ruleType}: ${value} - No emails found`, 0);
        }
        
        return {
          success: true,
          message: `Rule ${ruleIndex + 1}/${activeRules.length} complete`,
          ruleIndex: ruleIndex + 1,
          batchStart: 0,
          emailsProcessed: 0,
          hasMoreInRule: false,
          hasMoreRules: (ruleIndex + 1) < activeRules.length,
          currentRuleName: `${ruleType}: ${value}`,
          totalRules: activeRules.length
        };
      }
      
      switch (action) {
        case 'Trash':
          threads.forEach(thread => thread.moveToTrash());
          break;
        case 'Archive':
          threads.forEach(thread => thread.moveToArchive());
          break;
        case 'Mark Read':
          threads.forEach(thread => thread.markRead());
          break;
      }
      
      if (batchStart === 0) {
        logAction(`Cleanup Rule ${originalIndex + 1}`, `${action} - ${ruleType}: ${value} - Started`, threads.length);
      }
      
      const hasMoreInRule = threads.length === batchSize;
      
      return {
        success: true,
        message: `Processing rule ${ruleIndex + 1}/${activeRules.length}...`,
        ruleIndex: ruleIndex,
        batchStart: hasMoreInRule ? batchStart + batchSize : 0,
        emailsProcessed: threads.length,
        hasMoreInRule: hasMoreInRule,
        hasMoreRules: !hasMoreInRule && ((ruleIndex + 1) < activeRules.length),
        currentRuleName: `${ruleType}: ${value}`,
        currentAction: action,
        totalRules: activeRules.length,
        nextRuleIndex: hasMoreInRule ? ruleIndex : ruleIndex + 1
      };
      
    } catch (error) {
      logAction(`Cleanup Rule ${originalIndex + 1}`, `Error: ${error.message}`, 0);
      
      return {
        success: true,
        message: `Rule ${ruleIndex + 1} failed, continuing...`,
        ruleIndex: ruleIndex + 1,
        batchStart: 0,
        emailsProcessed: 0,
        hasMoreInRule: false,
        hasMoreRules: (ruleIndex + 1) < activeRules.length,
        error: error.message,
        totalRules: activeRules.length
      };
    }
    
  } catch (error) {
    return {
      success: false,
      message: `Error during cleanup: ${error.message}`,
      canResume: true,
      ruleIndex: ruleIndex,
      batchStart: batchStart
    };
  }
}

function logAction(action, details, count) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName('Logs');
    
    if (!logsSheet) return;
    
    const timestamp = new Date();
    logsSheet.appendRow([
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      action,
      details,
      count || 0
    ]);
    
  } catch (error) {
    console.error('Error logging action:', error);
  }
}

/**
 * Sets up scheduled triggers for automation
 * @param {string} frequency - 'daily', 'weekly', or 'none'
 */
function setupTriggers(frequency) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runScheduledCleanup') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    if (frequency === 'none') {
      return { success: true, message: 'Automated cleanup disabled.' };
    }
    
    const trigger = ScriptApp.newTrigger('runScheduledCleanup');
    
    if (frequency === 'daily') {
      trigger.timeBased().everyDays(1).atHour(2).create();
    } else if (frequency === 'weekly') {
      trigger.timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).create();
    } else {
      return { success: false, message: 'Invalid frequency. Use "daily", "weekly", or "none".' };
    }
    
    return {
      success: true,
      message: `Automated cleanup scheduled to run ${frequency} at 2:00 AM.`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `Error setting up triggers: ${error.message}`
    };
  }
}

function runScheduledCleanup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName('Rules');
    
    if (!rulesSheet) {
      logAction('Scheduled Cleanup', 'Rules sheet not found', 0);
      return;
    }
    
    const lastRow = rulesSheet.getLastRow();
    if (lastRow <= 1) {
      logAction('Scheduled Cleanup', 'No rules found', 0);
      return;
    }
    
    const rules = rulesSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let totalProcessed = 0;
    
    rules.forEach((rule, index) => {
      const [ruleType, value, action, status] = rule;
      
      if (status !== 'Active' || !value || !action) {
        return;
      }
      
      try {
        let searchQuery = '';
        switch (ruleType) {
          case 'Sender':
            searchQuery = `from:${value}`;
            break;
          case 'Subject':
            searchQuery = `subject:${value}`;
            break;
          case 'Content':
            searchQuery = value;
            break;
          default:
            return;
        }
        
        const threads = GmailApp.search(searchQuery, 0, 100);
        
        if (threads.length === 0) return;
        
        switch (action) {
          case 'Trash':
            threads.forEach(thread => thread.moveToTrash());
            break;
          case 'Archive':
            threads.forEach(thread => thread.moveToArchive());
            break;
          case 'Mark Read':
            threads.forEach(thread => thread.markRead());
            break;
        }
        
        totalProcessed += threads.length;
        logAction(`Cleanup Rule ${index + 1}`, `${action} - ${ruleType}: ${value}`, threads.length);
        
      } catch (error) {
        logAction(`Cleanup Rule ${index + 1}`, `Error: ${error.message}`, 0);
      }
    });
    
    logAction('Scheduled Cleanup Complete', `Processed ${totalProcessed} emails`, totalProcessed);
    
  } catch (error) {
    logAction('Scheduled Cleanup Error', error.message, 0);
  }
}

/**
 * Gets recent log entries for UI display
 * @param {number} limit - Number of entries to return
 */
function getRecentLogs(limit = 10) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName('Logs');
    
    if (!logsSheet) {
      return { success: false, logs: [] };
    }
    
    const lastRow = logsSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, logs: [] };
    }
    
    const numRows = Math.min(limit, lastRow - 1);
    const logs = logsSheet.getRange(lastRow - numRows + 1, 1, numRows, 4).getValues();
    
    logs.reverse();
    
    return {
      success: true,
      logs: logs.map(log => ({
        timestamp: log[0],
        action: log[1],
        details: log[2],
        count: log[3]
      }))
    };
    
  } catch (error) {
    return {
      success: false,
      logs: [],
      message: error.message
    };
  }
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const cleanupTrigger = triggers.find(t => t.getHandlerFunction() === 'runScheduledCleanup');
  
  if (!cleanupTrigger) {
    return { enabled: false, frequency: 'none' };
  }
  
  const eventType = cleanupTrigger.getEventType();
  if (eventType === ScriptApp.EventType.CLOCK) {
    return { enabled: true, frequency: 'daily' };
  }
  
  return { enabled: false, frequency: 'none' };
}
