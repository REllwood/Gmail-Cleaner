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
    .addItem('Analyse Inbox Now', 'analyseFromMenu')
    .addItem('Run Cleanup Now', 'cleanupFromMenu')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gmail Cleaner')
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Analysis and cleanup run in batches driven by the sidebar, so the menu
// items open the sidebar and leave it a note of which job to start
const PENDING_ACTION_KEY = 'pendingSidebarAction';

function analyseFromMenu() {
  openSidebarWithAction('analyse');
}

function cleanupFromMenu() {
  openSidebarWithAction('cleanup');
}

function openSidebarWithAction(action) {
  CacheService.getUserCache().put(PENDING_ACTION_KEY, action, 120);
  showSidebar();
}

/**
 * Returns the job a menu item asked the sidebar to start, and clears it
 * @return {string|null} 'analyse', 'cleanup', or null
 */
function takePendingSidebarAction() {
  const cache = CacheService.getUserCache();
  const action = cache.get(PENDING_ACTION_KEY);
  if (action) {
    cache.remove(PENDING_ACTION_KEY);
  }
  return action;
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
 * @param {number} startIndex - Inbox thread to start from (0, 50, 100, etc.)
 * @param {boolean} clearSheet - Whether to clear existing data first
 * @param {number} maxEmails - Maximum emails to scan (default: 1000)
 * @param {number} emailsSoFar - Emails already scanned in earlier batches
 */
function scanInbox(startIndex = 0, clearSheet = true, maxEmails = DEFAULT_SCAN_LIMIT, emailsSoFar = 0) {
  // Ensure the counts are valid numbers
  maxEmails = Number(maxEmails) || DEFAULT_SCAN_LIMIT;
  emailsSoFar = Number(emailsSoFar) || 0;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analysisSheet = ss.getSheetByName('Analysis');
    
    if (!analysisSheet) {
      return { success: false, message: 'Analysis sheet not found. Please run Setup Sheets first.' };
    }
    
    if (emailsSoFar >= maxEmails) {
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
    
    // The limit counts emails, so stop part-way through a batch (or a thread)
    // once it's reached
    let emailsInThisChunk = 0;
    let threadsScanned = 0;
    const limitHit = () => emailsSoFar + emailsInThisChunk >= maxEmails;
    
    for (const messages of GmailApp.getMessagesForThreads(threads)) {
      if (limitHit()) break;
      threadsScanned++;
      
      for (const message of messages) {
        if (limitHit()) break;
        emailsInThisChunk++;
        
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
      }
    }
    
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
    
    const nextIndex = startIndex + threadsScanned;
    const totalEmailsSoFar = emailsSoFar + emailsInThisChunk;
    const limitReached = totalEmailsSoFar >= maxEmails;
    const hasMore = !limitReached && threads.length === chunkSize;
    
    return {
      success: true,
      message: limitReached ? `Scan limit reached (${maxEmails.toLocaleString()} emails)` : `Processing...`,
      emailsProcessed: emailsInThisChunk,
      totalEmailsSoFar: totalEmailsSoFar,
      senderCount: senderArray.length,
      hasMore: hasMore,
      nextIndex: nextIndex,
      limitReached: limitReached
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

// Leaves out mail the action has already been applied to, so each batch drops
// out of the results and the next search can start from the top again
const ACTION_FILTERS = {
  'Trash': '-in:trash',
  'Archive': 'in:inbox',
  'Mark Read': 'is:unread'
};

/**
 * Builds the Gmail search query for a rule
 * @param {string} ruleType - 'Sender', 'Subject', or 'Content'
 * @param {string} value - The rule's value from the Rules sheet
 * @param {string} action - 'Trash', 'Archive', or 'Mark Read'
 * @return {string|null} The query, or null for an unknown rule type or action, or a blank value
 */
function buildSearchQuery(ruleType, value, action) {
  const text = String(value).trim();
  const actionFilter = ACTION_FILTERS[action];
  if (!text || !actionFilter) return null;
  
  let query;
  switch (ruleType) {
    case 'Sender':
      query = `from:${quoteSearchTerm(text)}`;
      break;
    case 'Subject':
      query = `subject:${quoteSearchTerm(text)}`;
      break;
    case 'Content':
      query = `(${text})`;
      break;
    default:
      return null;
  }
  
  return `${query} ${actionFilter}`;
}

/**
 * Quotes a value containing spaces or brackets so Gmail matches it as one
 * phrase in the field, rather than matching the extra words anywhere
 */
function quoteSearchTerm(text) {
  const clean = text.replace(/"/g, '');
  return /[\s(){}]/.test(clean) ? `"${clean}"` : clean;
}

const CLEANUP_BATCH_SIZE = 50;

// Scheduled runs stop starting new batches after this long, to stay inside
// Apps Script's 6-minute execution limit
const SCHEDULED_TIME_BUDGET_MS = 5 * 60 * 1000;

/**
 * Reads the active rules from the Rules sheet
 * @return {{rules: Array<Object>}|{error: string}} The rules, or why there are none
 */
function getActiveRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName('Rules');
  
  if (!rulesSheet) {
    return { error: 'Rules sheet not found. Please run Setup Sheets first.' };
  }
  
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) {
    return { error: 'No rules found. Please add rules to the Rules sheet.' };
  }
  
  const rules = rulesSheet.getRange(2, 1, lastRow - 1, 4).getValues()
    .map(([ruleType, value, action, status], idx) => ({ ruleType, value, action, status, ruleNumber: idx + 1 }))
    .filter(rule => rule.status === 'Active' && rule.value && rule.action);
  
  if (rules.length === 0) {
    return { error: 'No active rules found. Please add active rules to the Rules sheet.' };
  }
  
  return { rules: rules };
}

/**
 * Applies a rule's action to the next batch of matching threads
 * @param {Object} rule - A rule from getActiveRules()
 * @return {number|null} Threads processed, or null if the rule can't be run
 */
function cleanupRuleBatch(rule) {
  const searchQuery = buildSearchQuery(rule.ruleType, rule.value, rule.action);
  if (!searchQuery) return null;
  
  const threads = GmailApp.search(searchQuery, 0, CLEANUP_BATCH_SIZE);
  if (threads.length === 0) return 0;
  
  // One call per batch instead of one per thread
  switch (rule.action) {
    case 'Trash':
      GmailApp.moveThreadsToTrash(threads);
      break;
    case 'Archive':
      GmailApp.moveThreadsToArchive(threads);
      break;
    case 'Mark Read':
      GmailApp.markThreadsRead(threads);
      break;
  }
  
  return threads.length;
}

/**
 * Runs cleanup in batches to handle large numbers of emails. Each batch
 * searches from the top, because the previous batch has already dropped
 * out of the results.
 * @param {number} ruleIndex - Which rule to process (0-based)
 * @param {number} ruleTotal - Emails this rule has already processed in this run
 */
function runCleanup(ruleIndex = 0, ruleTotal = 0) {
  try {
    const { rules: activeRules, error } = getActiveRules();
    if (error) {
      return { success: false, message: error };
    }
    
    if (ruleIndex >= activeRules.length) {
      return {
        success: true,
        message: 'All rules processed!',
        isComplete: true,
        ruleIndex: ruleIndex,
        ruleTotal: 0
      };
    }
    
    const rule = activeRules[ruleIndex];
    const { ruleType, value, action, ruleNumber } = rule;
    const logRuleTotal = total => logAction(
      `Cleanup Rule ${ruleNumber}`,
      `${action} - ${ruleType}: ${value}${total === 0 ? ' - No emails found' : ''}`,
      total
    );
    
    try {
      const processed = cleanupRuleBatch(rule);
      if (processed === null) {
        return runCleanup(ruleIndex + 1, 0);
      }
      
      if (processed === 0) {
        logRuleTotal(ruleTotal);
        
        return {
          success: true,
          message: `Rule ${ruleIndex + 1}/${activeRules.length} complete`,
          ruleIndex: ruleIndex + 1,
          ruleTotal: 0,
          emailsProcessed: 0,
          hasMoreInRule: false,
          hasMoreRules: (ruleIndex + 1) < activeRules.length,
          currentRuleName: `${ruleType}: ${value}`,
          totalRules: activeRules.length,
          nextRuleIndex: ruleIndex + 1
        };
      }
      
      const hasMoreInRule = processed === CLEANUP_BATCH_SIZE;
      if (!hasMoreInRule) {
        logRuleTotal(ruleTotal + processed);
      }
      
      return {
        success: true,
        message: `Processing rule ${ruleIndex + 1}/${activeRules.length}...`,
        ruleIndex: ruleIndex,
        ruleTotal: hasMoreInRule ? ruleTotal + processed : 0,
        emailsProcessed: processed,
        hasMoreInRule: hasMoreInRule,
        hasMoreRules: !hasMoreInRule && ((ruleIndex + 1) < activeRules.length),
        currentRuleName: `${ruleType}: ${value}`,
        currentAction: action,
        totalRules: activeRules.length,
        nextRuleIndex: hasMoreInRule ? ruleIndex : ruleIndex + 1
      };
      
    } catch (error) {
      logAction(`Cleanup Rule ${ruleNumber}`, `Error: ${error.message}`, ruleTotal);
      
      return {
        success: true,
        message: `Rule ${ruleIndex + 1} failed, continuing...`,
        ruleIndex: ruleIndex + 1,
        ruleTotal: 0,
        emailsProcessed: 0,
        hasMoreInRule: false,
        hasMoreRules: (ruleIndex + 1) < activeRules.length,
        error: error.message,
        totalRules: activeRules.length,
        nextRuleIndex: ruleIndex + 1
      };
    }
    
  } catch (error) {
    return {
      success: false,
      message: `Error during cleanup: ${error.message}`,
      canResume: true,
      ruleIndex: ruleIndex,
      ruleTotal: ruleTotal
    };
  }
}

/**
 * Logs the overall result of a cleanup run from the sidebar
 * @param {number} totalEmails - Emails processed across all rules
 * @param {boolean} wasCancelled - Whether the run was cancelled part-way
 */
function logCleanupRun(totalEmails, wasCancelled) {
  logAction(
    wasCancelled ? 'Cleanup Cancelled' : 'Cleanup Complete',
    `Processed ${totalEmails} emails`,
    totalEmails
  );
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

// Triggers don't expose their schedule, so setupTriggers records it here
const SCHEDULE_FREQUENCY_KEY = 'scheduleFrequency';

/**
 * Sets up scheduled triggers for automation
 * @param {string} frequency - 'daily', 'weekly', or 'none'
 * @param {string} timeZone - IANA time zone to run in, e.g. 'Australia/Sydney'.
 *     Falls back to the script's time zone if missing or invalid.
 */
function setupTriggers(frequency, timeZone) {
  if (['daily', 'weekly', 'none'].indexOf(frequency) === -1) {
    return { success: false, message: 'Invalid frequency. Use "daily", "weekly", or "none".' };
  }
  
  const zone = isValidTimeZone(timeZone) ? timeZone : Session.getScriptTimeZone();
  
  try {
    const userProperties = PropertiesService.getUserProperties();
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runScheduledCleanup') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    if (frequency === 'none') {
      userProperties.deleteProperty(SCHEDULE_FREQUENCY_KEY);
      return { success: true, message: 'Automated cleanup disabled.' };
    }
    
    const trigger = ScriptApp.newTrigger('runScheduledCleanup');
    
    if (frequency === 'daily') {
      trigger.timeBased().everyDays(1).atHour(2).inTimezone(zone).create();
    } else {
      trigger.timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).inTimezone(zone).create();
    }
    userProperties.setProperty(SCHEDULE_FREQUENCY_KEY, frequency);
    
    return {
      success: true,
      message: `Automated cleanup scheduled to run ${frequency} at 2:00 AM (${zone} time).`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `Error setting up triggers: ${error.message}`
    };
  }
}

/**
 * Runs every active rule to completion from a time-driven trigger. Stops
 * early if it gets close to the execution time limit; the next scheduled
 * run picks up whatever is left.
 */
function runScheduledCleanup() {
  try {
    const { rules, error } = getActiveRules();
    if (error) {
      logAction('Scheduled Cleanup', error, 0);
      return;
    }
    
    const startTime = Date.now();
    const outOfTime = () => Date.now() - startTime > SCHEDULED_TIME_BUDGET_MS;
    let totalProcessed = 0;
    let stoppedEarly = false;
    
    for (const rule of rules) {
      const { ruleType, value, action, ruleNumber } = rule;
      let ruleTotal = 0;
      
      try {
        let processed;
        do {
          if (outOfTime()) {
            stoppedEarly = true;
            break;
          }
          processed = cleanupRuleBatch(rule);
          ruleTotal += processed || 0;
        } while (processed === CLEANUP_BATCH_SIZE);
        
        if (ruleTotal > 0) {
          logAction(`Cleanup Rule ${ruleNumber}`, `${action} - ${ruleType}: ${value}`, ruleTotal);
        }
        
      } catch (error) {
        logAction(`Cleanup Rule ${ruleNumber}`, `Error: ${error.message}`, ruleTotal);
      }
      
      totalProcessed += ruleTotal;
      if (stoppedEarly) break;
    }
    
    if (stoppedEarly) {
      logAction('Scheduled Cleanup Stopped', `Reached the time limit after ${totalProcessed} emails. The rest will be cleaned up on the next scheduled run.`, totalProcessed);
    } else {
      logAction('Scheduled Cleanup Complete', `Processed ${totalProcessed} emails`, totalProcessed);
    }
    
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

function isValidTimeZone(timeZone) {
  if (typeof timeZone !== 'string' || !timeZone) return false;
  try {
    Intl.DateTimeFormat('en-AU', { timeZone: timeZone });
    return true;
  } catch (error) {
    return false;
  }
}

/**
 * Reports whether automated cleanup is on, and how often it runs
 * @return {{enabled: boolean, frequency: string|null}} frequency is null for
 *     a schedule set up before it was recorded
 */
function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const cleanupTrigger = triggers.find(t => t.getHandlerFunction() === 'runScheduledCleanup');
  
  if (!cleanupTrigger || cleanupTrigger.getEventType() !== ScriptApp.EventType.CLOCK) {
    return { enabled: false, frequency: 'none' };
  }
  
  const frequency = PropertiesService.getUserProperties().getProperty(SCHEDULE_FREQUENCY_KEY);
  return { enabled: true, frequency: frequency };
}
