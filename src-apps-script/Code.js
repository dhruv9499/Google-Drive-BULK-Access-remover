/**
 * Google Drive Email Access Remover
 * 
 * Automatically removes specified email addresses from all Google Drive files
 * accessible by your account. Perfect for cleaning up access when employees,
 * freelancers, or associates leave your organization.
 * 
 * Author: [Dhruv Barot](mailto:dhruvbarot579@gmail.com)
 * Version: 1.0
 * License: MIT
 * 
 * SETUP REQUIRED:
 * 1. Enable Drive API v3 in Advanced Google Services
 * 2. Update TARGET_EMAILS array below with emails to remove
 * 3. Run startEmailCleanup() to begin the process
 */

// ==================== CONFIGURATION ====================

const TARGET_EMAILS = [
  'dhruv@wezake.com'
  // 'former.employee@company.com',
  // 'freelancer@example.com'
  // Add emails to remove here - one per line
];

const CONFIG = {
  BATCH_SIZE: 15,                    // Files to process per batch
  MAX_EXECUTION_TIME: 280000,        // 4.5 minutes (safety margin)
  RETRY_DELAY: 5000,                 // 5 seconds delay between batches
  MAX_LOG_SIZE: 8000                 // Max log size before truncation
};

// ==================== STATE MANAGEMENT ====================

const STATE_KEYS = {
  CURRENT_EMAIL_INDEX: 'currentEmailIndex',
  NEXT_PAGE_TOKEN: 'nextPageToken',
  PROCESSED_COUNT: 'processedCount',
  LOGS: 'processLogs',
  START_TIME: 'startTime',
  IS_RUNNING: 'isRunning',
  CURRENT_EMAIL: 'currentEmail'
};

// ==================== MAIN FUNCTIONS ====================

/**
 * Main function to start the email cleanup process
 * Call this function once to begin the entire process
 */
function startEmailCleanup() {
  console.log('🚀 Starting Google Drive Email Access Remover v1.0');
  
  // Validate configuration
  if (!validateConfiguration()) {
    return;
  }
  
  // Check if already running
  if (isProcessRunning()) {
    console.log('⚠️ Process is already running. Use checkStatus() to monitor progress.');
    return;
  }
  
  // Initialize and start
  initializeState();
  console.log(`🎯 Target emails: ${TARGET_EMAILS.join(', ')}`);
  console.log('📋 Starting batch processing...');
  
  processBatch();
}

/**
 * Check the current status of the cleanup process
 */
function checkStatus() {
  const properties = PropertiesService.getScriptProperties();
  const isRunning = properties.getProperty(STATE_KEYS.IS_RUNNING) === 'true';
  
  if (!isRunning) {
    console.log('💤 No process currently running');
    console.log('💡 Run startEmailCleanup() to begin cleanup');
    return { isRunning: false };
  }
  
  const processedCount = properties.getProperty(STATE_KEYS.PROCESSED_COUNT) || '0';
  const currentEmail = properties.getProperty(STATE_KEYS.CURRENT_EMAIL) || 'None';
  const currentEmailIndex = parseInt(properties.getProperty(STATE_KEYS.CURRENT_EMAIL_INDEX) || '0');
  const startTime = properties.getProperty(STATE_KEYS.START_TIME);
  
  console.log('📊 Process Status:');
  console.log(`   📧 Current email: ${currentEmail} (${currentEmailIndex + 1}/${TARGET_EMAILS.length})`);
  console.log(`   📁 Files processed: ${processedCount}`);
  console.log(`   ⏰ Started: ${new Date(startTime).toLocaleString()}`);
  console.log(`   ⏳ Running time: ${Math.round((Date.now() - new Date(startTime)) / 60000)} minutes`);
  
  return {
    isRunning: true,
    currentEmail,
    currentEmailIndex: currentEmailIndex + 1,
    totalEmails: TARGET_EMAILS.length,
    filesProcessed: parseInt(processedCount),
    startTime: startTime,
    runningMinutes: Math.round((Date.now() - new Date(startTime)) / 60000)
  };
}

/**
 * Emergency stop function - use only if needed
 */
function stopProcess() {
  clearState();
  console.log('🛑 Process manually stopped');
  console.log('💡 Run startEmailCleanup() to restart if needed');
}

// ==================== CORE PROCESSING ====================

/**
 * Processes a batch of files and schedules the next batch
 */
function processBatch() {
  const startTime = Date.now();
  const properties = PropertiesService.getScriptProperties();
  
  try {
    const currentEmailIndex = parseInt(properties.getProperty(STATE_KEYS.CURRENT_EMAIL_INDEX) || '0');
    const nextPageToken = properties.getProperty(STATE_KEYS.NEXT_PAGE_TOKEN);
    const processedCount = parseInt(properties.getProperty(STATE_KEYS.PROCESSED_COUNT) || '0');
    
    // Check if all emails processed
    if (currentEmailIndex >= TARGET_EMAILS.length) {
      completeProcess();
      return;
    }
    
    const currentEmail = TARGET_EMAILS[currentEmailIndex];
    console.log(`📧 Processing: ${currentEmail} (${currentEmailIndex + 1}/${TARGET_EMAILS.length})`);
    
    // Search for files shared with current email
    const searchQuery = `'${currentEmail}' in readers or '${currentEmail}' in writers or '${currentEmail}' in owners`;
    const filesResponse = Drive.Files.list({
      q: searchQuery,
      pageSize: CONFIG.BATCH_SIZE,
      pageToken: nextPageToken || undefined,
      fields: 'nextPageToken, files(id, name, mimeType)'
    });
    
    const files = filesResponse.files || [];
    console.log(`📋 Found ${files.length} files shared with ${currentEmail}`);
    
    // Process each file
    const batchLogs = [];
    let filesProcessedInBatch = 0;
    
    for (const file of files) {
      if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME) {
        console.log('⏰ Approaching execution time limit, stopping batch');
        break;
      }
      
      const result = processFile(file, currentEmail);
      batchLogs.push(result);
      filesProcessedInBatch++;
    }
    
    // Update state
    updateState(
      filesResponse.nextPageToken,
      processedCount + filesProcessedInBatch,
      batchLogs,
      currentEmailIndex,
      currentEmail
    );
    
    // Schedule next batch or move to next email
    if (filesResponse.nextPageToken) {
      scheduleNextBatch();
    } else {
      moveToNextEmail();
    }
    
  } catch (error) {
    console.error('❌ Error in processBatch:', error);
    handleProcessError(error);
  }
}

/**
 * Process a single file to remove target email permissions
 */
function processFile(file, targetEmail) {
  const fileInfo = {
    id: file.id,
    title: file.name,
    mimeType: file.mimeType,
    fileType: getFileType(file.mimeType),
    targetEmail: targetEmail,
    removed: false,
    error: null,
    skipped: false,
    foundButCantRemove: false,  // New field for permission issues
    targetPermissionRole: null,  // What access the target email had
    webViewLink: `https://drive.google.com/file/d/${file.id}/view`  // Direct link
  };
  
  try {
    // Get file permissions
    const permissionsResponse = Drive.Permissions.list(file.id, {
      fields: 'permissions(id, emailAddress, role, type)'
    });
    
    const permissions = permissionsResponse.permissions || [];
    
    // Find target email permission
    const targetPermission = permissions.find(p => p.emailAddress === targetEmail);
    
    if (targetPermission) {
      fileInfo.targetPermissionRole = targetPermission.role;
      
      try {
        // Try to remove the permission
        Drive.Permissions.remove(file.id, targetPermission.id);
        fileInfo.removed = true;
        console.log(`✅ Removed ${targetEmail} (${targetPermission.role}) from "${file.name}"`);
        
      } catch (removeError) {
        // This is the key scenario you asked about
        fileInfo.foundButCantRemove = true;
        fileInfo.error = `Found ${targetEmail} as ${targetPermission.role}, but insufficient permissions to remove`;
        
        console.error(`❌ Found ${targetEmail} as ${targetPermission.role} in "${file.name}" but cannot remove - insufficient permissions`);
        console.error(`🔗 File link: https://drive.google.com/file/d/${file.id}/view`);
      }
    } else {
      console.log(`ℹ️ ${targetEmail} not found in "${file.name}" permissions`);
    }
    
  } catch (permError) {
    if (permError.message.includes('insufficient permissions') || 
        permError.message.includes('not found') ||
        permError.message.includes('Forbidden')) {
      fileInfo.skipped = true;
      fileInfo.error = 'Cannot access file permissions';
      console.log(`⚠️ Skipped "${file.name}": Cannot access permissions`);
    } else {
      fileInfo.error = `Permission error: ${permError.message}`;
      console.error(`❌ Error accessing "${file.name}":`, permError.message);
    }
  }
  
  return fileInfo;
}

// ==================== STATE & WORKFLOW ====================

function validateConfiguration() {
  const currentUserEmail = Session.getActiveUser().getEmail();
  
  if (TARGET_EMAILS.length === 0) {
    console.error('❌ No target emails specified!');
    console.error('💡 Please update the TARGET_EMAILS array in the script');
    return false;
  }
  
  if (TARGET_EMAILS.includes(currentUserEmail)) {
    console.error(`❌ Cannot include your own email (${currentUserEmail}) in target list!`);
    console.error('💡 Remove your email from TARGET_EMAILS array');
    return false;
  }
  
  // Validate emails format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const invalidEmails = TARGET_EMAILS.filter(email => !emailRegex.test(email));
  
  if (invalidEmails.length > 0) {
    console.error(`❌ Invalid email format(s): ${invalidEmails.join(', ')}`);
    return false;
  }
  
  console.log(`✅ Configuration validated: ${TARGET_EMAILS.length} target email(s)`);
  return true;
}

function isProcessRunning() {
  return PropertiesService.getScriptProperties().getProperty(STATE_KEYS.IS_RUNNING) === 'true';
}

function initializeState() {
  const properties = PropertiesService.getScriptProperties();
  
  properties.setProperties({
    [STATE_KEYS.CURRENT_EMAIL_INDEX]: '0',
    [STATE_KEYS.NEXT_PAGE_TOKEN]: '',
    [STATE_KEYS.PROCESSED_COUNT]: '0',
    [STATE_KEYS.LOGS]: JSON.stringify([]),
    [STATE_KEYS.START_TIME]: new Date().toISOString(),
    [STATE_KEYS.IS_RUNNING]: 'true',
    [STATE_KEYS.CURRENT_EMAIL]: TARGET_EMAILS[0]
  });
}

function updateState(nextPageToken, processedCount, batchLogs, currentEmailIndex, currentEmail) {
  const properties = PropertiesService.getScriptProperties();
  
  // Handle logs with size limit
  let existingLogs = [];
  try {
    const logsJson = properties.getProperty(STATE_KEYS.LOGS);
    if (logsJson) {
      existingLogs = JSON.parse(logsJson);
    }
  } catch (error) {
    console.warn('⚠️ Could not parse existing logs, starting fresh');
  }
  
  const updatedLogs = existingLogs.concat(batchLogs);
  const logsJson = JSON.stringify(updatedLogs);
  
  if (logsJson.length > CONFIG.MAX_LOG_SIZE) {
    const truncatedLogs = updatedLogs.slice(-Math.floor(updatedLogs.length * 0.7));
    properties.setProperty(STATE_KEYS.LOGS, JSON.stringify(truncatedLogs));
  } else {
    properties.setProperty(STATE_KEYS.LOGS, logsJson);
  }
  
  // Update state
  properties.setProperties({
    [STATE_KEYS.NEXT_PAGE_TOKEN]: nextPageToken || '',
    [STATE_KEYS.PROCESSED_COUNT]: processedCount.toString(),
    [STATE_KEYS.CURRENT_EMAIL_INDEX]: currentEmailIndex.toString(),
    [STATE_KEYS.CURRENT_EMAIL]: currentEmail
  });
}

function moveToNextEmail() {
  const properties = PropertiesService.getScriptProperties();
  const currentEmailIndex = parseInt(properties.getProperty(STATE_KEYS.CURRENT_EMAIL_INDEX) || '0');
  const nextEmailIndex = currentEmailIndex + 1;
  
  if (nextEmailIndex >= TARGET_EMAILS.length) {
    completeProcess();
  } else {
    properties.setProperties({
      [STATE_KEYS.CURRENT_EMAIL_INDEX]: nextEmailIndex.toString(),
      [STATE_KEYS.NEXT_PAGE_TOKEN]: '',
      [STATE_KEYS.CURRENT_EMAIL]: TARGET_EMAILS[nextEmailIndex]
    });
    
    console.log(`➡️ Moving to next email: ${TARGET_EMAILS[nextEmailIndex]}`);
    scheduleNextBatch();
  }
}

function scheduleNextBatch() {
  console.log(`⏰ Scheduling next batch in ${CONFIG.RETRY_DELAY/1000} seconds...`);
  
  ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(CONFIG.RETRY_DELAY)
    .create();
}

function completeProcess() {
  console.log('🎉 All emails processed! Generating summary...');
  
  const properties = PropertiesService.getScriptProperties();
  const startTimeStr = properties.getProperty(STATE_KEYS.START_TIME);
  const processedCount = parseInt(properties.getProperty(STATE_KEYS.PROCESSED_COUNT) || '0');
  
  let logs = [];
  try {
    const logsJson = properties.getProperty(STATE_KEYS.LOGS);
    if (logsJson) {
      logs = JSON.parse(logsJson);
    }
  } catch (error) {
    console.warn('⚠️ Could not parse logs for summary');
  }
  
  // Generate and send summary
  const summary = generateSummary(logs, processedCount, startTimeStr);
  sendSummaryEmail(summary);
  
  // Clean up
  clearState();
  
  console.log('✅ Cleanup process completed successfully!');
}

function clearState() {
  const properties = PropertiesService.getScriptProperties();
  
  // Clear all state properties
  for (const key of Object.values(STATE_KEYS)) {
    properties.deleteProperty(key);
  }
  
  // Remove any pending triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function handleProcessError(error) {
  console.error('❌ Critical error occurred:', error);
  
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(STATE_KEYS.IS_RUNNING, 'false');
  
  const currentUser = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: currentUser,
    subject: '❌ Drive Email Cleanup - Process Error',
    body: `An error occurred during the Drive email cleanup process:\n\n${error.toString()}\n\nThe process has been stopped. Check the execution logs for more details.\n\nYou can restart by running startEmailCleanup() again.`
  });
}

// ==================== UTILITIES ====================

function getFileType(mimeType) {
  const typeMap = {
    'application/vnd.google-apps.document': 'Google Docs',
    'application/vnd.google-apps.spreadsheet': 'Google Sheets',
    'application/vnd.google-apps.presentation': 'Google Slides',
    'application/vnd.google-apps.folder': 'Folders',
    'application/vnd.google-apps.form': 'Google Forms',
    'application/vnd.google-apps.drawing': 'Google Drawings',
    'application/vnd.google-apps.map': 'Google My Maps',
    'application/vnd.google-apps.site': 'Google Sites',
    'application/pdf': 'PDF Files'
  };
  
  for (const [key, value] of Object.entries(typeMap)) {
    if (mimeType.startsWith(key)) {
      return value;
    }
  }
  
  if (mimeType.startsWith('image/')) return 'Images';
  if (mimeType.startsWith('video/')) return 'Videos';
  if (mimeType.startsWith('audio/')) return 'Audio Files';
  
  return 'Other Files';
}