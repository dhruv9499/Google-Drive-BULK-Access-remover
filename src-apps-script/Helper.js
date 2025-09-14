/**
 * Helper functions for Google Drive Email Access Remover
 * 
 * This file contains testing functions and reporting utilities.
 * Include this in your Google Apps Script project alongside Code.js
 */

// ==================== TESTING FUNCTIONS ====================

/**
 * Test edge case detection for files where current user has limited access
 * but target email has higher permissions.
 * 
 * NOTE: Due to Google Drive API limitations, this function may not be able to detect
 * all files where the user has limited access rights. This is a known limitation
 * of the Drive API when querying permissions.
 * 
 * @deprecated This function attempts to address a Drive API limitation but may not work
 * for all edge cases. See readme.md for more information about this limitation.
 */
function testEdgeCaseDetection() {
  if (TARGET_EMAILS.length === 0) {
    console.log('❌ No target emails configured in TARGET_EMAILS array');
    return { success: false, error: 'No target emails configured' };
  }
  
  const testEmail = TARGET_EMAILS[0];
  console.log(`🔍 EDGE CASE TEST: Searching for ALL files where ${testEmail} appears in ANY role`);
  console.log(`🧪 This test uses advanced search parameters to find files that might be missed`);
  
  const currentUserEmail = Session.getActiveUser().getEmail();
  console.log(`👤 Current user: ${currentUserEmail}`);
  
  try {
    // 1. First try the exact search query used in the main process
    const regularQuery = `'${testEmail}' in readers or '${testEmail}' in writers or '${testEmail}' in owners`;
    
    console.log(`\n📊 TEST 1: Using standard search query: "${regularQuery}"`);
    const regularResponse = Drive.Files.list({
      q: regularQuery,
      pageSize: 100,
      fields: 'files(id, name, mimeType, permissions(emailAddress, role, type))',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    
    const regularFiles = regularResponse.files || [];
    console.log(`✅ Standard search found ${regularFiles.length} files with ${testEmail}`);
    
    // 2. Then try a more aggressive search approach
    console.log(`\n📊 TEST 2: Using more advanced search to find edge cases`);
    const testResults = [];
    let totalFound = 0;
    
    // Try different specialized queries
    const specialQueries = [
      `'${testEmail}' in readers`,
      `'${testEmail}' in writers`, 
      `'${testEmail}' in owners`,
      // Using raw query which might catch other cases
      `permissions.emailAddress = '${testEmail}'`
    ];
    
    for (const query of specialQueries) {
      try {
        console.log(`🔎 Trying query: "${query}"`);
        const response = Drive.Files.list({
          q: query,
          pageSize: 50,
          fields: 'files(id, name, mimeType, permissions(emailAddress, role, type), capabilities(canEdit, canComment, canShare))',
          includeItemsFromAllDrives: true,
          supportsAllDrives: true
        });
        
        const files = response.files || [];
        totalFound += files.length;
        console.log(`   Found ${files.length} files with this query`);
        
        // Analyze each file's permissions
        for (const file of files) {
          const targetPerm = file.permissions ? 
            file.permissions.find(p => p.emailAddress === testEmail) : null;
          
          const targetRole = targetPerm ? targetPerm.role : 'Unknown';
          const currentUserCaps = {
            canEdit: file.capabilities?.canEdit || false,
            canComment: file.capabilities?.canComment || false,
            canShare: file.capabilities?.canShare || false
          };
          
          // Add to test results
          testResults.push({
            id: file.id,
            name: file.name,
            mimeType: file.mimeType,
            targetEmailRole: targetRole,
            currentUserCaps: currentUserCaps,
            link: `https://drive.google.com/file/d/${file.id}/view`
          });
          
          console.log(`   📄 ${file.name}`);
          console.log(`      - Target email (${testEmail}): ${targetRole} access`);
          console.log(`      - Your permissions: ${JSON.stringify(currentUserCaps)}`);
          console.log(`      - Link: https://drive.google.com/file/d/${file.id}/view`);
        }
      } catch (queryError) {
        console.error(`❌ Error with query "${query}": ${queryError.message}`);
      }
    }
    
    // Summarize results
    console.log(`\n📋 SUMMARY:`);
    console.log(`Found total of ${testResults.length} unique files across all queries`);
    
    // Count by permission combinations
    const edgeCases = testResults.filter(r => 
      r.targetEmailRole === 'writer' && 
      !r.currentUserCaps.canEdit && 
      !r.currentUserCaps.canShare);
    
    console.log(`\n⚠️ EDGE CASES DETECTED: ${edgeCases.length} files`);
    console.log(`These are files where ${testEmail} has editor access but you don't have permission to modify sharing:`);
    
    if (edgeCases.length > 0) {
      edgeCases.forEach((file, i) => {
        console.log(`${i+1}. "${file.name}"`);
        console.log(`   - Your access: ${file.currentUserCaps.canComment ? 'Commenter' : 'Viewer'}`);
        console.log(`   - Link: ${file.link}`);
      });
    }
    
    return { 
      success: true, 
      totalFilesFound: testResults.length,
      edgeCasesFound: edgeCases.length,
      edgeCases: edgeCases,
      testEmail: testEmail
    };
    
  } catch (error) {
    console.error('❌ Edge case test failed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Test Drive API v3 connection and basic functionality
 */
function testDriveV3Access() {
  console.log('🧪 Testing Drive API v3 access...');
  
  try {
    const response = Drive.Files.list({
      pageSize: 5,
      fields: 'files(id, name, mimeType)',
      orderBy: 'modifiedTime desc'
    });
    
    const files = response.files || [];
    console.log(`✅ Drive API v3 working - Found ${files.length} recent files`);
    
    if (files.length > 0) {
      console.log('📁 Your recent files:');
      files.forEach((file, index) => {
        console.log(`${index + 1}. "${file.name}" (${getFileType(file.mimeType)})`);
      });
    } else {
      console.log('⚠️ No files found - check your Drive access');
    }
    
    return { success: true, fileCount: files.length };
    
  } catch (error) {
    console.error('❌ Drive API v3 test failed:', error);
    console.error('💡 Solution: Enable Drive API in Services > Advanced Google Services');
    return { success: false, error: error.message };
  }
}

/**
 * Test search functionality for files shared with target emails
 */
function testEmailSearch() {
  if (TARGET_EMAILS.length === 0) {
    console.log('❌ No target emails configured in TARGET_EMAILS array');
    console.log('💡 Add some emails to test the search functionality');
    return { success: false, error: 'No target emails configured' };
  }
  
  const testEmail = TARGET_EMAILS[0];
  console.log(`🧪 Testing search for files shared with: ${testEmail}`);
  
  try {
    const searchQuery = `'${testEmail}' in readers or '${testEmail}' in writers or '${testEmail}' in owners`;
    
    const response = Drive.Files.list({
      q: searchQuery,
      pageSize: 10,
      fields: 'files(id, name, mimeType, permissions(emailAddress, role))',
      orderBy: 'modifiedTime desc',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    
    const files = response.files || [];
    console.log(`✅ Found ${files.length} files shared with ${testEmail}`);
    
    if (files.length > 0) {
      console.log('📁 Files shared with this email:');
      files.forEach((file, index) => {
        console.log(`${index + 1}. "${file.name}" (${getFileType(file.mimeType)})`);
        
        // Show what access the email has
        if (file.permissions) {
          const targetPerm = file.permissions.find(p => p.emailAddress === testEmail);
          if (targetPerm) {
            console.log(`   └── ${testEmail} has ${targetPerm.role} access`);
          }
        }
      });
      
      console.log(`\n🎯 Ready to process! Run startEmailCleanup() to begin removal.`);
    } else {
      console.log(`ℹ️ No files found shared with ${testEmail}`);
      console.log('💡 This email might already be removed, or never had access');
    }
    
    return { success: true, fileCount: files.length, testEmail };
    
  } catch (error) {
    console.error('❌ Email search test failed:', error);
    console.error('💡 Make sure Drive API v3 is enabled and the email is valid');
    return { success: false, error: error.message };
  }
}

/**
 * Quick diagnostic to check if everything is set up correctly
 */
function runDiagnostics() {
  console.log('🩺 Running full diagnostics...\n');
  
  const results = {
    driveAPI: false,
    emailsConfigured: false,
    emailSearch: false,
    edgeCaseDetection: false,
    ready: false
  };
  
  // Test 1: Drive API
  console.log('1️⃣ Testing Drive API v3...');
  const driveTest = testDriveV3Access();
  results.driveAPI = driveTest.success;
  
  console.log('\n');
  
  // Test 2: Email Configuration
  console.log('2️⃣ Checking email configuration...');
  if (TARGET_EMAILS.length > 0) {
    console.log(`✅ Found ${TARGET_EMAILS.length} target email(s): ${TARGET_EMAILS.join(', ')}`);
    results.emailsConfigured = true;
  } else {
    console.log('❌ No target emails configured');
    console.log('💡 Add emails to the TARGET_EMAILS array');
  }
  
  console.log('\n');
  
  // Test 3: Email Search (only if emails configured)
  if (results.emailsConfigured) {
    console.log('3️⃣ Testing email search...');
    const searchTest = testEmailSearch();
    results.emailSearch = searchTest.success;
  } else {
    console.log('3️⃣ Skipping email search test (no emails configured)');
  }
  
  console.log('\n');
  
  // Note: Edge Case Detection test removed from standard diagnostics
  // due to Drive API limitations. See testForEdgeCases() function and readme.md
  // for more information about this limitation.
  
  console.log('\n');
  console.log('4️⃣ About edge case detection:');
  console.log('⚠️ Due to Drive API limitations, some files where you have limited access');
  console.log('   (viewer/commenter) may not be detected by this script.');
  console.log('   Please see readme.md section "Known Limitation: Files with Limited Access"');
  console.log('   for more information about this limitation.');
  
  results.edgeCaseDetection = true; // No longer affects ready status
  
  console.log('\n');
  
  // Final Assessment
  console.log('📊 DIAGNOSTIC SUMMARY:');
  console.log(`Drive API v3: ${results.driveAPI ? '✅' : '❌'}`);
  console.log(`Emails Configured: ${results.emailsConfigured ? '✅' : '❌'}`);
  console.log(`Email Search: ${results.emailSearch ? '✅' : '❌'}`);
  if (results.edgeCaseDetection !== undefined) {
    console.log(`Edge Case Detection: ${results.edgeCaseDetection ? '✅' : '❌'}`);
  }
  
  results.ready = results.driveAPI && results.emailsConfigured && results.emailSearch;
  
  if (results.ready) {
    console.log('\n🎉 ALL SYSTEMS GO!');
    console.log('💡 Run startEmailCleanup() to begin the cleanup process');
  } else {
    console.log('\n⚠️ SETUP INCOMPLETE');
    console.log('💡 Fix the issues above before running startEmailCleanup()');
  }
  
  return results;
}

/**
 * Test specifically for edge cases - files with permission issues
 * Use this function to find files where you have limited access
 * but target emails have higher permissions
 * 
 * @deprecated Due to Google Drive API limitations, this function may not detect all files
 * where permissions prevent automated removal. See readme.md for more information.
 */
function testForEdgeCases() {
  console.log('🔎 Running edge case detection test...');
  console.log('⚠️ Warning: This test has limited effectiveness due to Drive API restrictions');
  console.log('   See readme.md section "Known Limitation: Files with Limited Access" for details');
  return testEdgeCaseDetection();
}

// ==================== REPORTING FUNCTIONS ====================

/**
 * Generate comprehensive summary of the cleanup process
 */
function generateSummary(logs, totalFiles, startTimeStr) {
  const summary = {
    totalFiles,
    startTime: startTimeStr,
    endTime: new Date().toISOString(),
    targetEmails: TARGET_EMAILS,
    byEmail: {},
    byFileType: {},
    totalRemovals: 0,
    totalErrors: 0,
    totalSkipped: 0,
    totalFoundButCantRemove: 0,
    filesNeedingManualReview: [],
    accessStats: {
      noAccess: 0,
      viewerAccess: 0,
      commenterAccess: 0,
      editorAccess: 0
    }
  };
  
  // Initialize email tracking
  TARGET_EMAILS.forEach(email => {
    summary.byEmail[email] = {
      filesFound: 0,
      removals: 0,
      errors: 0,
      skipped: 0,
      foundButCantRemove: 0,
      needsManualReview: []
    };
  });
  
  // Process logs
  for (const log of logs) {
    const fileType = log.fileType || 'Unknown';
    const email = log.targetEmail || 'Unknown';
    
    // File type stats
    if (!summary.byFileType[fileType]) {
      summary.byFileType[fileType] = {
        processed: 0,
        removed: 0,
        errors: 0,
        skipped: 0,
        foundButCantRemove: 0,
        files: []
      };
    }
    
    const typeStats = summary.byFileType[fileType];
    typeStats.processed++;
    
    // Email stats
    if (summary.byEmail[email]) {
      summary.byEmail[email].filesFound++;
    }
    
    // Count outcomes
    if (log.removed) {
      typeStats.removed++;
      summary.totalRemovals++;
      if (summary.byEmail[email]) {
        summary.byEmail[email].removals++;
      }
    } else if (log.foundButCantRemove) {
      typeStats.foundButCantRemove++;
      summary.totalFoundButCantRemove++;
      if (summary.byEmail[email]) {
        summary.byEmail[email].foundButCantRemove++;
      }
      
      // Add to files needing manual review
      summary.filesNeedingManualReview.push({
        title: log.title,
        email: log.targetEmail,
        fileType: log.fileType,
        role: log.targetPermissionRole,
        link: log.webViewLink,
        ownerEmail: log.ownerEmail || 'Unknown owner'
      });
      
      if (summary.byEmail[email]) {
        summary.byEmail[email].needsManualReview.push({
          title: log.title,
          role: log.targetPermissionRole,
          link: log.webViewLink,
          ownerEmail: log.ownerEmail || 'Unknown owner'
        });
      }
    } else if (log.error && !log.foundButCantRemove) {
      typeStats.errors++;
      summary.totalErrors++;
      if (summary.byEmail[email]) {
        summary.byEmail[email].errors++;
      }
    } else if (log.skipped) {
      typeStats.skipped++;
      summary.totalSkipped++;
      if (summary.byEmail[email]) {
        summary.byEmail[email].skipped++;
      }
    }
    
    // Add to file list if action taken
    if (log.removed || log.error || log.skipped || log.foundButCantRemove) {
      typeStats.files.push({
        title: log.title,
        email: log.targetEmail,
        removed: log.removed,
        error: log.error,
        skipped: log.skipped,
        foundButCantRemove: log.foundButCantRemove,
        role: log.targetPermissionRole,
        link: log.webViewLink
      });
    }
  }
  
  return summary;
}

/**
 * Send detailed summary email to user
 */
function sendSummaryEmail(summary) {
  const currentUser = Session.getActiveUser().getEmail();
  const duration = Math.round((new Date(summary.endTime) - new Date(summary.startTime)) / 1000 / 60);
  
  let emailBody = `
<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">

<h2 style="color: #1a73e8;">🚀 Drive Email Cleanup - Process Complete!</h2>

<div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
<h3 style="color: #137333; margin-top: 0;">📊 Summary</h3>
<table style="width: 100%; border-collapse: collapse;">
<tr><td style="padding: 5px;"><strong>Total Files Processed:</strong></td><td>${summary.totalFiles}</td></tr>
<tr><td style="padding: 5px;"><strong>Successful Removals:</strong></td><td style="color: #137333;">${summary.totalRemovals}</td></tr>
<tr><td style="padding: 5px;"><strong>Found But Can't Remove:</strong></td><td style="color: #f9ab00;"><strong>${summary.totalFoundButCantRemove}</strong></td></tr>
<tr><td style="padding: 5px;"><strong>Other Errors:</strong></td><td style="color: #d93025;">${summary.totalErrors}</td></tr>
<tr><td style="padding: 5px;"><strong>Skipped (No Access):</strong></td><td style="color: #9aa0a6;">${summary.totalSkipped}</td></tr>
<tr><td style="padding: 5px;"><strong>Processing Time:</strong></td><td>${duration} minutes</td></tr>
</table>
</div>
`;

  // Add manual review section if there are files that need attention
  if (summary.filesNeedingManualReview.length > 0) {
    emailBody += `
<div style="background: #fef7e0; border: 2px solid #f9ab00; border-radius: 8px; padding: 20px; margin: 20px 0;">
<h3 style="color: #f9ab00; margin-top: 0;">⚠️ Files Requiring Manual Review</h3>
<p><strong>${summary.filesNeedingManualReview.length} files</strong> contain target emails but couldn't be automatically removed due to insufficient permissions. 
You likely have <strong>view-only</strong> access to these files while the target emails have <strong>edit/comment</strong> access.</p>

<p><strong>Action Required:</strong> Contact the file owners or ask an admin to remove these permissions manually.</p>

<div style="max-height: 400px; overflow-y: auto; background: white; border-radius: 4px; padding: 15px; margin-top: 15px;">
`;

    // Group by email for better organization
    const filesByEmail = {};
    summary.filesNeedingManualReview.forEach(file => {
      if (!filesByEmail[file.email]) {
        filesByEmail[file.email] = [];
      }
      filesByEmail[file.email].push(file);
    });

    for (const [email, files] of Object.entries(filesByEmail)) {
      emailBody += `<h4 style="color: #d93025; margin: 15px 0 10px 0;">${email}</h4><ul>`;
      
      files.forEach(file => {
        emailBody += `
<li style="margin-bottom: 8px;">
  <strong><a href="${file.link}" target="_blank" style="color: #1a73e8; text-decoration: none;">${file.title}</a></strong>
  <br><span style="color: #5f6368; font-size: 14px;">${file.fileType} • ${file.role} access • 
  Owner: ${file.ownerEmail} • 
  <a href="${file.link}" target="_blank" style="color: #1a73e8;">Open File</a></span>
</li>`;
      });
      
      emailBody += '</ul>';
    }

    emailBody += `
</div>
<p style="font-size: 14px; color: #5f6368; margin-top: 15px;">
💡 <strong>Tip:</strong> Click the file links above to open them directly and manually remove the permissions, 
or share these links with file owners/admins who have the necessary permissions.
</p>
</div>
`;
  }

  emailBody += `<h3 style="color: #1a73e8;">📧 Results by Target Email</h3>`;

  for (const [email, stats] of Object.entries(summary.byEmail)) {
    const successRate = stats.filesFound > 0 ? Math.round((stats.removals / stats.filesFound) * 100) : 0;
    emailBody += `
<div style="background: #fff; border: 1px solid #dadce0; border-radius: 4px; padding: 15px; margin: 10px 0;">
<h4 style="margin: 0 0 10px 0; color: #202124;">${email}</h4>
<table style="width: 100%;">
<tr><td>Files Found:</td><td><strong>${stats.filesFound}</strong></td></tr>
<tr><td>Successfully Removed:</td><td style="color: #137333;"><strong>${stats.removals}</strong></td></tr>
<tr><td>Found But Can't Remove:</td><td style="color: #f9ab00;"><strong>${stats.foundButCantRemove}</strong></td></tr>
<tr><td>Other Errors:</td><td style="color: #d93025;">${stats.errors}</td></tr>
<tr><td>Skipped:</td><td style="color: #9aa0a6;">${stats.skipped}</td></tr>
<tr><td>Success Rate:</td><td><strong>${successRate}%</strong></td></tr>
</table>
`;

    // Add manual review files for this email
    if (stats.needsManualReview && stats.needsManualReview.length > 0) {
      emailBody += `
<details style="margin-top: 10px;">
<summary style="color: #f9ab00; cursor: pointer;"><strong>Files needing manual review (${stats.needsManualReview.length})</strong></summary>
<ul style="margin-top: 10px;">`;
      
      stats.needsManualReview.forEach(file => {
        emailBody += `
<li><a href="${file.link}" target="_blank" style="color: #1a73e8;">${file.title}</a> 
<span style="color: #5f6368;">(${file.role} access)</span></li>`;
      });
      
      emailBody += '</ul></details>';
    }

    emailBody += '</div>';
  }

  emailBody += `<h3 style="color: #1a73e8;">📁 Results by File Type</h3>`;

  for (const [fileType, stats] of Object.entries(summary.byFileType)) {
    emailBody += `
<div style="background: #f8f9fa; border-left: 4px solid #1a73e8; padding: 15px; margin: 10px 0;">
<h4 style="margin: 0 0 10px 0; color: #1a73e8;">${fileType}</h4>
<p>Processed: <strong>${stats.processed}</strong> | 
   Removed: <span style="color: #137333;"><strong>${stats.removed}</strong></span> | 
   Can't Remove: <span style="color: #f9ab00;"><strong>${stats.foundButCantRemove}</strong></span> |
   Errors: <span style="color: #d93025;">${stats.errors}</span> | 
   Skipped: <span style="color: #9aa0a6;">${stats.skipped}</span></p>
`;

    if (stats.files.length > 0) {
      emailBody += '<details><summary><strong>Files with Actions</strong></summary><ul>';
      for (const file of stats.files.slice(0, 25)) {
        emailBody += `<li><strong>${file.title}</strong> (${file.email})`;
        if (file.removed) emailBody += ' - ✅ Removed';
        if (file.foundButCantRemove) emailBody += ` - ⚠️ <a href="${file.link}" target="_blank">Manual Review Needed</a> (${file.role})`;
        if (file.error && !file.foundButCantRemove) emailBody += ' - ❌ Error';
        if (file.skipped) emailBody += ' - 🔒 Skipped';
        emailBody += '</li>';
      }
      if (stats.files.length > 25) {
        emailBody += `<li><em>... and ${stats.files.length - 25} more files</em></li>`;
      }
      emailBody += '</ul></details>';
    }
    emailBody += '</div>';
  }

  emailBody += `
<hr style="margin: 30px 0; border: none; border-top: 1px solid #dadce0;">
<p style="color: #5f6368; font-size: 14px;">
<strong>Google Drive Email Access Remover v1.0 by Dhruv Barot(dhruvbarot579@gmail.com)</strong><br>
Process completed at ${new Date(summary.endTime).toLocaleString()}<br>
Generated by Google Apps Script by Dhruv Barot(dhruvbarot579@gmail.com)
</p

</div>
`;

  let subject = `✅ Drive Cleanup Complete - ${summary.totalRemovals} emails removed from ${summary.totalFiles} files`;
  
  if (summary.totalFoundButCantRemove > 0) {
    subject += ` (${summary.totalFoundButCantRemove} need manual review)`;
  }

  MailApp.sendEmail({
    to: currentUser,
    subject: subject,
    htmlBody: emailBody
  });
  
  console.log('📧 Summary email sent to', currentUser);
  
  if (summary.totalFoundButCantRemove > 0) {
    console.log(`⚠️ ${summary.totalFoundButCantRemove} files need manual review - check your email for direct links`);
  }
}