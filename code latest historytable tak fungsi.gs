// ============================================
// KAWAD KAKI SYSTEM - CLEAN VERSION
// ============================================

const SPREADSHEET_ID = '1Hoa0jEP85ppNYBZhsASDTQI__kDGzyfvbSoXs76b3uY';

// ============================================
// HANDLE PAGE REQUESTS
// ============================================
function doGet(e) {
  try {
    const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'login';
    
    Logger.log('doGet called with page: ' + page);
    
    return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('Sistem Penilian Kawad Kaki')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    Logger.log('doGet error: ' + error.toString());
    
    const errorHtml = '<html><body style="padding:50px;font-family:Arial;">' +
      '<h1>Error Loading Page</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<p><a href="?page=login">Back to Login</a></p>' +
      '</body></html>';
    
    return HtmlService.createHtmlOutput(errorHtml);
  }
}

// ============================================
// LOGIN FUNCTION
// ============================================
function doLogin(username, password) {
  Logger.log('=== LOGIN ATTEMPT START ===');
  Logger.log('Username received: [' + username + ']');
  Logger.log('Password received: [' + password + ']');
  
    // Try database lookup
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName('USERS');
    
    if (!usersSheet) {
      Logger.log('USERS sheet not found');
      return { 
        success: false, 
        message: 'Username atau password salah' 
      };
    }
    
    const data = usersSheet.getDataRange().getValues();
    Logger.log('Checking database with ' + (data.length - 1) + ' users');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      const userId = String(row[0]).trim();
      const dbUsername = String(row[1]).trim();
      const dbPassword = String(row[2]).trim();
      const role = String(row[3]).trim();
      const fullName = String(row[4]).trim();
      const isActive = row[5];
      
      if (dbUsername.toLowerCase() === username.toLowerCase() && dbPassword === password) {
        if (isActive !== true && isActive !== 'TRUE' && isActive !== 'true') {
          return { success: false, message: 'User tidak aktif' };
        }
        
        Logger.log('Database login successful: ' + fullName);
        return { 
          success: true, 
          role: role, 
          full_name: fullName,
          user_id: userId,
          username: dbUsername
        };
      }
    }
    
    Logger.log('No matching credentials in database');
    return { 
      success: false, 
      message: 'Username atau password salah' 
    };
    
  } catch (error) {
    Logger.log('Database error: ' + error.toString());
    return { 
      success: false, 
      message: 'Username atau password salah' 
    };
  }
}

// ============================================
// GET DATA FROM SHEETS
// ============================================
function getTeams() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TEAMS');
    
    if (!sheet) {
      Logger.log('ERROR: TEAMS sheet not found!');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('TEAMS sheet - Total rows: ' + data.length);
    
    const teams = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        teams.push({
          team_id: data[i][0],
          team_code: data[i][1],
          school_name: data[i][2],
          kategori: data[i][3],
          gender: data[i][4],
          school_level: data[i][5],
          ketua_platun_name: data[i][6],
          is_active: data[i][7]
        });
        Logger.log('Team ' + i + ': ' + data[i][1] + ' - is_active: ' + data[i][7] + ' (type: ' + typeof data[i][7] + ')');
      }
    }
    
    Logger.log('Teams loaded: ' + teams.length);
    return teams;
    
  } catch (e) {
    Logger.log('Error loading teams: ' + e.toString());
    return [];
  }
}

function getJudges() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('JUDGES');
    const data = sheet.getDataRange().getValues();
    
    const judges = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        judges.push({
          judge_id: data[i][0],
          judge_name: data[i][1],
          judge_number: data[i][2],
          judge_role: data[i][3],
          is_active: data[i][4]
        });
      }
    }
    
    Logger.log('Judges loaded: ' + judges.length);
    return judges;
    
  } catch (e) {
    Logger.log('Error loading judges: ' + e.toString());
    return [];
  }
}

// ============================================
// JUDGE CRUD FUNCTIONS
// ============================================

function addJudge(judgeData) {
  try {
    Logger.log('=== ADD JUDGE START ===');
    Logger.log('Judge data: ' + JSON.stringify(judgeData));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('JUDGES');
    
    if (!sheet) {
      return { success: false, message: 'JUDGES sheet not found' };
    }
    
    // Validate required fields
    if (!judgeData.judge_name || !judgeData.judge_name.trim()) {
      return { success: false, message: 'Nama hakim diperlukan' };
    }
    
    if (!judgeData.judge_number || !judgeData.judge_number.trim()) {
      return { success: false, message: 'Nombor hakim diperlukan' };
    }
    
    if (!judgeData.judge_role) {
      return { success: false, message: 'Peranan hakim diperlukan' };
    }
    
    // Generate new judge_id
    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const idNum = parseInt(data[i][0].toString().replace('J', ''));
        if (!isNaN(idNum) && idNum > maxId) {
          maxId = idNum;
        }
      }
    }
    
    const newId = 'J' + String(maxId + 1).padStart(3, '0');
    Logger.log('New judge_id: ' + newId);
    
    // Check duplicate judge_number (active only)
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && 
          data[i][2].toString().toUpperCase() === judgeData.judge_number.toUpperCase() && 
          data[i][4] !== false) {
        return { 
          success: false, 
          message: 'Nombor hakim "' + judgeData.judge_number + '" sudah wujud!' 
        };
      }
    }
    
    // Add new judge
    const newRow = [
      newId,
      judgeData.judge_name.trim(),
      judgeData.judge_number.trim().toUpperCase(),
      judgeData.judge_role,
      true
    ];
    
    sheet.appendRow(newRow);
    
    // Log to audit
    const userId = judgeData.user_id || 'SYSTEM';
    logToAudit(
      userId,
      'CREATE',
      'JUDGES',
      newId,
      '',
      'Name: ' + judgeData.judge_name.trim() + ', Number: ' + judgeData.judge_number.trim().toUpperCase() + ', Role: ' + judgeData.judge_role,
      'Added new judge'
    );
    
    Logger.log('Judge added: ' + newId);
    return { 
      success: true, 
      message: 'Hakim berjaya ditambah!',
      judge_id: newId
    };
    
  } catch (e) {
    Logger.log('Error adding judge: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function updateJudge(judgeData) {
  try {
    Logger.log('=== UPDATE JUDGE START ===');
    Logger.log('Judge data: ' + JSON.stringify(judgeData));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('JUDGES');
    
    if (!sheet) {
      return { success: false, message: 'JUDGES sheet not found' };
    }
    
    // Validate
    if (!judgeData.judge_id) {
      return { success: false, message: 'Judge ID missing' };
    }
    
    if (!judgeData.judge_name || !judgeData.judge_name.trim()) {
      return { success: false, message: 'Nama hakim diperlukan' };
    }
    
    if (!judgeData.judge_number || !judgeData.judge_number.trim()) {
      return { success: false, message: 'Nombor hakim diperlukan' };
    }
    
    if (!judgeData.judge_role) {
      return { success: false, message: 'Peranan hakim diperlukan' };
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Find the judge
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === judgeData.judge_id) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Hakim tidak dijumpai' };
    }
    
    // Check duplicate judge_number (excluding current judge)
    for (let i = 1; i < data.length; i++) {
      if (i !== rowIndex && 
          data[i][2] && 
          data[i][2].toString().toUpperCase() === judgeData.judge_number.toUpperCase() &&
          data[i][4] !== false) {
        return { 
          success: false, 
          message: 'Nombor hakim "' + judgeData.judge_number + '" sudah digunakan!' 
        };
      }
    }
    
    // Update (rowIndex+1 because sheets are 1-indexed)
    const sheetRow = rowIndex + 1;
    
    // Get old values for audit
    const oldData = {
      name: data[rowIndex][1],
      number: data[rowIndex][2],
      role: data[rowIndex][3]
    };
    
    sheet.getRange(sheetRow, 2).setValue(judgeData.judge_name.trim());
    sheet.getRange(sheetRow, 3).setValue(judgeData.judge_number.trim().toUpperCase());
    sheet.getRange(sheetRow, 4).setValue(judgeData.judge_role);
    
    // Log to audit
    const userId = judgeData.user_id || 'SYSTEM';
    const newData = {
      name: judgeData.judge_name.trim(),
      number: judgeData.judge_number.trim().toUpperCase(),
      role: judgeData.judge_role
    };
    logToAudit(
      userId,
      'UPDATE',
      'JUDGES',
      judgeData.judge_id,
      'Old: Name=' + oldData.name + ', Number=' + oldData.number + ', Role=' + oldData.role,
      'New: Name=' + newData.name + ', Number=' + newData.number + ', Role=' + newData.role,
      'Updated judge info'
    );
    
    Logger.log('Judge updated: ' + judgeData.judge_id);
    return { success: true, message: 'Hakim berjaya dikemaskini!' };
    
  } catch (e) {
    Logger.log('Error updating judge: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function deleteJudge(judgeId, userId) {
  try {
    Logger.log('=== DELETE JUDGE START ===');
    Logger.log('Judge ID: ' + judgeId);
    
    userId = userId || 'SYSTEM';
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('JUDGES');
    
    if (!sheet) {
      return { success: false, message: 'JUDGES sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Find the judge
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === judgeId) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Hakim tidak dijumpai' };
    }
    
    // Get judge info for audit
    const judgeName = data[rowIndex][1];
    const judgeNumber = data[rowIndex][2];
    const judgeRole = data[rowIndex][3];
    
    // Soft delete - set is_active to false
    const sheetRow = rowIndex + 1;
    sheet.getRange(sheetRow, 5).setValue(false);
    
    // Log to audit
    logToAudit(
      userId,
      'DELETE',
      'JUDGES',
      judgeId,
      'Judge: ' + judgeName + ' (' + judgeNumber + '), Role: ' + judgeRole + ', Active: TRUE',
      'Active: FALSE (Soft deleted)',
      'Soft deleted judge'
    );
    
    Logger.log('Judge deleted (soft): ' + judgeId);
    return { success: true, message: 'Hakim berjaya dipadam!' };
    
  } catch (e) {
    Logger.log('Error deleting judge: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
function getConfig() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('CONFIG_SETTINGS');
    const data = sheet.getDataRange().getValues();
    
    const config = {};
    for (let i = 1; i < data.length; i++) {
      const [key, value] = data[i];
      if (key) {
        config[key] = value;
      }
    }
    
    Logger.log('Config loaded: ' + Object.keys(config).length + ' settings');
    return config;
    
  } catch (e) {
    Logger.log('Error loading config: ' + e.toString());
    return {};
  }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// SAVE SCORES TO SHEETS
// ============================================
function saveScores(submission) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return { 
      success: false, 
      message: 'Sistem sedang sibuk. Sila cuba lagi.' 
    };
  }

  try {
    Logger.log('=== SAVE SCORES START ===');
    Logger.log('Submission received: ' + JSON.stringify(submission));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const formType = submission.formType;
    
    Logger.log('Form type extracted: [' + formType + ']');
    
    if (!formType) {
      Logger.log('ERROR: formType is null or undefined');
      lock.releaseLock();
      return { success: false, message: 'Form type missing in submission' };
    }
    
    // Determine target sheet name
    let sheetName = '';
    switch(formType) {
      case 'kawad_kp': 
        sheetName = 'SCORES_KAWAD_KP'; 
        break;
      case 'kawad_platun': 
        sheetName = 'SCORES_KAWAD_PLATUN'; 
        break;
      case 'pakaian_kp': 
        sheetName = 'SCORES_PAKAIAN_KP'; 
        break;
      case 'pakaian_platun': 
        sheetName = 'SCORES_PAKAIAN_PLATUN'; 
        break;
      case 'formasi': 
        sheetName = 'SCORES_FORMASI'; 
        break;
      default:
        Logger.log('ERROR: Invalid form type: [' + formType + ']');
        lock.releaseLock();
        return { success: false, message: 'Invalid form type: ' + formType };
    }
    
    Logger.log('Target sheet: ' + sheetName);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('ERROR: Sheet not found: ' + sheetName);
      lock.releaseLock();
      return { success: false, message: 'Sheet not found: ' + sheetName };
    }
    
    const timestamp = new Date();
    const entryId = 'ENTRY_' + timestamp.getTime() + '_' + Math.floor(Math.random() * 1000);
    
    Logger.log('Entry ID: ' + entryId);
    Logger.log('Number of scores: ' + submission.scores.length);
    Logger.log('Number of penalties: ' + submission.penalties.length);
    
    // Calculate totals
    let totalScore = 0;
    let totalPenalty = 0;
    
    // Save scores based on form type
    if (formType === 'kawad_kp') {
      // KAWAD KP: score_pergerakan & score_bahasa in separate columns + edited
      Logger.log('Saving KAWAD KP scores...');
      
      submission.scores.forEach(function(scoreItem) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          scoreItem.code,
          scoreItem.score_pergerakan || 0,
          scoreItem.score_bahasa || 0,
          '', // penalty_code
          0,  // penalty_value
          submission.keyed_by,
          false // edited - NEW!
        ];
        sheet.appendRow(row);
        
        totalScore += parseInt(scoreItem.score_pergerakan || 0);
        totalScore += parseInt(scoreItem.score_bahasa || 0);
      });
      
      // Save penalties
      submission.penalties.forEach(function(penalty) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          '', // item_code
          0, 0, // scores
          penalty.code,
          penalty.value,
          submission.keyed_by,
          false // edited - NEW!
        ];
        sheet.appendRow(row);
        totalPenalty += parseInt(penalty.value || 0);
      });
      
    } else if (formType === 'formasi') {
      // FORMASI: score_formasi1-4 in separate columns + edited
      Logger.log('Saving FORMASI scores...');
      
      submission.scores.forEach(function(scoreItem) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          scoreItem.code,
          scoreItem.score_formasi1 || 0,
          scoreItem.score_formasi2 || 0,
          scoreItem.score_formasi3 || 0,
          scoreItem.score_formasi4 || 0,
          '', // penalty_code
          0,  // penalty_value
          submission.keyed_by,
          false // edited - NEW!
        ];
        sheet.appendRow(row);
        
        totalScore += parseInt(scoreItem.score_formasi1 || 0);
        totalScore += parseInt(scoreItem.score_formasi2 || 0);
        totalScore += parseInt(scoreItem.score_formasi3 || 0);
        totalScore += parseInt(scoreItem.score_formasi4 || 0);
      });
      
      // Save penalties
      submission.penalties.forEach(function(penalty) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          '', // item_code
          0, 0, 0, 0, // formasi scores (empty for penalty rows)
          penalty.code,
          penalty.value,  // penalty_value — ADDED (was missing before)
          submission.keyed_by,
          false // edited
        ];
        sheet.appendRow(row);
        totalPenalty += parseInt(penalty.value || 0);
      });
      
    } else if (formType === 'kawad_platun') {
      // KAWAD PLATUN: Single score + separate penalty columns
      Logger.log('Saving KAWAD PLATUN scores...');
      
      submission.scores.forEach(function(scoreItem) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          scoreItem.code,
          scoreItem.score || 0,
          '', // penalty_code
          0,  // penalty_value
          submission.keyed_by,
          false // edited
        ];
        sheet.appendRow(row);
        totalScore += parseInt(scoreItem.score || 0);
      });
      
      // Save penalties
      submission.penalties.forEach(function(penalty) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          '', // item_code (empty for penalty rows)
          0,  // score
          penalty.code,
          penalty.value,
          submission.keyed_by,
          false
        ];
        sheet.appendRow(row);
        totalPenalty += parseInt(penalty.value || 0);
      });
      
    } else {
      // PAKAIAN KP, PAKAIAN PLATUN: NO penalties, NO edited column
      Logger.log('Saving PAKAIAN scores...');
      
      submission.scores.forEach(function(scoreItem) {
        const row = [
          entryId,
          timestamp,
          submission.team_id,
          submission.judge_id,
          scoreItem.code,
          scoreItem.score || 0,
          '', // penalty (empty - NOTA instead)
          submission.keyed_by
          // NO edited column!
        ];
        sheet.appendRow(row);
        totalScore += parseInt(scoreItem.score || 0);
      });
    }
    
    const finalScore = totalScore - totalPenalty;
    
    Logger.log('=== SAVE COMPLETE ===');
    Logger.log('Total score: ' + totalScore);
    Logger.log('Total penalty: ' + totalPenalty);
    Logger.log('Final score: ' + finalScore);

    // Update calculated results
    updateCalculatedResults(submission.team_id, entryId);
    
    // Log to audit
    logToAudit(
      submission.keyed_by,
      'SUBMIT_SCORES',
      sheetName,
      entryId,
      '',
      'Score: ' + totalScore + ', Penalty: ' + totalPenalty + ', Final: ' + finalScore
    );
    
    lock.releaseLock();
    
    return { 
      success: true, 
      message: 'Markah berjaya disimpan',
      entry_id: entryId,
      total_score: finalScore,
      raw_score: totalScore,
      penalties: totalPenalty
    };
    
  } catch (error) {
    Logger.log('=== ERROR IN SAVESCORES ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    
    lock.releaseLock();
    
    return { 
      success: false, 
      message: 'Error: ' + error.toString() 
    };
  }
}

// ============================================
// UPDATE CALCULATED RESULTS
// ============================================
function updateCalculatedResults(teamId, entryId) {
  try {
    Logger.log('=== UPDATE CALCULATED RESULTS START ===');
    Logger.log('Team ID: ' + teamId);
    Logger.log('Entry ID: ' + entryId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const resultsSheet = ss.getSheetByName('CALCULATED_RESULTS');
    
    if (!resultsSheet) {
      Logger.log('CALCULATED_RESULTS sheet not found');
      return false;
    }
    
    // Get team info
    const teamsSheet = ss.getSheetByName('TEAMS');
    const teamsData = teamsSheet.getDataRange().getValues();
    let teamInfo = null;
    
    for (let i = 1; i < teamsData.length; i++) {
      if (teamsData[i][1] === teamId) {  // Column B = team_code (L1, P1)
        teamInfo = {
          team_id: teamsData[i][0],
          team_code: teamsData[i][1],
          kategori: teamsData[i][3],
          gender: teamsData[i][4]
        };
        break;
      }
    }
    
    if (!teamInfo) {
      Logger.log('Team not found: ' + teamId);
      return false;
    }
    
    // Calculate scores for each form type
    const scores = {
      pakaian_kp: calculateFormTotal(teamId, 'SCORES_PAKAIAN_KP'),
      pakaian_platun: calculateFormTotal(teamId, 'SCORES_PAKAIAN_PLATUN'),
      kawad_kp: calculateFormTotal(teamId, 'SCORES_KAWAD_KP'),
      kawad_platun: calculateFormTotal(teamId, 'SCORES_KAWAD_PLATUN'),
      formasi: calculateFormTotal(teamId, 'SCORES_FORMASI')
    };
    
    // Get weightage from CONFIG_SETTINGS
    const weightage = getWeightageConfig();
    
    // Maximum scores for each form (from borang)
    const MAX_PAKAIAN_KP = 32;
    const MAX_PAKAIAN_PLATUN = 288;
    const MAX_PAKAIAN_TOTAL = MAX_PAKAIAN_KP + MAX_PAKAIAN_PLATUN; // 320
    const MAX_KAWAD_KP = 270;
    const MAX_KAWAD_PLATUN = 190;
    const MAX_FORMASI = 90;
    
    // Calculate raw totals (these go to columns E-I)
    const pakaian_kp_total = scores.pakaian_kp;
    const pakaian_platun_total = scores.pakaian_platun;
    const kawad_kp_total = scores.kawad_kp;
    const kawad_platun_total = scores.kawad_platun;
    const formasi_total = scores.formasi;
    
    // Calculate weighted percentages (these go to columns J-M)
    const pakaian_combined = pakaian_kp_total + pakaian_platun_total;
    const pakaian_platun_total_10pct = (pakaian_combined / MAX_PAKAIAN_TOTAL) * weightage.pakaian;
    const kawad_kp_total_10pct = (kawad_kp_total / MAX_KAWAD_KP) * weightage.kawad_kp;
    const kawad_platun_total_50pct = (kawad_platun_total / MAX_KAWAD_PLATUN) * weightage.kawad_platun;
    const formasi_total_30pct = (formasi_total / MAX_FORMASI) * weightage.formasi;
    
    // Calculate grand total (sum of weighted percentages)
    const grandTotal = pakaian_platun_total_10pct + kawad_kp_total_10pct + 
                       kawad_platun_total_50pct + formasi_total_30pct;
    
    Logger.log('=== SCORE CALCULATION ===');
    Logger.log('Raw scores:', {
      pakaian_kp: pakaian_kp_total,
      pakaian_platun: pakaian_platun_total,
      kawad_kp: kawad_kp_total,
      kawad_platun: kawad_platun_total,
      formasi: formasi_total
    });
    Logger.log('Weighted scores:', {
      pakaian_10pct: pakaian_platun_total_10pct.toFixed(2),
      kawad_kp_10pct: kawad_kp_total_10pct.toFixed(2),
      kawad_platun_50pct: kawad_platun_total_50pct.toFixed(2),
      formasi_30pct: formasi_total_30pct.toFixed(2)
    });
    Logger.log('Grand total:', grandTotal.toFixed(2));
    
    // Check if team already exists in results
    const resultsData = resultsSheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < resultsData.length; i++) {
    // NOTE: resultsData Column A contains team_id (UUID), not team_code
      if (resultsData[i][0] === teamInfo.team_id) {  // Compare UUID, not team_code
        rowIndex = i + 1; // +1 because row index starts at 1
        break;
      }
    }
    
    const timestamp = new Date();
    
    if (rowIndex > 0) {
      // Update existing row
      Logger.log('Updating existing row: ' + rowIndex);
      
      resultsSheet.getRange(rowIndex, 1, 1, 16).setValues([[
        teamInfo.team_id,                    // A
        teamInfo.team_code,                  // B
        teamInfo.kategori,                   // C
        teamInfo.gender,                     // D
        pakaian_kp_total,                    // E - Raw score
        pakaian_platun_total,                // F - Raw score
        kawad_kp_total,                      // G - Raw score
        kawad_platun_total,                  // H - Raw score
        formasi_total,                       // I - Raw score
        pakaian_platun_total_10pct,          // J - Weighted 10%
        kawad_kp_total_10pct,                // K - Weighted 10%
        kawad_platun_total_50pct,            // L - Weighted 50%
        formasi_total_30pct,                 // M - Weighted 30%
        grandTotal,                          // N - Grand total
        '',                                  // O - Rank (updated by updateRankings)
        timestamp                            // P - Last updated
      ]]);
    } else {
      // Add new row
      Logger.log('Adding new row');
      
      resultsSheet.appendRow([
        teamInfo.team_id,                    // A
        teamInfo.team_code,                  // B
        teamInfo.kategori,                   // C
        teamInfo.gender,                     // D
        pakaian_kp_total,                    // E - Raw score
        pakaian_platun_total,                // F - Raw score
        kawad_kp_total,                      // G - Raw score
        kawad_platun_total,                  // H - Raw score
        formasi_total,                       // I - Raw score
        pakaian_platun_total_10pct,          // J - Weighted 10%
        kawad_kp_total_10pct,                // K - Weighted 10%
        kawad_platun_total_50pct,            // L - Weighted 50%
        formasi_total_30pct,                 // M - Weighted 30%
        grandTotal,                          // N - Grand total
        '',                                  // O - Rank
        timestamp                            // P - Last updated
      ]);
    }
    
    Logger.log('CALCULATED_RESULTS updated successfully');
    
    // Update rankings for all teams
    updateRankings();
    
    return true;
    
  } catch (error) {
    Logger.log('Error updating calculated results: ' + error.toString());
    return false;
  }
}

// ============================================
// UPDATE RANKINGS
// ============================================
function updateRankings() {
  try {
    Logger.log('=== UPDATE RANKINGS START ===');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const resultsSheet = ss.getSheetByName('CALCULATED_RESULTS');
    
    if (!resultsSheet) {
      Logger.log('CALCULATED_RESULTS sheet not found');
      return false;
    }
    
    const data = resultsSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('No teams to rank');
      return true;
    }
    
    // Group teams by kategori and gender
    const groups = {};
    
    for (let i = 1; i < data.length; i++) {
      const teamId = data[i][0];
      const kategori = data[i][2];
      const gender = data[i][3];
      const grandTotal = data[i][13]; // Column N (14th column, index 13)
      
      if (!teamId) continue;
      
      const groupKey = kategori + '_' + gender;
      
      if (!groups[groupKey]) {
        groups[groupKey] = [];
      }
      
      groups[groupKey].push({
        row: i + 1, // +1 for 1-based row index
        teamId: teamId,
        grandTotal: grandTotal
      });
    }
    
    // Rank each group
    for (const groupKey in groups) {
      const teams = groups[groupKey];
      
      // Sort by grandTotal descending
      teams.sort(function(a, b) {
        return b.grandTotal - a.grandTotal;
      });
      
      // Assign ranks
      for (let j = 0; j < teams.length; j++) {
        const rank = j + 1;
        const row = teams[j].row;
        
        // Update rank in column O (15th column)
        resultsSheet.getRange(row, 15).setValue(rank);
      }
    }
    
    Logger.log('Rankings updated successfully');
    return true;
    
  } catch (error) {
    Logger.log('Error updating rankings: ' + error.toString());
    return false;
  }
}

// Helper function to calculate total for a specific form
function calculateFormTotal(teamId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetName);
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Group scores by entry_id (each judge creates unique entry_id)
    const entriesMap = {};
    
    for (let i = 1; i < data.length; i++) {
      const rowTeamId = data[i][2]; // Column C (team_id)
      
      if (rowTeamId === teamId) {
        const entryId = data[i][0]; // Column A (entry_id)
        
        if (!entriesMap[entryId]) {
          entriesMap[entryId] = 0;
        }
        
        // Calculate score based on sheet structure
        if (sheetName === 'SCORES_KAWAD_KP') {
          // Columns F & G (score_pergerakan, score_bahasa)
          entriesMap[entryId] += (parseInt(data[i][5]) || 0);
          entriesMap[entryId] += (parseInt(data[i][6]) || 0);
          // Penalty in column I
          entriesMap[entryId] -= (parseInt(data[i][8]) || 0);
          
        } else if (sheetName === 'SCORES_KAWAD_PLATUN') {
          // KAWAD_PLATUN: Single score (F) + penalty (H)
          entriesMap[entryId] += (parseInt(data[i][5]) || 0);
          // Subtract penalty from column H
          entriesMap[entryId] -= (parseInt(data[i][7]) || 0);
          
        } else if (sheetName === 'SCORES_FORMASI') {
          // Columns F-I (formasi 1-4)
          entriesMap[entryId] += (parseInt(data[i][5]) || 0);
          entriesMap[entryId] += (parseInt(data[i][6]) || 0);
          entriesMap[entryId] += (parseInt(data[i][7]) || 0);
          entriesMap[entryId] += (parseInt(data[i][8]) || 0);
          // Penalty value in column K [10] — subtract
          entriesMap[entryId] -= (parseInt(data[i][10]) || 0);
          
        } else {
          // PAKAIAN: Single score column F
          entriesMap[entryId] += (parseInt(data[i][5]) || 0);
        }
      }
    }
    
    // Calculate average across all entries (judges)
    const entryScores = Object.values(entriesMap);
    
    if (entryScores.length === 0) {
      return 0;
    }
    
    const totalScore = entryScores.reduce((sum, score) => sum + score, 0);
    const averageScore = totalScore / entryScores.length;
    
    Logger.log(`${sheetName} - Team ${teamId}: ${entryScores.length} judge(s), Average: ${averageScore.toFixed(2)}`);
    
    return averageScore;
    
  } catch (error) {
    Logger.log('Error calculating form total: ' + error.toString());
    return 0;
  }
}

// ============================================
// LOG TO AUDIT LOG
// ============================================
function logToAudit(userId, action, tableName, recordId, oldValue, newValue, notes) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const auditSheet = ss.getSheetByName('AUDIT_LOG');
    
    if (!auditSheet) {
      Logger.log('AUDIT_LOG sheet not found');
      return false;
    }
    
    const timestamp = new Date();
    const logId = 'LOG_' + timestamp.getTime();
    
    auditSheet.appendRow([
      logId,
      timestamp,
      userId,
      action,
      tableName,
      recordId,
      oldValue || '',
      newValue || '',
      notes || ''       // optional notes — default '' if not passed
    ]);
    
    Logger.log('Audit log created: ' + logId);
    return true;
    
  } catch (error) {
    Logger.log('Error logging to audit: ' + error.toString());
    return false;
  }
}

// ============================================
// GET WEIGHTAGE FROM CONFIG
// ============================================
function getWeightageConfig() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName('CONFIG_SETTINGS');
    
    if (!configSheet) {
      Logger.log('CONFIG_SETTINGS not found, using defaults');
      return {
        pakaian: 10,
        kawad_kp: 10,
        kawad_platun: 50,
        formasi: 30
      };
    }
    
    const data = configSheet.getDataRange().getValues();
    const config = {
      pakaian: 10,      // default
      kawad_kp: 10,     // default
      kawad_platun: 50, // default
      formasi: 30       // default
    };
    
    // Read from sheet
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      const value = parseFloat(data[i][1]) || 0;
      
      if (key === 'pemberat_elit_pakaian_kp') config.pakaian = value;
      if (key === 'pemberat_elit_kawad_kp') config.kawad_kp = value;
      if (key === 'pemberat_elit_kawad_platun') config.kawad_platun = value;
      if (key === 'pemberat_elit_formasi') config.formasi = value;
    }
    
    Logger.log('Weightage config loaded:', config);
    return config;
    
  } catch (error) {
    Logger.log('Error loading weightage: ' + error.toString());
    return {
      pakaian: 10,
      kawad_kp: 10,
      kawad_platun: 50,
      formasi: 30
    };
  }
}

// ============================================
// GET LEADERBOARD DATA (CLEAN VERSION)
// ============================================
// Pastikan variable ini wujud di bahagian paling atas fail anda
// const SPREADSHEET_ID = 'MASUKKAN_ID_SHEET_ANDA_DI_SINI'; 

function getLeaderboardData() {
  try {
    // Pastikan ID sheet betul. Kalau variable global tak detect, masukkan manual ID di sini
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName('CALCULATED_RESULTS');
    
    if (!sheet) {
      Logger.log('CALCULATED_RESULTS sheet not found');
      return []; // Return array kosong, JANGAN return null
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return []; 
    
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const leaderboard = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      leaderboard.push({
        team_id: String(row[0]),     // Paksa jadi String
        team_code: String(row[1]),   
        kategori: String(row[2]),    
        gender: String(row[3]),      
        pakaian_kp_total: Number(row[4]) || 0,
        pakaian_platun_total: Number(row[5]) || 0,
        kawad_kp_total: Number(row[6]) || 0,
        kawad_platun_total: Number(row[7]) || 0,
        formasi_total: Number(row[8]) || 0,
        pakaian_weighted: Number(row[9]) || 0,
        kawad_kp_weighted: Number(row[10]) || 0,
        kawad_platun_weighted: Number(row[11]) || 0,
        formasi_weighted: Number(row[12]) || 0,
        grand_total: Number(row[13]) || 0,
        rank: Number(row[14]) || 0,
        // PEMBAIKAN UTAMA: Tukar tarikh ke String. Jika tidak, data boleh jadi NULL.
        last_updated: String(row[15]) 
      });
    }
    
    Logger.log('Leaderboard data fetched: ' + leaderboard.length);
    return leaderboard; // Pastikan return array
    
  } catch (e) {
    Logger.log('Error getting leaderboard: ' + e.toString());
    // Jangan throw error, return array kosong supaya UI tak stuck
    return []; 
  }
}

// ============================================
// RECALCULATE ALL TEAMS (triggered by Admin)
// ============================================
function recalculateAllTeams() {
  const startTime = new Date();

  try {
    Logger.log('=== RECALCULATE ALL TEAMS START ===');

    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teamsSheet  = ss.getSheetByName('TEAMS');
    const resultsSheet = ss.getSheetByName('CALCULATED_RESULTS');

    if (!teamsSheet || !resultsSheet) {
      return { success: false, message: 'Sheet TEAMS atau CALCULATED_RESULTS tidak ditemui.' };
    }

    // 1. Read all teams
    const teamsData = teamsSheet.getDataRange().getValues();
    const teamIds   = [];
    for (var i = 1; i < teamsData.length; i++) {
      if (teamsData[i][0]) teamIds.push(teamsData[i][0]);   // column A = team_id
    }

    if (teamIds.length === 0) {
      return { success: false, message: 'Tiada pasukan ditemui dalam sheet TEAMS.' };
    }

    Logger.log('Teams to recalculate: ' + teamIds.length);

    // 2. Get weightage once (reuse for all teams)
    const weightage = getWeightageConfig();

    // 3. Max scores (from borang)
    const MAX_PAKAIAN_KP      = 32;
    const MAX_PAKAIAN_PLATUN  = 288;
    const MAX_PAKAIAN_TOTAL   = MAX_PAKAIAN_KP + MAX_PAKAIAN_PLATUN; // 320
    const MAX_KAWAD_KP        = 270;
    const MAX_KAWAD_PLATUN    = 190;
    const MAX_FORMASI         = 90;

    // 4. Read current CALCULATED_RESULTS to build a map: team_id → row number
    const resultsData  = resultsSheet.getDataRange().getValues();
    const existingRows = {};                          // { team_id : rowIndex (1-based) }
    for (var r = 1; r < resultsData.length; r++) {
      if (resultsData[r][0]) existingRows[resultsData[r][0]] = r + 1;
    }

    // 5. Loop every team — recalculate & write
    const timestamp = new Date();
    var processed   = 0;

    for (var t = 0; t < teamIds.length; t++) {
      var teamId = teamIds[t];

      // --- get team info from TEAMS sheet ---
      var teamInfo = null;
      for (var ti = 1; ti < teamsData.length; ti++) {
        if (teamsData[ti][0] === teamId) {
          teamInfo = {
            team_id   : teamsData[ti][0],
            team_code : teamsData[ti][1],
            kategori  : teamsData[ti][3],
            gender    : teamsData[ti][4]
          };
          break;
        }
      }
      if (!teamInfo) { Logger.log('Team not found in TEAMS: ' + teamId); continue; }

      // --- calculate raw totals from SCORES sheets ---
      // NOTE: SCORES sheets now use team_code (L1, P1) not team_id (UUID)
      var pakaian_kp_total     = calculateFormTotal(teamInfo.team_code, 'SCORES_PAKAIAN_KP');
      var pakaian_platun_total = calculateFormTotal(teamInfo.team_code, 'SCORES_PAKAIAN_PLATUN');
      var kawad_kp_total       = calculateFormTotal(teamInfo.team_code, 'SCORES_KAWAD_KP');
      var kawad_platun_total   = calculateFormTotal(teamInfo.team_code, 'SCORES_KAWAD_PLATUN');
      var formasi_total        = calculateFormTotal(teamInfo.team_code, 'SCORES_FORMASI');

      // --- SKIP team if ALL scores are zero (tiada scores dalam mana-mana tab) ---
      if (pakaian_kp_total === 0 && pakaian_platun_total === 0 &&
          kawad_kp_total === 0 && kawad_platun_total === 0 && formasi_total === 0) {
        Logger.log('SKIP ' + teamId + ' — no scores found in any SCORES sheet');
        continue;
      }

      // --- weighted percentages ---
      var pakaian_combined          = pakaian_kp_total + pakaian_platun_total;
      var pakaian_platun_total_10pct  = (pakaian_combined     / MAX_PAKAIAN_TOTAL)  * weightage.pakaian;
      var kawad_kp_total_10pct       = (kawad_kp_total        / MAX_KAWAD_KP)       * weightage.kawad_kp;
      var kawad_platun_total_50pct   = (kawad_platun_total    / MAX_KAWAD_PLATUN)   * weightage.kawad_platun;
      var formasi_total_30pct        = (formasi_total         / MAX_FORMASI)        * weightage.formasi;

      // --- grand total ---
      var grandTotal = pakaian_platun_total_10pct + kawad_kp_total_10pct +
                       kawad_platun_total_50pct  + formasi_total_30pct;

      // --- row payload (16 columns, A–P) ---
      var rowValues = [
        teamInfo.team_id,                  // A
        teamInfo.team_code,                // B
        teamInfo.kategori,                 // C
        teamInfo.gender,                   // D
        pakaian_kp_total,                  // E
        pakaian_platun_total,              // F
        kawad_kp_total,                    // G
        kawad_platun_total,                // H
        formasi_total,                     // I
        pakaian_platun_total_10pct,        // J
        kawad_kp_total_10pct,              // K
        kawad_platun_total_50pct,          // L
        formasi_total_30pct,               // M
        grandTotal,                        // N
        '',                                // O – rank (updateRankings fills this)
        timestamp                          // P
      ];

      // --- read OLD grand_total from resultsData (before overwrite) ---
      var oldGrandTotal = 0;
      if (existingRows[teamId]) {
        oldGrandTotal = Number(resultsData[existingRows[teamId] - 1][13]) || 0; // col N = index 13
      }

      // --- write: update existing row OR append new row ---
      if (existingRows[teamId]) {
        resultsSheet.getRange(existingRows[teamId], 1, 1, 16).setValues([rowValues]);
        Logger.log('Updated row ' + existingRows[teamId] + ' for ' + teamId);
      } else {
        resultsSheet.appendRow(rowValues);
        Logger.log('Appended new row for ' + teamId);
      }

      // --- audit log: one row per team ---
      logToAudit(
        'ADMIN',
        'RECALCULATE_ALL',
        'SCORES_*',
        teamInfo.team_id,
        'Score: ' + oldGrandTotal.toFixed(2) + ', Final: ' + oldGrandTotal.toFixed(2),
        'Score: ' + grandTotal.toFixed(2)    + ', Final: ' + grandTotal.toFixed(2),
        'Recalculate by ADMIN'
      );

      processed++;
    }

    // 6. Recalculate rankings for ALL teams at once
    updateRankings();

    // 7. Summary audit log (1 row after all teams done)
    var duration = ((new Date() - startTime) / 1000).toFixed(2);

    logToAudit('ADMIN', 'RECALCULATE_ALL', 'CALCULATED_RESULTS', 'SUMMARY',
               '',
               '',
               'Recalculated ' + processed + ' pasukan | Masa: ' + duration + 's');

    Logger.log('=== RECALCULATE ALL DONE. Teams: ' + processed + ' | Duration: ' + duration + 's ===');

    return {
      success        : true,
      teamsProcessed : processed,
      duration       : duration
    };

  } catch (error) {
    Logger.log('Error in recalculateAllTeams: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ============================================
// DAFTAR PASUKAN & USER MANAGEMENT (CRUD)
// ============================================

function addTeam(teamData) {
  try {
    Logger.log('=== ADD TEAM START ===');
    Logger.log('Team data: ' + JSON.stringify(teamData));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    if (!teamsSheet) {
      return { success: false, message: 'TEAMS sheet not found' };
    }
    
    // Validate required fields
    if (!teamData.team_code || !teamData.school_name || !teamData.kategori || !teamData.gender) {
      return { success: false, message: 'Sila isi semua medan wajib (Kod Pasukan, Nama Sekolah, Kategori, Jantina)' };
    }
    
    // Validate team_code format (alphanumeric only, max 10 chars)
    const codeRegex = /^[A-Za-z0-9]+$/;
    if (!codeRegex.test(teamData.team_code) || teamData.team_code.length > 10) {
      return { success: false, message: 'Kod Pasukan hanya boleh huruf/nombor, max 10 aksara' };
    }
    
    // Check team_code uniqueness
    const data = teamsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === teamData.team_code && data[i][7] === true) { // col B = team_code, H = is_active
        return { success: false, message: 'Kod Pasukan "' + teamData.team_code + '" sudah wujud' };
      }
    }
    
    // Generate team_id
    const timestamp = new Date().getTime();
    const teamId = 'TEAM_' + timestamp;
    
    // Append row
    teamsSheet.appendRow([
      teamId,                           // A
      teamData.team_code,               // B
      teamData.school_name,             // C
      teamData.kategori,                // D
      teamData.gender,                  // E
      teamData.school_level || '',      // F
      teamData.ketua_platun_name || '', // G
      true                              // H - is_active
    ]);
    
    // Audit log
    logToAudit(
      'ADMIN',
      'ADD_TEAM',
      'TEAMS',
      teamId,
      '',
      'Team: ' + teamData.team_code + ' | School: ' + teamData.school_name
    );
    
    Logger.log('Team added: ' + teamId);
    return { success: true, message: 'Pasukan berjaya ditambah', team_id: teamId };
    
  } catch (error) {
    Logger.log('Error adding team: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateTeam(teamId, teamData) {
  try {
    Logger.log('=== UPDATE TEAM START ===');
    Logger.log('Team ID: ' + teamId);
    Logger.log('Team data: ' + JSON.stringify(teamData));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    if (!teamsSheet) {
      return { success: false, message: 'TEAMS sheet not found' };
    }
    
    // Validate required fields
    if (!teamData.team_code || !teamData.school_name || !teamData.kategori || !teamData.gender) {
      return { success: false, message: 'Sila isi semua medan wajib' };
    }
    
    // Validate team_code format
    const codeRegex = /^[A-Za-z0-9]+$/;
    if (!codeRegex.test(teamData.team_code) || teamData.team_code.length > 10) {
      return { success: false, message: 'Kod Pasukan hanya boleh huruf/nombor, max 10 aksara' };
    }
    
    // Find row
    const data = teamsSheet.getDataRange().getValues();
    let rowIndex = -1;
    let oldData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === teamId) {
        rowIndex = i + 1; // 1-based
        oldData = data[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Pasukan tidak dijumpai' };
    }
    
    // Check team_code uniqueness (exclude current team)
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === teamData.team_code && data[i][0] !== teamId && data[i][7] === true) {
        return { success: false, message: 'Kod Pasukan "' + teamData.team_code + '" sudah wujud' };
      }
    }
    
    // Update row (keep team_id and is_active)
    teamsSheet.getRange(rowIndex, 1, 1, 8).setValues([[
      teamId,                           // A - unchanged
      teamData.team_code,               // B
      teamData.school_name,             // C
      teamData.kategori,                // D
      teamData.gender,                  // E
      teamData.school_level || '',      // F
      teamData.ketua_platun_name || '', // G
      oldData[7]                        // H - is_active unchanged
    ]]);
    
    // Audit log
    logToAudit(
      'ADMIN',
      'UPDATE_TEAM',
      'TEAMS',
      teamId,
      'Code: ' + oldData[1] + ' | School: ' + oldData[2],
      'Code: ' + teamData.team_code + ' | School: ' + teamData.school_name
    );
    
    Logger.log('Team updated: ' + teamId);
    return { success: true, message: 'Pasukan berjaya dikemaskini' };
    
  } catch (error) {
    Logger.log('Error updating team: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteTeam(teamId) {
  try {
    Logger.log('=== DELETE TEAM START ===');
    Logger.log('Team ID: ' + teamId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    if (!teamsSheet) {
      return { success: false, message: 'TEAMS sheet not found' };
    }
    
    // Find row
    const data = teamsSheet.getDataRange().getValues();
    let rowIndex = -1;
    let teamCode = '';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === teamId) {
        rowIndex = i + 1; // 1-based
        teamCode = data[i][1];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Pasukan tidak dijumpai' };
    }
    
    // Soft delete: set is_active = FALSE
    teamsSheet.getRange(rowIndex, 8).setValue(false); // col H
    
    // Audit log
    logToAudit(
      'ADMIN',
      'DELETE_TEAM',
      'TEAMS',
      teamId,
      'Active: true',
      'Active: false, Team: ' + teamCode
    );
    
    Logger.log('Team deleted (soft): ' + teamId);
    return { success: true, message: 'Pasukan berjaya dipadam' };
    
  } catch (error) {
    Logger.log('Error deleting team: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// NEW CRUD FUNCTIONS - TEAMS MANAGEMENT
// ============================================

/**
 * Add new team to TEAMS sheet
 * @param {Object} teamData - {team_code, school_name, kategori, gender, school_level, ketua_platun_name}
 * @return {Object} {success, message, team_id}
 */
function addTeam(teamData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    // Validation
    if (!teamData.team_code || !teamData.school_name || !teamData.kategori || !teamData.gender || !teamData.school_level) {
      return {success: false, message: 'Kod pasukan, nama sekolah, kategori, jantina, dan tahap sekolah adalah wajib!'};
    }
    
    // Validate team_code format (alphanumeric only, max 10 chars)
    if (!/^[A-Za-z0-9]+$/.test(teamData.team_code) || teamData.team_code.length > 10) {
      return {success: false, message: 'Kod pasukan mesti alphanumeric sahaja, maksimum 10 aksara!'};
    }
    
    // Validate kategori
    const validKategori = ['ELIT', 'AMATUR'];
    if (!validKategori.includes(teamData.kategori)) {
      return {success: false, message: 'Kategori tidak sah!'};
    }
    
    // Validate gender
    const validGender = ['LELAKI', 'PEREMPUAN'];
    if (!validGender.includes(teamData.gender)) {
      return {success: false, message: 'Jantina tidak sah!'};
    }
    
    // Validate school_level
    const validSchoolLevel = ['RENDAH', 'MENENGAH'];
    if (!validSchoolLevel.includes(teamData.school_level)) {
      return {success: false, message: 'Tahap sekolah tidak sah!'};
    }
    
    // ✅ FIX: Check uniqueness safely even if sheet only has header
    const lastRow = teamsSheet.getLastRow();
    if (lastRow > 1) {
      const existingCodes = teamsSheet
        .getRange(2, 2, lastRow - 1, 1)
        .getValues()
        .flat();
        
      if (existingCodes.includes(teamData.team_code)) {
        return {success: false, message: 'Kod pasukan "' + teamData.team_code + '" sudah wujud!'};
      }
    }
    
    // Generate team_id
    const timestamp = new Date().getTime();
    const teamId = 'TEAM_' + timestamp;
    
    // Prepare row data (8 columns: A-H)
    const newRow = [
      teamId,                                    // A: team_id
      teamData.team_code,                        // B: team_code
      teamData.school_name,                      // C: school_name
      teamData.kategori,                         // D: kategori
      teamData.gender,                           // E: gender
      teamData.school_level,                     // F: school_level
      teamData.ketua_platun_name || '',          // G: ketua_platun_name
      true                                       // H: is_active
    ];
    
    // Append row
    teamsSheet.appendRow(newRow);
    
    // Audit log
    logToAudit(Session.getActiveUser().getEmail(), 'ADD_TEAM', 'TEAMS', teamId, '', teamData.team_code + ' - ' + teamData.school_name, 'Added new team');
    
    return {
      success: true, 
      message: 'Pasukan "' + teamData.team_code + '" berjaya ditambah!',
      team_id: teamId
    };
    
  } catch (e) {
    Logger.log('Error in addTeam: ' + e.toString());
    return {success: false, message: 'Ralat: ' + e.toString()};
  }
}


/**
 * Update existing team in TEAMS sheet
 * @param {String} teamId - team_id to update
 * @param {Object} teamData - {team_code, school_name, kategori, gender, school_level, ketua_platun_name}
 * @return {Object} {success, message}
 */
function updateTeam(teamId, teamData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    // Validation
    if (!teamData.team_code || !teamData.school_name || !teamData.kategori || !teamData.gender || !teamData.school_level) {
      return {success: false, message: 'Kod pasukan, nama sekolah, kategori, jantina, dan tahap sekolah adalah wajib!'};
    }
    
    // Validate team_code format
    if (!/^[A-Za-z0-9]+$/.test(teamData.team_code) || teamData.team_code.length > 10) {
      return {success: false, message: 'Kod pasukan mesti alphanumeric sahaja, maksimum 10 aksara!'};
    }
    
    // Validate kategori
    const validKategori = ['ELIT', 'AMATUR'];
    if (!validKategori.includes(teamData.kategori)) {
      return {success: false, message: 'Kategori tidak sah!'};
    }
    
    // Validate gender
    const validGender = ['LELAKI', 'PEREMPUAN'];
    if (!validGender.includes(teamData.gender)) {
      return {success: false, message: 'Jantina tidak sah!'};
    }
    
    // Validate school_level
    const validSchoolLevel = ['RENDAH', 'MENENGAH'];
    if (!validSchoolLevel.includes(teamData.school_level)) {
      return {success: false, message: 'Tahap sekolah tidak sah!'};
    }
    
    // Find row by team_id (column A)
    const teamIds = teamsSheet.getRange(2, 1, teamsSheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = teamIds.indexOf(teamId);
    
    if (rowIndex === -1) {
      return {success: false, message: 'Pasukan tidak dijumpai!'};
    }
    
    const actualRow = rowIndex + 2; // +2 because array is 0-indexed and sheet has header row
    
    // Check if team_code is being changed to an existing code (exclude current row)
    const existingCodes = teamsSheet.getRange(2, 2, teamsSheet.getLastRow() - 1, 1).getValues().flat();
    const currentCode = teamsSheet.getRange(actualRow, 2).getValue();
    
    if (teamData.team_code !== currentCode && existingCodes.includes(teamData.team_code)) {
      return {success: false, message: 'Kod pasukan "' + teamData.team_code + '" sudah wujud!'};
    }
    
    // Update row (columns B-G, preserve A and H)
    teamsSheet.getRange(actualRow, 2).setValue(teamData.team_code);              // B: team_code
    teamsSheet.getRange(actualRow, 3).setValue(teamData.school_name);            // C: school_name
    teamsSheet.getRange(actualRow, 4).setValue(teamData.kategori);               // D: kategori
    teamsSheet.getRange(actualRow, 5).setValue(teamData.gender);                 // E: gender
    teamsSheet.getRange(actualRow, 6).setValue(teamData.school_level);     // F: school_level (required)
    teamsSheet.getRange(actualRow, 7).setValue(teamData.ketua_platun_name || '');// G: ketua_platun_name
    
    // Audit log
    logToAudit(Session.getActiveUser().getEmail(), 'UPDATE_TEAM', 'TEAMS', teamId, '', teamData.team_code + ' - ' + teamData.school_name, 'Updated team');
    
    return {
      success: true,
      message: 'Pasukan "' + teamData.team_code + '" berjaya dikemaskini!'
    };
    
  } catch (e) {
    Logger.log('Error in updateTeam: ' + e.toString());
    return {success: false, message: 'Ralat: ' + e.toString()};
  }
}


/**
 * Soft delete team (set is_active = FALSE)
 * @param {String} teamId - team_id to delete
 * @return {Object} {success, message}
 */
function deleteTeam(teamId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    // Find row by team_id (column A)
    const teamIds = teamsSheet.getRange(2, 1, teamsSheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = teamIds.indexOf(teamId);
    
    if (rowIndex === -1) {
      return {success: false, message: 'Pasukan tidak dijumpai!'};
    }
    
    const actualRow = rowIndex + 2;
    const teamCode = teamsSheet.getRange(actualRow, 2).getValue();
    
    // Soft delete: set is_active = FALSE (column H)
    teamsSheet.getRange(actualRow, 8).setValue(false);
    
    // Audit log
    logToAudit(Session.getActiveUser().getEmail(), 'DELETE_TEAM', 'TEAMS', teamId, 'is_active=TRUE', 'is_active=FALSE', 'Soft deleted team: ' + teamCode);
    
    return {
      success: true,
      message: 'Pasukan "' + teamCode + '" berjaya dipadam!'
    };
    
  } catch (e) {
    Logger.log('Error in deleteTeam: ' + e.toString());
    return {success: false, message: 'Ralat: ' + e.toString()};
  }
}


// ============================================
// NEW CRUD FUNCTIONS - USER MANAGEMENT
// ============================================

/**
 * Add new user to USERS sheet
 * @param {Object} userData - {username, password, role, full_name}
 * @return {Object} {success, message, user_id}
 */
function addUser(userData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('USERS');
    
    // Check if USERS sheet exists
    if (!usersSheet) {
      Logger.log('ERROR: USERS sheet not found!');
      return {success: false, message: 'USERS sheet tidak dijumpai! Sila hubungi admin.'};
    }
    
    // Validation
    if (!userData.username || !userData.password || !userData.role || !userData.full_name) {
      return {success: false, message: 'Semua medan adalah wajib!'};
    }
    
    // Validate username format (alphanumeric + underscore, 3-20 chars)
    if (!/^[A-Za-z0-9_]{3,20}$/.test(userData.username)) {
      return {success: false, message: 'Username mesti alphanumeric/underscore sahaja, 3-20 aksara!'};
    }
    
    // Validate password length
    if (userData.password.length < 4) {
      return {success: false, message: 'Password mesti sekurang-kurangnya 4 aksara!'};
    }
    
    // Validate role
    const validRoles = ['ADMIN', 'STATISTIK'];
    if (!validRoles.includes(userData.role)) {
      return {success: false, message: 'Role tidak sah! Pilih ADMIN atau STATISTIK.'};
    }
    
    // Check uniqueness of username (case-insensitive, column B)
    const lastRow = usersSheet.getLastRow();
    Logger.log('addUser - Checking username uniqueness. LastRow: ' + lastRow);
    
    if (lastRow > 1) {
      const existingUsernames = usersSheet.getRange(2, 2, lastRow - 1, 1)
        .getValues()
        .flat()
        .filter(u => u && u.toString().trim() !== ''); // Filter out empty values
      
      Logger.log('Existing usernames: ' + JSON.stringify(existingUsernames));
      
      // Case-insensitive comparison
      const usernameLower = userData.username.toLowerCase().trim();
      const existingLower = existingUsernames.map(u => u.toString().toLowerCase().trim());
      
      if (existingLower.includes(usernameLower)) {
        Logger.log('Username conflict: ' + userData.username + ' matches existing user');
        return {success: false, message: 'Username "' + userData.username + '" sudah wujud!'};
      }
    }
    
    // Generate user_id
    const timestamp = new Date().getTime();
    const userId = 'USER_' + timestamp;
    
    // Prepare row data (6 columns: A-F)
    const newRow = [
      userId,                     // A: user_id
      userData.username,          // B: username
      userData.password,          // C: password (plaintext for now)
      userData.role,              // D: role
      userData.full_name,         // E: full_name
      true                        // F: is_active (default TRUE)
    ];
    
    // Append row
    usersSheet.appendRow(newRow);
    
    // Audit log
    logToAudit(Session.getActiveUser().getEmail(), 'ADD_USER', 'USERS', userId, '', userData.username + ' (' + userData.role + ')', 'Added new user');
    
    return {
      success: true,
      message: 'User "' + userData.username + '" berjaya ditambah!',
      user_id: userId
    };
    
  } catch (e) {
    Logger.log('Error in addUser: ' + e.toString());
    return {success: false, message: 'Ralat: ' + e.toString()};
  }
}

// ============================================
// GET SUBMISSION HISTORY
// ============================================
// ============================================
// GET SUBMISSION HISTORY - SIMPLE VERSION (NO DATE FILTER)
// ============================================

// ============================================
// PRINT BORANG - GET ENTRY DATA FOR PRINTING
// ============================================
function getBorangDataForPrint(entryId) {
  try {
    Logger.log('=== GET BORANG DATA FOR PRINT ===');
    Logger.log('Entry ID: ' + entryId);
    
    if (!entryId) {
      return {
        success: false,
        message: 'Entry ID diperlukan'
      };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Search all SCORES sheets to find this entry
    const sheetNames = [
      'SCORES_PAKAIAN_KP',
      'SCORES_PAKAIAN_PLATUN', 
      'SCORES_KAWAD_KP',
      'SCORES_KAWAD_PLATUN',
      'SCORES_FORMASI'
    ];
    
    let foundData = null;
    let formType = '';
    
    // Find the entry
    for (let i = 0; i < sheetNames.length; i++) {
      const sheetName = sheetNames[i];
      const sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      
      // Check if entry exists in this sheet
      for (let row = 1; row < data.length; row++) {
        if (data[row][0] === entryId) { // Column A = entry_id
          // Found the entry! Now get all related data
          formType = sheetName.replace('SCORES_', '');
          
          const teamCode = data[row][2]; // Column C
          const judgeId = data[row][3];  // Column D
          const timestamp = data[row][1]; // Column B
          
          // Determine keyed_by column based on form type
          let keyedByColumn = 7;
          if (sheetName === 'SCORES_KAWAD_KP') keyedByColumn = 9;
          else if (sheetName === 'SCORES_KAWAD_PLATUN') keyedByColumn = 8;
          else if (sheetName === 'SCORES_FORMASI') keyedByColumn = 11;
          
          const keyedBy = data[row][keyedByColumn];
          
          // Get judge info
          const judgesSheet = ss.getSheetByName('JUDGES');
          const judgesData = judgesSheet.getDataRange().getValues();
          let judgeNumber = 'N/A';
          let judgeName = 'N/A';
          
          for (let j = 1; j < judgesData.length; j++) {
            if (judgesData[j][0] === judgeId) {
              judgeNumber = judgesData[j][2]; // Column C = judge_number
              judgeName = judgesData[j][1];   // Column B = full_name
              break;
            }
          }
          
          // Get team info
          const teamsSheet = ss.getSheetByName('TEAMS');
          const teamsData = teamsSheet.getDataRange().getValues();
          let teamName = teamCode;
          
          for (let t = 1; t < teamsData.length; t++) {
            if (teamsData[t][1] === teamCode) { // Column B = team_code
              teamName = teamsData[t][2]; // Column C = team_name
              break;
            }
          }
          
          // Now collect all rows with this entry_id
          const items = [];
          let totalScore = 0;
          
          for (let r = 1; r < data.length; r++) {
            if (data[r][0] === entryId) {
              
              // Structure depends on form type
              if (formType === 'PAKAIAN_KP' || formType === 'PAKAIAN_PLATUN') {
                // Columns: A=entry_id, B=timestamp, C=team_code, D=judge_id, E=item_code, F=score_pergerakan, G=score_bahasa, H=penalty_code, I=penalty_value, J=keyed_by
                const itemCode = data[r][4];
                const penaltyCode = data[r][6];
                
                if (itemCode) {
                  items.push({
                    code: itemCode,
                    score: data[r][5] || 0
                  });
                  totalScore += (data[r][5] || 0);
                }
                
                if (penaltyCode) {
                  items.push({
                    code: penaltyCode,
                    score: data[r][7] || 0,
                    isPenalty: true
                  });
                  totalScore += (data[r][7] || 0);
                }
                
              } else if (formType === 'KAWAD_KP') {
                // Columns: E=item_code, F=score_pergerakan, G=score_bahasa, H=penalty_code, I=penalty_value
                const itemCode = data[r][4];
                const penaltyCode = data[r][7];
                
                if (itemCode) {
                  items.push({
                    code: itemCode,
                    score_pergerakan: data[r][5] || 0,
                    score_bahasa: data[r][6] || 0
                  });
                  totalScore += (data[r][5] || 0) + (data[r][6] || 0);
                }
                
                if (penaltyCode) {
                  items.push({
                    code: penaltyCode,
                    score: data[r][8] || 0,
                    isPenalty: true
                  });
                  totalScore += (data[r][8] || 0);
                }
                
              } else if (formType === 'KAWAD_PLATUN') {
                // Columns: E=item_code, F=score, G=penalty_code, H=penalty_value
                const itemCode = data[r][4];
                const penaltyCode = data[r][6];
                
                if (itemCode) {
                  items.push({
                    code: itemCode,
                    score: data[r][5] || 0
                  });
                  totalScore += (data[r][5] || 0);
                }
                
                if (penaltyCode) {
                  items.push({
                    code: penaltyCode,
                    score: data[r][7] || 0,
                    isPenalty: true
                  });
                  totalScore += (data[r][7] || 0);
                }
                
              } else if (formType === 'FORMASI') {
                // Columns: E=item_code, F-I=scores for formasi 1-4, J=penalty_code, K=penalty_value
                const itemCode = data[r][4];
                const penaltyCode = data[r][9];
                
                if (itemCode) {
                  items.push({
                    code: itemCode,
                    score_f1: data[r][5] || 0,
                    score_f2: data[r][6] || 0,
                    score_f3: data[r][7] || 0,
                    score_f4: data[r][8] || 0
                  });
                  totalScore += (data[r][5] || 0) + (data[r][6] || 0) + (data[r][7] || 0) + (data[r][8] || 0);
                }
                
                if (penaltyCode) {
                  items.push({
                    code: penaltyCode,
                    score: data[r][10] || 0,
                    isPenalty: true
                  });
                  totalScore += (data[r][10] || 0);
                }
              }
            }
          }
          
          // Get max markah from CONFIG_SETTINGS
          const configSheet = ss.getSheetByName('CONFIG_SETTINGS');
          let maxMarkah = 0;
          if (configSheet) {
            const configData = configSheet.getDataRange().getValues();
            for (let c = 1; c < configData.length; c++) {
              const key = configData[c][0];
              const value = configData[c][1];
              
              if (formType === 'KAWAD_KP' && key === 'max_markah_kawad_kp') {
                maxMarkah = value;
              } else if (formType === 'KAWAD_PLATUN' && key === 'max_markah_kawad_platun') {
                maxMarkah = value;
              } else if (formType === 'FORMASI' && key === 'max_markah_formasi') {
                maxMarkah = value;
              }
            }
          }
          
          foundData = {
            entryId: entryId,
            formType: formType,
            teamCode: teamCode,
            teamName: teamName,
            judgeId: judgeId,
            judgeNumber: judgeNumber,
            judgeName: judgeName,
            timestamp: Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
            keyedBy: keyedBy,
            items: items,
            totalScore: totalScore,
            maxMarkah: maxMarkah
          };
          
          Logger.log('Found entry data: ' + JSON.stringify(foundData));
          
          return {
            success: true,
            data: foundData
          };
        }
      }
    }
    
    // Not found
    return {
      success: false,
      message: 'Entry tidak dijumpai'
    };
    
  } catch (error) {
    Logger.log('ERROR in getBorangDataForPrint: ' + error.toString());
    return {
      success: false,
      message: 'Ralat: ' + error.toString()
    };
  }
}

// ============================================
// LIST ALL ENTRIES FOR PRINT SELECTION
// ============================================
function listEntriesForPrint(filterRole, filterUserId) {
  try {
    Logger.log('=== LIST ENTRIES FOR PRINT ===');
    Logger.log('Role: ' + filterRole + ', User ID: ' + filterUserId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entries = [];
    
    const sheetNames = [
      'SCORES_PAKAIAN_KP',
      'SCORES_PAKAIAN_PLATUN', 
      'SCORES_KAWAD_KP',
      'SCORES_KAWAD_PLATUN',
      'SCORES_FORMASI'
    ];
    
    // Get judges info for display
    const judgesSheet = ss.getSheetByName('JUDGES');
    const judgesData = judgesSheet.getDataRange().getValues();
    const judgeMap = {};
    for (let j = 1; j < judgesData.length; j++) {
      judgeMap[judgesData[j][0]] = judgesData[j][2]; // judge_id → judge_number
    }
    
    // Scan all sheets
    sheetNames.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const formType = sheetName.replace('SCORES_', '');
      
      // Determine keyed_by column
      let keyedByColumn = 7;
      if (sheetName === 'SCORES_KAWAD_KP') keyedByColumn = 9;
      else if (sheetName === 'SCORES_KAWAD_PLATUN') keyedByColumn = 8;
      else if (sheetName === 'SCORES_FORMASI') keyedByColumn = 11;
      
      const seenEntries = {}; // Track unique entries
      
      for (let i = 1; i < data.length; i++) {
        const entryId = data[i][0];
        const keyedBy = data[i][keyedByColumn];
        
        // Skip if already seen (each entry has multiple rows)
        if (seenEntries[entryId]) continue;
        
        // Filter by user if not admin
        if (filterRole !== 'ADMIN' && keyedBy !== filterUserId) continue;
        
        seenEntries[entryId] = true;
        
        const teamCode = data[i][2];
        const judgeId = data[i][3];
        const timestamp = data[i][1];
        const judgeNumber = judgeMap[judgeId] || 'N/A';
        
        entries.push({
          entryId: entryId,
          teamCode: teamCode,
          formType: formType,
          judgeNumber: judgeNumber,
          timestamp: Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          keyedBy: keyedBy
        });
      }
    });
    
    // Sort by timestamp (newest first)
    entries.sort(function(a, b) {
      return b.timestamp.localeCompare(a.timestamp);
    });
    
    Logger.log('Found ' + entries.length + ' entries');
    
    return {
      success: true,
      entries: entries
    };
    
  } catch (error) {
    Logger.log('ERROR in listEntriesForPrint: ' + error.toString());
    return {
      success: false,
      message: 'Ralat: ' + error.toString()
    };
  }
}

// ============================================
// LIST ENTRIES FOR PRINT - WITH FILTERS
// ============================================
function listEntriesForPrintFiltered(filterRole, filterUserId, teamFilter, formFilter, judgeFilter) {
  try {
    Logger.log('=== LIST ENTRIES FOR PRINT (FILTERED) ===');
    Logger.log('Role: ' + filterRole + ', User ID: ' + filterUserId);
    Logger.log('Filters - Team: ' + teamFilter + ', Form: ' + formFilter + ', Judge: ' + judgeFilter);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entries = [];
    
    // Determine which sheets to search based on form filter
    let sheetNames = [
      'SCORES_PAKAIAN_KP',
      'SCORES_PAKAIAN_PLATUN', 
      'SCORES_KAWAD_KP',
      'SCORES_KAWAD_PLATUN',
      'SCORES_FORMASI'
    ];
    
    // If form filter specified, only search that sheet
    if (formFilter) {
      sheetNames = ['SCORES_' + formFilter];
    }
    
    // Get judges info
    const judgesSheet = ss.getSheetByName('JUDGES');
    const judgesData = judgesSheet.getDataRange().getValues();
    const judgeMap = {};
    for (let j = 1; j < judgesData.length; j++) {
      judgeMap[judgesData[j][0]] = judgesData[j][2]; // judge_id → judge_number
    }
    
    // Scan sheets
    sheetNames.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const formType = sheetName.replace('SCORES_', '');
      
      // Determine keyed_by column
      let keyedByColumn = 7;
      if (sheetName === 'SCORES_KAWAD_KP') keyedByColumn = 9;
      else if (sheetName === 'SCORES_KAWAD_PLATUN') keyedByColumn = 8;
      else if (sheetName === 'SCORES_FORMASI') keyedByColumn = 11;
      
      const seenEntries = {};
      
      for (let i = 1; i < data.length; i++) {
        const entryId = data[i][0];
        const keyedBy = data[i][keyedByColumn];
        
        if (seenEntries[entryId]) continue;
        
        // Filter by user role
        if (filterRole !== 'ADMIN' && keyedBy !== filterUserId) continue;
        
        const teamCode = data[i][2];
        const judgeId = data[i][3];
        const timestamp = data[i][1];
        const judgeNumber = judgeMap[judgeId] || 'N/A';
        
        // Apply filters
        if (teamFilter && teamCode !== teamFilter) continue;
        if (judgeFilter && judgeId !== judgeFilter) continue;
        // formFilter already applied by limiting sheetNames
        
        seenEntries[entryId] = true;
        
        entries.push({
          entryId: entryId,
          teamCode: teamCode,
          formType: formType,
          judgeNumber: judgeNumber,
          timestamp: Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          keyedBy: keyedBy
        });
      }
    });
    
    // Sort by timestamp
    entries.sort(function(a, b) {
      return b.timestamp.localeCompare(a.timestamp);
    });
    
    Logger.log('Found ' + entries.length + ' filtered entries');
    
    return {
      success: true,
      entries: entries
    };
    
  } catch (error) {
    Logger.log('ERROR in listEntriesForPrintFiltered: ' + error.toString());
    return {
      success: false,
      message: 'Ralat: ' + error.toString()
    };
  }
}

// ============================================
// GET ALL TEAMS (for filter dropdown)
// ============================================
function getAllTeams() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teamsSheet = ss.getSheetByName('TEAMS');
    
    if (!teamsSheet) {
      return [];
    }
    
    const data = teamsSheet.getDataRange().getValues();
    const teams = [];
    
    for (let i = 1; i < data.length; i++) {
      teams.push({
        team_id: data[i][0],
        team_code: data[i][1],
        team_name: data[i][2]
      });
    }
    
    return teams;
    
  } catch (error) {
    Logger.log('ERROR in getAllTeams: ' + error.toString());
    return [];
  }
}

// ============================================
// GET ALL JUDGES (for filter dropdown)
// ============================================
function getAllJudges() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const judgesSheet = ss.getSheetByName('JUDGES');
    
    if (!judgesSheet) {
      return [];
    }
    
    const data = judgesSheet.getDataRange().getValues();
    const judges = [];
    
    for (let i = 1; i < data.length; i++) {
      judges.push({
        judge_id: data[i][0],
        full_name: data[i][1],
        judge_number: data[i][2]
      });
    }
    
    return judges;
    
  } catch (error) {
    Logger.log('ERROR in getAllJudges: ' + error.toString());
    return [];
  }
}
