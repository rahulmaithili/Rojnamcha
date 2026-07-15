/**
 * SHIV SHAKTI HP GAS AGENCY - Rojnamcha Ledger Web App
 * Server-side Apps Script code
 */

/* ─── Cache helpers ─────────────────────────────────────────────────── */
var CACHE_TTL = 300; // 5 minutes in seconds

function _cGet(key) {
  var raw = CacheService.getScriptCache().get(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function _cPut(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), CACHE_TTL);
  } catch(e) { /* ignore oversized cache items */ }
}

function _cDel() {
  // Bust all ledger caches after a save
  CacheService.getScriptCache().removeAll(['rj_dates','rj_profile','rj_options']);
}

function doGet(e) {
  initializeDatabase();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var initData;
  try {
    initData = getInitialData(today);
  } catch(err) {
    initData = { success: false, error: err.toString() };
  }
  
  var template = HtmlService.createTemplateFromFile('Index_GAS');
  template.initDataJson = JSON.stringify(initData);
  
  return template.evaluate()
      .setTitle('Rojnamcha Ledger - Shiv Shakti HP Gas Agency')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper to include HTML file contents in templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Batches and retrieves all startup data in a single network call
 */
function getInitialData(dateStr) {
  try {
    initializeDatabase();
    var ssMaster = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();

    // ── 1. Company Profile (cached) ───────────────────────────────────
    var profile = _cGet('rj_profile');
    if (!profile) {
      var profileSheet = ssMaster.getSheetByName('CompanyProfile');
      var profileData = profileSheet.getDataRange().getValues();
      profile = {
        companyName: 'SHIV SHAKTI HP GAS AGENCY',
        companyAddress: 'PANDAUL BAZAR, MADHUBANI-847234',
        companyGSTIN: '10AEYPK3674C2ZH',
        companyLogo: '',
        mailDefaultRecipient: '',
        mailDefaultSubject: '',
        mailDefaultBody: ''
      };
      for (var i = 1; i < profileData.length; i++) {
        var key = profileData[i][0]; var val = profileData[i][1];
        if (key === 'CompanyName')          profile.companyName = val;
        if (key === 'CompanyAddress')       profile.companyAddress = val;
        if (key === 'CompanyGSTIN')         profile.companyGSTIN = val;
        if (key === 'CompanyLogo')          profile.companyLogo = val;
        if (key === 'MailDefaultRecipient') profile.mailDefaultRecipient = val;
        if (key === 'MailDefaultSubject')   profile.mailDefaultSubject = val;
        if (key === 'MailDefaultBody')      profile.mailDefaultBody = val;
      }
      _cPut('rj_profile', profile);
    }

    // ── 2. Dropdown Options (cached) ──────────────────────────────────
    var options = _cGet('rj_options');
    if (!options) {
      var dropdownSheet = ssMaster.getSheetByName('DropdownOptions');
      var dropdownData = dropdownSheet.getDataRange().getValues();
      var dColMap = getHeaderMapping(dropdownSheet);
      var dSecIdx = dColMap['Section']     !== undefined ? dColMap['Section']     : 0;
      var dDetIdx = dColMap['Details']     !== undefined ? dColMap['Details']     : 1;
      var dRateIdx= dColMap['DefaultRate'] !== undefined ? dColMap['DefaultRate'] : 2;
      options = { receipts: [], expenses: [] };
      for (var i = 1; i < dropdownData.length; i++) {
        var sec  = dropdownData[i][dSecIdx];
        var det  = dropdownData[i][dDetIdx];
        var rate = parseFloat(dropdownData[i][dRateIdx]) || 0;
        if (sec === 'Receipt') options.receipts.push({ details: det, defaultRate: rate });
        else if (sec === 'Expense') options.expenses.push({ details: det, defaultRate: rate });
      }
      _cPut('rj_options', options);
    }

    // ── 3. Existing Dates + current-day data (sharded) ──────
    var yearStr = getYearFromDateStr(dateStr);
    var ssYear = getSpreadsheetForYear(yearStr);
    var rSheet = ssYear.getSheetByName('Rojnamcha');
    var rData  = rSheet.getDataRange().getValues();
    var colMap = getHeaderMapping(rSheet);
    var dateIdx    = colMap['Date']               !== undefined ? colMap['Date']               : 0;
    var receiptIdx = colMap['ReceiptTotal']        !== undefined ? colMap['ReceiptTotal']        : 1;
    var expenseIdx = colMap['ExpenseTotal']        !== undefined ? colMap['ExpenseTotal']        : 2;
    var netIdx     = colMap['NetBalance']          !== undefined ? colMap['NetBalance']          : 3;
    var homeIdx    = colMap['CashSendToHome']      !== undefined ? colMap['CashSendToHome']      : 4;
    var bankIdx    = colMap['CashDepositToBank']   !== undefined ? colMap['CashDepositToBank']   : 5;
    var closingIdx = colMap['ClosingCashInHand']   !== undefined ? colMap['ClosingCashInHand']   : 6;
    var homeNoteIdx= colMap['CashSendToHomeNote']  !== undefined ? colMap['CashSendToHomeNote']  : 8;
    var bankNoteIdx= colMap['CashDepositToBankNote']!==undefined ? colMap['CashDepositToBankNote']: 9;

    var dates = getExistingDates();
    var summary = null;

    for (var i = 1; i < rData.length; i++) {
      var raw = rData[i][dateIdx];
      if (!raw) continue;
      var fd = raw instanceof Date
        ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
        : raw.toString();

      if (fd === dateStr) {
        summary = {
          date: dateStr,
          receiptTotal:      parseFloat(rData[i][receiptIdx]) || 0,
          expenseTotal:      parseFloat(rData[i][expenseIdx]) || 0,
          netBalance:        parseFloat(rData[i][netIdx])     || 0,
          cashSendToHome:    parseFloat(rData[i][homeIdx])    || 0,
          cashDepositToBank: parseFloat(rData[i][bankIdx])    || 0,
          closingCashInHand: parseFloat(rData[i][closingIdx]) || 0,
          cashSendToHomeNote:    rData[i][homeNoteIdx] || '',
          cashDepositToBankNote: rData[i][bankNoteIdx] || ''
        };
        break;
      }
    }
    
    var previousClosingCash = getPreviousClosingCash(dateStr);

    // ── 4. Details for dateStr ────────────────────────────────────────
    var dSheet  = ssYear.getSheetByName('RojnamchaDetails');
    var dData   = dSheet.getDataRange().getValues();
    var dColMap = getHeaderMapping(dSheet);
    var dDateIdx= dColMap['Date']    !== undefined ? dColMap['Date']    : 0;
    var dSecIdx2= dColMap['Section'] !== undefined ? dColMap['Section'] : 1;
    var dSlIdx  = dColMap['SL']      !== undefined ? dColMap['SL']      : 2;
    var dDetIdx2= dColMap['Details'] !== undefined ? dColMap['Details'] : 3;
    var dRateIdx2=dColMap['Rate']    !== undefined ? dColMap['Rate']    : 4;
    var dQtyIdx = dColMap['Qty']     !== undefined ? dColMap['Qty']     : 5;
    var dAmtIdx = dColMap['Amount']  !== undefined ? dColMap['Amount']  : 6;

    var details = [];
    for (var j = 1; j < dData.length; j++) {
      var rowDate = dData[j][dDateIdx];
      if (!rowDate) continue;
      var frd = rowDate instanceof Date
        ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
        : rowDate.toString();
      if (frd === dateStr) {
        details.push({
          section: dData[j][dSecIdx2],
          sl:      parseInt(dData[j][dSlIdx])    || 0,
          details: dData[j][dDetIdx2],
          rate:    parseFloat(dData[j][dRateIdx2])|| 0,
          qty:     parseFloat(dData[j][dQtyIdx])  || 0,
          amount:  parseFloat(dData[j][dAmtIdx])  || 0
        });
      }
    }

    var rojnamchaData = (summary || details.length > 0) ? { summary: summary, details: details } : null;
    if (!rojnamchaData) { rojnamchaData = { isNew: true }; }

    return {
      success: true,
      profile: profile,
      options: options,
      dates: dates,
      rojnamchaData: rojnamchaData,
      previousClosingCash: previousClosingCash
    };
  } catch (e) {
    return { success: false, error: e.toString() + '\n' + e.stack };
  }
}

/**
 * Initializes sheets if they don't exist and populates initial values
 */
function initializeDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Rojnamcha Summary Sheet
  var rojnamchaSheet = ss.getSheetByName('Rojnamcha');
  if (!rojnamchaSheet) {
    rojnamchaSheet = ss.insertSheet('Rojnamcha');
    rojnamchaSheet.appendRow([
      'Date', 
      'ReceiptTotal', 
      'ExpenseTotal', 
      'NetBalance', 
      'CashSendToHome', 
      'CashDepositToBank', 
      'ClosingCashInHand', 
      'Timestamp'
    ]);
    rojnamchaSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  } else {
    var headers = rojnamchaSheet.getRange(1, 1, 1, rojnamchaSheet.getLastColumn() || 1).getValues()[0];
    if (headers.indexOf('CashSendToHomeNote') === -1) {
      rojnamchaSheet.getRange(1, 9).setValue('CashSendToHomeNote').setFontWeight('bold');
      rojnamchaSheet.getRange(1, 10).setValue('CashDepositToBankNote').setFontWeight('bold');
    }
  }
  
  // 2. Rojnamcha Details Sheet
  var detailsSheet = ss.getSheetByName('RojnamchaDetails');
  if (!detailsSheet) {
    detailsSheet = ss.insertSheet('RojnamchaDetails');
    detailsSheet.appendRow([
      'Date', 
      'Section', 
      'SL', 
      'Details', 
      'Rate', 
      'Qty', 
      'Amount'
    ]);
    detailsSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  
  // 3. Dropdown Options Sheet
  var dropdownSheet = ss.getSheetByName('DropdownOptions');
  if (!dropdownSheet) {
    dropdownSheet = ss.insertSheet('DropdownOptions');
    dropdownSheet.appendRow(['Section', 'Details', 'DefaultRate']);
    dropdownSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // Populate default values from standard agency ledger
    var defaultOptions = [
      // Receipts
      ['Receipt', 'NAME TRANSFER A TO B', 1777.00],
      ['Receipt', 'REGULATOR LEAK', 100.00],
      ['Receipt', 'DEATH CASE name transfer', 177.00],
      ['Receipt', 'NAME TRANSFER PUMY', 143.00],
      ['Receipt', 'CTA IN', 177.00],
      ['Receipt', 'TV OUT (615102)', 118.00],
      ['Receipt', '19 KG SALE', 3500.00],
      ['Receipt', 'REFILLING 14.2 KG', 1013.00],
      ['Receipt', 'REFILLING HD SALE 14.2 KG', 1013.00],
      
      // Expenses
      ['Expense', 'TV OUT CASH RETURN (615102)', 1800.00],
      ['Expense', '19 KG SUMIT ENG', 3500.00],
      ['Expense', 'DISCOUNT 19 KG', 100.00],
      ['Expense', '19 KG SECURITY RETURN (RAKSH jHA)', 26000.00],
      ['Expense', 'REFIL PAY ONLINE 14.2 KG', 1013.00],
      ['Expense', 'NAME TRANSFER ONLINE PAID', 1949.00],
      ['Expense', 'Expy By Godam Labour Under Construction', 12800.00]
    ];
    
    for (var i = 0; i < defaultOptions.length; i++) {
      dropdownSheet.appendRow(defaultOptions[i]);
    }
  }
  
  // 4. Company Profile Sheet
  var profileSheet = ss.getSheetByName('CompanyProfile');
  if (!profileSheet) {
    profileSheet = ss.insertSheet('CompanyProfile');
    profileSheet.appendRow(['Key', 'Value']);
    profileSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    
    var defaultProfile = [
      ['CompanyName', 'SHIV SHAKTI HP GAS AGENCY'],
      ['CompanyAddress', 'PANDAUL BAZAR, MADHUBANI-847234'],
      ['CompanyGSTIN', '10AEYPK3674C2ZH']
    ];
    
    for (var j = 0; j < defaultProfile.length; j++) {
      profileSheet.appendRow(defaultProfile[j]);
    }
  }
  
  // 5. Users Sheet
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['Username', 'PasswordHash', 'Name', 'Email', 'Role', 'Status', 'ResetRequest', 'CreatedAt']);
    usersSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    
    // Add default admin user (Password: admin)
    usersSheet.appendRow([
      'admin',
      '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918',
      'Administrator',
      'admin@agency.com',
      'Admin',
      'Approved',
      false,
      new Date()
    ]);
  }
}

/**
 * Returns Company Profile data
 */
function getCompanyProfile() {
  initializeDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CompanyProfile');
  var data = sheet.getDataRange().getValues();
  
  var profile = {
    companyName: 'SHIV SHAKTI HP GAS AGENCY',
    companyAddress: 'PANDAUL BAZAR, MADHUBANI-847234',
    companyGSTIN: '10AEYPK3674C2ZH',
    companyLogo: '',
    mailDefaultRecipient: '',
    mailDefaultSubject: '',
    mailDefaultBody: ''
  };
  
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key === 'CompanyName') profile.companyName = val;
    if (key === 'CompanyAddress') profile.companyAddress = val;
    if (key === 'CompanyGSTIN') profile.companyGSTIN = val;
    if (key === 'CompanyLogo') profile.companyLogo = val;
    if (key === 'MailDefaultRecipient') profile.mailDefaultRecipient = val;
    if (key === 'MailDefaultSubject') profile.mailDefaultSubject = val;
    if (key === 'MailDefaultBody') profile.mailDefaultBody = val;
  }
  
  if (profile.companyLogo && profile.companyLogo.indexOf('data:image') === 0) {
    profile.companyLogo = profile.companyLogo.replace(/\s/g, '');
  }
  
  // If companyLogo is a Google Drive ID or URL, convert it to direct lh3 link so it works in PDF
  if (profile.companyLogo && profile.companyLogo.indexOf('data:image') !== 0) {
    var driveId = getDriveFileId(profile.companyLogo);
    if (driveId) {
      profile.companyLogo = 'https://lh3.googleusercontent.com/d/' + driveId;
    }
  }
  
  console.log({
    message: "Loaded company profile",
    logoLength: profile.companyLogo ? profile.companyLogo.length : 0
  });
  return profile;
}

/**
 * Saves Company Profile data
 */
function saveCompanyProfile(profile) {
  console.log({
    message: "Saving company profile",
    logoLength: profile.companyLogo ? profile.companyLogo.length : 0
  });
  _cDel(); // bust profile cache
  CacheService.getScriptCache().remove('rj_profile');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CompanyProfile');
  var data = sheet.getDataRange().getValues();
  
  var logoToSave = profile.companyLogo || '';
  
  var keys = ['CompanyName', 'CompanyAddress', 'CompanyGSTIN', 'CompanyLogo', 'MailDefaultRecipient', 'MailDefaultSubject', 'MailDefaultBody'];
  var vals = [
    profile.companyName, 
    profile.companyAddress, 
    profile.companyGSTIN, 
    logoToSave, 
    profile.mailDefaultRecipient || '', 
    profile.mailDefaultSubject || '', 
    profile.mailDefaultBody || ''
  ];
  
  for (var k = 0; k < keys.length; k++) {
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === keys[k]) {
        foundIndex = i + 1;
        break;
      }
    }
    
    if (foundIndex > -1) {
      sheet.getRange(foundIndex, 2).setValue(vals[k]);
    } else {
      sheet.appendRow([keys[k], vals[k]]);
    }
  }
  
  return { success: true };
}

/**
 * Returns list of dropdown options grouped by section
 */
function getDropdownOptions() {
  initializeDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DropdownOptions');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var sectionIdx = colMap['Section'] !== undefined ? colMap['Section'] : 0;
  var detailsIdx = colMap['Details'] !== undefined ? colMap['Details'] : 1;
  var rateIdx = colMap['DefaultRate'] !== undefined ? colMap['DefaultRate'] : 2;
  
  var options = {
    receipts: [],
    expenses: []
  };
  
  for (var i = 1; i < data.length; i++) {
    var section = data[i][sectionIdx];
    var details = data[i][detailsIdx];
    var rate = parseFloat(data[i][rateIdx]) || 0;
    
    if (section === 'Receipt') {
      options.receipts.push({ details: details, defaultRate: rate });
    } else if (section === 'Expense') {
      options.expenses.push({ details: details, defaultRate: rate });
    }
  }
  
  return options;
}

/**
 * Saves or updates a custom dropdown option
 */
function saveDropdownOption(section, details, defaultRate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DropdownOptions');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var sectionIdx = colMap['Section'] !== undefined ? colMap['Section'] : 0;
  var detailsIdx = colMap['Details'] !== undefined ? colMap['Details'] : 1;
  var rateIdx = colMap['DefaultRate'] !== undefined ? colMap['DefaultRate'] : 2;
  
  var foundIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][sectionIdx] === section && data[i][detailsIdx] === details) {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex > -1) {
    sheet.getRange(foundIndex, rateIdx + 1).setValue(defaultRate);
  } else {
    var newRow = [];
    newRow[sectionIdx] = section;
    newRow[detailsIdx] = details;
    newRow[rateIdx] = defaultRate;
    for (var k = 0; k < newRow.length; k++) {
      if (newRow[k] === undefined) newRow[k] = '';
    }
    sheet.appendRow(newRow);
  }
  return { success: true };
}

/**
 * Deletes a custom dropdown option
 */
function deleteDropdownOption(section, details) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DropdownOptions');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var sectionIdx = colMap['Section'] !== undefined ? colMap['Section'] : 0;
  var detailsIdx = colMap['Details'] !== undefined ? colMap['Details'] : 1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][sectionIdx] === section && data[i][detailsIdx] === details) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Option not found' };
}

/**
 * Updates an existing dropdown option details and rate
 */
function updateDropdownOption(section, oldDetails, newDetails, newRate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DropdownOptions');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var sectionIdx = colMap['Section'] !== undefined ? colMap['Section'] : 0;
  var detailsIdx = colMap['Details'] !== undefined ? colMap['Details'] : 1;
  var rateIdx = colMap['DefaultRate'] !== undefined ? colMap['DefaultRate'] : 2;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][sectionIdx] === section && data[i][detailsIdx] === oldDetails) {
      sheet.getRange(i + 1, detailsIdx + 1).setValue(newDetails);
      sheet.getRange(i + 1, rateIdx + 1).setValue(newRate);
      return { success: true };
    }
  }
  return { success: false, message: 'Option not found' };
}


/**
 * Gets sorted list of existing dates with entries
 */
function getExistingDates() {
  initializeDatabase();
  var dates = [];
  
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var masterRSheet = masterSs.getSheetByName('Rojnamcha');
  if (masterRSheet) {
    var masterData = masterRSheet.getDataRange().getValues();
    for (var i = 1; i < masterData.length; i++) {
      var d = masterData[i][0];
      if (d) {
        var fd = d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : d.toString();
        if (fd && dates.indexOf(fd) === -1) {
          dates.push(fd);
        }
      }
    }
  }
  
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var tz = Session.getScriptTimeZone();
  
  for (var key in allProps) {
    if (key.indexOf("SS_YEAR_") === 0) {
      var ssId = allProps[key];
      try {
        var ssYear = SpreadsheetApp.openById(ssId);
        var rSheet = ssYear.getSheetByName('Rojnamcha');
        if (rSheet) {
          var data = rSheet.getDataRange().getValues();
          for (var i = 1; i < data.length; i++) {
            var d = data[i][0];
            if (d) {
              var fd = d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : d.toString();
              if (fd && dates.indexOf(fd) === -1) {
                dates.push(fd);
              }
            }
          }
        }
      } catch(e) {
        console.log("Error reading dates from spreadsheet ID " + ssId + ": " + e.toString());
      }
    }
  }
  
  dates.sort(function(a, b) {
    return new Date(b) - new Date(a);
  });
  
  return dates;
}

function getPreviousClosingCash(dateStr) {
  var targetDate = new Date(dateStr);
  targetDate.setHours(0,0,0,0);
  
  var year = getYearFromDateStr(dateStr);
  var ssYear = getSpreadsheetForYear(year);
  var rSheet = ssYear.getSheetByName('Rojnamcha');
  var data = rSheet.getDataRange().getValues();
  
  var bestDate = null;
  var bestCash = 0;
  
  for (var i = 1; i < data.length; i++) {
    var cellDate = data[i][0];
    if (cellDate) {
      var d = cellDate instanceof Date ? cellDate : new Date(cellDate);
      d.setHours(0,0,0,0);
      if (d < targetDate) {
        if (!bestDate || d > bestDate) {
          bestDate = d;
          bestCash = parseFloat(data[i][6]) || 0;
        }
      }
    }
  }
  
  if (bestDate) {
    return bestCash;
  }
  
  var currentYearInt = parseInt(year);
  var props = PropertiesService.getScriptProperties();
  
  for (var y = currentYearInt - 1; y >= currentYearInt - 10; y--) {
    var ssId = props.getProperty("SS_YEAR_" + y);
    if (ssId) {
      try {
        var ssPrev = SpreadsheetApp.openById(ssId);
        var rSheetPrev = ssPrev.getSheetByName('Rojnamcha');
        var dataPrev = rSheetPrev.getDataRange().getValues();
        
        var bestDatePrev = null;
        var bestCashPrev = 0;
        
        for (var i = 1; i < dataPrev.length; i++) {
          var cellDate = dataPrev[i][0];
          if (cellDate) {
            var d = cellDate instanceof Date ? cellDate : new Date(cellDate);
            d.setHours(0,0,0,0);
            if (d < targetDate) {
              if (!bestDatePrev || d > bestDatePrev) {
                bestDatePrev = d;
                bestCashPrev = parseFloat(dataPrev[i][6]) || 0;
              }
            }
          }
        }
        if (bestDatePrev) {
          return bestCashPrev;
        }
      } catch(e) {
        console.log("Error checking previous year " + y + ": " + e.toString());
      }
    }
  }
  
  return 0;
}

function getRojnamchaData(dateStr) {
  var year = getYearFromDateStr(dateStr);
  var ss = getSpreadsheetForYear(year);
  var tz = Session.getScriptTimeZone();

  var cacheKey = 'rj_day_' + dateStr;
  var cached = _cGet(cacheKey);
  if (cached) return cached;

  var rSheet = ss.getSheetByName('Rojnamcha');
  var rData  = rSheet.getDataRange().getValues();
  var summary = null;
  var bestDate = null;
  var previousClosingCash = 0;
  var targetDate = new Date(dateStr);
  targetDate.setHours(0,0,0,0);

  for (var i = 1; i < rData.length; i++) {
    var raw = rData[i][0];
    if (!raw) continue;
    var fd = raw instanceof Date
      ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
      : raw.toString();
    if (fd === dateStr) {
      summary = {
        date: dateStr,
        receiptTotal:          parseFloat(rData[i][1]) || 0,
        expenseTotal:          parseFloat(rData[i][2]) || 0,
        netBalance:            parseFloat(rData[i][3]) || 0,
        cashSendToHome:        parseFloat(rData[i][4]) || 0,
        cashDepositToBank:     parseFloat(rData[i][5]) || 0,
        closingCashInHand:     parseFloat(rData[i][6]) || 0,
        cashSendToHomeNote:    rData[i][8] || '',
        cashDepositToBankNote: rData[i][9] || ''
      };
    } else {
      var rowD = raw instanceof Date ? raw : new Date(raw);
      rowD.setHours(0,0,0,0);
      if (rowD < targetDate && (!bestDate || rowD > bestDate)) {
        bestDate = rowD;
        previousClosingCash = parseFloat(rData[i][6]) || 0;
      }
    }
  }

  var dSheet = ss.getSheetByName('RojnamchaDetails');
  var dData  = dSheet.getDataRange().getValues();
  var details = [];
  for (var j = 1; j < dData.length; j++) {
    var rowDate = dData[j][0];
    if (!rowDate) continue;
    var frd = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
      : rowDate.toString();
    if (frd === dateStr) {
      details.push({
        section: dData[j][1],
        sl:      parseInt(dData[j][2])   || 0,
        details: dData[j][3],
        rate:    parseFloat(dData[j][4]) || 0,
        qty:     parseFloat(dData[j][5]) || 0,
        amount:  parseFloat(dData[j][6]) || 0
      });
    }
  }

  var result;
  if (!summary && details.length === 0) {
    var resolvedPreviousClosing = getPreviousClosingCash(dateStr);
    result = { isNew: true, previousClosingCash: resolvedPreviousClosing };
  } else {
    result = { summary: summary, details: details };
    _cPut(cacheKey, result);
  }
  return result;
}

function saveRojnamcha(dateStr, summary, details) {
  var year = getYearFromDateStr(dateStr);
  var ss = getSpreadsheetForYear(year);
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    throw new Error('Could not obtain lock to save ledger. Please try again.');
  }
  
  try {
    var rSheet = ss.getSheetByName('Rojnamcha');
    var rData = rSheet.getDataRange().getValues();
    var summaryRowIndex = -1;
    
    var colMap = getHeaderMapping(rSheet);
    var dateIdx = colMap['Date'] !== undefined ? colMap['Date'] : 0;
    var receiptIdx = colMap['ReceiptTotal'] !== undefined ? colMap['ReceiptTotal'] : 1;
    var expenseIdx = colMap['ExpenseTotal'] !== undefined ? colMap['ExpenseTotal'] : 2;
    var netIdx = colMap['NetBalance'] !== undefined ? colMap['NetBalance'] : 3;
    var homeIdx = colMap['CashSendToHome'] !== undefined ? colMap['CashSendToHome'] : 4;
    var bankIdx = colMap['CashDepositToBank'] !== undefined ? colMap['CashDepositToBank'] : 5;
    var closingIdx = colMap['ClosingCashInHand'] !== undefined ? colMap['ClosingCashInHand'] : 6;
    var tsIdx = colMap['Timestamp'] !== undefined ? colMap['Timestamp'] : 7;
    var homeNoteIdx = colMap['CashSendToHomeNote'] !== undefined ? colMap['CashSendToHomeNote'] : 8;
    var bankNoteIdx = colMap['CashDepositToBankNote'] !== undefined ? colMap['CashDepositToBankNote'] : 9;
    
    for (var i = 1; i < rData.length; i++) {
      var d = rData[i][dateIdx];
      var formattedDate = d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : d.toString();
      if (formattedDate === dateStr) {
        summaryRowIndex = i + 1;
        break;
      }
    }
    
    var timestamp = new Date();
    var lastCol = Math.max(10, rSheet.getLastColumn());
    var rowValues = [];
    for (var c = 0; c < lastCol; c++) {
      rowValues.push('');
    }
    
    if (summaryRowIndex > -1) {
      rowValues = rSheet.getRange(summaryRowIndex, 1, 1, lastCol).getValues()[0];
    }
    
    rowValues[dateIdx] = dateStr;
    rowValues[receiptIdx] = summary.receiptTotal;
    rowValues[expenseIdx] = summary.expenseTotal;
    rowValues[netIdx] = summary.netBalance;
    rowValues[homeIdx] = summary.cashSendToHome;
    rowValues[bankIdx] = summary.cashDepositToBank;
    rowValues[closingIdx] = summary.closingCashInHand;
    rowValues[tsIdx] = timestamp;
    rowValues[homeNoteIdx] = summary.cashSendToHomeNote || '';
    rowValues[bankNoteIdx] = summary.cashDepositToBankNote || '';
    
    if (summaryRowIndex > -1) {
      rSheet.getRange(summaryRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      rSheet.appendRow(rowValues);
    }
    
    var dSheet = ss.getSheetByName('RojnamchaDetails');
    var dData = dSheet.getDataRange().getValues();
    
    var dColMap = getHeaderMapping(dSheet);
    var dDateIdx = dColMap['Date'] !== undefined ? dColMap['Date'] : 0;
    var dSecIdx = dColMap['Section'] !== undefined ? dColMap['Section'] : 1;
    var dSlIdx = dColMap['SL'] !== undefined ? dColMap['SL'] : 2;
    var dDetIdx = dColMap['Details'] !== undefined ? dColMap['Details'] : 3;
    var dRateIdx = dColMap['Rate'] !== undefined ? dColMap['Rate'] : 4;
    var dQtyIdx = dColMap['Qty'] !== undefined ? dColMap['Qty'] : 5;
    var dAmtIdx = dColMap['Amount'] !== undefined ? dColMap['Amount'] : 6;
    
    var dLastCol = Math.max(7, dSheet.getLastColumn());
    var keeperRows = [];
    
    if (dData.length > 0) {
      var headerRow = dData[0];
      while (headerRow.length < dLastCol) headerRow.push('');
      keeperRows.push(headerRow);
    }
    
    for (var j = 1; j < dData.length; j++) {
      var rowDate = dData[j][dDateIdx];
      if (!rowDate) continue;
      var formattedRowDate = rowDate instanceof Date ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : rowDate.toString();
      if (formattedRowDate !== dateStr) {
        var row = dData[j];
        while (row.length < dLastCol) row.push('');
        if (row.length > dLastCol) row = row.slice(0, dLastCol);
        keeperRows.push(row);
      }
    }
    
    for (var k = 0; k < details.length; k++) {
      var item = details[k];
      if (item.details && item.details.trim() !== '') {
        var detailRow = [];
        for (var c = 0; c < dLastCol; c++) {
          detailRow.push('');
        }
        detailRow[dDateIdx] = dateStr;
        detailRow[dSecIdx] = item.section;
        detailRow[dSlIdx] = item.sl;
        detailRow[dDetIdx] = item.details.trim();
        detailRow[dRateIdx] = item.rate;
        detailRow[dQtyIdx] = item.qty;
        detailRow[dAmtIdx] = item.amount;
        
        keeperRows.push(detailRow);
      }
    }
    
    dSheet.clearContents();
    dSheet.getRange(1, 1, keeperRows.length, dLastCol).setValues(keeperRows);
    
    CacheService.getScriptCache().remove('rj_day_' + dateStr);
    _cDel();
    
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function getRangeReport(startDateStr, endDateStr) {
  initializeDatabase();
  
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);
  startDate.setHours(0,0,0,0);
  endDate.setHours(23,59,59,999);
  
  var startYear = startDate.getFullYear();
  var endYear = endDate.getFullYear();
  
  var openingCashMap = {};
  var rawSummaries = [];
  var tz = Session.getScriptTimeZone();
  
  for (var y = startYear; y <= endYear; y++) {
    var ssYear = getSpreadsheetForYear(y.toString());
    
    var dSheet = ssYear.getSheetByName('RojnamchaDetails');
    if (dSheet) {
      var dData = dSheet.getDataRange().getValues();
      var dColMap = getHeaderMapping(dSheet);
      var dDateIdx = dColMap['Date'] !== undefined ? dColMap['Date'] : 0;
      var dSectionIdx = dColMap['Section'] !== undefined ? dColMap['Section'] : 1;
      var dSlIdx = dColMap['SL'] !== undefined ? dColMap['SL'] : 2;
      var dAmountIdx = dColMap['Amount'] !== undefined ? dColMap['Amount'] : 6;
      
      for (var j = 1; j < dData.length; j++) {
        var rDate = dData[j][dDateIdx];
        if (rDate) {
          var formattedRDate = rDate instanceof Date ? Utilities.formatDate(rDate, tz, 'yyyy-MM-dd') : rDate.toString();
          var section = dData[j][dSectionIdx];
          var sl = parseInt(dData[j][dSlIdx]) || 0;
          var amount = parseFloat(dData[j][dAmountIdx]) || 0;
          
          if (section === 'Receipt' && sl === 1) {
            openingCashMap[formattedRDate] = amount;
          }
        }
      }
    }
    
    var rSheet = ssYear.getSheetByName('Rojnamcha');
    if (rSheet) {
      var rData = rSheet.getDataRange().getValues();
      var rColMap = getHeaderMapping(rSheet);
      var rDateIdx = rColMap['Date'] !== undefined ? rColMap['Date'] : 0;
      
      for (var i = 1; i < rData.length; i++) {
        var cellDate = rData[i][rDateIdx];
        if (cellDate) {
          var formattedDate = cellDate instanceof Date ? Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd') : cellDate.toString();
          var dateVal = new Date(formattedDate);
          dateVal.setHours(12,0,0,0);
          
          if (dateVal >= startDate && dateVal <= endDate) {
            rawSummaries.push({
              rowValues: rData[i],
              formattedDate: formattedDate,
              rColMap: rColMap
            });
          }
        }
      }
    }
  }
  
  var records = [];
  for (var k = 0; k < rawSummaries.length; k++) {
    var item = rawSummaries[k];
    var formattedDate = item.formattedDate;
    var rColMap = item.rColMap;
    var rowValues = item.rowValues;
    
    var rReceiptIdx = rColMap['ReceiptTotal'] !== undefined ? rColMap['ReceiptTotal'] : 1;
    var rExpenseIdx = rColMap['ExpenseTotal'] !== undefined ? rColMap['ExpenseTotal'] : 2;
    var rHomeIdx = rColMap['CashSendToHome'] !== undefined ? rColMap['CashSendToHome'] : 4;
    var rBankIdx = rColMap['CashDepositToBank'] !== undefined ? rColMap['CashDepositToBank'] : 5;
    var rClosingIdx = rColMap['ClosingCashInHand'] !== undefined ? rColMap['ClosingCashInHand'] : 6;
    var rHomeNoteIdx = rColMap['CashSendToHomeNote'] !== undefined ? rColMap['CashSendToHomeNote'] : 8;
    var rBankNoteIdx = rColMap['CashDepositToBankNote'] !== undefined ? rColMap['CashDepositToBankNote'] : 9;
    
    var opCash = openingCashMap[formattedDate] || 0;
    var rTotal = parseFloat(rowValues[rReceiptIdx]) || 0;
    var eTotal = parseFloat(rowValues[rExpenseIdx]) || 0;
    var cashHome = parseFloat(rowValues[rHomeIdx]) || 0;
    var cashBank = parseFloat(rowValues[rBankIdx]) || 0;
    var closingCash = parseFloat(rowValues[rClosingIdx]) || 0;
    
    var dynamicReceipts = rTotal - opCash;
    if (dynamicReceipts < 0) dynamicReceipts = 0;
    
    records.push({
      date: formattedDate,
      openingCash: opCash,
      receipts: dynamicReceipts,
      expenses: eTotal,
      cashSendToHome: cashHome,
      cashSendToHomeNote: rowValues[rHomeNoteIdx] || '',
      cashDepositToBank: cashBank,
      cashDepositToBankNote: rowValues[rBankNoteIdx] || '',
      closingCashInHand: closingCash
    });
  }
  
  records.sort(function(a, b) {
    return new Date(a.date) - new Date(b.date);
  });
  
  return records;
}

/**
 * SHA-256 password hashing utility for backend security
 */
function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    hash += byteString;
  }
  return hash;
}

/**
 * Helper to map header text to column index (0-based)
 */
function getHeaderMapping(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h !== null && h !== undefined) {
      map[h.toString().trim()] = i;
    }
  }
  return map;
}

/**
 * Validates user credentials and returns session details
 */
function loginUser(username, password) {
  initializeDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var inputHash = hashPassword(password);
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var passIdx = colMap['PasswordHash'] !== undefined ? colMap['PasswordHash'] : 1;
  var nameIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
  var emailIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
  var roleIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
  var statusIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
  
  var normalizedUsername = username.trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString().toLowerCase() === normalizedUsername) {
      if (data[i][passIdx] === inputHash) {
        var status = data[i][statusIdx] ? data[i][statusIdx].toString().trim() : 'Pending';
        if (status === 'Approved') {
          return {
            success: true,
            user: {
              username: data[i][usernameIdx].toString().trim(),
              name: data[i][nameIdx] ? data[i][nameIdx].toString().trim() : '',
              email: data[i][emailIdx] ? data[i][emailIdx].toString().trim() : '',
              role: data[i][roleIdx] ? data[i][roleIdx].toString().trim() : 'User',
              status: status,
              passwordHash: inputHash
            }
          };
        } else if (status === 'Pending') {
          return { success: false, message: 'Your account is pending admin approval.' };
        } else {
          return { success: false, message: 'Your account access has been restricted/rejected.' };
        }
      } else {
        return { success: false, message: 'Invalid password. Please try again.' };
      }
    }
  }
  return { success: false, message: 'Username not found.' };
}

/**
 * Registers a new user with status 'Pending'
 */
function registerUser(username, password, name, email) {
  initializeDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  
  var normalizedUsername = username.trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString().toLowerCase() === normalizedUsername) {
      return { success: false, message: 'Username is already taken.' };
    }
  }
  
  var hashedPassword = hashPassword(password);
  var row = [];
  var lastCol = Math.max(7, sheet.getLastColumn() - 1);
  for (var c = 0; c <= lastCol; c++) {
    row.push('');
  }
  
  var uIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var pIdx = colMap['PasswordHash'] !== undefined ? colMap['PasswordHash'] : 1;
  var nIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
  var eIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
  var rIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
  var sIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
  var reIdx = colMap['ResetRequest'] !== undefined ? colMap['ResetRequest'] : 6;
  var cIdx = colMap['CreatedAt'] !== undefined ? colMap['CreatedAt'] : 7;
  
  row[uIdx] = username.trim();
  row[pIdx] = hashedPassword;
  row[nIdx] = name.trim();
  row[eIdx] = email.trim();
  row[rIdx] = 'User';
  row[sIdx] = 'Pending';
  row[reIdx] = false;
  row[cIdx] = new Date();
  
  sheet.appendRow(row);
  
  return { success: true, message: 'Registration submitted. Awaiting Admin approval.' };
}

/**
 * Flags a user account as requesting a password reset
 */
function requestPasswordReset(username) {
  initializeDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var reIdx = colMap['ResetRequest'] !== undefined ? colMap['ResetRequest'] : 6;
  
  var normalizedUsername = username.trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString().toLowerCase() === normalizedUsername) {
      sheet.getRange(i + 1, reIdx + 1).setValue(true); // ResetRequest = true
      return { success: true, message: 'Reset request sent to Admin. Contact Admin to get your new password.' };
    }
  }
  return { success: false, message: 'Username not found.' };
}

/**
 * Validates persistent session on load
 */
function validateSession(username, passwordHash) {
  try {
    initializeDatabase();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return { success: false, error: "Users sheet not found" };
    var data = sheet.getDataRange().getValues();
    
    var colMap = getHeaderMapping(sheet);
    var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
    var passIdx = colMap['PasswordHash'] !== undefined ? colMap['PasswordHash'] : 1;
    var nameIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
    var emailIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
    var roleIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
    var statusIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
    
    var normalizedUsername = username.trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][usernameIdx] && data[i][usernameIdx].toString().toLowerCase() === normalizedUsername && data[i][passIdx] === passwordHash) {
        var status = data[i][statusIdx] ? data[i][statusIdx].toString().trim() : '';
        if (status === 'Approved') {
          return {
            success: true,
            user: {
              username: data[i][usernameIdx].toString().trim(),
              name: data[i][nameIdx] ? data[i][nameIdx].toString().trim() : '',
              email: data[i][emailIdx] ? data[i][emailIdx].toString().trim() : '',
              role: data[i][roleIdx] ? data[i][roleIdx].toString().trim() : 'User',
              status: status
            }
          };
        }
      }
    }
    return { success: false, error: "Invalid credentials or account pending approval" };
  } catch (e) {
    return { success: false, error: e.toString() + "\n" + e.stack };
  }
}

/**
 * Fetches all users for admin panel (Admin only)
 */
function adminGetUserList(callerUsername, callerPasswordHash) {
  var check = validateSession(callerUsername, callerPasswordHash);
  if (!check.success || check.user.role !== 'Admin') {
    throw new Error('Unauthorized access');
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  
  var colMap = getHeaderMapping(sheet);
  var uIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var nIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
  var eIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
  var rIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
  var sIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
  var reIdx = colMap['ResetRequest'] !== undefined ? colMap['ResetRequest'] : 6;
  var cIdx = colMap['CreatedAt'] !== undefined ? colMap['CreatedAt'] : 7;
  
  for (var i = 1; i < data.length; i++) {
    var username = data[i][uIdx];
    if (!username || username.toString().trim() === '') continue; // Skip empty rows
    
    var dateVal = data[i][cIdx];
    var formattedCreated = '';
    if (dateVal) {
      if (dateVal instanceof Date) {
        formattedCreated = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      } else {
        formattedCreated = dateVal.toString();
      }
    }
    
    users.push({
      username: username.toString().trim(),
      name: data[i][nIdx] ? data[i][nIdx].toString().trim() : '',
      email: data[i][eIdx] ? data[i][eIdx].toString().trim() : '',
      role: data[i][rIdx] ? data[i][rIdx].toString().trim() : 'User',
      status: data[i][sIdx] ? data[i][sIdx].toString().trim() : 'Pending',
      resetRequest: data[i][reIdx] === true || data[i][reIdx] === 'true',
      createdAt: formattedCreated
    });
  }
  return users;
}

/**
 * Adds a new user directly (Admin only)
 */
function adminAddUser(callerUsername, callerPasswordHash, user) {
  var check = validateSession(callerUsername, callerPasswordHash);
  if (!check.success || check.user.role !== 'Admin') {
    throw new Error('Unauthorized access');
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  
  var normalizedUsername = user.username.trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString().toLowerCase() === normalizedUsername) {
      return { success: false, message: 'Username already exists.' };
    }
  }
  
  var hashedPassword = hashPassword(user.password || '123456'); // Default password if empty
  var row = [];
  var lastCol = Math.max(7, sheet.getLastColumn() - 1);
  for (var c = 0; c <= lastCol; c++) {
    row.push('');
  }
  
  var uIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var pIdx = colMap['PasswordHash'] !== undefined ? colMap['PasswordHash'] : 1;
  var nIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
  var eIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
  var rIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
  var sIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
  var reIdx = colMap['ResetRequest'] !== undefined ? colMap['ResetRequest'] : 6;
  var cIdx = colMap['CreatedAt'] !== undefined ? colMap['CreatedAt'] : 7;
  
  row[uIdx] = user.username.trim();
  row[pIdx] = hashedPassword;
  row[nIdx] = user.name.trim();
  row[eIdx] = user.email.trim();
  row[rIdx] = user.role;
  row[sIdx] = user.status;
  row[reIdx] = false;
  row[cIdx] = new Date();
  
  sheet.appendRow(row);
  return { success: true };
}

/**
 * Updates a user's role or status (Admin only)
 */
function adminUpdateUser(callerUsername, callerPasswordHash, targetUsername, updatedData) {
  var check = validateSession(callerUsername, callerPasswordHash);
  if (!check.success || check.user.role !== 'Admin') {
    throw new Error('Unauthorized access');
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var nameIdx = colMap['Name'] !== undefined ? colMap['Name'] : 2;
  var emailIdx = colMap['Email'] !== undefined ? colMap['Email'] : 3;
  var roleIdx = colMap['Role'] !== undefined ? colMap['Role'] : 4;
  var statusIdx = colMap['Status'] !== undefined ? colMap['Status'] : 5;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString() === targetUsername) {
      sheet.getRange(i + 1, nameIdx + 1).setValue(updatedData.name.trim());
      sheet.getRange(i + 1, emailIdx + 1).setValue(updatedData.email.trim());
      sheet.getRange(i + 1, roleIdx + 1).setValue(updatedData.role);
      sheet.getRange(i + 1, statusIdx + 1).setValue(updatedData.status);
      return { success: true };
    }
  }
  return { success: false, message: 'User not found' };
}

/**
 * Deletes a user (Admin only)
 */
function adminDeleteUser(callerUsername, callerPasswordHash, targetUsername) {
  var check = validateSession(callerUsername, callerPasswordHash);
  if (!check.success || check.user.role !== 'Admin') {
    throw new Error('Unauthorized access');
  }
  
  if (callerUsername === targetUsername) {
    return { success: false, message: 'You cannot delete your own admin account.' };
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString() === targetUsername) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'User not found' };
}

/**
 * Resets a user's password to a new temporary password (Admin only)
 */
function adminResetPassword(callerUsername, callerPasswordHash, targetUsername, newPassword) {
  var check = validateSession(callerUsername, callerPasswordHash);
  if (!check.success || check.user.role !== 'Admin') {
    throw new Error('Unauthorized access');
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var colMap = getHeaderMapping(sheet);
  var usernameIdx = colMap['Username'] !== undefined ? colMap['Username'] : 0;
  var passIdx = colMap['PasswordHash'] !== undefined ? colMap['PasswordHash'] : 1;
  var reIdx = colMap['ResetRequest'] !== undefined ? colMap['ResetRequest'] : 6;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][usernameIdx] && data[i][usernameIdx].toString() === targetUsername) {
      var hashedNewPassword = hashPassword(newPassword);
      sheet.getRange(i + 1, passIdx + 1).setValue(hashedNewPassword); // Set PasswordHash
      sheet.getRange(i + 1, reIdx + 1).setValue(false); // Clear ResetRequest flag
      return { success: true };
    }
  }
  return { success: false, message: 'User not found' };
}

/**
 * Fetches summary records and online payment details for the audit report
 */
function getRangeAuditData(startDateStr, endDateStr) {
  initializeDatabase();
  
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);
  startDate.setHours(0,0,0,0);
  endDate.setHours(23,59,59,999);
  
  var startYear = startDate.getFullYear();
  var endYear = endDate.getFullYear();
  
  var summaries = [];
  var onlinePayments = [];
  var tz = Session.getScriptTimeZone();
  var ONLINE_KEYWORDS = ['online', 'gpay', 'phonepe', 'paytm', 'upi', 'digital', 'neft', 'rtgs', 'imps', 'bhim', 'netbanking', 'internet pay', 'net pay'];
  
  for (var y = startYear; y <= endYear; y++) {
    var ssYear = getSpreadsheetForYear(y.toString());
    
    // 1. Get summaries
    var rSheet = ssYear.getSheetByName('Rojnamcha');
    if (rSheet) {
      var rData = rSheet.getDataRange().getValues();
      var colMap = getHeaderMapping(rSheet);
      var dateIdx = colMap['Date'] !== undefined ? colMap['Date'] : 0;
      var homeIdx = colMap['CashSendToHome'] !== undefined ? colMap['CashSendToHome'] : 4;
      var bankIdx = colMap['CashDepositToBank'] !== undefined ? colMap['CashDepositToBank'] : 5;
      var closingIdx = colMap['ClosingCashInHand'] !== undefined ? colMap['ClosingCashInHand'] : 6;
      var homeNoteIdx = colMap['CashSendToHomeNote'] !== undefined ? colMap['CashSendToHomeNote'] : 8;
      var bankNoteIdx = colMap['CashDepositToBankNote'] !== undefined ? colMap['CashDepositToBankNote'] : 9;
      
      for (var i = 1; i < rData.length; i++) {
        var cellDate = rData[i][dateIdx];
        if (cellDate) {
          var formattedDate = cellDate instanceof Date ? Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd') : cellDate.toString();
          var dateVal = new Date(formattedDate);
          dateVal.setHours(12,0,0,0);
          
          if (dateVal >= startDate && dateVal <= endDate) {
            summaries.push({
              date: formattedDate,
              cashSendToHome: parseFloat(rData[i][homeIdx]) || 0,
              cashSendToHomeNote: rData[i][homeNoteIdx] || '',
              cashDepositToBank: parseFloat(rData[i][bankIdx]) || 0,
              cashDepositToBankNote: rData[i][bankNoteIdx] || '',
              closingCashInHand: parseFloat(rData[i][closingIdx]) || 0
            });
          }
        }
      }
    }
    
    // 2. Get details for online payments
    var dSheet = ssYear.getSheetByName('RojnamchaDetails');
    if (dSheet) {
      var dData = dSheet.getDataRange().getValues();
      var dColMap = getHeaderMapping(dSheet);
      var dDateIdx = dColMap['Date'] !== undefined ? dColMap['Date'] : 0;
      var dSecIdx = dColMap['Section'] !== undefined ? dColMap['Section'] : 1;
      var dDetIdx = dColMap['Details'] !== undefined ? dColMap['Details'] : 3;
      var dRateIdx = dColMap['Rate'] !== undefined ? dColMap['Rate'] : 4;
      var dQtyIdx = dColMap['Qty'] !== undefined ? dColMap['Qty'] : 5;
      var dAmtIdx = dColMap['Amount'] !== undefined ? dColMap['Amount'] : 6;
      
      for (var j = 1; j < dData.length; j++) {
        var rowDate = dData[j][dDateIdx];
        if (rowDate) {
          var formattedRowDate = rowDate instanceof Date ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') : rowDate.toString();
          var dateVal = new Date(formattedRowDate);
          dateVal.setHours(12,0,0,0);
          
          if (dateVal >= startDate && dateVal <= endDate) {
            var section = dData[j][dSecIdx];
            var details = dData[j][dDetIdx] || '';
            
            if (section === 'Expense') {
              var lowerDetails = details.toLowerCase();
              var isOnline = false;
              for (var k = 0; k < ONLINE_KEYWORDS.length; k++) {
                if (lowerDetails.indexOf(ONLINE_KEYWORDS[k]) > -1) {
                  isOnline = true;
                  break;
                }
              }
              
              if (isOnline) {
                onlinePayments.push({
                  date: formattedRowDate,
                  details: details,
                  rate: parseFloat(dData[j][dRateIdx]) || 0,
                  qty: parseFloat(dData[j][dQtyIdx]) || 0,
                  amount: parseFloat(dData[j][dAmtIdx]) || 0
                });
              }
            }
          }
        }
      }
    }
  }
  
  summaries.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
  onlinePayments.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
  
  return {
    summaries: summaries,
    onlinePayments: onlinePayments
  };
}

/**
 * Diagnostic helper to print spreadsheet details and active Users list
 */
function debugUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl();
  var name = ss.getName();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    return "No Users sheet found in spreadsheet: " + name + " (" + url + ")";
  }
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push(data[i]);
  }
  return {
    spreadsheetName: name,
    spreadsheetUrl: url,
    rowsCount: data.length,
    headers: data[0],
    rows: rows
  };
}

/**
 * Generates detailed ledger records (summary + detail rows) for all dates between startDateStr and endDateStr
 */
function getDetailedRangeLedger(startDateStr, endDateStr) {
  initializeDatabase();
  
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);
  startDate.setHours(0,0,0,0);
  endDate.setHours(23,59,59,999);
  
  var startYear = startDate.getFullYear();
  var endYear = endDate.getFullYear();
  
  var summaryRecords = getRangeReport(startDateStr, endDateStr);
  
  var detailsByDate = {};
  var tz = Session.getScriptTimeZone();
  
  for (var y = startYear; y <= endYear; y++) {
    var ssYear = getSpreadsheetForYear(y.toString());
    var dSheet = ssYear.getSheetByName('RojnamchaDetails');
    if (dSheet) {
      var dData = dSheet.getDataRange().getValues();
      var dColMap = getHeaderMapping(dSheet);
      var dDateIdx = dColMap['Date'] !== undefined ? dColMap['Date'] : 0;
      
      for (var j = 1; j < dData.length; j++) {
        var rowDate = dData[j][dDateIdx];
        if (rowDate) {
          var formattedRowDate = rowDate instanceof Date ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') : rowDate.toString();
          var dateVal = new Date(formattedRowDate);
          dateVal.setHours(12,0,0,0);
          
          if (dateVal >= startDate && dateVal <= endDate) {
            if (!detailsByDate[formattedRowDate]) {
              detailsByDate[formattedRowDate] = [];
            }
            detailsByDate[formattedRowDate].push({
              section: dData[j][1],
              sl: parseInt(dData[j][2]) || 0,
              details: dData[j][3],
              rate: parseFloat(dData[j][4]) || 0,
              qty: parseFloat(dData[j][5]) || 0,
              amount: parseFloat(dData[j][6]) || 0
            });
          }
        }
      }
    }
  }
  
  var result = [];
  for (var i = 0; i < summaryRecords.length; i++) {
    var s = summaryRecords[i];
    var dateKey = s.date;
    var dayDetails = detailsByDate[dateKey] || [];
    
    dayDetails.sort(function(a, b) {
      if (a.section !== b.section) {
        return a.section === 'Receipt' ? -1 : 1;
      }
      return a.sl - b.sl;
    });
    
    result.push({
      summary: s,
      details: dayDetails
    });
  }
  
  return result;
}

/**
 * Sends a PDF of the ledger via Gmail to multiple comma-separated email addresses
 */
function sendLedgerEmail(recipientStr, subject, bodyText, startDateStr, endDateStr, preparedBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Could not obtain script lock. Please try again.' };
  }
  
  var tempFile = null;
  try {
    // 1. Validate recipients
    if (!recipientStr || recipientStr.trim() === '') {
      throw new Error('Recipient email addresses are required.');
    }
    
    var emails = recipientStr.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e !== ''; });
    if (emails.length === 0) {
      throw new Error('No valid recipient email addresses found.');
    }
    
    // 2. Fetch detailed ledger data
    var ledgerList = getDetailedRangeLedger(startDateStr, endDateStr);
    if (ledgerList.length === 0) {
      throw new Error('No ledger data found for the selected date range: ' + startDateStr + ' to ' + endDateStr);
    }
    
    // 3. Load company profile settings
    var profile = getCompanyProfile();
    
    // 4. Prepare logo for (a) email inline CID image and (b) PDF via temp Drive URL
    var logoUrlForPdf   = '';   // URL that the Apps Script PDF renderer can fetch
    var logoBlobForEmail = null; // Blob used as inline CID attachment in the email body
    var tempLogoFile     = null; // Temporary public Drive file for PDF; deleted after

    if (profile.companyLogo) {
      try {
        var rawLogo = profile.companyLogo.replace(/\s/g, '');

        var logoBlob;
        if (rawLogo.indexOf('data:image') === 0) {
          // Stored as base64 data URL
          var mimeMatch = rawLogo.match(/data:(image\/[^;]+);base64,/);
          var mime      = mimeMatch ? mimeMatch[1] : 'image/jpeg';
          var b64Data   = rawLogo.replace(/^data:[^;]+;base64,/, '');
          logoBlob = Utilities.newBlob(Utilities.base64Decode(b64Data), mime, 'company_logo');
        } else {
          // Drive URL / ID
          var dId = getDriveFileId(rawLogo);
          if (dId) {
            try { logoBlob = DriveApp.getFileById(dId).getBlob(); } catch(e) {}
          }
        }

        if (logoBlob) {
          // (a) Blob for inline CID image in the email body
          logoBlobForEmail = logoBlob;

          // (b) Upload a temporary copy to Drive so the PDF renderer can fetch it
          try {
            var folder = DriveApp.getRootFolder();
            tempLogoFile = folder.createFile(logoBlob.copyBlob().setName('rj_tmp_logo'));
            tempLogoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            // lh3 thumbnail URL – reliably loaded by the Apps Script PDF engine
            logoUrlForPdf = 'https://lh3.googleusercontent.com/d/' + tempLogoFile.getId();
          } catch(uploadErr) {
            console.log('Temp logo upload failed: ' + uploadErr.message);
            logoUrlForPdf = '';
          }
        }
      } catch (logoErr) {
        console.log('Logo preparation error: ' + logoErr.message);
      }
    }

    console.log({
      message: "Generating email PDF",
      logoUrlForPdf: logoUrlForPdf ? logoUrlForPdf.substring(0, 80) : 'none',
      hasEmailBlob: !!logoBlobForEmail
    });

    // 5. Create HTML template
    var template = HtmlService.createTemplateFromFile('EmailTemplate');
    template.companyName       = profile.companyName    || '';
    template.companyAddress    = profile.companyAddress || '';
    template.companyGSTIN      = profile.companyGSTIN   || '';
    template.companyLogoBase64 = logoUrlForPdf;   // Drive URL for PDF renderer
    template.startDate         = startDateStr;
    template.endDate           = endDateStr;
    template.ledgerList        = ledgerList;
    template.formatDateReadableServer = formatDateReadableServer;
    template.preparedBy        = preparedBy || 'System';
    template.formatNoteForSummary = formatNoteForSummary;
    
    var htmlContent = template.evaluate().getContent();
    
    // 6. Convert to PDF Blob
    var pdfName = 'Rojnamcha_Ledger_' + startDateStr + '_to_' + endDateStr + '.pdf';
    var pdfInputBlob = Utilities.newBlob(htmlContent, 'text/html', pdfName);
    var pdfBlob = pdfInputBlob.getAs('application/pdf');
    pdfBlob.setName(pdfName);
    
    // Clean up temp logo Drive file now that PDF is generated
    if (tempLogoFile) {
      try { tempLogoFile.setTrashed(true); tempLogoFile = null; } catch(e) {}
    }
    
    // 6. Build the beautiful HTML email body
    var htmlEmailBody = 
      '<div style="font-family: \'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif; background-color: #f3f4f6; padding: 25px; margin: 0; color: #1f2937;">' +
      '  <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); overflow: hidden; border: 1px solid #e5e7eb;">' +
      '    <div style="background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%); padding: 25px 35px; text-align: center; color: #ffffff;">' +
      (logoBlobForEmail ? '      <img src="cid:logoImg" style="max-height: 45px; max-width: 120px; margin-bottom: 12px; display: inline-block; vertical-align: middle;" /><br/>' : '') +
      '      <h2 style="margin: 0; font-size: 22px; font-weight: bold; letter-spacing: 0.5px;">' + (profile.companyName || 'SHIV SHAKTI HP GAS AGENCY') + '</h2>' +
      '      <p style="margin: 5px 0 0 0; font-size: 12px; color: #e0e7ff; opacity: 0.9;">' + (profile.companyAddress || '') + '</p>' +
      (profile.companyGSTIN ? '      <p style="margin: 5px 0 0 0; font-size: 11px; font-weight: bold; color: #c7d2fe;">GSTIN: ' + profile.companyGSTIN + '</p>' : '') +
      '    </div>' +
      '    <div style="padding: 30px; line-height: 1.6; font-size: 14px;">' +
      '      <p style="color: #4b5563; margin-top: 0;">' + bodyText.replace(/\n/g, '<br>') + '</p>' +
      '      <div style="margin: 25px 0; padding: 15px; background-color: #f9fafb; border-left: 4px solid #4f46e5; border-radius: 4px;">' +
      '        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">' +
      '          <tr>' +
      '            <td style="padding: 4px 0; color: #6b7280; font-weight: 500;">Report Type:</td>' +
      '            <td style="padding: 4px 0; color: #111827; font-weight: bold;">Daily Ledger (Rojnamcha)</td>' +
      '          </tr>' +
      '          <tr>' +
      '            <td style="padding: 4px 0; color: #6b7280; font-weight: 500;">Date Range:</td>' +
      '            <td style="padding: 4px 0; color: #111827; font-weight: bold;">' + (startDateStr === endDateStr ? startDateStr : startDateStr + ' to ' + endDateStr) + '</td>' +
      '          </tr>' +
      '        </table>' +
      '      </div>' +
      '      <p style="color: #4b5563; font-size: 12px; font-style: italic; margin-bottom: 0;">Note: The detailed ledger PDF is attached to this email.</p>' +
      '    </div>' +
      '    <div style="background-color: #f9fafb; padding: 15px 30px; text-align: center; font-size: 11px; color: #9ca3af; border-top: 1px solid #f3f4f6;">' +
      '      <p style="margin: 0;">This is an automated email sent from Rojnamcha Management System.</p>' +
      '      <p style="margin: 4px 0 0 0;">&copy; ' + new Date().getFullYear() + ' ' + (profile.companyName || 'SHIV SHAKTI HP GAS AGENCY') + '. All rights reserved.</p>' +
      '    </div>' +
      '  </div>' +
      '</div>';
      
    // 7. Send the email with HTML body and PDF attachment
    var mailOptions = {
      htmlBody: htmlEmailBody,
      attachments: [pdfBlob],
      name: profile.companyName || 'Rojnamcha System'
    };
    if (logoBlobForEmail) {
      mailOptions.inlineImages = {
        logoImg: logoBlobForEmail
      };
    }
    
    for (var i = 0; i < emails.length; i++) {
      GmailApp.sendEmail(emails[i], subject, bodyText, mailOptions);
    }
    
    return { success: true, message: 'Email sent successfully to ' + emails.length + ' recipient(s)!' };
  } catch (e) {
    return { success: false, message: 'Email failed: ' + e.toString() };
  } finally {
    // Clean up temp logo file if not already deleted
    if (tempLogoFile) {
      try { tempLogoFile.setTrashed(true); } catch(e) {}
    }
    lock.releaseLock();
  }
}

/**
 * Formats a yyyy-MM-dd date into dd-MMM-yyyy format
 */
function formatDateReadableServer(dateStr) {
  if (!dateStr) return '';
  var parts = dateStr.split('-');
  if (parts.length === 3) {
    var date = new Date(parts[0], parts[1] - 1, parts[2]);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }
  return dateStr;
}

/**
 * Helper to extract file ID from a Google Drive URL
 */
function getDriveFileId(url) {
  if (!url) return null;
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Cleans and formats breakdown notes for the printed Daily Ledger summary
 */
function formatNoteForSummary(noteStr) {
  if (!noteStr) return '';
  noteStr = noteStr.trim();
  if (noteStr.indexOf('[Breakdown:') === 0) {
    var match = noteStr.match(/^\[Breakdown:\s*(.*?)\]\s*(.*)$/);
    if (match) {
      var breakdownStr = match[1];
      var generalNote = match[2].trim();
      var parts = [];
      var items = breakdownStr.split(';').map(function(s) { return s.trim(); }).filter(Boolean);
      items.forEach(function(item) {
        var amount = 0;
        var note = '';
        var dashIdx = item.indexOf(' - ');
        if (dashIdx !== -1) {
          amount = parseFloat(item.substring(0, dashIdx).trim()) || 0;
          note = item.substring(dashIdx + 3).trim();
        } else {
          amount = parseFloat(item) || 0;
          note = item.replace(/^[0-9.]+\s*/, '').trim();
        }
        
        var cleanText = note;
        var directMatch = note.match(/^((?:\s*(?:\d+|coins)x\d+\s*)(?:,\s*(?:\d+|coins)x\d+\s*)*)(.*)$/i);
        if (directMatch) {
          cleanText = directMatch[2].trim();
        }
        
        var amtStr = 'Rs.' + amount.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        if (cleanText) {
          parts.push(amtStr + ' (' + cleanText + ')');
        } else {
          parts.push(amtStr);
        }
      });
      var result = parts.join(' + ');
      if (generalNote) {
        result += ' — ' + generalNote;
      }
      return result;
    }
  }
  
  var directMatch = noteStr.match(/^((?:\s*(?:\d+|coins)x\d+\s*)(?:,\s*(?:\d+|coins)x\d+\s*)*)(.*)$/i);
  if (directMatch) {
    return directMatch[2].trim();
  }
  var bracketMatch = noteStr.match(/^\[(.*?)\](.*)$/);
  if (bracketMatch) {
    return bracketMatch[2].trim();
  }
  
  return noteStr;
}

/* ─── API Router and Sharding Logic ──────────────────────────────────── */

function doPost(e) {
  var output = { success: false, error: 'Unknown request' };
  try {
    var rawData = e.postData.contents;
    var payload = JSON.parse(rawData);
    var action = payload.action;
    var args = payload.args || [];
    
    var whitelist = [
      'loginUser', 'registerUser', 'requestPasswordReset', 'validateSession',
      'getInitialData', 'getRojnamchaData', 'saveRojnamcha', 'getPreviousClosingCash',
      'getDropdownOptions', 'saveDropdownOption', 'deleteDropdownOption', 'updateDropdownOption',
      'getExistingDates', 'getRangeReport', 'getDetailedRangeLedger', 'getRangeAuditData',
      'getCompanyProfile', 'saveCompanyProfile', 'sendLedgerEmail',
      'adminGetUserList', 'adminAddUser', 'adminUpdateUser', 'adminDeleteUser', 'adminResetPassword',
      'exportAllDataJSON', 'importAllDataJSON'
    ];
    
    if (whitelist.indexOf(action) > -1 && typeof this[action] === 'function') {
      var result = this[action].apply(this, args);
      output = result;
    } else {
      output = { success: false, error: 'Unauthorized or missing API action: ' + action };
    }
  } catch (err) {
    output = { success: false, error: err.toString() + '\n' + err.stack };
  }
  
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJsonValue(val) {
  if (typeof val === 'string') {
    var isoDateReg = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/;
    if (isoDateReg.test(val)) {
      return new Date(val);
    }
  }
  return val;
}

function exportAllDataJSON() {
  initializeDatabase();
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  
  var backup = {
    version: "1.1",
    exportedAt: new Date().toISOString(),
    master: {
      DropdownOptions: [],
      CompanyProfile: [],
      Users: []
    },
    rojnamcha: [],
    rojnamchaDetails: []
  };
  
  // 1. Export master sheets
  var masterSheets = ['DropdownOptions', 'CompanyProfile', 'Users'];
  for (var i = 0; i < masterSheets.length; i++) {
    var sheet = masterSs.getSheetByName(masterSheets[i]);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            if (values[r][c] instanceof Date) {
              values[r][c] = values[r][c].toISOString();
            }
          }
        }
        backup.master[masterSheets[i]] = values;
      }
    }
  }
  
  // Helper to extract data from a specific spreadsheet's Rojnamcha and RojnamchaDetails
  function extractLedgerData(ss) {
    var rSheet = ss.getSheetByName('Rojnamcha');
    if (rSheet) {
      var lastRow = rSheet.getLastRow();
      var lastCol = rSheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        var values = rSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            if (values[r][c] instanceof Date) {
              values[r][c] = values[r][c].toISOString();
            }
          }
          backup.rojnamcha.push(values[r]);
        }
      }
    }
    
    var dSheet = ss.getSheetByName('RojnamchaDetails');
    if (dSheet) {
      var lastRow = dSheet.getLastRow();
      var lastCol = dSheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        var values = dSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            if (values[r][c] instanceof Date) {
              values[r][c] = values[r][c].toISOString();
            }
          }
          backup.rojnamchaDetails.push(values[r]);
        }
      }
    }
  }
  
  // 2. Extract from master spreadsheet's ledger sheets
  extractLedgerData(masterSs);
  
  // 3. Extract from all year-specific spreadsheets
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  for (var key in allProps) {
    if (key.indexOf("SS_YEAR_") === 0) {
      var ssId = allProps[key];
      try {
        var ssYear = SpreadsheetApp.openById(ssId);
        extractLedgerData(ssYear);
      } catch (e) {
        console.log("Error exporting from spreadsheet ID " + ssId + ": " + e.toString());
      }
    }
  }
  
  return JSON.stringify(backup);
}

function importAllDataJSON(backupDataStr) {
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Could not obtain lock to import data. Please try again.');
  }
  
  try {
    var backup = JSON.parse(backupDataStr);
    if (!backup || !backup.master) {
      return { success: false, error: 'Invalid backup file format' };
    }
    
    // 1. Verify master sheets are in backup
    var masterSheets = ['DropdownOptions', 'CompanyProfile', 'Users'];
    for (var i = 0; i < masterSheets.length; i++) {
      if (!backup.master.hasOwnProperty(masterSheets[i])) {
        return { success: false, error: 'Backup is missing master sheet: ' + masterSheets[i] };
      }
    }
    
    // 2. Clear master sheets
    var sheetsToClear = ['DropdownOptions', 'CompanyProfile', 'Users', 'Rojnamcha', 'RojnamchaDetails'];
    for (var i = 0; i < sheetsToClear.length; i++) {
      var sheet = masterSs.getSheetByName(sheetsToClear[i]);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow > 1) {
          sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        }
      }
    }
    
    // 3. Clear all yearly spreadsheets
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    for (var key in allProps) {
      if (key.indexOf("SS_YEAR_") === 0) {
        var ssId = allProps[key];
        try {
          var ssYear = SpreadsheetApp.openById(ssId);
          var rSheet = ssYear.getSheetByName('Rojnamcha');
          if (rSheet && rSheet.getLastRow() > 1) {
            rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).clearContent();
          }
          var dSheet = ssYear.getSheetByName('RojnamchaDetails');
          if (dSheet && dSheet.getLastRow() > 1) {
            dSheet.getRange(2, 1, dSheet.getLastRow() - 1, dSheet.getLastColumn()).clearContent();
          }
        } catch (e) {
          console.log("Error clearing yearly spreadsheet ID " + ssId + ": " + e.toString());
        }
      }
    }
    
    // 4. Restore master sheets
    for (var i = 0; i < masterSheets.length; i++) {
      var sheetName = masterSheets[i];
      var rows = backup.master[sheetName];
      if (rows && rows.length > 0) {
        var sheet = masterSs.getSheetByName(sheetName);
        if (sheet) {
          for (var r = 0; r < rows.length; r++) {
            for (var c = 0; c < rows[r].length; c++) {
              rows[r][c] = parseJsonValue(rows[r][c]);
            }
          }
          sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        }
      }
    }
    
    // 5. Restore Rojnamcha daily records (split by year)
    var rojnamchaRows = backup.rojnamcha || [];
    var rojnamchaDetailsRows = backup.rojnamchaDetails || [];
    
    var rojnamchaByYear = {};
    var detailsByYear = {};
    
    function getYearFromRow(row) {
      var dateVal = row[0];
      if (!dateVal) return new Date().getFullYear().toString();
      if (dateVal instanceof Date) {
        return dateVal.getFullYear().toString();
      }
      var dateStr = dateVal.toString();
      if (dateStr.indexOf('-') > -1) {
        var parts = dateStr.split('T')[0].split('-');
        return parts[0];
      }
      var d = new Date(dateStr);
      if (!isNaN(d.getTime())) {
        return d.getFullYear().toString();
      }
      return new Date().getFullYear().toString();
    }
    
    // Group Rojnamcha rows by year
    for (var r = 0; r < rojnamchaRows.length; r++) {
      var row = rojnamchaRows[r];
      for (var c = 0; c < row.length; c++) {
        row[c] = parseJsonValue(row[c]);
      }
      var year = getYearFromRow(row);
      if (!rojnamchaByYear[year]) {
        rojnamchaByYear[year] = [];
      }
      rojnamchaByYear[year].push(row);
    }
    
    // Group Details rows by year
    for (var r = 0; r < rojnamchaDetailsRows.length; r++) {
      var row = rojnamchaDetailsRows[r];
      for (var c = 0; c < row.length; c++) {
        row[c] = parseJsonValue(row[c]);
      }
      var year = getYearFromRow(row);
      if (!detailsByYear[year]) {
        detailsByYear[year] = [];
      }
      detailsByYear[year].push(row);
    }
    
    // Write Rojnamcha rows to corresponding yearly spreadsheets
    for (var year in rojnamchaByYear) {
      var ssYear = getSpreadsheetForYear(year);
      var sheet = ssYear.getSheetByName('Rojnamcha');
      if (sheet) {
        var rows = rojnamchaByYear[year];
        if (sheet.getMaxColumns() < rows[0].length) {
          sheet.insertColumnsAfter(sheet.getMaxColumns(), rows[0].length - sheet.getMaxColumns());
        }
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
    }
    
    // Write Details rows to corresponding yearly spreadsheets
    for (var year in detailsByYear) {
      var ssYear = getSpreadsheetForYear(year);
      var sheet = ssYear.getSheetByName('RojnamchaDetails');
      if (sheet) {
        var rows = detailsByYear[year];
        if (sheet.getMaxColumns() < rows[0].length) {
          sheet.insertColumnsAfter(sheet.getMaxColumns(), rows[0].length - sheet.getMaxColumns());
        }
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
    }
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() + "\n" + err.stack };
  } finally {
    lock.releaseLock();
  }
}


function getYearFromDateStr(dateStr) {
  if (!dateStr) return new Date().getFullYear().toString();
  var parts = dateStr.split('-');
  if (parts.length > 0) {
    return parts[0];
  }
  return new Date().getFullYear().toString();
}

function getSpreadsheetForYear(yearStr) {
  var propKey = "SS_YEAR_" + yearStr;
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(propKey);
  
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      console.log("Failed to open spreadsheet by ID: " + ssId + ". Re-creating...");
    }
  }
  
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var masterFile = DriveApp.getFileById(masterSs.getId());
  
  var newFileName = "Rojnamcha_Ledger_" + yearStr;
  var newFile = masterFile.makeCopy(newFileName);
  var newSs = SpreadsheetApp.openById(newFile.getId());
  
  var sheetNamesToClear = ['Rojnamcha', 'RojnamchaDetails'];
  for (var i = 0; i < sheetNamesToClear.length; i++) {
    var sheet = newSs.getSheetByName(sheetNamesToClear[i]);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    }
  }
  
  props.setProperty(propKey, newSs.getId());
  return newSs;
}
