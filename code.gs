/**
 * Huiswerk & Toetsen App - Backend (Code.gs)
 * Ondersteunt registreren, inloggen, uitloggen, klassen, huiswerk en toetsen.
 */

// Opent de app als Web App (via URL)
function doGet() {
  return getHtmlOutput_()
    .setTitle('Huiswerk & Toetsen App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Voegt een menu toe als het script aan een Google Spreadsheet gekoppeld is
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📚 Huiswerk App')
      .addItem('App openen', 'showDialog')
      .addToUi();
  } catch (e) {
    // Alleen actief wanneer gekoppeld aan een Google Spreadsheet
  }
}

// Opent de app als pop-up dialoogvenster vanuit Google Spreadsheets
function showDialog() {
  var html = getHtmlOutput_().setWidth(950).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Huiswerk & Toetsen App');
}

function getHtmlOutput_() {
  var possibleNames = ['Dialog', 'Dialog (1) (1)', 'Index'];
  for (var i = 0; i < possibleNames.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(possibleNames[i]);
    } catch (e) {}
  }
  throw new Error('Geen HTML-bestand gevonden. Noem je HTML-bestand in Apps Script "Dialog" of "Index".');
}

/* ==========================================================================
   DATABASE HELPERS (GOOGLE SPREADSHEETS)
   ========================================================================== */

function getDb_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var ssId = props.getProperty('APP_SPREADSHEET_ID');
    if (ssId) {
      try {
        ss = SpreadsheetApp.openById(ssId);
      } catch (e) {
        ss = null;
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create('Huiswerk & Toetsen Database');
      props.setProperty('APP_SPREADSHEET_ID', ss.getId());
    }
  }
  return ss;
}

function getSheet_(name, headers) {
  var ss = getDb_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getRows_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function formatIsoDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).slice(0, 10);
}

function formatDisplayDate_(val) {
  var iso = formatIsoDate_(val);
  var parts = iso.split('-');
  if (parts.length === 3) {
    return parts[2] + '-' + parts[1] + '-' + parts[0];
  }
  return iso;
}

/* ==========================================================================
   ACCOUNT AANMAKEN, INLOGGEN & UITLOGGEN
   ========================================================================== */

function hashPassword_(password, salt) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ':' + password,
    Utilities.Charset.UTF_8
  );
  return raw.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

/**
 * Maakt een nieuw leerling- of docentaccount aan.
 */
function registerAccount(role, username, password, name) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    role = String(role || '').trim();
    username = String(username || '').trim();
    password = String(password || '');
    name = String(name || '').trim();

    if (role !== 'student' && role !== 'teacher') {
      return { success: false, message: 'Kies eerst of je een leerling of docent bent.' };
    }
    if (!username || !name || !password) {
      return { success: false, message: 'Vul alle velden in.' };
    }
    if (username.length < 3) {
      return { success: false, message: 'Gebruikersnaam moet minimaal 3 tekens lang zijn.' };
    }
    if (password.length < 4) {
      return { success: false, message: 'Wachtwoord moet minimaal 4 tekens lang zijn.' };
    }

    var sheet = getSheet_('Accounts', ['username', 'role', 'name', 'salt', 'passwordHash', 'createdAt']);
    var rows = getRows_(sheet);

    var exists = rows.some(function(r) {
      return String(r[0]).trim().toLowerCase() === username.toLowerCase();
    });
    if (exists) {
      return { success: false, message: 'Deze gebruikersnaam is al in gebruik.' };
    }

    var salt = Utilities.getUuid();
    var passwordHash = hashPassword_(password, salt);
    sheet.appendRow([username, role, name, salt, passwordHash, new Date().toISOString()]);

    return {
      success: true,
      message: 'Account aangemaakt! Je kunt nu inloggen.'
    };
  } catch (err) {
    return { success: false, message: 'Fout bij aanmaken: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Controleert gebruikersnaam + wachtwoord en geeft een sessietoken terug.
 */
function loginAccount(role, username, password) {
  role = String(role || '').trim();
  username = String(username || '').trim();
  password = String(password || '');

  if (!role || !username || !password) {
    return { success: false, message: 'Vul je gebruikersnaam en wachtwoord in.' };
  }

  var sheet = getSheet_('Accounts', ['username', 'role', 'name', 'salt', 'passwordHash', 'createdAt']);
  var rows = getRows_(sheet);

  var account = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === username.toLowerCase()) {
      account = {
        username: String(rows[i][0]).trim(),
        role: String(rows[i][1]).trim(),
        name: String(rows[i][2]).trim(),
        salt: String(rows[i][3]),
        passwordHash: String(rows[i][4])
      };
      break;
    }
  }

  if (!account) {
    return { success: false, message: 'Onbekende gebruikersnaam of onjuist wachtwoord.' };
  }

  if (account.role !== role) {
    var rolNaam = account.role === 'teacher' ? 'docent' : 'leerling';
    return { success: false, message: 'Dit account hoort bij de rol "' + rolNaam + '".' };
  }

  var checkHash = hashPassword_(password, account.salt);
  if (checkHash !== account.passwordHash) {
    return { success: false, message: 'Onbekende gebruikersnaam of onjuist wachtwoord.' };
  }

  var token = Utilities.getUuid();
  var sessionData = {
    username: account.username,
    role: account.role,
    name: account.name
  };

  // Sla sessie op in zowel Cache als het tabblad Sessions
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(sessionData), 21600);
  var sessSheet = getSheet_('Sessions', ['token', 'username', 'role', 'name', 'createdAt']);
  sessSheet.appendRow([token, account.username, account.role, account.name, new Date().toISOString()]);

  return {
    success: true,
    token: token,
    username: account.username,
    name: account.name
  };
}

/**
 * Verwijdert de actieve sessie bij uitloggen.
 */
function logoutAccount(token) {
  if (!token) return { success: true };
  CacheService.getScriptCache().remove('sess_' + token);

  var sessSheet = getSheet_('Sessions', ['token', 'username', 'role', 'name', 'createdAt']);
  var rows = getRows_(sessSheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === String(token)) {
      sessSheet.deleteRow(i + 2);
    }
  }
  return { success: true };
}

function getSession_(token, requiredRole) {
  if (!token) throw new Error('Niet ingelogd. Log opnieuw in.');

  var session = null;
  var cached = CacheService.getScriptCache().get('sess_' + token);
  if (cached) {
    session = JSON.parse(cached);
  } else {
    var sessSheet = getSheet_('Sessions', ['token', 'username', 'role', 'name', 'createdAt']);
    var rows = getRows_(sessSheet);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(token)) {
        session = {
          username: String(rows[i][1]),
          role: String(rows[i][2]),
          name: String(rows[i][3])
        };
        CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session), 21600);
        break;
      }
    }
  }

  if (!session) throw new Error('Sessie verlopen. Log opnieuw in.');
  if (requiredRole && session.role !== requiredRole) {
    throw new Error('Geen toegang voor deze rol.');
  }
  return session;
}

/* ==========================================================================
   KLASSEN & LEERLINGEN (DOCENT)
   ========================================================================== */

function getStudentAccounts(token) {
  getSession_(token, 'teacher');
  var sheet = getSheet_('Accounts', ['username', 'role', 'name', 'salt', 'passwordHash', 'createdAt']);
  return getRows_(sheet)
    .filter(function(r) { return String(r[1]).trim() === 'student'; })
    .map(function(r) {
      return { username: String(r[0]).trim(), name: String(r[2]).trim() };
    });
}

function getClasses(token) {
  var user = getSession_(token, 'teacher');
  var sheet = getSheet_('Classes', ['className', 'teacherUsername', 'studentsJson']);
  return getRows_(sheet)
    .filter(function(r) { return String(r[1]).trim().toLowerCase() === user.username.toLowerCase(); })
    .map(function(r) { return String(r[0]).trim(); });
}

function addClass(token, className, studentUsernames) {
  var user = getSession_(token, 'teacher');
  className = String(className || '').trim();
  if (!className) return { success: false, message: 'Vul een klasnaam in.' };

  var sheet = getSheet_('Classes', ['className', 'teacherUsername', 'studentsJson']);
  var exists = getRows_(sheet).some(function(r) {
    return String(r[0]).trim().toLowerCase() === className.toLowerCase() &&
           String(r[1]).trim().toLowerCase() === user.username.toLowerCase();
  });
  if (exists) return { success: false, message: 'Je hebt al een klas met deze naam.' };

  sheet.appendRow([className, user.username, JSON.stringify(studentUsernames || [])]);
  return { success: true };
}

function openClass(token, className) {
  var user = getSession_(token, 'teacher');
  var sheet = getSheet_('Classes', ['className', 'teacherUsername', 'studentsJson']);
  var rows = getRows_(sheet);
  var allStudents = getStudentAccounts(token);
  var studentMap = {};
  allStudents.forEach(function(s) { studentMap[s.username.toLowerCase()] = s; });

  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === String(className).trim().toLowerCase() &&
        String(rows[i][1]).trim().toLowerCase() === user.username.toLowerCase()) {
      var usernames = JSON.parse(rows[i][2] || '[]');
      var students = usernames
        .map(function(u) { return studentMap[String(u).toLowerCase()]; })
        .filter(Boolean);
      return { success: true, className: String(rows[i][0]).trim(), students: students };
    }
  }
  return { success: false, message: 'Klas niet gevonden.' };
}

function getStudentsInClass(token, className) {
  var res = openClass(token, className);
  return res.success ? res.students : [];
}

/* ==========================================================================
   HUISWERK (DOCENT & LEERLING)
   ========================================================================== */

function saveHomeworkForClass(token, data) {
  var user = getSession_(token, 'teacher');
  var students = data.students || [];
  if (!students.length) return { success: false, message: 'Selecteer minimaal één leerling.' };

  var allStudents = getStudentAccounts(token);
  var nameMap = {};
  allStudents.forEach(function(s) { nameMap[s.username.toLowerCase()] = s.name; });

  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);

  students.forEach(function(stuUser) {
    var id = 'hw_' + Utilities.getUuid();
    var stuName = nameMap[String(stuUser).toLowerCase()] || stuUser;
    sheet.appendRow([
      id, data.selectedClass, user.username, user.name,
      stuUser, stuName, data.subject, data.description, data.dueDate, 'Niet af'
    ]);
  });

  return { success: true, message: 'Huiswerk opgeslagen voor ' + students.length + ' leerling(en).' };
}

function getClassHomework(token, className) {
  var user = getSession_(token, 'teacher');
  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);

  return getRows_(sheet)
    .filter(function(r) {
      return String(r[1]).trim().toLowerCase() === String(className).trim().toLowerCase() &&
             String(r[2]).trim().toLowerCase() === user.username.toLowerCase();
    })
    .map(function(r) {
      return {
        id: String(r[0]),
        leerling: String(r[5]),
        vak: String(r[6]),
        omschrijving: String(r[7]),
        inleverdatum: formatDisplayDate_(r[8]),
        inleverdatumRaw: formatIsoDate_(r[8]),
        status: String(r[9] || 'Niet af')
      };
    });
}

function toggleHomeworkStatus(token, id, currentStatus) {
  getSession_(token, 'teacher');
  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);
  var rows = getRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      var nextStatus = currentStatus === 'Af' ? 'Niet af' : 'Af';
      sheet.getRange(i + 2, 10).setValue(nextStatus);
      return { success: true };
    }
  }
  return { success: false, message: 'Huiswerk niet gevonden.' };
}

function updateHomework(token, id, subject, description, dueDate) {
  getSession_(token, 'teacher');
  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);
  var rows = getRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 2, 7, 1, 3).setValues([[subject, description, dueDate]]);
      return { success: true };
    }
  }
  return { success: false, message: 'Huiswerk niet gevonden.' };
}

function deleteHomework(token, id) {
  getSession_(token, 'teacher');
  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);
  var rows = getRows_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, message: 'Huiswerk niet gevonden.' };
}

function getStudentHomework(token) {
  var user = getSession_(token, 'student');
  var sheet = getSheet_('Homework', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'studentUsername', 'studentName', 'subject', 'description', 'dueDate', 'status'
  ]);

  return getRows_(sheet)
    .filter(function(r) { return String(r[4]).trim().toLowerCase() === user.username.toLowerCase(); })
    .map(function(r) {
      return {
        vak: String(r[6]),
        omschrijving: String(r[7]),
        docent: String(r[3]),
        klas: String(r[1]),
        inleverdatum: formatDisplayDate_(r[8]),
        status: String(r[9]) === 'Af' ? 'Af' : 'Nog niet af'
      };
    });
}

/* ==========================================================================
   TOETSEN & CIJFERS (DOCENT & LEERLING)
   ========================================================================== */

function saveTest(token, data) {
  var user = getSession_(token, 'teacher');
  var sheet = getSheet_('Tests', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'testName', 'subject', 'testDate', 'questionsJson', 'resultsJson', 'permissionsJson'
  ]);
  var rows = getRows_(sheet);

  if (data.id) {
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 2, 5, 1, 6).setValues([[
          data.testName,
          data.subject,
          data.testDate,
          JSON.stringify(data.questions || []),
          JSON.stringify(data.results || {}),
          JSON.stringify(data.permissions || [])
        ]]);
        return { success: true };
      }
    }
  }

  var newId = 'test_' + Utilities.getUuid();
  sheet.appendRow([
    newId,
    data.selectedClass,
    user.username,
    user.name,
    data.testName,
    data.subject,
    data.testDate,
    JSON.stringify(data.questions || []),
    JSON.stringify(data.results || {}),
    JSON.stringify(data.permissions || [])
  ]);
  return { success: true };
}

function getClassTests(token, className) {
  var user = getSession_(token, 'teacher');
  var sheet = getSheet_('Tests', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'testName', 'subject', 'testDate', 'questionsJson', 'resultsJson', 'permissionsJson'
  ]);

  return getRows_(sheet)
    .filter(function(r) {
      return String(r[1]).trim().toLowerCase() === String(className).trim().toLowerCase() &&
             String(r[2]).trim().toLowerCase() === user.username.toLowerCase();
    })
    .map(function(r) {
      return {
        id: String(r[0]),
        testName: String(r[4]),
        subject: String(r[5]),
        testDate: formatDisplayDate_(r[6]),
        testDateRaw: formatIsoDate_(r[6]),
        questions: JSON.parse(r[7] || '[]'),
        results: JSON.parse(r[8] || '{}'),
        permissions: JSON.parse(r[9] || '[]')
      };
    });
}

function deleteTest(token, id) {
  getSession_(token, 'teacher');
  var sheet = getSheet_('Tests', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'testName', 'subject', 'testDate', 'questionsJson', 'resultsJson', 'permissionsJson'
  ]);
  var rows = getRows_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, message: 'Toets niet gevonden.' };
}

function calcStudentScore_(questions, studentScores) {
  var max = questions.reduce(function(s, q) { return s + Number(q.maxPoints || 0); }, 0);
  var total = 0;
  var hasScore = false;

  questions.forEach(function(q) {
    var val = studentScores ? studentScores[q.id] : '';
    if (val !== '' && val !== undefined && val !== null) {
      total += Number(val);
      hasScore = true;
    }
  });

  var percentage = max > 0 && hasScore ? Math.round((total / max) * 100) : 0;
  var grade = max > 0 && hasScore ? Math.round(((total / max) * 9 + 1) * 10) / 10 : null;
  return { total: total, max: max, percentage: percentage, grade: grade };
}

function getStudentTests(token) {
  var user = getSession_(token, 'student');
  var classSheet = getSheet_('Classes', ['className', 'teacherUsername', 'studentsJson']);
  var studentClasses = {};
  getRows_(classSheet).forEach(function(r) {
    var members = JSON.parse(r[2] || '[]').map(function(u) { return String(u).toLowerCase(); });
    if (members.indexOf(user.username.toLowerCase()) !== -1) {
      studentClasses[String(r[0]).trim().toLowerCase() + '|' + String(r[1]).trim().toLowerCase()] = true;
    }
  });

  var testSheet = getSheet_('Tests', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'testName', 'subject', 'testDate', 'questionsJson', 'resultsJson', 'permissionsJson'
  ]);

  return getRows_(testSheet)
    .filter(function(r) {
      var key = String(r[1]).trim().toLowerCase() + '|' + String(r[2]).trim().toLowerCase();
      return !!studentClasses[key];
    })
    .map(function(r) {
      var questions = JSON.parse(r[7] || '[]');
      var results = JSON.parse(r[8] || '{}');
      var permissions = JSON.parse(r[9] || '[]').map(function(u) { return String(u).toLowerCase(); });

      var myScores = results[user.username] || results[user.username.toLowerCase()] || {};
      var score = calcStudentScore_(questions, myScores);

      return {
        id: String(r[0]),
        testName: String(r[4]),
        subject: String(r[5]),
        testDate: formatDisplayDate_(r[6]),
        docent: String(r[3]),
        klas: String(r[1]),
        max: score.max,
        total: score.total,
        percentage: score.percentage,
        grade: score.grade,
        canViewTest: permissions.indexOf(user.username.toLowerCase()) !== -1
      };
    });
}

function getStudentTestDetails(token, testId) {
  var user = getSession_(token, 'student');
  var testSheet = getSheet_('Tests', [
    'id', 'className', 'teacherUsername', 'teacherName',
    'testName', 'subject', 'testDate', 'questionsJson', 'resultsJson', 'permissionsJson'
  ]);
  var rows = getRows_(testSheet);

  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(testId)) {
      var permissions = JSON.parse(rows[i][9] || '[]').map(function(u) { return String(u).toLowerCase(); });
      if (permissions.indexOf(user.username.toLowerCase()) === -1) {
        return { success: false, message: 'Je hebt geen toestemming om deze toets in te zien.' };
      }
      var questions = JSON.parse(rows[i][7] || '[]');
      var results = JSON.parse(rows[i][8] || '{}');
      var myScores = results[user.username] || results[user.username.toLowerCase()] || {};
      var score = calcStudentScore_(questions, myScores);

      return {
        success: true,
        testName: String(rows[i][4]),
        subject: String(rows[i][5]),
        testDate: formatDisplayDate_(rows[i][6]),
        score: score,
        questions: questions
      };
    }
  }
  return { success: false, message: 'Toets niet gevonden.' };
}