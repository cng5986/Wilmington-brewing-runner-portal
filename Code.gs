function setupPortal() {
  const ss = SpreadsheetApp.create('Wilmington Brewing Runner Portal Database');

  ['Settings', 'Members', 'Attendance', 'Rewards', 'Events', 'Email Log']
    .forEach((name, index) => {
      const sheet = index === 0 ? ss.getActiveSheet() : ss.insertSheet();
      sheet.setName(name);
    });

  setupSettings_(ss);
  setupMembers_(ss);
  setupAttendance_(ss);
  setupRewards_(ss);
  setupEvents_(ss);
  setupEmailLog_(ss);

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Logger.log('Portal database created: ' + ss.getUrl());
}

function setupSettings_(ss) {
  ss.getSheetByName('Settings').getRange('A1:B12').setValues([
    ['Setting', 'Value'],
    ['Portal Name', 'Wilmington Brewing Runner Portal'],
    ['Club Name', 'Wilmington Brewing Co. Run Club'],
    ['Runner ID Prefix', 'WBRC'],
    ['Reward Name', 'Fist Bumps'],
    ['Primary Color', '#7A263A'],
    ['Background Color', '#F7F5F0'],
    ['Logo URL', ''],
    ['Check-In Mode', 'Volunteer scans member QR codes'],
    ['Default Fist Bumps Per Check-In', 1],
    ['Rewards Editable', 'Yes'],
    ['Created At', new Date()]
  ]);
}

function setupMembers_(ss) {
  ss.getSheetByName('Members').getRange('A1:O1').setValues([[
    'Runner ID', 'First Name', 'Last Name', 'Email', 'Phone',
    'Shirt Size', 'Emergency Contact Name', 'Emergency Contact Phone',
    'Waiver Accepted', 'Waiver Timestamp', 'Join Date', 'Fist Bumps',
    'QR Code URL', 'Status', 'Notes'
  ]]);
}

function setupAttendance_(ss) {
  ss.getSheetByName('Attendance').getRange('A1:I1').setValues([[
    'Timestamp', 'Event Name', 'Runner ID', 'First Name', 'Last Name',
    'Fist Bumps Earned', 'Checked In By', 'Duplicate?', 'Notes'
  ]]);
}

function setupRewards_(ss) {
  ss.getSheetByName('Rewards').getRange('A1:C8').setValues([
    ['Fist Bumps Needed', 'Reward', 'Active'],
    [5, 'TBD', 'Yes'],
    [10, 'TBD', 'Yes'],
    [20, 'TBD', 'Yes'],
    [35, 'TBD', 'Yes'],
    [50, 'TBD', 'Yes'],
    [75, 'TBD', 'Yes'],
    [100, 'TBD', 'Yes']
  ]);
}

function setupEvents_(ss) {
  ss.getSheetByName('Events').getRange('A1:F1').setValues([[
    'Event ID', 'Event Name', 'Date', 'Location', 'Active', 'Notes'
  ]]);
}

function setupEmailLog_(ss) {
  ss.getSheetByName('Email Log').getRange('A1:E1').setValues([[
    'Timestamp', 'Email', 'Subject', 'Status', 'Notes'
  ]]);
}

function getDatabase_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('No spreadsheet ID found. Run setupPortal first.');
  return SpreadsheetApp.openById(id);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.checkinId = e && e.parameter && e.parameter.checkin
    ? e.parameter.checkin
    : '';

  return template
    .evaluate()
    .setTitle('Wilmington Brewing Runner Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function registerMember(data) {
  if (!data.firstName || !data.lastName || !data.email || !data.waiverAccepted) {
    throw new Error('Missing required registration fields.');
  }

  const ss = getDatabase_();
  const members = ss.getSheetByName('Members');

  const email = String(data.email).trim().toLowerCase();
  const emails = members.getRange('D2:D').getValues().flat().map(e => String(e).trim().toLowerCase());

  if (emails.includes(email)) {
    throw new Error('This email is already registered.');
  }

  const runnerNumber = members.getLastRow();
  const runnerId = 'WBRC-' + String(runnerNumber).padStart(4, '0');
  const webAppUrl = 'https://script.google.com/macros/s/AKfycbw39le-POb1vpuYYlj9DHSUqk5zO3bscAAd8WSoF3I2ddI0O90uDuXISIK0ShKVrAnX/exec';
const checkInUrl = webAppUrl + '?checkin=' + encodeURIComponent(runnerId);
const qrCodeUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(checkInUrl) + '&size=300';

  members.appendRow([
    runnerId,
    data.firstName,
    data.lastName,
    email,
    data.phone || '',
    data.shirtSize || '',
    data.emergencyName || '',
    data.emergencyPhone || '',
    'Yes',
    new Date(),
    new Date(),
    0,
    qrCodeUrl,
    'Active',
    ''
  ]);
sendWelcomeEmail_(email, data.firstName, data.lastName, runnerId, qrCodeUrl);
  return {
    runnerId,
    qrCodeUrl
  };
}

function getMemberByLookup(lookup) {
  const ss = getDatabase_();
  const members = ss.getSheetByName('Members');
  const rows = members.getDataRange().getValues();
  const search = String(lookup).trim().toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const runnerId = String(row[0]).trim().toLowerCase();
    const email = String(row[3]).trim().toLowerCase();

    if (runnerId === search || email === search) {
      return {
        runnerId: row[0],
        firstName: row[1],
        lastName: row[2],
        email: row[3],
        fistBumps: row[11] || 0,
        qrCodeUrl: row[12],
        status: row[13]
      };
    }
  }

  throw new Error('No member found. Check the Runner ID or email.');
}
function checkInMember(runnerIdInput, notes) {
  const ss = getDatabase_();
  const members = ss.getSheetByName('Members');
  const attendance = ss.getSheetByName('Attendance');

  const rows = members.getDataRange().getValues();
  const search = String(runnerIdInput).trim().toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const runnerId = String(row[0]).trim().toLowerCase();

    if (runnerId === search) {
      const today = new Date();
      const todayKey = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const attendanceRows = attendance.getDataRange().getValues();

      for (let j = 1; j < attendanceRows.length; j++) {
        const attendanceDate = attendanceRows[j][0];
        const attendanceRunnerId = String(attendanceRows[j][2]).trim().toLowerCase();

        if (attendanceDate && attendanceRunnerId === search) {
          const attendanceKey = Utilities.formatDate(
            new Date(attendanceDate),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          );

          if (attendanceKey === todayKey) {
            return {
              duplicate: true,
              runnerId: row[0],
              firstName: row[1],
              lastName: row[2],
              currentTotal: row[11] || 0
            };
          }
        }
      }

      const oldTotal = Number(row[11]) || 0;
      const newTotal = oldTotal + 1;

      members.getRange(i + 1, 12).setValue(newTotal);

      attendance.appendRow([
        new Date(),
        'Weekly Run',
        row[0],
        row[1],
        row[2],
        1,
        Session.getActiveUser().getEmail() || 'Volunteer',
        'No',
        notes || ''
      ]);

      return {
        duplicate: false,
        runnerId: row[0],
        firstName: row[1],
        lastName: row[2],
        oldTotal,
        newTotal
      };
    }
  }

  throw new Error('Runner ID not found.');
}
function sendWelcomeEmail_(email, firstName, lastName, runnerId, qrCodeUrl) {
  const subject = 'Welcome to Wilmington Brewing Co. Run Club';

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;background:#f7f5f0;padding:24px;">' +
      '<div style="max-width:520px;margin:auto;background:#7a263a;color:white;border-radius:20px;padding:28px;text-align:center;">' +
        '<h2 style="margin-top:0;">Wilmington Brewing Co. Run Club</h2>' +
        '<h3>Welcome, ' + firstName + '!</h3>' +
        '<p style="font-size:26px;font-weight:bold;">' + runnerId + '</p>' +
        '<p>Save this QR code and bring it to every run.</p>' +
        '<img src="' + qrCodeUrl + '" style="background:white;padding:12px;border-radius:14px;width:220px;">' +
        '<p style="font-size:18px;"><strong>👊 0 Fist Bumps</strong></p>' +
        '<p>Next reward: TBD</p>' +
        '<p>Run • Walk • Jog • Hang Out • Earn Fist Bumps</p>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}
function getAdminStats() {
  const ss = getDatabase_();

  const members = ss.getSheetByName('Members');
  const attendance = ss.getSheetByName('Attendance');

  const totalMembers = Math.max(members.getLastRow() - 1, 0);
  const totalCheckIns = Math.max(attendance.getLastRow() - 1, 0);

  let totalFistBumps = 0;
  if (members.getLastRow() > 1) {
    const values = members.getRange(2, 12, members.getLastRow() - 1, 1).getValues();
    totalFistBumps = values.reduce(function(sum, row) {
      return sum + (Number(row[0]) || 0);
    }, 0);
  }
const ADMIN_PASSWORD = 'wbrc-admin';

function verifyAdminPassword(password) {
  return {
    allowed: password === ADMIN_PASSWORD
  };
}
  return {
    totalMembers,
    totalCheckIns,
    totalFistBumps
  };
}
function verifyAdminPassword(password) {
  if (String(password).trim() === 'wbrc-admin') {
    return { allowed: true };
  }

  return { allowed: false };
}