function setupPortal() {
  const ss = SpreadsheetApp.create('Wilmington Brewing Runner Portal Database');

  const tabs = [
    'Settings',
    'Members',
    'Attendance',
    'Rewards',
    'Events',
    'Email Log'
  ];

  tabs.forEach((name, index) => {
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
  Logger.log('Spreadsheet ID saved: ' + ss.getId());
}

function setupSettings_(ss) {
  const sheet = ss.getSheetByName('Settings');

  sheet.getRange('A1:B12').setValues([
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

  sheet.autoResizeColumns(1, 2);
}

function setupMembers_(ss) {
  ss.getSheetByName('Members').getRange('A1:O1').setValues([[
    'Runner ID',
    'First Name',
    'Last Name',
    'Email',
    'Phone',
    'Shirt Size',
    'Emergency Contact Name',
    'Emergency Contact Phone',
    'Waiver Accepted',
    'Waiver Timestamp',
    'Join Date',
    'Fist Bumps',
    'QR Code URL',
    'Status',
    'Notes'
  ]]);
}

function setupAttendance_(ss) {
  ss.getSheetByName('Attendance').getRange('A1:I1').setValues([[
    'Timestamp',
    'Event Name',
    'Runner ID',
    'First Name',
    'Last Name',
    'Fist Bumps Earned',
    'Checked In By',
    'Duplicate?',
    'Notes'
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
    'Event ID',
    'Event Name',
    'Date',
    'Location',
    'Active',
    'Notes'
  ]]);
}

function setupEmailLog_(ss) {
  ss.getSheetByName('Email Log').getRange('A1:E1').setValues([[
    'Timestamp',
    'Email',
    'Subject',
    'Status',
    'Notes'
  ]]);
}

function getDatabase_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('No spreadsheet ID found. Run setupPortal first.');
  }
  return SpreadsheetApp.openById(id);
}
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
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

  const existingEmails = members.getRange('D2:D').getValues().flat();
  if (existingEmails.includes(data.email)) {
    throw new Error('This email is already registered.');
  }

  const nextRow = members.getLastRow() + 1;
  const runnerId = 'WBRC-' + String(nextRow - 1).padStart(4, '0');
  const qrCodeUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(runnerId) + '&size=300';

  members.appendRow([
    runnerId,
    data.firstName,
    data.lastName,
    data.email,
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

  return {
    runnerId: runnerId,
    qrCodeUrl: qrCodeUrl
  };
}
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
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

  const emails = members.getRange('D2:D').getValues().flat().filter(String);
  if (emails.includes(data.email)) {
    throw new Error('This email is already registered.');
  }

  const runnerNumber = members.getLastRow();
  const runnerId = 'WBRC-' + String(runnerNumber).padStart(4, '0');
  const qrCodeUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(runnerId) + '&size=300';

  members.appendRow([
    runnerId,
    data.firstName,
    data.lastName,
    data.email,
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

  return {
    runnerId: runnerId,
    qrCodeUrl: qrCodeUrl
  };
}