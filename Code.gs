const PORTAL = {
  clubName: 'Wilmington Brewing Co. Run Club',
  portalName: 'Wilmington Brewing Runner Portal',
  idPrefix: 'WBRC',
  primary: '#7A263A',
  snowbound: '#F7F5F0',
  charcoal: '#2F2F2F',
  waiver: `By signing above, I warrant that I am an adult over eighteen (18) years of age and competent to execute this release and indemnity agreement. In consideration of participation in Run Cub (running/walking group) I release and hold harmless Wilmington Brewing Company, LLC and/or Wilmington Homebrew Supply Corp., their respective officers, shareholders, members, managers, and employees (the "Sponsor Companies") from any and all liability, claims, and demands of whatever kind of nature, either in law or in equity, which may arise from my volunteer participation in Run Club. I execute this agreement knowingly, having read and understood the provisions hereof. My heirs, successors, assigns, representatives, and I hereby release and shall indemnify (including for each breach of warranty by me), defend and hold harmless the Sponsor Companies and their respective members, managers, representatives, agents, and employees (individually and collectively), from and against any and all claims of any nature, whether sounding in tort, statute, contract or otherwise, costs and expenses (including attorneys' fees), demands, suits or judgments brought by me or any third party claiming by , on behalf of, or through me, against the Sponsoring Companies or any third party relating to my participation in Run Club.\n\nThis Agreement shall be construed and enforced under the laws of the State of North Carolina, excepting conflicts of laws principles thereof. The invalidity of any provision hereof shall not affect the validity of the remaining provisions`
};

function setupPortal() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create(PORTAL.portalName);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheet_(ss, 'Settings', [['Setting','Value'],['Club Name',PORTAL.clubName],['Portal Name',PORTAL.portalName],['Runner ID Prefix',PORTAL.idPrefix],['Reward Currency','Fist Bumps'],['Primary Color',PORTAL.primary],['Background Color',PORTAL.snowbound],['Logo URL',''],['Web App URL',''],['Waiver Text',PORTAL.waiver]]);
  ensureSheet_(ss, 'Members', [['Runner ID','First Name','Last Name','Email','Phone','Shirt Size','Emergency Contact Name','Emergency Contact Phone','Waiver Accepted','Waiver Timestamp','Join Date','Fist Bumps','QR Code URL','Status']]);
  ensureSheet_(ss, 'Attendance', [['Timestamp','Runner ID','First Name','Last Name','Event Name','Fist Bumps Earned','Checked In By','Notes']]);
  ensureSheet_(ss, 'Rewards', [['Fist Bumps Needed','Reward Name','Active'],[5,'TBD','Yes'],[10,'TBD','Yes'],[20,'TBD','Yes'],[35,'TBD','Yes'],[50,'TBD','Yes'],[75,'TBD','Yes'],[100,'TBD','Yes']]);
  ensureSheet_(ss, 'Events', [['Event Name','Event Date','Location','Active','Notes'],['Weekly Run',new Date(),'Wilmington Brewing Company','Yes','Default active event']]);
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Next: Deploy as Web App, then run setWebAppUrlFromPrompt.');
}

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.page = e && e.parameter && e.parameter.page ? e.parameter.page : 'home';
  return t.evaluate().setTitle(PORTAL.portalName).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function setWebAppUrlFromPrompt() {
  const ui = SpreadsheetApp.getUi ? null : null;
  const url = Browser.inputBox('Paste the Web App URL');
  setSetting_('Web App URL', url);
  Logger.log('Saved Web App URL: ' + url);
}

function getAppData() {
  return { settings: getSettings_(), activeEvent: getActiveEvent_(), rewards: getRewards_() };
}

function registerMember(form) {
  if (!form.waiverAccepted) throw new Error('Waiver agreement is required.');
  const ss = getSs_();
  const members = ss.getSheetByName('Members');
  const rows = getRows_(members);
  const email = String(form.email || '').trim().toLowerCase();
  if (!email) throw new Error('Email is required.');
  const existing = rows.find(r => String(r['Email']).toLowerCase() === email);
  if (existing) return buildMemberResponse_(existing['Runner ID']);
  const runnerId = nextRunnerId_(rows.length);
  const qrUrl = createQrUrl_(runnerId);
  members.appendRow([runnerId, form.firstName, form.lastName, email, form.phone || '', form.shirtSize || '', form.emergencyName || '', form.emergencyPhone || '', 'Yes', new Date(), new Date(), 0, qrUrl, 'Active']);
  sendWelcomeEmail_(email, form.firstName, runnerId, qrUrl);
  return buildMemberResponse_(runnerId);
}

function lookupMember(query) {
  const q = String(query || '').trim().toLowerCase();
  const rows = getRows_(getSs_().getSheetByName('Members'));
  const m = rows.find(r => String(r['Runner ID']).toLowerCase() === q || String(r['Email']).toLowerCase() === q);
  if (!m) throw new Error('Member not found.');
  return buildMemberResponse_(m['Runner ID']);
}

function checkInMember(runnerId, notes) {
  const ss = getSs_();
  const member = getMemberById_(runnerId);
  if (!member) throw new Error('Runner ID not found.');
  const event = getActiveEvent_();
  if (!event) throw new Error('No active event. Create one in Admin first.');
  const attendance = ss.getSheetByName('Attendance');
  const existing = getRows_(attendance).find(r => String(r['Runner ID']) === String(runnerId) && String(r['Event Name']) === String(event.name));
  if (existing) throw new Error('Already checked in for this event.');
  attendance.appendRow([new Date(), runnerId, member['First Name'], member['Last Name'], event.name, 1, Session.getActiveUser().getEmail() || 'Volunteer', notes || '']);
  recalcMemberFistBumps_(runnerId);
  return buildMemberResponse_(runnerId);
}

function adminCreateEvent(name, location, notes) {
  const ss = getSs_();
  const events = ss.getSheetByName('Events');
  const data = events.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) events.getRange(i+1,4).setValue('No');
  events.appendRow([name || 'Run Club Event', new Date(), location || 'Wilmington Brewing Company', 'Yes', notes || '']);
  return getActiveEvent_();
}

function getAdminStats() {
  const ss = getSs_();
  const members = getRows_(ss.getSheetByName('Members'));
  const attendance = getRows_(ss.getSheetByName('Attendance'));
  return { totalMembers: members.length, totalCheckIns: attendance.length, activeEvent: getActiveEvent_(), leaderboard: members.sort((a,b)=>Number(b['Fist Bumps']||0)-Number(a['Fist Bumps']||0)).slice(0,10) };
}

function ensureSheet_(ss, name, values) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,values.length,values[0].length).setValues(values);
  sh.autoResizeColumns(1, Math.max(1, sh.getLastColumn()));
  return sh;
}
function getSs_() { const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); if (!id) throw new Error('Run setupPortal first.'); return SpreadsheetApp.openById(id); }
function getRows_(sh) { const v = sh.getDataRange().getValues(); if (v.length < 2) return []; const h = v[0]; return v.slice(1).filter(r => r.some(Boolean)).map(r => Object.fromEntries(h.map((k,i)=>[k,r[i]]))); }
function getSettings_() { const rows = getRows_(getSs_().getSheetByName('Settings')); return Object.fromEntries(rows.map(r => [r.Setting, r.Value])); }
function setSetting_(name, value) { const sh = getSs_().getSheetByName('Settings'); const data = sh.getDataRange().getValues(); for (let i=1;i<data.length;i++) if (data[i][0]===name) { sh.getRange(i+1,2).setValue(value); return; } sh.appendRow([name,value]); }
function nextRunnerId_(count) { return PORTAL.idPrefix + '-' + String(count + 1).padStart(4,'0'); }
function createQrUrl_(runnerId) { const web = getSettings_()['Web App URL'] || ''; const payload = web ? web + '?page=volunteer&runner=' + encodeURIComponent(runnerId) : runnerId; return 'https://quickchart.io/qr?size=300&text=' + encodeURIComponent(payload); }
function getMemberById_(runnerId) { return getRows_(getSs_().getSheetByName('Members')).find(r => String(r['Runner ID']) === String(runnerId)); }
function getRewards_() { return getRows_(getSs_().getSheetByName('Rewards')).filter(r => r.Active === 'Yes'); }
function getActiveEvent_() { const e = getRows_(getSs_().getSheetByName('Events')).find(r => r.Active === 'Yes'); return e ? { name:e['Event Name'], date:e['Event Date'], location:e.Location, notes:e.Notes } : null; }
function recalcMemberFistBumps_(runnerId) { const ss=getSs_(); const total=getRows_(ss.getSheetByName('Attendance')).filter(r=>String(r['Runner ID'])===String(runnerId)).reduce((s,r)=>s+Number(r['Fist Bumps Earned']||0),0); const sh=ss.getSheetByName('Members'); const ids=sh.getRange('A2:A').getValues().flat(); const idx=ids.findIndex(id=>String(id)===String(runnerId)); if(idx>=0) sh.getRange(idx+2,12).setValue(total); }
function buildMemberResponse_(runnerId) { const m=getMemberById_(runnerId); const rewards=getRewards_(); const bumps=Number(m['Fist Bumps']||0); const next=rewards.find(r=>Number(r['Fist Bumps Needed'])>bumps) || null; return { runnerId:m['Runner ID'], firstName:m['First Name'], lastName:m['Last Name'], email:m.Email, fistBumps:bumps, qrUrl:m['QR Code URL'], nextReward: next ? {needed:Number(next['Fist Bumps Needed']), name:next['Reward Name'], remaining:Number(next['Fist Bumps Needed'])-bumps} : null }; }
function sendWelcomeEmail_(email, firstName, runnerId, qrUrl) { MailApp.sendEmail(email, 'Welcome to Wilmington Brewing Co. Run Club', `Hi ${firstName},\n\nWelcome to Wilmington Brewing Co. Run Club!\n\nYour Runner ID is ${runnerId}.\n\nYour personal QR code: ${qrUrl}\n\nShow this QR code at each run to earn Fist Bumps.\n\nCheers,\nWilmington Brewing Runner Portal`); }
