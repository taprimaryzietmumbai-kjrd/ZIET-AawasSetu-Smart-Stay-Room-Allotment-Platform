/**
 * ZIET AawasSetu — Google Apps Script backend
 * Install in a NEW standalone Apps Script project.
 * This script creates a separate AawasSetu Master Database.
 * It never accesses the PRTM Music / Nameplate spreadsheet.
 *
 * SETUP:
 * 1. Change INITIAL_ADMIN_PIN below to a private PIN (at least 8 digits).
 * 2. Run setupAawasSetu() once and authorize.
 * 3. Run showAawasSetuDatabaseUrl() to see the new database URL.
 * 4. Deploy > New deployment > Web app; Execute as Me; access as Anyone.
 * 5. Copy the Web app URL into the AawasSetu frontend configuration.
 */

const INITIAL_ADMIN_PIN = 'CHANGE-THIS-TO-A-PRIVATE-8-DIGIT-OR-LONGER-PIN';
const DB_PROPERTY = 'AWAASSETU_DATABASE_ID';
const PIN_HASH_PROPERTY = 'AWAASSETU_ADMIN_PIN_SHA256';
const TOKEN_PROPERTY_PREFIX = 'AWAASSETU_TOKEN_';

const SHEETS = {
  Rooms: ['Building','Floor','Room','Beds','AC','Geyser','Attached Washroom','Facility','Allotable','Status'],
  Participants: ['Participant ID','Name','Designation','School','Region','Gender','DOB','Employee Code','Role','Preference','Special Requirement','Verification Status','Source'],
  Allotments: ['Allotment ID','Participant ID','Participant Name','Building','Floor','Room','Bed','Allotment Status','Allotted At','Updated At','Remarks'],
  Verification: ['Record ID','Participant ID','Participant Name','Field','Submitted Value','Suggested Value','Reason','Review Status','Reviewed By','Reviewed At'],
  AuditLog: ['Timestamp','Actor','Action','Record Type','Record ID','Details']
};

function setupAawasSetu() {
  if (!INITIAL_ADMIN_PIN || INITIAL_ADMIN_PIN.indexOf('CHANGE-THIS') === 0 || INITIAL_ADMIN_PIN.length < 8) {
    throw new Error('First set a private INITIAL_ADMIN_PIN (minimum 8 characters) in this script.');
  }
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(DB_PROPERTY);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('ZIET AawasSetu — Master Database');
    id = ss.getId();
    props.setProperty(DB_PROPERTY, id);
  }
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,SHEETS[name].length).setFontWeight('bold').setBackground('#dceeff');
    }
  });
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) ss.deleteSheet(defaultSheet);
  props.setProperty(PIN_HASH_PROPERTY, sha256_(INITIAL_ADMIN_PIN));
  seedRooms_(ss);
  audit_(ss, 'SYSTEM', 'INITIALIZE', 'DATABASE', id, 'AawasSetu database initialized');
  Logger.log('AawasSetu database: ' + ss.getUrl());
  Logger.log('Database ID: ' + id);
  return ss.getUrl();
}

function showAawasSetuDatabaseUrl() {
  const id = PropertiesService.getScriptProperties().getProperty(DB_PROPERTY);
  if (!id) throw new Error('Run setupAawasSetu() first.');
  Logger.log(SpreadsheetApp.openById(id).getUrl());
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'health');
  let result;
  try {
    if (action === 'health') {
      const id = PropertiesService.getScriptProperties().getProperty(DB_PROPERTY);
      result = {success: !!id, service: 'ZIET AawasSetu API', databaseConnected: !!id, timestamp: new Date().toISOString()};
    } else if (action === 'summary') {
      const auth = authorize_(e.parameter.token);
      result = summary_(db_());
    } else if (action === 'rooms') {
      authorize_(e.parameter.token);
      result = {success:true, rooms: readObjects_(db_().getSheetByName('Rooms'))};
    } else if (action === 'participants') {
      authorize_(e.parameter.token);
      result = {success:true, participants: readObjects_(db_().getSheetByName('Participants'))};
    } else {
      result = {success:false, error:'Unknown action'};
    }
  } catch (err) {
    result = {success:false, error:err.message};
  }
  return jsonp_(result, e && e.parameter && e.parameter.callback);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (_) { return json_({success:false,error:'Invalid JSON'}); }
  try {
    const action = String(body.action || '');
    if (action === 'login') {
      const expected = PropertiesService.getScriptProperties().getProperty(PIN_HASH_PROPERTY);
      if (!expected || sha256_(String(body.pin || '')) !== expected) throw new Error('Invalid administrator PIN');
      const token = Utilities.getUuid() + Utilities.getUuid();
      CacheService.getScriptCache().put(TOKEN_PROPERTY_PREFIX + token, 'admin', 21600);
      return json_({success:true, token:token, expiresIn:21600});
    }
    authorize_(body.token);
    if (action === 'saveParticipant') {
      const p = body.participant || {};
      if (!String(p.name || '').trim()) throw new Error('Participant name is required');
      const sh = db_().getSheetByName('Participants');
      const id = String(p.id || Utilities.getUuid());
      upsert_(sh, 'Participant ID', id, {
        'Participant ID':id,'Name':p.name || '', 'Designation':p.designation || '',
        'School':p.school || '', 'Region':p.region || '', 'Gender':p.gender || '',
        'DOB':p.dob || '', 'Employee Code':p.employeeCode || '', 'Role':p.role || 'Participant',
        'Preference':p.preference || '', 'Special Requirement':p.specialRequirement || '',
        'Verification Status':p.verificationStatus || 'Pending','Source':p.source || 'Admin entry'
      });
      audit_(db_(), 'ADMIN', 'SAVE', 'PARTICIPANT', id, p.name);
      return json_({success:true,id:id});
    }
    if (action === 'saveAllotment') {
      const a = body.allotment || {};
      ['participantId','participantName','building','floor','room','bed'].forEach(k => { if (!String(a[k] || '').trim()) throw new Error('Missing allotment field: ' + k); });
      const sh = db_().getSheetByName('Allotments');
      const id = String(a.id || Utilities.getUuid());
      const rows = readObjects_(sh);
      const duplicate = rows.some(r => String(r['Allotment ID']) !== id && String(r['Building'])===String(a.building) && String(r['Floor'])===String(a.floor) && String(r['Room'])===String(a.room) && String(r['Bed'])===String(a.bed) && String(r['Allotment Status'])==='Active');
      if (duplicate) throw new Error('This bed is already occupied.');
      upsert_(sh, 'Allotment ID', id, {'Allotment ID':id,'Participant ID':a.participantId,'Participant Name':a.participantName,'Building':a.building,'Floor':a.floor,'Room':a.room,'Bed':a.bed,'Allotment Status':'Active','Allotted At':new Date(),'Updated At':new Date(),'Remarks':a.remarks || ''});
      audit_(db_(), 'ADMIN', 'SAVE', 'ALLOTMENT', id, a.participantName);
      return json_({success:true,id:id});
    }
    if (action === 'releaseAllotment') {
      const id = String(body.id || '');
      const sh = db_().getSheetByName('Allotments');
      const rows = sh.getDataRange().getValues();
      const headers = rows[0];
      const idCol = headers.indexOf('Allotment ID');
      const statusCol = headers.indexOf('Allotment Status');
      const updatedCol = headers.indexOf('Updated At');
      for (let i=1;i<rows.length;i++) if (String(rows[i][idCol])===id) {
        sh.getRange(i+1,statusCol+1).setValue('Released');
        sh.getRange(i+1,updatedCol+1).setValue(new Date());
        audit_(db_(),'ADMIN','RELEASE','ALLOTMENT',id,'Released by administrator');
        return json_({success:true});
      }
      throw new Error('Allotment record not found');
    }
    throw new Error('Unknown action');
  } catch (err) {
    return json_({success:false,error:err.message});
  }
}

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(DB_PROPERTY);
  if (!id) throw new Error('Database not initialized. Run setupAawasSetu().');
  return SpreadsheetApp.openById(id);
}
function authorize_(token) {
  if (!token || CacheService.getScriptCache().get(TOKEN_PROPERTY_PREFIX + token) !== 'admin') throw new Error('Administrator session expired or invalid. Please sign in again.');
  return true;
}
function seedRooms_(ss) {
  const sh = ss.getSheetByName('Rooms');
  if (sh.getLastRow() > 1) return;
  const rows = [];
  ['A1','A2','A3'].forEach(building => {
    ['Ground','First','Second','Third'].forEach(floor => {
      [1,2,3].forEach(room => {
        const excluded = building==='A1' && floor==='Ground' && room===1;
        rows.push([building,floor,room,3,'Yes','Yes','Yes',floor==='Second'?'RO purifier available':'',excluded?'No':'Yes',excluded?'Store Room cum Guard Room':'Available']);
      });
    });
  });
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}
function summary_(ss) {
  const rooms = readObjects_(ss.getSheetByName('Rooms'));
  const people = readObjects_(ss.getSheetByName('Participants'));
  const allotments = readObjects_(ss.getSheetByName('Allotments')).filter(r => String(r['Allotment Status'])==='Active');
  return {success:true,capacity:{buildings:3,floors:12,rooms:35,beds:105},participants:people.length,activeAllotments:allotments.length,occupiedBeds:allotments.length,availableBeds:105-allotments.length,rooms:rooms};
}
function readObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o={}; headers.forEach((h,i)=>o[h]=r[i]); return o;
  });
}
function upsert_(sh, keyHeader, key, obj) {
  const values=sh.getDataRange().getValues(), headers=values[0].map(String), keyCol=headers.indexOf(keyHeader);
  let row=-1;
  for(let i=1;i<values.length;i++) if(String(values[i][keyCol])===String(key)){row=i+1;break;}
  const output=headers.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:(row>0?values[row-1][headers.indexOf(h)]:''));
  if(row>0) sh.getRange(row,1,1,headers.length).setValues([output]);
  else sh.appendRow(output);
}
function audit_(ss, actor, action, type, id, details) {
  ss.getSheetByName('AuditLog').appendRow([new Date(),actor,action,type,id,details]);
}
function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function jsonp_(o, callback) {
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback+'('+JSON.stringify(o)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(o);
}
