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

// Explicitly exclude the two Mumbai City school participants from accommodation.
function isLocalParticipant_(p) {
  const name = String(p.name || p.Name || '').toLowerCase().replace(/[^a-z]/g, '');
  const school = String(p.school || p.School || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  // Explicit local-name exceptions plus Mumbai-area school/location rules.
  return name === 'nikhilkesharwani' || name === 'sachinhpardeshi' ||
    /kendriyavidyalaya(no2|number2)colaba/.test(school) ||
    /kendriyavidyalayaiitpowai/.test(school) ||
    /(panvel|ambernath|hamla|thane|koliwada)/.test(school);
}
function isAccommodationEligible_(p) {
  const category = String(p['Participation Category'] || p.participationCategory || '').trim().toLowerCase();
  const role = String(p.Role || p.role || '').trim().toLowerCase();
  return category !== 'local participants' && role !== 'local participant' && !isLocalParticipant_(p);
}
function participantKey_(p) {
  return String(p['Participant ID'] || p.id || '').trim();
}
function normalizedPersonName_(p) {
  return String(p.Name || p.name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

const SHEETS = {
  Rooms: ['Building','Floor','Room','Beds','AC','Geyser','Attached Washroom','Facility','Allotable','Status'],
  Participants: ['Participant ID','Name','Designation','School','Region','Gender','DOB','Employee Code','Role','Preference','Special Requirement','Verification Status','Source','Participation Category'],
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
    if (sh.getLastRow() > 0) ensureHeaders_(sh, SHEETS[name]);
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
    } else if (action === 'allotments') {
      authorize_(e.parameter.token);
      result = {success:true, allotments: readObjects_(db_().getSheetByName('Allotments')).filter(r => String(r['Allotment Status'])==='Active')};
    } else if (action === 'allocationPreview') {
      authorize_(e.parameter.token);
      result = allocationPreview_(db_());
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
    if (action === 'bulkImport') {
      const list = Array.isArray(body.participants) ? body.participants : [];
      if (!list.length) throw new Error('No participant records supplied');
      if (list.length > 500) throw new Error('Import limit is 500 records per batch');
      const sh = db_().getSheetByName('Participants');
      let imported = 0, local = 0;
      list.forEach((p, i) => {
        if (!String(p.name || '').trim()) return;
        const id = String(p.id || (String(p.source || 'IMPORT') + '-' + (p.employeeCode || i+1)));
        const category = (p.participationCategory === 'Local Participants' || isLocalParticipant_(p)) ? 'Local Participants' : 'Accommodation Required';
        if (category === 'Local Participants') local++;
        upsert_(sh, 'Participant ID', id, {'Participant ID':id,'Name':p.name,'Designation':p.designation||'','School':p.school||'','Region':p.region||'','Gender':p.gender||'','DOB':p.dob||'','Employee Code':p.employeeCode||'','Role':p.role||'Participant','Preference':p.preference||'','Special Requirement':p.specialRequirement||'','Verification Status':p.verificationStatus||'Imported','Source':p.source||'Bulk workbook import','Participation Category':category});
        imported++;
      });
      audit_(db_(),'ADMIN','BULK_IMPORT','PARTICIPANTS','BATCH',JSON.stringify({imported:imported,local:local}));
      return json_({success:true,imported:imported,localParticipants:local,eligible:imported-local});
    }
    if (action === 'markLocalParticipants') {
      const sh = db_().getSheetByName('Participants');
      const values = sh.getDataRange().getValues();
      const headers = values[0].map(String);
      const nameCol = headers.indexOf('Name');
      const schoolCol = headers.indexOf('School');
      const categoryCol = headers.indexOf('Participation Category');
      let marked = 0;
      const names = [];
      for (let i = 1; i < values.length; i++) {
        const p = {name: values[i][nameCol], school: values[i][schoolCol]};
        if (isLocalParticipant_(p) && String(values[i][categoryCol]) !== 'Local Participants') {
          sh.getRange(i + 1, categoryCol + 1).setValue('Local Participants');
          marked++;
          names.push(String(values[i][nameCol]));
        }
      }
      audit_(db_(), 'ADMIN', 'MARK_LOCAL_PARTICIPANTS', 'PARTICIPANTS', 'MUMBAI_CITY', JSON.stringify({marked:marked,names:names}));
      return json_({success:true,marked:marked,names:names});
    }
    if (action === 'autoAllocate') {
      const result = autoAllocate_(db_());
      return json_({success:true, ...result});
    }
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
        'Verification Status':p.verificationStatus || 'Pending','Source':p.source || 'Admin entry','Participation Category':p.participationCategory || 'Accommodation Required'
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
  const local = people.filter(p => !isAccommodationEligible_(p)).length;
  const eligible = people.length-local;
  const occ = {};
  allotments.forEach(a => { const k=[a.Building,a.Floor,a.Room].join('|'); occ[k]=(occ[k]||0)+1; });
  rooms.forEach(r => { const k=[r.Building,r.Floor,r.Room].join('|'); r.Occupied=occ[k]||0; r.Available=Number(r.Beds||3)-r.Occupied; r.Status=String(r.Allotable)==='No'?'Excluded':(r.Occupied?'Occupied '+r.Occupied+'/'+Number(r.Beds||3):'Available'); });
  return {success:true,capacity:{buildings:3,floors:12,rooms:35,beds:105},participants:people.length,localParticipants:local,eligibleParticipants:eligible,activeAllotments:allotments.length,occupiedBeds:allotments.length,availableBeds:105-allotments.length,rooms:rooms,allotments:allotments};
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


function ensureHeaders_(sh, required) {
  const last = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1,1,1,last).getValues()[0].map(String);
  required.forEach(h => { if (existing.indexOf(h) < 0) { sh.getRange(1,sh.getLastColumn()+1).setValue(h); } });
}
function bedKeysForPreview_(allotments) {
  const counts={};
  allotments.forEach(a=>{
    const key=[a.Building,a.Floor,a.Room,a.Bed].join('|');
    counts[key]=(counts[key]||0)+1;
  });
  return counts;
}
function allocationPreview_(ss) {
  const allPeople=readObjects_(ss.getSheetByName('Participants'));
  const eligible=allPeople.filter(isAccommodationEligible_);
  const local=allPeople.filter(p=>!isAccommodationEligible_(p));
  const active=readObjects_(ss.getSheetByName('Allotments')).filter(a=>String(a['Allotment Status'])==='Active');
  const ids={}; eligible.forEach(p=>{const id=String(p['Participant ID']||'');ids[id]=(ids[id]||0)+1;});
  const names={}; eligible.forEach(p=>{const n=String(p.Name||'').trim().toLowerCase();names[n]=(names[n]||0)+1;});
  const eligibleById={}; eligible.forEach(p=>eligibleById[String(p['Participant ID'])]=p);
  const allottedIds={}; active.forEach(a=>{const id=String(a['Participant ID']||'');allottedIds[id]=(allottedIds[id]||0)+1;});
  const duplicateParticipantIds=Object.keys(ids).filter(id=>id && ids[id]>1);
  const duplicateNames=Object.keys(names).filter(n=>n && names[n]>1);
  const unallotted=eligible.filter(p=>!allottedIds[String(p['Participant ID'])]);
  const orphanAllotments=active.filter(a=>!eligibleById[String(a['Participant ID'])]);
  const duplicateActiveAllotments=Object.keys(allottedIds).filter(id=>allottedIds[id]>1);
  const roomCounts={}; active.forEach(a=>{const k=[a.Building,a.Floor,a.Room].join('|');roomCounts[k]=(roomCounts[k]||0)+1;});
  const overCapacity=Object.keys(roomCounts).filter(k=>roomCounts[k]>3);
  const localIds={}; local.forEach(p=>{localIds[String(p['Participant ID'])]=true;});
  const localAllotted=active.filter(a=>localIds[String(a['Participant ID'])]);
  const genderConflicts=[];
  const roomPeople={};
  active.forEach(a=>{const k=[a.Building,a.Floor,a.Room].join('|');(roomPeople[k]=roomPeople[k]||[]).push(a);});
  Object.keys(roomPeople).forEach(k=>{
    const genders={};
    roomPeople[k].forEach(a=>{
      const p=eligibleById[String(a['Participant ID'])];
      if(p){const g=String(p.Gender||'').trim().toLowerCase();if(g)genders[g]=true;}
    });
    if(Object.keys(genders).length>1) genderConflicts.push(k);
  });
  const rps=eligible.filter(p=>/resource person|resource associate/i.test(String(p.Role||'')));
  const rooms=readObjects_(ss.getSheetByName('Rooms')).filter(r=>String(r.Allotable)==='Yes');
  const reservedRoomKey='A1|Ground|2';
  const reservedRoomOccupants=active.filter(a=>[a.Building,a.Floor,String(a.Room)].join('|')===reservedRoomKey);
  const reservedRoomViolations=reservedRoomOccupants.filter(a=>normalizedPersonName_({Name:a['Participant Name']})!=='manishavivekmahajan');
  const duplicateActiveBeds=Object.keys(bedKeysForPreview_(active)).filter(k=>bedKeysForPreview_(active)[k]>1);
  const result={
    success:true,
    counts:{totalRecords:allPeople.length,localParticipants:local.length,eligibleParticipants:eligible.length,activeAllotments:active.length,unallottedEligible:unallotted.length,resourcePersons:rps.length,allotableRooms:rooms.length,allotableBeds:rooms.reduce((n,r)=>n+Number(r.Beds||3),0)},
    duplicateParticipantIds:duplicateParticipantIds,
    duplicateNames:duplicateNames,
    duplicateActiveAllotments:duplicateActiveAllotments,
    duplicateActiveBeds:duplicateActiveBeds,
    orphanAllotments:orphanAllotments,
    localParticipantsAllotted:localAllotted,
    overCapacityRooms:overCapacity,
    genderConflictRooms:genderConflicts,
    reservedRoomViolations:reservedRoomViolations,
    unallottedParticipants:unallotted.map(p=>({id:p['Participant ID'],name:p.Name,school:p.School,gender:p.Gender,role:p.Role})),
    activeAllotments:active
  };
  result.safeForFreshAllocation=duplicateParticipantIds.length===0 && duplicateActiveAllotments.length===0 && duplicateActiveBeds.length===0 && orphanAllotments.length===0 && localAllotted.length===0 && overCapacity.length===0 && genderConflicts.length===0 && reservedRoomViolations.length===0;
  return result;
}

function autoAllocate_(ss) {
  // Incremental, non-destructive allocation. Refuse to add beds until existing data passes reconciliation.
  const allPeople=readObjects_(ss.getSheetByName('Participants'));
  const people=allPeople.filter(isAccommodationEligible_);
  if(!people.length) throw new Error('Import participant records before automatic allocation.');
  const norm=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const genderOf=p=>{
    const v=String(p.Gender||'').trim().toLowerCase();
    if(/^(f|female|woman|women)$/.test(v)) return 'Female';
    if(/^(m|male|man|men)$/.test(v)) return 'Male';
    throw new Error('Gender needs verification before allocation: '+p.Name+' ('+(p.Gender||'blank')+'). No records were changed.');
  };
  const manisha=p=>norm(p.Name)==='manishavivekmahajan';
  const rp=p=>/resource person|resource associate/i.test(String(p.Role||''));
  const ground=p=>['rdparaskar','santoshvishramnathe','manishavivekmahajan'].includes(norm(p.Name));
  const participantById={}; people.forEach(p=>{const id=String(p['Participant ID']||'');if(!id)throw new Error('Missing Participant ID for '+p.Name);if(participantById[id])throw new Error('Duplicate Participant ID: '+id);participantById[id]=p;});
  const sh=ss.getSheetByName('Allotments');
  const allAllotments=readObjects_(sh), active=allAllotments.filter(a=>String(a['Allotment Status'])==='Active');
  const activeById={}, bedKeys={}, roomMap={}, activeNames={};
  active.forEach(a=>{
    const id=String(a['Participant ID']||''), p=participantById[id];
    if(!p) {
      const sourcePerson=allPeople.find(x=>String(x['Participant ID']||'')===id);
      if(sourcePerson && !isAccommodationEligible_(sourcePerson))
        throw new Error('Local participant still has an active allotment: '+(a['Participant Name']||id)+'. Run Local Participant marking and reconcile this existing allotment before allocating. No records changed.');
      throw new Error('Existing allotment has no eligible participant match: '+(a['Participant Name']||id)+'. Reconcile manually; no records changed.');
    }
    if(activeById[id]) throw new Error('Participant has multiple active allotments: '+p.Name+'. No records changed.');
    activeById[id]=true;
    const personName=normalizedPersonName_(p);
    if(personName) {
      if(!activeNames[personName]) activeNames[personName]=[];
      activeNames[personName].push(a);
    }
    const bed=[a.Building,a.Floor,a.Room,a.Bed].join('|');
    if(bedKeys[bed]) throw new Error('Duplicate occupied bed: '+bed+'. No records changed.');
    bedKeys[bed]=true;
    const key=[a.Building,a.Floor,a.Room].join('|');
    (roomMap[key]=roomMap[key]||[]).push({person:p,allotment:a});
  });
  const roomRows=readObjects_(ss.getSheetByName('Rooms'));
  const rooms=roomRows.filter(r=>String(r.Allotable)==='Yes');
  const roomByKey={}; rooms.forEach(r=>roomByKey[[r.Building,r.Floor,r.Room].join('|')]=r);
  Object.keys(roomMap).forEach(k=>{
    const occupants=roomMap[k], room=roomByKey[k];
    if(!room) throw new Error('Existing allotment points to a non-allotable/unknown room: '+k+'. No records changed.');
    if(occupants.length>Number(room.Beds||3)) throw new Error('Existing room exceeds bed capacity: '+k+'. No records changed.');
    const gs={}; occupants.forEach(x=>gs[genderOf(x.person)]=true);
    if(Object.keys(gs).length>1) throw new Error('Existing room has mixed genders: '+k+'. Reconcile before allocation; no records changed.');
  });
  const duplicateActiveNames=Object.keys(activeNames).filter(n=>activeNames[n].length>1);
  if(duplicateActiveNames.length) throw new Error('Multiple active allotments share the same normalized participant name: '+duplicateActiveNames.join(', ')+'. Verify whether these are duplicate imports or different people before allocating. No records changed.');
  const reservedExisting=(roomMap['A1|Ground|2']||[]).filter(x=>!manisha(x.person));
  if(reservedExisting.length) throw new Error('A1 Ground Floor Room 2 is not exclusively reserved for Manisha Vivek Mahajan. Reconcile existing occupants first. No records changed.');
  const pending=people.filter(p=>!activeById[String(p['Participant ID'])]);
  if(!pending.length) return {success:true,assigned:0,alreadyAllotted:active.length,roomsUsed:Object.keys(roomMap).length,occupiedBeds:active.length,localParticipantsExcluded:allPeople.length-people.length,message:'All eligible participants already have active allotments.'};
  // Validate all genders before making any writes.
  const byGender={Female:[],Male:[]};
  pending.forEach(p=>byGender[genderOf(p)].push(p));
  const assignments=[], now=new Date();
  const floorOrder=['Ground','First','Second','Third'], buildingOrder=['A1','A2','A3'];
  const roomKey=r=>[r.Building,r.Floor,r.Room].join('|');
  const roomSort=(a,b)=>floorOrder.indexOf(a.Floor)-floorOrder.indexOf(b.Floor)||buildingOrder.indexOf(a.Building)-buildingOrder.indexOf(b.Building)||Number(a.Room)-Number(b.Room);
  function currentCount(r){return (roomMap[roomKey(r)]||[]).length;}
  function currentGender(r){const xs=roomMap[roomKey(r)]||[];return xs.length?genderOf(xs[0].person):'';}
  function freeBeds(r){return Number(r.Beds||3)-currentCount(r);}
  function addToRoom(p,r,remark){
    const key=roomKey(r), n=currentCount(r)+1;
    const a={'Allotment ID':Utilities.getUuid(),'Participant ID':p['Participant ID'],'Participant Name':p.Name,'Building':r.Building,'Floor':r.Floor,'Room':r.Room,'Bed':n,'Allotment Status':'Active','Allotted At':now,'Updated At':now,'Remarks':remark};
    assignments.push(a); (roomMap[key]=roomMap[key]||[]).push({person:p,allotment:a});
  }
  function placeGroup(group,g,preferGround){
    const school=String(group[0].School||'');
    const resourceGroup=group.some(rp);
    function candidatesFor(minBeds){
      return rooms.filter(r=>freeBeds(r)>=minBeds&&(!currentGender(r)||currentGender(r)===g)&&(!resourceGroup||currentCount(r)===0));
    }
    function sortCandidates(candidates){
      candidates.sort((a,b)=>{
        const ag=preferGround&&a.Floor==='Ground'?0:1, bg=preferGround&&b.Floor==='Ground'?0:1;
        const ac=currentCount(a)>0?0:1, bc=currentCount(b)>0?0:1;
        const as=school&&roomMap[roomKey(a)]&&roomMap[roomKey(a)].some(x=>String(x.person.School||'')===school)?0:1;
        const bs=school&&roomMap[roomKey(b)]&&roomMap[roomKey(b)].some(x=>String(x.person.School||'')===school)?0:1;
        return ag-bg||as-bs||ac-bc||roomSort(a,b);
      });
      return candidates;
    }
    // Resource persons retain their protected double-occupancy grouping in an otherwise empty room.
    if(resourceGroup){
      const room=sortCandidates(candidatesFor(group.length))[0];
      if(!room) throw new Error('No compatible double-occupancy room for '+group.map(p=>p.Name).join(', ')+'. Existing allotments were preserved; no new records were written.');
      group.forEach(p=>addToRoom(p,room,'Resource Person — double occupancy'));
      return;
    }
    // Regular participants may be split across compatible rooms. This avoids
    // rejecting a group of three merely because the remaining beds are fragmented.
    group.forEach(p=>{
      const room=sortCandidates(candidatesFor(1))[0];
      if(!room) throw new Error('No compatible room space for '+p.Name+'. Existing allotments were preserved; no new records were written.');
      let note=manisha(p)?'Single occupancy — A1 Ground Floor Room 2':group.length===1?'Single occupancy':group.length===2?'Double occupancy':'Triple occupancy (maximum 3 per room)';
      if(ground(p)) note+='; Ground Floor Priority';
      addToRoom(p,room,note);
    });
  }
  // Manisha's reserved single room is protected and never shared.
  const manishaPending=pending.find(manisha);
  if(manishaPending){
    const target=rooms.find(r=>r.Building==='A1'&&r.Floor==='Ground'&&String(r.Room)==='2');
    if(!target) throw new Error('Reserved room A1 Ground Room 2 is unavailable. No records changed.');
    const existing=roomMap[roomKey(target)]||[];
    if(existing.length) throw new Error('A1 Ground Room 2 is occupied by another participant; Manisha single-room exception needs review. No records changed.');
    addToRoom(manishaPending,target,'Single occupancy — A1 Ground Floor Room 2; Ground Floor Priority');
    byGender[genderOf(manishaPending)]=byGender[genderOf(manishaPending)].filter(p=>!manisha(p));
  }
  ['Female','Male'].forEach(g=>{
    let list=byGender[g].filter(p=>!manisha(p));
    const priority=list.filter(ground), rest=list.filter(p=>!ground(p));
    // Put named ground-floor priority participants into a ground-floor room first.
    priority.forEach(p=>placeGroup([p],g,true));
    // RPs are allocated in pairs; an odd RP shares with a same-gender participant.
    const rps=rest.filter(rp), nonRps=rest.filter(p=>!rp(p)), groups=[];
    for(let i=0;i+1<rps.length;i+=2) groups.push([rps[i],rps[i+1]]);
    let pool=nonRps.slice();
    if(rps.length%2){
      if(!pool.length) throw new Error('Cannot provide double occupancy for '+rps[rps.length-1].Name+': no same-gender participant is available. Existing allotments preserved; no new records written.');
      groups.push([rps[rps.length-1],pool.shift()]);
    }
    groups.forEach(gp=>placeGroup(gp,g,false));
    // Keep schoolmates together where feasible; groups of 3 first, then 2.
    const schools={}; pool.forEach(p=>{const k=String(p.School||'Unknown');(schools[k]=schools[k]||[]).push(p);});
    let ordered=[];Object.keys(schools).sort().forEach(k=>ordered=ordered.concat(schools[k]));
    while(ordered.length){
      const size=Math.min(3,ordered.length);
      placeGroup(ordered.splice(0,size),g,false);
    }
  });
  if(assignments.length!==pending.length) throw new Error('Allocation validation failed; no new records were written.');
  // One bulk write only after all validations and placement decisions succeed.
  const rows=assignments.map(a=>[a['Allotment ID'],a['Participant ID'],a['Participant Name'],a.Building,a.Floor,a.Room,a.Bed,a['Allotment Status'],a['Allotted At'],a['Updated At'],a.Remarks]);
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  const localCount=allPeople.filter(p=>!isAccommodationEligible_(p)).length;
  audit_(ss,'ADMIN','INCREMENTAL_AUTO_ALLOCATE','ALLOTMENTS','BATCH',JSON.stringify({newlyAssigned:assignments.length,previouslyActive:active.length,roomsUsed:Object.keys(roomMap).length,localExcluded:localCount}));
  return {success:true,assigned:assignments.length,alreadyAllotted:active.length,roomsUsed:Object.keys(roomMap).length,occupiedBeds:active.length+assignments.length,localParticipantsExcluded:localCount,assignments:assignments};
}
