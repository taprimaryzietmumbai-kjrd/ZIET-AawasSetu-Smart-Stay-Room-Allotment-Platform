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
        const category = p.participationCategory === 'Local Participants' ? 'Local Participants' : 'Accommodation Required';
        if (category === 'Local Participants') local++;
        upsert_(sh, 'Participant ID', id, {'Participant ID':id,'Name':p.name,'Designation':p.designation||'','School':p.school||'','Region':p.region||'','Gender':p.gender||'','DOB':p.dob||'','Employee Code':p.employeeCode||'','Role':p.role||'Participant','Preference':p.preference||'','Special Requirement':p.specialRequirement||'','Verification Status':p.verificationStatus||'Imported','Source':p.source||'Bulk workbook import','Participation Category':category});
        imported++;
      });
      audit_(db_(),'ADMIN','BULK_IMPORT','PARTICIPANTS','BATCH',JSON.stringify({imported:imported,local:local}));
      return json_({success:true,imported:imported,localParticipants:local,eligible:imported-local});
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
  const local = people.filter(p => String(p['Participation Category'])==='Local Participants').length;
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
function allocationPreview_(ss) {
  const allPeople=readObjects_(ss.getSheetByName('Participants'));
  const eligible=allPeople.filter(p=>String(p['Participation Category']||'')!=='Local Participants' && String(p.Role||'Participant')!=='Local Participant');
  const local=allPeople.filter(p=>String(p['Participation Category']||'')==='Local Participants');
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
  const localAllotted=active.filter(a=>local.some(p=>String(p['Participant ID'])===String(a['Participant ID'])));
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
  const result={
    success:true,
    counts:{totalRecords:allPeople.length,localParticipants:local.length,eligibleParticipants:eligible.length,activeAllotments:active.length,unallottedEligible:unallotted.length,resourcePersons:rps.length,allotableRooms:rooms.length,allotableBeds:rooms.reduce((n,r)=>n+Number(r.Beds||3),0)},
    duplicateParticipantIds:duplicateParticipantIds,
    duplicateNames:duplicateNames,
    duplicateActiveAllotments:duplicateActiveAllotments,
    orphanAllotments:orphanAllotments,
    localParticipantsAllotted:localAllotted,
    overCapacityRooms:overCapacity,
    genderConflictRooms:genderConflicts,
    unallottedParticipants:unallotted.map(p=>({id:p['Participant ID'],name:p.Name,school:p.School,gender:p.Gender,role:p.Role})),
    activeAllotments:active
  };
  result.safeForFreshAllocation=active.length===0 && duplicateParticipantIds.length===0 && localAllotted.length===0 && overCapacity.length===0;
  return result;
}

function autoAllocate_(ss) {
  const allPeople=readObjects_(ss.getSheetByName('Participants'));
  const people=allPeople.filter(p=>String(p['Participation Category']||'')!=='Local Participants' && String(p.Role||'Participant')!=='Local Participant');
  const rooms=readObjects_(ss.getSheetByName('Rooms')).filter(r=>String(r.Allotable)==='Yes');
  const sh=ss.getSheetByName('Allotments');
  if(readObjects_(sh).some(a=>String(a['Allotment Status'])==='Active')) throw new Error('Active allotments already exist. Release them before a fresh allocation.');
  if(!people.length) throw new Error('Import participant records before automatic allocation.');
  const norm=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const female=p=>/female|f$/.test(String(p.Gender||'').toLowerCase());
  const gender=p=>female(p)?'Female':'Male';
  const manisha=p=>norm(p.Name)==='manishavivekmahajan';
  const rp=p=>/resource person|resource associate/i.test(String(p.Role||''));
  const ground=p=>['rdparaskar','santoshvishramnathe','manishavivekmahajan'].includes(norm(p.Name));
  const floorOrder=['Ground','First','Second','Third'];
  const roomOrder=rooms.sort((a,b)=>{
    const ap=a.Building==='A1'&&a.Floor==='Ground'&&String(a.Room)==='2'?-1:0;
    const bp=b.Building==='A1'&&b.Floor==='Ground'&&String(b.Room)==='2'?-1:0;
    return ap-bp||['A1','A2','A3'].indexOf(a.Building)-['A1','A2','A3'].indexOf(b.Building)||floorOrder.indexOf(a.Floor)-floorOrder.indexOf(b.Floor)||Number(a.Room)-Number(b.Room);
  });
  const assignments=[],used=new Set();
  function takeRoom(g,occupants,forceManisha){
    let room;
    if(forceManisha) room=roomOrder.find(r=>!used.has([r.Building,r.Floor,r.Room].join('|'))&&r.Building==='A1'&&r.Floor==='Ground'&&String(r.Room)==='2');
    if(!room&&occupants.some(ground)) room=roomOrder.find(r=>!used.has([r.Building,r.Floor,r.Room].join('|'))&&r.Floor==='Ground');
    if(!room) room=roomOrder.find(r=>!used.has([r.Building,r.Floor,r.Room].join('|')));
    if(!room) throw new Error('Capacity conflict: '+(used.size+1)+' rooms are required, but only '+roomOrder.length+' are allotable. Review gender-wise double occupancy for Resource Persons.');
    used.add([room.Building,room.Floor,room.Room].join('|'));
    occupants.forEach((p,i)=>{
      let remark=manisha(p)?'Single occupancy — A1 Ground Floor Room 2':rp(p)?'Resource Person — double occupancy':occupants.length===2?'Double occupancy':'Triple occupancy';
      if(ground(p)) remark+='; Ground Floor Priority';
      assignments.push({'Allotment ID':Utilities.getUuid(),'Participant ID':p['Participant ID'],'Participant Name':p.Name,'Building':room.Building,'Floor':room.Floor,'Room':room.Room,'Bed':i+1,'Allotment Status':'Active','Allotted At':new Date(),'Updated At':new Date(),'Remarks':remark});
    });
  }
  const byGender={Female:people.filter(p=>gender(p)==='Female'),Male:people.filter(p=>gender(p)==='Male')};
  if(byGender.Female.some(manisha)) takeRoom('Female',[byGender.Female.find(manisha)],true);
  else if(byGender.Male.some(manisha)) takeRoom('Male',[byGender.Male.find(manisha)],true);

  ['Female','Male'].forEach(g=>{
    let list=byGender[g].filter(p=>!manisha(p));
    const rps=list.filter(rp), nonRps=list.filter(p=>!rp(p));
    // Pair RPs with another same-gender RP; an odd RP is paired with a same-gender participant.
    const special=[];
    for(let i=0;i+1<rps.length;i+=2) special.push([rps[i],rps[i+1]]);
    let remainingNonRps=nonRps.slice();
    if(rps.length%2){
      if(!remainingNonRps.length) throw new Error('Cannot provide double occupancy for '+rps[rps.length-1].Name+': no same-gender participant is available to share.');
      special.push([rps[rps.length-1],remainingNonRps.shift()]);
    }
    special.forEach(group=>takeRoom(g,group,false));
    // Group remaining participants in triples, using the minimum number of doubles needed.
    const schoolGroups={};
    remainingNonRps.forEach(p=>{const k=String(p.School||'Unknown');(schoolGroups[k]=schoolGroups[k]||[]).push(p);});
    let pool=[];Object.keys(schoolGroups).sort().forEach(k=>pool=pool.concat(schoolGroups[k]));
    if(pool.length===1) throw new Error('One '+g+' participant remains after Resource Person pairing; double occupancy cannot be satisfied without another same-gender participant.');
    const roomCount=Math.ceil(pool.length/3), doubleRooms=roomCount*3-pool.length;
    const tripleCount=pool.length-doubleRooms*2;
    const triples=pool.splice(0,tripleCount), doubles=pool;
    while(triples.length) takeRoom(g,triples.splice(0,3),false);
    while(doubles.length) takeRoom(g,doubles.splice(0,2),false);
  });
  if(assignments.length!==people.length) throw new Error('Allocation validation failed: assigned '+assignments.length+' of '+people.length+'.');
  if(used.size>rooms.length) throw new Error('Room capacity exceeded.');
  const rows=assignments.map(a=>[a['Allotment ID'],a['Participant ID'],a['Participant Name'],a.Building,a.Floor,a.Room,a.Bed,a['Allotment Status'],a['Allotted At'],a['Updated At'],a.Remarks]);
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  const localCount=allPeople.filter(p=>String(p['Participation Category'])==='Local Participants').length;
  audit_(ss,'ADMIN','AUTO_ALLOCATE','ALLOTMENTS','BATCH',JSON.stringify({assigned:assignments.length,rooms:used.size,localExcluded:localCount,resourcePersonOccupancy:'double'}));
  return {success:true,assigned:assignments.length,roomsUsed:used.size,occupiedBeds:assignments.length,localParticipantsExcluded:localCount,assignments:assignments};
}
