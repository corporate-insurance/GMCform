// Google Apps Script – paste into Extensions > Apps Script of your Google Sheet.
var ADMIN_KEY="CHANGE-THIS-KEY"; // choose your own secret key; type the same in the admin page
var HEAD=["Total No. of Lifes","Employees","Emp ID","Employee Name","Gender","Relationship","D.O.B","Age","Designation","Mobile No","Mail id"];

function doPost(e){
  var lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var d=JSON.parse(e.postData.contents);
    var name=String(d.company||"General").replace(/[^\w \-]/g,"").trim().slice(0,60)||"General";
    var empId=String(d.empId||"").trim();
    var rows=d.rows;
    if(!empId||!Array.isArray(rows)||!rows.length||rows.length>8)return out("invalid");
    rows=rows.map(function(r){var x=[];for(var i=0;i<9;i++)x.push(String(r[i]==null?"":r[i]).slice(0,120));return x});

    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName(name)||ss.insertSheet(name);

    // Existing entries grouped per employee (a group starts at a "Self" row)
    var groups=[];
    if(sh.getLastRow()>1){
      var vals=sh.getRange(2,1,sh.getLastRow()-1,11).getValues();
      vals.forEach(function(r){
        if(String(r[5])==="Self")groups.push([]);
        if(groups.length)groups[groups.length-1].push(r.slice(2).map(String));
      });
    }
    var done=false;
    groups=groups.map(function(g){if(String(g[0][0])===empId){done=true;return rows}return g});
    if(!done)groups.push(rows);

    // Rebuild with running numbers
    var all=[HEAD],lives=0;
    groups.forEach(function(g,gi){
      g.forEach(function(r,ri){all.push([++lives,ri===0?gi+1:""].concat(r))});
    });
    sh.clear();
    var rng=sh.getRange(1,1,all.length,11);
    rng.setNumberFormat("@");          // plain text: keeps leading zeros, blocks formulas
    rng.setValues(all);
    sh.getRange(1,1,1,11).setFontWeight("bold");
    sh.setFrozenRows(1);
    return out("ok");
  }catch(err){return out("error: "+err)}
  finally{lock.releaseLock()}
}
function doGet(e){
  var p=(e&&e.parameter)||{};
  if(!p.action)return out("APNA GMC collector is running");
  if(!ADMIN_KEY||p.key!==ADMIN_KEY)return json({error:"unauthorized"});
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(p.action==="list"){
    var list=[];
    ss.getSheets().forEach(function(sh){
      var n=sh.getLastRow();
      if(n<1)return;
      var emp=0;
      if(n>1)sh.getRange(2,6,n-1,1).getValues().forEach(function(r){if(r[0]==="Self")emp++});
      list.push({name:sh.getName(),employees:emp,lives:Math.max(n-1,0)});
    });
    return json({companies:list});
  }
  if(p.action==="data"){
    var s=ss.getSheetByName(String(p.c||""));
    if(!s)return json({error:"No such company"});
    return json({rows:s.getDataRange().getValues()});
  }
  return json({error:"bad request"});
}
function json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function out(t){return ContentService.createTextOutput(t)}
