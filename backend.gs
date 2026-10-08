// Google Apps Script — stores orders in a Google Sheet.
const ADMIN_PIN = 'CHANGE-ME-1234'; // <- set your own PIN

function sheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('orders') || ss.insertSheet('orders');
  if (s.getLastRow() === 0) s.appendRow(['ref','status','createdAt','json']);
  return s;
}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function doPost(e){
  const b = JSON.parse(e.postData.contents), s = sheet_();
  if (b.action === 'add'){
    const o = b.order; o.status = 'pending';
    s.appendRow([o.ref, 'pending', o.createdAt, JSON.stringify(o)]);
    return out_({ok:true});
  }
  if (b.pin !== ADMIN_PIN) return out_({ok:false});
  const rows = s.getDataRange().getValues();
  if (b.action === 'list'){
    return out_({ok:true, orders: rows.slice(1).map(r => { const o = JSON.parse(r[3]); o.status = r[1]; return o; })});
  }
  for (let i = 1; i < rows.length; i++){
    if (rows[i][0] === b.ref){
      if (b.action === 'status') s.getRange(i+1, 2).setValue(b.status);
      if (b.action === 'delete') s.deleteRow(i+1);
      return out_({ok:true});
    }
  }
  return out_({ok:false});
}
