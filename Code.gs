// Vaishnavi Jan Sewa Kendra - Google Apps Script backend starter
// 1) Create a Google Sheet, copy its ID from the URL, and paste below.
// 2) In Apps Script: Project Settings > Script Properties, set SHEET_ID to that ID.
// 3) Deploy > New deployment > Web app; Execute as: Me; access: Anyone (only if appropriate for your data).
// NOTE: Public web app endpoints are not suitable for sensitive customer documents.
function doGet(){return json_({ok:true,message:'Vaishnavi Jan Sewa Kendra API ready'});}
function doPost(e){
  try{
    var body=JSON.parse((e.postData&&e.postData.contents)||'{}');
    var type=String(body.type||'');
    var allowed=['customers','applications','accounts'];
    if(allowed.indexOf(type)<0) throw new Error('Invalid record type');
    var id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if(!id) throw new Error('SHEET_ID Script Property is missing');
    var ss=SpreadsheetApp.openById(id);
    var sh=ss.getSheetByName(type)||ss.insertSheet(type);
    if(sh.getLastRow()===0) sh.appendRow(['Saved At','Record JSON']);
    sh.appendRow([new Date(),JSON.stringify(body.record||{})]);
    return json_({ok:true});
  }catch(err){return json_({ok:false,error:String(err.message||err)});}
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
