//файл для создания и управления настройками системы
function createFileList(id)   //файл с перечнем всех файлов в папке/подпапках системы
{
  var array=['Files','Folders'];
  var headers=['Name','Id','MimeType','Author','Creation date']; //заголовки столбцов
  var ss= SpreadsheetApp.create('FilesList');
  ss=createSheetsInExitingTable(ss,array);
  for(var i in array){
    ss.getSheetByName(array[i]).appendRow(headers);
    ss.getSheetByName(array[i]).setFrozenRows(1);
    }
  var res=DriveApp.getFileById(ss.getId()).makeCopy(ss.getName(),DriveApp.getFolderById(id)); 
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return res;
}

function createFormImage(){
  var bytes=HtmlService.createHtmlOutputFromFile('icon').getBlob().getDataAsString().split(',');
  var blob=Utilities.newBlob(bytes,MimeType.PNG).setName('icon');
  SitesApp.getActivePage().addHostedAttachment(blob);
}

function autoWidth(ss){  //автоматичская установка ширины колонки по содержимому SpreadSheetApp ss
  for(var i=1;i<ss.getLastColumn();i++){
    ss.autoResizeColumn(i);
  }
}

function createDomainFile(id,domains){
var domain=SpreadsheetApp.openById(createFile(id,'Domain',['Name','Status']).getId());
  if(domains!==undefined){
    for(var i in domains){
    domain.getSheetByName(mainSheet).appendRow([domains[i].name,domains[i].status]).setFrozenRows(1); //записать домены
    domain.insertSheet(domains[i].status.toString()); //создать для каждого домена свой лист
    domain.getSheetByName(domains[i].status.toString()).appendRow(['Name','ID','Author','Category']).setFrozenRows(1);
    }
  }
  return domain;
}

function createGeneralFile(id){
var general=SpreadsheetApp.openById(createFile(id,'General').getId());
general.getSheetByName(mainSheet).appendRow(['Name','Value']).appendRow(['answerLimit','5']).appendRow(['numberForms',12]).setFrozenRows(1);
general.insertSheet('Categories').appendRow(['Name']).appendRow(['Загальна']).setFrozenRows(1);
general.insertSheet('UAID').appendRow(['id']).setFrozenRows(1); //UAID-unique answer id
return general;
}

function createManagersFile(id){
var admins=SpreadsheetApp.openById(createFile(id,'Managers',['email','id','Destination','Creation date']).getId());
  admins.insertSheet(mngSheet).appendRow(['e-mail']).setFrozenRows(1);
  admins.insertSheet('Admins').appendRow(['e-mail']).setFrozenRows(1);
  admins.appendRow([Session.getActiveUser().getEmail()]);
  return admins;
}

function install(id,domains){ 
  createFormImage();
  var filesList=createFileList(id);
  var settings={
    rootId:id,
    filesListId:filesList.getId(),
  };
  var blob= Utilities.newBlob(JSON.stringify(settings),MimeType.PLAIN_TEXT).setName('setup');
  SitesApp.getActivePage().addHostedAttachment(blob);
  createFolders(id,domains);
  createGeneralFile(id);
  createManagersFile(id);
  createFile(id,'Translate',['Original','Translate']);
  createDomainFile(id,domains);
  PropertiesService.getUserProperties().deleteAllProperties();
  DriveApp.getFolderById(id).setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
  //установить дополнительно права доступа для менеджеров как редактор
  return 'Успішно встановлено!';
}

function createFolders(id,domains){  //FolderId id 
  var root=DriveApp.getFolderById(id);
  var folders=[];
  var generalFolders=['Documents','Settings'];
  var docFolders=['Templates','Users'];
  var templFolders=[];
  for(var j in domains) templFolders.push(domains[j].status);
    folders.push(root.createFolder(generalFolders[1]));
  var docs=root.createFolder(generalFolders[0]);
  folders.push(docs);
  folders.push(docs.createFolder(docFolders[1]));
  var templ=docs.createFolder(docFolders[0]);
  folders.push(templ);
  for(var s in templFolders)
    folders.push(templ.createFolder(templFolders[s]));
  for(var i in folders)
    addToFilesList(folders[i],id);
}

function createFile(id,fileName,columnHeaders){ 
  var file=SpreadsheetApp.create(fileName);
  file.getSheets()[0].setName(mainSheet);
  if(columnHeaders!==undefined) file.getSheetByName(mainSheet).appendRow(columnHeaders).setFrozenRows(1);
  var temp=DriveApp.getFileById(file.getId()).makeCopy(file.getName(),DriveApp.getFolderById(Data.GetElementInList('Settings',MimeType.FOLDER,id)[0].ID));
  addToFilesList(temp,id);
  DriveApp.getFileById(file.getId()).setTrashed(true);
  return temp;
}

function isFirstAdmin(email){
  var admins;
  var res=false;
  var m=Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS);
  if(m===undefined){
  admins=SitesApp.getActiveSite().getOwners();
  for(var i in admins){
    if(admins[i].getEmail()===email) res=true;
    }
  }
  return res;
}

function setup() {
  if(isFirstAdmin(userEmail)===true){
    if(attachementExist('setup')!==true) return true;
  } 
}

