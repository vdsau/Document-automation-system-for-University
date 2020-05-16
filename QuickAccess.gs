//-------------------------GLOBAL---------------------------------
var root=getRoot();//@ID - rootId, FILESLIST - fileslist id

var userEmail=Session.getActiveUser().getEmail();

var user=(setup()!==true)?defineUser(userEmail):'';

var mngSheet="Managers";

var nowTime=Utilities.formatDate(new Date(),"GMT+3", "HH:mm:ss dd.MM.yyyy");
var nowDate=Utilities.formatDate(new Date(),"GMT+3", "dd.MM.yyyy");

var _properties={
 lastDate:'LAST_DATE',
 score:'SCORE',
};

var adminLimit=150;

var _userProp=PropertiesService.getUserProperties();

var _sys_vars={
numberForms: 'numberForms',
answerLimit:'answerLimit',
};

var _headers={
note:'�������',
time:'��� �����������',
status:'������',
sender:'³��������',
answerId:'id ³�����',
lastEdit:'���� �������� 䳿',
};

var mainHeaders=['����� ���������',_headers.note,_headers.time,_headers.status,_headers.sender,_headers.answerId,_headers.lastEdit];

var emailSubject='������� ������������ �������������';

var additionHeaders=[_headers.status,_headers.note,_headers.time,_headers.sender,_headers.answerId,_headers.lastEdit];

var mainSheet='Main';

var site=SitesApp.getActiveSite().getUrl();

var gotoSystem="\n\n\n����� � �������:\n "+site;

var defaultText="-�� ������-";

var mngrFile='manager_';

var _status={
  ALL:'�� �������',
  INQUERY : '� ����',
  DONE: '������',
  INPROCESS: '� �������',
  DENIED:'³�������',
};

//-------------------------Settings-------------------------------
function getRoot(){
  return Settings.GetRoot();
}
function getAttachement(name){
  return Settings.GetAttachement(name);
}
function attachementExist(name){
  return Settings.AttachementExist(name);
}
function getAttach(name){
  return Settings.GetAttach(name);
}
function getSettingByName(name){
 return Settings.GetSettingByName(name);
}
function setSettingByName(name,value){
 return Settings.SetSettingByName(name,value);
}
function setDomains(domains){
 return Settings.SetDomains(domains);
}
function setCategorie(value){
 return Settings.SetCategorie(value);
}

//-------------------------DOCS---------------------------------

function generateDocument(params){ //������� ��������
  return QDoc.GenerateDocument(params);
}
function parseElem(doc) {  //Document doc-�������� �������� 
  return QDoc.ParseElem(doc);
}

//-------------------------SHEETS-------------------------------

function createSheetsInExitingTable(ss,sheetsArray){
  return QSheet.CreateSheetsInExitingTable(ss,sheetsArray);
}
function getSheet (sheet,listName){  //������� ������ �� ����� � ������������ ����� � ����� String sheet - url � ����� �������
  return QSheet.GetSheet(sheet,listName);
}
function getData(ss,sheetName) {   //�������� �������� �� ������� �� ss - string, id �������
  return QSheet.GetData(ss,sheetName);
}
function createSheetsInTable(tableName,sheetsArray) {// String tableName-��� �������, Array sheetsArray-�������� ������ �����, ������� � ������ ������������� ����
  return  QSheet.CreateSheetsInTable(tableName,sheetsArray);
}
function getHistory (sheetId,sheetName,columnsNames) {//String sheetId- id �������, String sheetName - ��� ����� � �������, Object columnsNames - ������,� ������� � ����������� ������������������ ���������� �������� �������:������ �������
  return QSheet.GetHistory(sheetId,sheetName,columnsNames);
}
function clearHistory(managerMode){  //@return String �������� ��� ������� ������� 
 return QSheet.ClearHistory(managerMode);
}
function getFormHeaders(headers){  //�������� ��������� �������� �������
return QSheet.GetFormHeaders(headers);
}

//-------------------------DATA-------------------------------
function getName(){
  return Data.GetName();
}
function createAnswerId (){
return Data.CreateAnswerId();
}
function toUpper(text,position){
  return Data.ToUpper(text,position);
}
function addManager(email){
  return Data.AddManager(email);
}
function addToFilesList(file,id){
  return Data.AddToFilesList(file,id);
}
function getElementInList(name,mimeType){
  return Data.GetElementInList(name,mimeType);
}
function addTemplate(title,destination,idForm,idTempl){
  return Data.AddTemplate(title,destination,idForm,idTempl);
}
function deleteTemplate(params){
 return Data.DeleteTemplate(params);
}
function getUserData(){ //�������� ������ ������������
  return Data.GetUserData();
}
function parseEmail(email){ //�������� ��������� ����� ������������
  return Data.ParseEmail(email);
}
function defineUser (email) {  //���������� ������������ �� ��������� ������
  return Data.DefineUser(email);  
}
function isAdmin(email) { //�������� �� ������
  return Data.IsAdmin(email);
}

//-------------------------Content-------------------------------
function createForm(items){
  return Content.CreateForm(items);
}
function defineItem(item){
  return Content.DefineItem(item);
}
function createHeader(header){
 return Content.CreateHeader(header);
}
function doGet(e) {
  return (setup()===true)?openPage('setup'):openPage('Page');
}
function include(filename) {   //������� html-��������
  return Content.Include(filename);
}
function openElement(name) { //������� ���������� �������� ����� ���������� gs
  return Content.OpenElement(name);
}
function openPage(name) { //�������� �������� � ����������� gs
  return Content.OpenPage(name);
}

function getDefaultText(){
return defaultText;
}