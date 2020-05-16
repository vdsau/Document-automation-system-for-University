Settings={
GetRoot:function(){
var res={};
var attach=Settings.GetAttachement('setup');
if(attach!==undefined){
res['ID']=attach.rootId;
res['FILESLIST']=attach.filesListId;
}
return res;
},  

AttachementExist:function(name){
var attach=SitesApp.getActiveSite().getAttachments();
for(var i in attach)
if(attach[i].getTitle()===name) return true;
return false;
},

GetAttach:function (name){
var attach=SitesApp.getActiveSite().getAttachments();
var res;
for(var i in attach){
if(attach[i].getTitle()===name){
res=attach[i];
break;}
}
return res;
}, 

GetAttachement:function (name){
var s=Settings.GetAttach(name);
return (s!==undefined)?JSON.parse(s.getBlob().getDataAsString()):undefined;
},

GetSettingByName:function(name){
var res;
if(name!==undefined){
 var settings=SpreadsheetApp.openById(Data.GetElementInList('General',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
 var data=settings.getDataRange().getValues();
 for(var row=1;row<data.length; row++){
  if(data[row][0].toString()===name){
   res=data[row][1].toString();
    break;
   }
  }
 }
 return res;
},

SetCategorie:function(value){ //Array domains{name, status}
var res='';
 if(value!==undefined){
 var settings=SpreadsheetApp.openById(Data.GetElementInList('General',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName('Categories').appendRow([value]);
  var data=settings.getDataRange().getValues();
  for(var row=1;row<data.length;row++){
  res+='<tr><td>'+row+'</td>';
    data[row].forEach(function (item,index,array){
     res+='<td>'+item.toString()+'</td>';
    });
   res+='</tr>';
  }
 }
  return res;
},

SetDomains:function(domains){ //Array domains{name, status}
var res='';
 if(domains!==undefined){
 var settings=SpreadsheetApp.openById(Data.GetElementInList('Domain',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
  domains.forEach(function (item,index,array){
     settings.appendRow([item.name.toString(),item.status.toString()]);
  });
  var data=settings.getDataRange().getValues();
  for(var row=1;row<data.length;row++){
  res+='<tr><td>'+row+'</td>';
   data[row].forEach(function (item,index,array){
     res+='<td>'+item.toString()+'</td>';
    });
   res+='</tr>';
  }
 }
  return res;
},

SetSettingByName:function(name,value){ 
 if(name!==undefined){
 var settings=SpreadsheetApp.openById(Data.GetElementInList('General',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
 var data=settings.getDataRange().getValues();
 for(var row=1;row<data.length; row++){
  if(data[row][0].toString()===name){
   settings.getRange(row+1,2).setValue(value);
    break;
   }
  }
 }
 return value;
},

};
       
QDoc={
DefaultAttributes:{
'FONT_FAMILY':'Times New Roman',
'FOREGROUND_COLOR':'#000000', //установить чёрный цвет текста
'BACKGROUND_COLOR':'#ffffff', //установить белый фон текста
'BOLD':'false' //Не выделять жирным
},

GetAnswers : function(params){
var responses=[];  //массив ответов
var title='';
var content='';
for(var i=0;i<params.headers.length; i++){
title='{'+params.headers[i].toString()+'}';
responses.push({'TITLE':title,
'CONTENT':params.content[i].toString(),});
}
responses.push({'TITLE':'{ДАТА}', //сегоднешняя дата
'CONTENT': Utilities.formatDate(new Date(),"GMT+2", "dd.MM.yyyy"),   
});
return responses;
},

GenerateDocument:function(params){ //Object params,fileId
var res='';
var responses=QDoc.GetAnswers(params);
var ss=SpreadsheetApp.openById(params.fileId);  //файл менеджера
QSheet.SetStatus(ss,params,_status.INPROCESS);
var user=Data.GetElementInList(params.sender,MimeType.GOOGLE_SHEETS)[0];
if(user!==undefined){
ss=SpreadsheetApp.openById(user.ID);
QSheet.SetStatus(ss,params,_status.INPROCESS);
}

var formData=getFormData(params.sheetName);

var name=params.answerId;
var newName=params.sheetName+params.sender;
var userFolder=Data.GetElementInList(newName,MimeType.FOLDER); //папка пользователя для данной формы
var fdr;
if(userFolder[0]===undefined){
  fdr=DriveApp.getFolderById(Data.GetElementInList(params.sender,MimeType.FOLDER)[0].ID).createFolder(newName);
  Data.AddToFilesList(fdr); }
  else {
   fdr=DriveApp.getFolderById(userFolder[0].ID);
  }
  var folder;
  var f=Data.GetElementInList(name,MimeType.FOLDER);
  if(f[0]===undefined){ 
  folder=fdr.createFolder(name);
  Data.AddToFilesList(folder);
  }
  else  folder=DriveApp.getFolderById(f[0].ID); 
var dc=DriveApp.getFileById(formData.doc.toString());
var doc=dc.makeCopy(dc.getName(), folder);
 Data.AddToFilesList(doc);
 doc=DocumentApp.openById(doc.getId());
QDoc.ParseElem(doc,responses,'header');
QDoc.ParseElem(doc,responses,'body');
QDoc.ParseElem(doc,responses,'footer');
 SpreadsheetApp.flush();
 res=params.closeButton+"<center><div style='margin-top:10%;'>Документ успішно створено!<br>Посилання на документ:<br><a target='_blank' href='"+doc.getUrl()+"'>Документ</a></div></center>";
 GmailApp.sendEmail(params.sender.toString(), emailSubject, "Поданий Вами документ опрацьовується.\nОчікуйте повідомлення про готовність."+gotoSystem);
return res;
},

CheckArea:function(doc,area){  // String area - проверить на header,footer,body
var parag;
area=area.toLowerCase();
switch(area){
case 'header':
if(doc.getHeader()!== null)
parag=doc.getHeader().getParagraphs();
break;
case 'body':
if(doc.getBody()!== null)
parag=doc.getBody().getParagraphs();
break;
case 'footer':
if(doc.getFooter()!== null)
parag=doc.getFooter().getParagraphs();
break;
default:
if(doc.getHeader()!== null)
parag=doc.getHeader().getParagraphs()[0];
break;
}
return parag;
},

ParseElem: function(doc,responses,area)//Document doc-открытый документ,String area - header,body,footer 
{  
var parag=QDoc.CheckArea(doc,area);
for(var x in  parag) {
for(var j in responses) { 
var temp=parag[x].findText(responses[j]['TITLE']);
if(temp!==null) {
parag[x].setText(temp.getElement().asText().getText().replace(RegExp(responses[j]['TITLE']),responses[j]['CONTENT']));
parag[x].setAttributes(QDoc.DefaultAttributes);
}
else continue;
} }
return doc;
}, };           

QSheet= {

GetData:function(ss,sheetName) {   //получить значения из таблицы по SpreadsheetApp ss таблица
if(sheetName===undefined) return ss.getActiveSheet().getDataRange().getValues(); 
return ss.getSheetByName(sheetName).getDataRange().getValues();
},

SetStatus:function(spreadSheet,params,status){
var array=[mainSheet,params.sheetName];
for(var i=0; i<array.length;i++){
var dt=spreadSheet.getSheetByName(array[i].toString()).getDataRange();
var data=dt.getValues();
var headers=QSheet.GetFormHeaders(data[0]);
for(var row=1;row<data.length;row++){
 for(var col=0;col<data[row].length; col++){
  if(data[row][headers.answerId].toString()===params.answerId.toString()){
  dt.getCell(row+1, headers.status+1).setValue(status.toString()); 
  dt.getCell(row+1, headers.lastEdit+1).setValue(nowTime.toString()); 
  break; }
   }
  }
 }
},

GetFormHeaders : function (headers){
var res={};
var formHeaders=[];
formHeaders.push('№');
for(var i=0; i<headers.length;i++){ 
var temp=headers[i].toString();
formHeaders.push(temp);
switch(temp){
case mainHeaders[0].toString(): 
res.document=i;//Название документа
break;
case _headers.status.toString(): 
res.status=i;  //статус
break;
case _headers.sender.toString(): 
res.sender=i;  //отправитель
break;
case _headers.answerId.toString(): 
res.answerId=i;//id ответа
break;
case _headers.lastEdit.toString():
res.lastEdit=i;
break;
}
}
res.formHeaders=formHeaders;
return res;
},

ExitingSheets:function(ss) // SpreadSheet ss - таблица в которой необходимо проверить наличие листов
{
var res=[];
var sheets=ss.getSheets();
if(sheets.length===0) return 0;
else
{
for(var i in sheets)
{
if(sheets[i].getName()!=='')
res.push(sheets[i].getName());
}
return res;
}
},

CreateSheetsInExitingTable:function(ss,sheetsArray){ //SpreadSheet ss, Array sheetsArray-массив с названиями листов
if(sheetsArray.length===1)
ss.getActiveSheet().setName(sheetsArray[0]); //убрать лист по умолчанию
else{
ss.getActiveSheet().setName(sheetsArray[0]); //убрать лист по умолчанию
QSheet.SheetExist(ss,ss.getSheets() ,sheetsArray);
}
return ss;
},

CreateSheetsInTable:function(tableName,sheetsArray)// String tableName-имя таблицы, Array sheetsArray-названия листов книги
{
var ss=(File.FileExist(tableName,MimeType.GOOGLE_SHEETS))?SpreadsheetApp.create(tableName):0;
return QSheet.CreateSheetsInExitingTable(ss,sheetsArray);
},

SheetExist:function(spreadSheet,sheets,sheetsArray){  //проверка существующих листов в книге
for(var i=0;i<sheetsArray.length;i++)
{
if(sheets[i]===undefined) spreadSheet.insertSheet(sheetsArray[i]);
else{
if(sheets[i].getName()!==sheetsArray[i]) spreadSheet.insertSheet(sheetsArray[i]);   //если отсутствует - добавить
else continue;
}
}
},

ClearHistory : function(managerMode){
var userSheet=SpreadsheetApp.openById(Data.GetElementInList((managerMode===true)?mngrFile+userEmail:userEmail,MimeType.GOOGLE_SHEETS)[0].ID);
for(var i=0; i<userSheet.getSheets().length; i++){
var data=userSheet.getSheets()[i].getDataRange().getValues();
var length=data.length;
userSheet.getSheets()[i].getRange(2, 1, length, data[length-1].length).clear();
}
return 'Історію документів успішно очищено!';
},
                                  
  DefineStatusCol : function(params){
  var res='';
  var row,col;
  var indexes=[]; // набор индексов с подходящими строками
  for(row=1;row<params.data.length; row++){
  if(params.data[row][params.filters.statusCol].toString()===params.status && params.data[row][params.filters.sender].toString()===userEmail) {
  indexes.push(params.data[row]);//записать строку с необходимым статусом
  }
  }
  var temp=(indexes.length<params.rowCount)?indexes.length:params.rowCount;
  for(row=0;row<temp;row++){res+='<tr class='+getStatus(indexes[row][params.filters.statusCol].toString())+'><td>'+(row+1).toString()+'</td>';
  for(col=0;col<params.filters.sender;col++)
  res+='<td>'+indexes[row][col].toString()+'</td>';
  res+='</tr>';
  }
  return res;
  },
  
  _GetHistory:function(data,params){
   var res='';  
    res="<thead>"; //создать заголовок таблицы
     for(var i=0; i<params._headers.sender+1;i++){
      res+="<td>"+params._headers.formHeaders[i].toString()+ "</td>";
      } res+="</thead>";  
     if(_status[params.status]===_status.ALL) {
       for(var row=1; row<=params.temp-1;row++){ //содержимое
        res+="<tr class=\'"+getStatus(data[row][params._headers.status].toString())+"\'><td>"+row+"</td>";
        for(var col=0; col<params._headers.sender;col++){
          if(data[row][params._headers.sender].toString()===userEmail){
             res+="<td>"+data[row][col].toString()+"</td>";
           }else continue;
        } res+="</tr>";
       }
     } else {
        res+=QSheet.DefineStatusCol({filters:{statusCol:params._headers.status,sender:params._headers.sender},
        rowCount: params.rowCount,
        data:data,
        status:_status[params.status],
         });
        }  
   return res;
  },
  
  _GetAnswers:function(data,params){
     var res='<thead>';
     for(var i=0; i<params._headers.sender+1;i++){
      res+="<td>"+params._headers.formHeaders[i].toString()+ "</td>";
      } 
        res+="<td>"+data[0][params._headers.sender].toString()+"</td>";
        res+="<td>"+data[0][params._headers.lastEdit].toString()+"</td>";
      res+="</thead>";  
      if(_status[params.status]===_status.ALL) {
       for(var row=1; row<=params.temp-1;row++){ //содержимое
        if(data[row][params._headers.document].toString()===params.sheetName){
          res+="<tr class=\'"+getStatus(data[row][params._headers.status].toString())+"\' onclick=\"details(event,\'"+params.ss.getId()+"\',\'"+data[row][params._headers.answerId].toString()+"\',\'"+data[row][params._headers.document].toString()+"\');\"><td>"+row+"</td>";
           for(var col=0; col<params._headers.sender;col++){
             res+="<td>"+data[row][col].toString()+"</td>";
            } 
         res+="<td>"+data[row][params._headers.sender].toString()+"</td>";
         res+="<td>"+data[row][params._headers.lastEdit].toString()+"</td>";
         res+="</tr>";
        }
       }
      }
      else{
       var row,col;
       var indexes=[]; // набор индексов с подходящими строками
            for(row=1;row<data.length; row++){
               if(data[row][params._headers.status].toString()===_status[params.status] && data[row][params._headers.document].toString()===params.sheetName) {
                    indexes.push(data[row]);//записать строку с необходимым статусом 
              }
            }
                var temp=(indexes.length<params.rowCount)?indexes.length:params.rowCount;
                for(row=0;row<temp;row++){
                    res+='<tr class='+getStatus(indexes[row][params._headers.status].toString())+'onclick=\"details(event,\''+params.ss.getId()+'\',\''+data[row][params._headers.answerId].toString()+'\',\''+data[row][params._headers.document].toString()+'\');\"><td>'+(row+1).toString()+'</td>';
                    for(col=0;col<params._headers.sender;col++)
                      res+='<td>'+indexes[row][col].toString()+'</td>';
                      res+="<td>"+indexes[row][params._headers.sender].toString()+"</td>";
                      res+="<td>"+indexes[row][params._headers.lastEdit].toString()+"</td>";
                      res+='</tr>';
       }
      }
     return res;
  },
  
GetHistory:function (params) {
var userSheet;
if(!params.adminMode) userSheet=Data.GetElementInList(userEmail,MimeType.GOOGLE_SHEETS)[0];
else userSheet=Data.GetElementInList(mngrFile+userEmail,MimeType.GOOGLE_SHEETS)[0];
var res;
var empty='Історія документів порожня.';
 if(userSheet===undefined) res=empty;
 else{ 
  var ss=SpreadsheetApp.openById(userSheet.ID);
    if(ss===null) res=empty;
     else{
       var data=ss.getSheetByName((params.adminMode)?mainSheet:params.sheetName).getDataRange().getValues();
        var headers=QSheet.GetFormHeaders(data[0]);
         var temp=(data.length<params.rowCount)?data.length:params.rowCount; //количество выводимых строк
         if(params.rowCount<data.length) temp=++params.rowCount;
         params._headers=headers;
         params.temp=temp;
         params.ss=ss;
         res="<table>";
       res+=(params.adminMode)?QSheet._GetAnswers(data,params):QSheet._GetHistory(data,params);
       res+="</table>";
       }
      }
return res;
},

};   
Data=
{
GetName : function (){
var result;
var params={
method: 'get',
contentType: 'application/json',               
headers: {
Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
},
muteHttpExceptions:true,
};
var response=UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v2/userinfo?alt=json',params);
result=JSON.parse(response.getContentText());
return result;
},

GenerateAnswerId:function(){  //создать ид ответа
 var result       = '';
      var words        = '0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM-';
        var max_position = words.length - 1;
            for( i = 0; i < max_position; i++ ) {
                position = Math.floor ( Math.random() * max_position );
                result = result + words.substring(position, position + 1);
            }
        return result;
},

CheckAnswerId:function(answerId,sheet){ //проверить ид по таблице, true если уже существует данный ид
var res=false;
var uaid=sheet.getDataRange().getValues();
for(var row=1;row<uaid.length;row++){
 if(uaid[row][0].toString()===answerId) res=true; 
}
return res;
},

CreateAnswerId:function(){
var file=Data.GetElementInList('General',MimeType.GOOGLE_SHEETS)[0];
if(file!==undefined){
var res='';
var uaid=SpreadsheetApp.openById(file.ID).getSheetByName('UAID');
 var answerId=Data.GenerateAnswerId();
 if(Data.CheckAnswerId(answerId,uaid)) Data.CreateAnswerId();
 else {
  uaid.appendRow([answerId]).sort(1,true);
  res=answerId;
   }
   return res;
 }else return undefined;
},

DefineTitle : function(item){
var name=Data.GetName();
var title=item.getTitle();
var result='';
switch(title){
case 'ПІБ':
result=name.family_name+" "+name.given_name;
break;
case "Прізвище Ім'я По-батькові":
result=name.family_name+" "+name.given_name;
break;
case "Ім'я":
result=name.given_name;
break;
case "Прізвище":
result=name.family_name;
break;
default:
result='';
break;
}
return result;
},

FindFilesList :function(rootId){
var res;
if(rootId===undefined ) return res;
var files=DriveApp.getFolderById(rootId).getFilesByName('FilesList');
while(files.hasNext()){
var file=files.next();
if(file.getMimeType()===MimeType.GOOGLE_SHEETS){
res=file.getId();
break;
}
}
return res;
},

AddToFilesList: function (file,rootId){ //Object file- объект файла который необходимо записать в таблицу,String rootId- id корневой папки
var fileId=file.getId();
var table=SpreadsheetApp.openById((root.FILESLIST===undefined)?Data.FindFilesList(rootId):root.FILESLIST);
var sheet;
if(DriveApp.getFileById(fileId).getMimeType()!== MimeType.FOLDER) {
table.getSheetByName('Files').appendRow([file.getName(),fileId,DriveApp.getFileById(fileId).getMimeType(),userEmail,nowTime]).sort(1,true);
}
else {
table.getSheetByName('Folders').appendRow([file.getName(),fileId,DriveApp.getFileById(fileId).getMimeType(),userEmail,nowTime]).sort(1,true);
}
autoWidth(table);
return 0;
},

SetRes:function(name,id,mimeType,author,rowIndex){ 
var res={};
res['NAME']=name;
res['ID']=id;
res['MIME']=mimeType;
res['AUTHOR']=author;
res['ROWINDEX']=rowIndex;
return res;
},

GetElementInList:function(name,mimeType,id){ //поиск по глобальной таблице файлов String name-название файла, MimeType mimeType - тип файлов
var file=(root.FILESLIST===undefined)?Data.FindFilesList(id):root.FILESLIST;
if(file===undefined) return undefined;
var f=SpreadsheetApp.openById(file);
var data;
var res=[];
if(mimeType!== MimeType.FOLDER) {
data=f.getSheetByName('Files').getDataRange().getValues();
for ( var i = 0; i < data.length;i++){
if (data[i][0] === name && data[i][2] === mimeType){
res.push(Data.SetRes(data[i][0],data[i][1],data[i][2],data[i][3],i));
break;
}
}
}else if(mimeType=== MimeType.FOLDER) {
data=f.getSheetByName('Folders').getDataRange().getValues();
for ( var i = 0; i < data.length;i++){
if (data[i][0] === name){
res.push(Data.SetRes(data[i][0],data[i][1],data[i][2],data[i][3],i));
break;
}
}
}else res=undefined;
return res;
},

AddDirector : function(email,sheetName){
var file=SpreadsheetApp.openById(Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(sheetName);
file.appendRow([email]).sort(1,true);
return "Успішно додано!\nОновіть сторінку.";
},

AddAdmin : function(email){
return Data.AddDirector(email,'Admins');
},

AddManager : function(email){
return Data.AddDirector(email,mngSheet);
},

DeleteTemplate : function(params) {
 var deleteIndex=0;
 var res={};
 var name=formatString(params.name);
 var file=SpreadsheetApp.openById(Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
  var data=file.getDataRange().getValues();
  for ( var i = 1; i < data.length;i++){
  var temp=JSON.parse(data[i][1].toString());
  if(temp.form.toString().search(RegExp(params.id))!==(-1) && data[i][2].toString()===params.destination && temp.name.toString()===name && data[i][0].toString()===params.author){
   deleteIndex=i+1;
   break;
   }
  }
  file.deleteRow(deleteIndex);
  file=SpreadsheetApp.openById(Data.GetElementInList('Domain',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(params.destination);
  data=file.getDataRange().getValues();
  for ( var i = 1; i < data.length;i++){
   if(data[i][0].toString()===params.name && data[i][1].toString()===params.id && data[i][2].toString()===params.author){
   deleteIndex=i+1;
   break;
   }
  }
   file.deleteRow(deleteIndex);
  file=DriveApp.getFolderById(Data.GetElementInList(params.name,MimeType.FOLDER)[0].ID).setTrashed(true);
  var form=Data.GetElementInList(params.name,MimeType.GOOGLE_FORMS)[0].ROWINDEX+1;
  var folder=Data.GetElementInList(params.name,MimeType.FOLDER)[0].ROWINDEX+1;
  var res={
   form:form,
   doc:doc,
   folder:folder,
  };
  file=SpreadsheetApp.openById(root.FILESLIST);
  file.getSheetByName('Files').deleteRow(form);
   file.getSheetByName('Folders').deleteRow(folder); 
  var doc=Data.GetElementInList(params.name,MimeType.GOOGLE_DOCS)[0].ROWINDEX+1;
 file.getSheetByName('Files').deleteRow(doc);
 return 'Успішно видалено!\nОновіть сторінку.';
},

NotifyAdmins: function(title,managerFile){
var admins=[];
var data=managerFile.getSheetByName('Admins').getDataRange().getValues();
for(var i=1; i<data.length; i++) admins.push(data[i][0].toString());
for(var f in admins) {
GmailApp.sendEmail(admins[f],emailSubject,'Менеджер: '+userEmail+' ,завантажив в систему документ:\n'+title+gotoSystem);
}
},

AddTemplate : function(params){ 
// String idForm/idTempl - id формы и шаблона,которые необходимо добавить,String title-нзвание папки и зоголовок документа,String destination - назначение
var f;
var temp=Data.GetElementInList(userEmail+'Templates'+params.destination,MimeType.FOLDER)[0];
if(temp===undefined) {
f=DriveApp.getFolderById(Data.GetElementInList(params.destination,MimeType.FOLDER)[0].ID).createFolder(userEmail+'Templates'+params.destination);
Data.AddToFilesList(f,root.ID);
}
else f=DriveApp.getFolderById(temp.ID);
if(Data.GetElementInList(params.title,MimeType.FOLDER)[0] === undefined ){//если не существует папка с заданным именем
f=f.createFolder(params.title); 
var managerFile=SpreadsheetApp.openById(Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS)[0].ID);
temp=DriveApp.getFileById(params.idForm);
var Form=temp.makeCopy(temp.getName(), f);
temp=DriveApp.getFileById(params.idTempl);
var Doc=temp.makeCopy(Form.getName(), f);
Data.AddToFilesList(f,root.ID);
Data.AddToFilesList(Form,root.ID);
Data.AddToFilesList(Doc,root.ID);
var formName=formatString(Form.getName());
var assoc=JSON.stringify({ name:formName,
form:Form.getId(),
doc:Doc.getId()});
var managerSs;
var headers=_getFormHeaders(FormApp.openById(Form.getId()).getItems());
var managerHistory=Data.GetElementInList(userEmail,MimeType.FOLDER)[0];
if(managerHistory===undefined){
var dst=DriveApp.getFolderById(Data.GetElementInList('Users',MimeType.FOLDER)[0].ID).createFolder(userEmail);
addToFilesList(dst);
managerSs=SpreadsheetApp.create(mngrFile+userEmail);
managerSs.getActiveSheet().setName(mainSheet).appendRow(mainHeaders).setFrozenRows(1);
managerSs.insertSheet(Form.getName()).appendRow(headers).setFrozenRows(1);
addToFilesList(DriveApp.getFileById(managerSs.getId()).makeCopy(managerSs.getName(), dst));
DriveApp.getFileById(managerSs.getId()).setTrashed(true);
}else{
managerSs=SpreadsheetApp.openById(Data.GetElementInList(mngrFile+userEmail,MimeType.GOOGLE_SHEETS)[0].ID);
if(managerSs.getSheetByName(Form.getName())===null) managerSs.insertSheet(Form.getName()).appendRow(headers).setFrozenRows(1);
}
var ss=managerFile.getSheetByName(mainSheet).appendRow([userEmail,assoc,params.destination,nowTime]).sort(1,true);
SpreadsheetApp.openById(Data.GetElementInList('Domain',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(params.destination).appendRow([Form.getName(),Form.getId(),userEmail,(params.category!==undefined)?params.category:""]).sort(1,true);
Data.NotifyAdmins(params.title,managerFile);
}else return "<strong style=\'color:red;\'>Помилка! Документ вже існує.</strong><br><br><button class=\'button_cl\' onclick=\"t(\'addTemplate\');\">Додати інший документ</button>";
return "<strong style=\'color:red;\'>Успішно додано</strong><br><br><button class=\'button_cl\' onclick=\"t(\'addTemplate\');\">Додати ще один документ</button>";
},

IsDirector : function(sheetName){
var managers=SpreadsheetApp.openById(Data.GetElementInList('Managers',MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(sheetName);
var data=managers.getDataRange().getValues();
for ( var i = 0; i < data.length;i++){
if (data[i][0].toString() === userEmail) return true;
}
return false;
},

IsAdmin : function(email) { //проверка на админа
return Data.IsDirector('Admins');
}, 

IsManager : function(email) { //проверка на менеджера
return Data.IsDirector('Managers');
},

SetUserStatus: function(status,ismanager,isadmin){ // String status - статус пользователя, bool ismanager/admin - менеджер/админ?
var res={};
res['STATUS']=status;
res['MANAGER']=ismanager;
res['ADMIN']=isadmin;
return res;
},

ToUpper: function(text,position){
var res='';
if(position===text.toString().length) res=text.toString()[position].toUpperCase();
else res=text.toString()[position].toUpperCase()+text.toString().substr(position+1,text.toString().length);
return res;
},

DefineUser:  function (email)   //проверить полномочия согласно файлу users
{
var res;
if(Data.IsAdmin(email) === true) res=Data.SetUserStatus('Адміністратор',true,true);  
else {
if(Data.IsManager(email) === true) res=Data.SetUserStatus('Менеджер',true,false);
else{
var data = QSheet.GetData(SpreadsheetApp.openById(Data.GetElementInList('Domain',MimeType.GOOGLE_SHEETS)[0].ID));
for(var i=0;i<data.length;i++){
if(parseEmail(email) === data[i][0]){
res = Data.SetUserStatus(data[i][1],false,false); //статус пользователя
break;
} else res = Data.SetUserStatus('',false,false);
}
}
}
return res;
},

ParseEmail : function (email) // получить правую часть email
{
var result='';
for(var i=email.length-1; i!=0;i--)
{
if(email[i]!=='@') result+=email[i];
else break;
}
result=result.split("").reverse().join("");
return result;
},

};

Content={
CreateTemplate : function(item,isRequired,title,helpText){//добавить закрывающий </div>
var res='';
if(item!==undefined){
res+=(isRequired)?"<div id='required' style='margin-bottom:2%;'><font style='color:red;'>*</font><span class='_itemTitle'> ":"<div style='margin-bottom:2%;'><span class='_itemTitle'>";
if(title===undefined) title=(item.getTitle()!=='')?item.getTitle():'Заголовок відсутній';
if(helpText===undefined) helpText=item.getHelpText();
res+=title+"</span>:<div style='margin-bottom:1%;font-size:0.9em;font-style:italic;'>"+helpText+"</div>";}
return res;
},

CreateComment: function (){
var res=Content.CreateTemplate('',false,'Примітка','наприклад, документ потрібно в 2-х екземплярах');
res+="<textarea id='_note' class='textarea_cl'></textarea></div>";
return res;
},

CreateRadio: function(item){  //radiobutton
var res=Content.CreateTemplate(item,item.asScaleItem().isRequired());
res+="<div class='_answer'>";
for(var i=item.asScaleItem().getLowerBound();i<= item.asScaleItem().getUpperBound();i++){
res+=i+"<input type='radio' name='radiobtn' class='_answerItem radiobutton_cl' style='margin-right:1%;' value='"+i+"'>";
res+=(i===item.asScaleItem().getUpperBound())?"<br>":"";
}
res+="</div></div>";
return res;
},

CreateText : function(item){   //input type='text'
var res=Content.CreateTemplate(item,item.asTextItem().isRequired());
res+="<input type='text' class='_answer textbox_cl' value='"+Data.DefineTitle(item)+"'></div>";
return res;
},

CreateTextBox : function(item){   //textarea
var res=Content.CreateTemplate(item,item.asParagraphTextItem().isRequired());
res+="<textarea class='_answer textarea_cl'>"+Data.DefineTitle(item)+"</textarea></div>";
return res;
},

CreateCheckBox : function(item){   //input type='checkbox'
var temp=(item.getType()!==FormApp.ItemType.MULTIPLE_CHOICE)?item.asCheckboxItem():item.asMultipleChoiceItem();
var res=Content.CreateTemplate(item,temp.isRequired());
res+="<div class='_answer'>";
for(var i=0; i<temp.getChoices().length;i++)
res+="<input type='checkbox' class='_answerItem checkbox_cl' value='"+temp.getChoices()[i].getValue()+"'>"+temp.getChoices()[i].getValue()+"<br>";
res+="</div></div>";
return res;
},

CreateTime : function(item){  
var res=Content.CreateTemplate(item,item.asTimeItem().isRequired());
var i=0;
res+="<div class='_answer' id='time'>Години : <select class='_answerItem select_cl' style='width:5%;' name='hours'>";
for(i=0;i<24;i++)
res+="<option value='"+i+"'>"+i+"</option>";
res+="</select> Хвилини : <select class='_answerItem select_cl' style='width:5%;' name='minutes'>";
for(i=0; i<60; i++ )
res+="<option value='"+i+"'>"+i+"</option>";
res+="</select></div></div>";
return res;
},

CreateSelect : function(item){
var temp=item.asListItem();
var res=Content.CreateTemplate(item,temp.isRequired());
res+="<select class='_answer select_cl' style='width:20%;'>";
for(var i=0; i< temp.getChoices().length; i++){
res+="<option value='"+temp.getChoices()[i].getValue()+"'>"+temp.getChoices()[i].getValue()+"</option>";
}
res+="</select></div>";
return res;
},

CreateDate : function(item){  
var res=Content.CreateTemplate(item,item.asDateItem().isRequired());
res+="<div id='date' class='_answer'>День : <input type='text' class='_answerItem select_cl' style='width:5%;' name='day'> Місяць : <select name='month' class='_answerItem select_cl' style='width:5%;'>";
for(var i =1;i<=12;i++)
res+="<option value='"+i+"'>"+i+"</option>";
res+="</select> Рік : <input type='text' class='_answerItem textbox_cl' style='width:5%;' name='year'></div></div>";
return res;
},

CreateLoadFile : function(item){  
var res=Content.CreateTemplate(item,false);
res+="<span>Оберіть файл на вашому Google Disk-у:<br><font style='font-size:0.8em;'>Для вибору декількох файлів:<br>Утримуючи Ctrl оберіть потрібні файли<br>або<br>Тапніть по потрібним файлам</font></span><br><select multiple class='_answer select_cl' size='5' style='width:25%; margin-top:1%;'>";
var temp=DriveApp.getRootFolder().getFiles();
while(temp.hasNext()){
var t=temp.next();
res+="<option value='"+t.getId()+"'>"+t.getName()+"</option>";
}
res+="</select></div>";
return res;
},

DefineItem: function (item){
var res='';
switch(item.getType()){
case FormApp.ItemType.MULTIPLE_CHOICE:
res=Content.CreateCheckBox(item);
break;
case FormApp.ItemType.CHECKBOX:
res=Content.CreateCheckBox(item);
break;
case FormApp.ItemType.DATE:
res=Content.CreateDate(item);
break;
case FormApp.ItemType.TIME:
res=Content.CreateTime(item);
break;
case FormApp.ItemType.TEXT:
res=Content.CreateText(item);
break;
case FormApp.ItemType.PARAGRAPH_TEXT:
res=Content.CreateTextBox(item);
break;
case FormApp.ItemType.SCALE:
res=Content.CreateRadio(item);
break;
case FormApp.ItemType.LIST:
res=Content.CreateSelect(item);
break;
default:
res='';
}
if(item.getType().ordinal()===16)  //FILE_UPLOAD:
res=Content.CreateLoadFile(item); 
return res;
},

CreateHeader : function(header){
 return "<center><span style='color:royalblue;font-size:1.4em;font-weight:bold;'>"+header+"</span></center><hr>";
},

CreateForm : function(formId){
var form=FormApp.openById(formId);
var header=(form.getTitle()!=='')?form.getTitle():DriveApp.getFileById(formId).getName();
var res="<form method='post' id='answerForm' name='_answer'><span id='_formId' style='visibility:hidden;'>"+formId+"</span>"+Content.CreateHeader(header);
for(var i in form.getItems())
res+=Content.DefineItem(form.getItems()[i]);
res+=Content.CreateComment();
res+="<input type='text' style='visibility:hidden;' id='fakeMsg'><br>";
if(triesScore()>0)res+="<input type='button' class='button_cl' value='Надіслати' onclick=\"sendForm(document.forms.namedItem(\'answerForm\'));\">";
else res+="<strong style='color:red'>На жаль, денний ліміт подачі документів вичерпано!</strong>";
res+="</form>";
return res;
},

Include:function(filename) {   //открыть html-страницу
return HtmlService.createHtmlOutputFromFile(filename).getContent();
},

OpenElement:function(name) { //вывести содержимое страницы после выполнения gs
return openPage(name).getContent();
},

OpenPage:function(name) { //получить страницу с выполнением gs
return HtmlService.createTemplateFromFile(name).evaluate();
}
}; 

function getFormData(name){
var res;
var query=formatString(name);
var managers=SpreadsheetApp.openById(Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
var data=managers.getDataRange().getValues();
for(var i=1;i<data.length; i++){
 var temp=JSON.parse(data[i][1].toString());
 if(temp.name.toString().search(RegExp(query))!==(-1)){
res=temp;
break; }
}
return res;
}

function defineManager(formid){ //определить менеджера по ид-формы
var res='';
var managers=SpreadsheetApp.openById(Data.GetElementInList(mngSheet,MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(mainSheet);
var data=managers.getDataRange().getValues();
for ( var i = 1; i < data.length;i++){
if(JSON.parse(data[i][1].toString()).form.toString().search(RegExp(formid))!==(-1)){
res=data[i][0].toString();
break;
}
}
return res;
} 

function fillMainSheet(ss,formName,mainData,res){
if(ss.getSheetByName(formName)!==null) insertRow(ss.getSheetByName(formName),res.results);
else ss.insertSheet(formName).appendRow(res.headers).appendRow(res.results).setFrozenRows(1);
if(ss.getSheetByName(mainSheet)!==null ) insertRow(ss.getSheetByName(mainSheet),mainData);
}  

function acceptAnswer(res,form,note,answerId){ //записать ответы по нужным файлам
var user=Data.GetElementInList(userEmail,MimeType.FOLDER)[0];
var formName=form.getName();
var mainData=[formName,note.toString(),nowTime,_status.INQUERY,userEmail,answerId,nowTime];
var temp;
var manager=defineManager(form.getId());
if(user!==undefined){
temp=user.ID;
}else{
if(userEmail!==manager){
var t=DriveApp.getFolderById(Data.GetElementInList('Users',MimeType.FOLDER)[0].ID).createFolder(userEmail);
addToFilesList(t);
temp=t.getId();
}else temp=DriveApp.getFolderById(Data.GetElementInList(userEmail,MimeType.FOLDER)[0].ID);
} //добавить права доступа к файлам
user=Data.GetElementInList(userEmail,MimeType.GOOGLE_SHEETS)[0];
var managerSs=SpreadsheetApp.openById(Data.GetElementInList(mngrFile.toString()+manager,MimeType.GOOGLE_SHEETS)[0].ID);
fillMainSheet(managerSs,formName,mainData,res);
if(user!==undefined){
temp=SpreadsheetApp.openById(user.ID); //записать ответы для определённой формы
fillMainSheet(temp,formName,mainData,res);
}else{  // запись последних ответов в вверх таблицы
var ss=SpreadsheetApp.create(userEmail);
ss.getActiveSheet().setName(mainSheet).appendRow(mainHeaders).setFrozenRows(1);
insertRow(ss.getSheetByName(mainSheet),mainData);
ss.insertSheet(formName);
var file=DriveApp.getFileById(ss.getId()).makeCopy(ss.getName(),DriveApp.getFolderById(temp));
DriveApp.getFileById(ss.getId()).setTrashed(true);
addToFilesList(file);
temp=SpreadsheetApp.openById(file.getId()).getSheetByName(formName).appendRow(res.headers); //записать ответы для определённой формы
temp.setFrozenRows(1);
insertRow(temp,res.results);
}
var nt=(note!=='')?"\n\nПримітка до відповіді:\n"+note:'';
GmailApp.sendEmail(manager,emailSubject,'На форму \"'+form+"\" була надана відповідь."+nt+gotoSystem);
}   

function _getFormHeaders(items){
var headersArray=[];
for(var i=0; i<items.length;i++){
headersArray.push(items[i].getTitle());
}
for(var s in additionHeaders) headersArray.push(additionHeaders[s].toString());
return headersArray; 
}   

function addMeta(id,res,note,answerId) {  //дописать необходимые данные к ответам формы и заголовкам формы
res.push({title:_headers.status.toString(),
value:_status.INQUERY});
res.push({title:_headers.note.toString(),
value:note});
res.push({title:_headers.time.toString(),
value:nowTime});
res.push({title:_headers.sender.toString(),
value:userEmail.toString()});
res.push({title:_headers.answerId.toString(),
value:answerId.toString()});
res.push({title:_headers.lastEdit.toString(),
value:nowTime.toString()});
var items=FormApp.openById(id).getItems(); //получить заголовки формы
var headersArray=_getFormHeaders(items);
var resultsArray=[];//переписать массив через splice
for(var s in res) resultsArray.push(res[s].value); //получить ответы с формы
return {headers:headersArray,results:resultsArray};
}

function sendForm(res,id,note){
var form=DriveApp.getFileById(id); 
var answerId=createAnswerId();
var res=addMeta(id,res,note,answerId);
acceptAnswer(res,form,note,answerId);
scoreDecrement();
return 'Ваш запит успішно збережено!<br><br><button class=\'button_cl\' onclick=\"t(\'forms\');\">Надіслати ще один документ</button>';
}
                                
function insertRow(sheet, rowData, optIndex) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { 
    var index = (optIndex===undefined)?1:optIndex;
    sheet.insertRowAfter(index).getRange(index+1, 1, 1, rowData.length).setValues([rowData]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function _getFormData(data,res){
for(var row=1;row<data.length; row++){
  res.push({name:data[row][0].toString(),id:data[row][1].toString(),});
 }
}

function getForms(status) {
var res=[];
var data;
var domain=SpreadsheetApp.openById(getElementInList('Domain',MimeType.GOOGLE_SHEETS)[0].ID);
if(status!==undefined){
 data=domain.getSheetByName(status).getDataRange().getValues();
 _getFormData(data,res);
}else {
 for(var i=0; i<domain.getSheets().length; i++){
  if(domain.getSheets()[i].getName()!==mainSheet){
   data=domain.getSheets()[i].getDataRange().getValues();
   _getFormData(data,res);
  }
 }
}
  return res;
}

function getStatus(text){
var res;
switch(text){
case _status.INQUERY:
res='_INQUERY';
break;
case _status.DONE:
res='_DONE';
break;
case _status.DENIED:
res='_DENIED';
break;
case _status.INPROCESS:
res='_INPROCESS';
break;
}
return res;
}

function createDataTable(data,params){
var res="<button id='_closeButton' class='button_cl' title='Попередня сторінка' onclick=\"details();\">&larr;</button><div style='margin:2% 0 2% 0;'>Назва документа: <span id='_sheetName'>"+params.sheetName.toString()+"</span></div><div><table><thead><tr id='headers'>";
var headers=getFormHeaders(data[0]);
var sender='';
var statusVal='';
for(var i=0; i<=headers.sender+1;i++){
res+="<td>"+headers.formHeaders[i].toString()+ "</td>";
} 
res+="<td>"+headers.formHeaders[headers.lastEdit+1].toString()+ "</td>";
res+="</tr></thead>";  
for(var row=1;row<data.length; row++){ if(data[row][headers.answerId].toString()===params.answerId.toString()) { 
statusVal=data[row][headers.status].toString();
res+="<tr id=\'"+data[row][headers.answerId].toString()+"\' class=\'"+getStatus(statusVal)+"\'><td>"+row+"</td>";
for(var col=0;col<=headers.sender;col++) res+="<td>"+data[row][col].toString()+"</td>";
res+="<td>"+data[row][headers.lastEdit].toString()+"</td>";
res+="</tr>";
sender=data[row][headers.sender].toString();
}
}
if(statusVal!==_status.DENIED){
res+="</table></div><span id='_sender' style='display:none;'>"+sender+"</span><div style='margin-top:2%;'><p>Виконати дію та встановити статус документа:</p>";
 res+="<button class='changeStatus button_cl' onclick=\"accept(\'"+params.fileId.toString()+"\');\">Згенерувати документ</button>";
 if(statusVal!==_status.DONE && statusVal!==_status.INPROCESS ) res+="<button class='button_cl changeStatus' onclick=\"denied(\'"+params.fileId.toString()+"\');\">Відхилити</button>";
 if(statusVal!==_status.INQUERY || statusVal===_status.DONE || statusVal===_status.INPROCESS) res+="<button class='button_cl changeStatus' onclick=\"done(\'"+params.fileId.toString()+"\');\">Готово</button>";
}
res+="<div id='reason'><p id='reasonHeader'>Причина:</p><textarea id='reasonText'></textarea><br><button id='sendButton' class='button_cl changeStatus' >Підтвердити</div></div>";
return res;
}

function createDetails(params){
var answerSheet=SpreadsheetApp.openById(Data.GetElementInList(mngrFile+userEmail,MimeType.GOOGLE_SHEETS)[0].ID).getSheetByName(params.sheetName.toString());
var data=answerSheet.getDataRange().getValues();
var res=createDataTable(data,params);
return res;
}

function setStatuses(params,status){ //установить статусы по необходимым файлам
var managerSs=SpreadsheetApp.openById(Data.GetElementInList(mngrFile.toString()+userEmail,MimeType.GOOGLE_SHEETS)[0].ID);
QSheet.SetStatus(managerSs,params,status);
var user=SpreadsheetApp.openById(Data.GetElementInList(params.sender,MimeType.GOOGLE_SHEETS)[0].ID);
if(user!==undefined) QSheet.SetStatus(user,params,status);
return {managerFile:managerSs,
userFile:user,};
}

function denied(params){
var temp=setStatuses(params,_status.DENIED);
var r=(params.reason===''||params.reason===undefined)?'':"Причина:\n"+params.reason;
GmailApp.sendEmail(params.sender, emailSubject, "Поданий Вами документ ("+params.sheetName+") був відхилений.\n\n"+r+gotoSystem);
//return createDataTable(temp.managerFile.getDataRange().getValues(),params);
return "processingAnswers";
}

function done(params){
var temp=setStatuses(params,_status.DONE);
var r=(params.reason===''||params.reason===undefined)?'':"Примітка:\n"+params.reason;
GmailApp.sendEmail(params.sender, emailSubject, "Поданий Вами документ ("+params.sheetName+") готовий.\n\n"+r+gotoSystem);
//return createDataTable(temp.managerFile.getDataRange().getValues(),params);
return "processingAnswers";
}

function formatString(text){
var str='';
for(var i=0;i<text.length;i++){
 if(text[i]==='(' || text[i]===')' || text[i]===' ') str+='';
 else str+=text[i];
}
return str;
}

function getPinnedFormsList(isAdmin){ //список закрепленных за менеджером форм
 var res=[];
 var file=Data.GetElementInList(mngSheet.toString(),MimeType.GOOGLE_SHEETS)[0];
 if(file!==undefined){
  var mngr=SpreadsheetApp.openById(file.ID).getSheetByName(mainSheet).getDataRange().getValues();
  for(var row=1;row<mngr.length; row++){
   if(user.ADMIN) res.push([JSON.parse(mngr[row][1].toString()),mngr[row][2].toString(),mngr[row][3].toString(),mngr[row][0].toString()]);
   else{ if(mngr[row][0].toString()===userEmail){
     res.push([JSON.parse(mngr[row][1].toString()),mngr[row][2].toString(),mngr[row][3].toString()]);
     }
    }
   }
 }
 return res;
}

function deleteSystem(){
DriveApp.getFolderById(root.ID).setTrashed(true);
var attach=SitesApp.getActiveSite().getAttachments();
for(var i=0;i<attach.length; i++){
 attach[i].deleteAttachment();
}
_userProp.deleteAllProperties();
return "Система успішно видалена.\nОновіть сторінку.";
}

function checkDate(){//false если совпадает дата
 var res=false;
 if(_userProp.getProperty(_properties.lastDate)!==null) {
    var lastDate=_userProp.getProperty(_properties.lastDate);
    if(lastDate!==nowDate) {
     _userProp.setProperty(_properties.lastDate, nowDate);
     res=true; } 
     else res=false;
  }
  else {
   _userProp.setProperty(_properties.lastDate,nowDate);
   res=true;
  }
  return res;
}

function scoreDecrement(){
  var limit=(user.ADMIN===true || user.MANAGER ===true)?adminLimit:getSettingByName(_sys_vars.answerLimit);
 var prop=_userProp.getProperty(_properties.score);
 if(prop!==null) {
  if(prop>1 && prop <= limit){
    prop--;
    _userProp.setProperty(_properties.score, prop);
     }
     else _userProp.setProperty(_properties.score, 0);
   }
  else _userProp.setProperty(_properties.score,limit);
  return Math.abs(_userProp.getProperty(_properties.score));
}

function triesScore(){
 // _userProp.deleteAllProperties();
  var limit=(user.ADMIN===true || user.MANAGER ===true)?adminLimit:getSettingByName(_sys_vars.answerLimit);
  if(checkDate()) _userProp.setProperty(_properties.score, limit);
  return Math.abs(_userProp.getProperty(_properties.score));
}
