function onGmailMessage(e) {
  var messageId = e.gmail.messageId;
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var message = GmailApp.getMessageById(messageId);
  var msgDate = message.getDate();
  var formattedDate = Utilities.formatDate(msgDate, "GMT+0", 'E MM/dd/yyyy');
  var rfcHeaderFrom = message.getHeader("From");
  var rfcHeaderFromMailRegex = /[^< ]+(?=>)/g;
  var obj = {};
  obj.message = message;
  obj.msgId = messageId;
  obj.id = message.getId();
  obj.Name = message.getFrom();
  obj.subj = message.getSubject();
  obj.Body = message.getPlainBody();
  obj.date = formattedDate;
  // var attachment = message.getAttachments()[0];
  if (message.getAttachments()[0] != undefined) {
    var att = message.getAttachments()[0];
    obj.atts = att.getName();
  } else {
    obj.atts = "No attachements"
  }
  if (rfcHeaderFrom.match(rfcHeaderFromMailRegex) != null) {
      obj.email = rfcHeaderFrom.match(rfcHeaderFromMailRegex)[0];
  }else{
      obj.email = rfcHeaderFrom;
  }
  return DApp(obj);
}
function onHomepage(e) { // cette fonction de demarrage  
      return DApp(e);
}
function DApp(obj) { 
 // PropertiesService.getUserProperties().deleteProperty("Sheet")

  var value = PropertiesService.getUserProperties().getProperty("Sheet");
  if(value != null){    
  var val = String(value);
  var ss = SpreadsheetApp.openById("1AtTPRZvdC_GF0yIk7jOLZf25YA3-oOkN-KNhJ1ngT5I"); 
  SpreadsheetApp.setActiveSpreadsheet(ss);
  sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  var textFinder = ss.createTextFinder(obj.id);
  var firstOccurrence = textFinder.findNext();
  var cardSection1Divider1 = CardService.newDivider();                     // Declaration d'un divider 

  var card = CardService.newCardBuilder()                                          // declaration d'une Card
  var section = CardService.newCardSection()                               // Declaration d'une CardSection 
  var Nothing = CardService.newAction().setFunctionName('Nothing');      // Cet function ne fait rien 
  section.addWidget(CardService.newTextButton()
    .setText("Contact Details")
    .setOnClickAction(Nothing)) 

  var names = obj.Name                                               // Ce variable contient le Nom et le prenom et l'email exite dans le champ From
  var reg = /(.+)( )/g;                                              // Ce variable est un regex qui regroupe une chaine de caractaire a deux groupes
  if(names.match(reg) != null){
      var name = names.match(reg)[0]                                    // Ce variable permet de separe le nom et le prenom dans un groupe 
  }else{
      var name = " "                                // Ce variable permet de separe le nom et le prenom dans un groupe 
  }
  var decoratedTextName = CardService.newDecoratedText()
      .setText(name)
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSUepaBdMZtoy5GmiKF_v1vkRbwo3MgxAiIwcaztDaqiYwLdV58jhq19hUX00btfdkBUF8&usqp=CAU"))

  section.addWidget(decoratedTextName)

  var decoratedTextEmail = CardService.newDecoratedText()
      .setText(obj.email)
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://cdn-icons-png.flaticon.com/512/2374/2374459.png"))

  section.addWidget(decoratedTextEmail)

  if(firstOccurrence != null){
      Logger.log(firstOccurrence.getValue())

    var row = firstOccurrence.getRow();
    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
    var data = rowData[0];
    Logger.log(data[1])

  section.addWidget(cardSection1Divider1)                             // Ce divider pour l'organisation et le design 
  var Nothing = CardService.newAction().setFunctionName('Nothing');      // Cet function ne fait rien 
  section.addWidget(CardService.newTextButton()
    .setText("Input Details")
    .setOnClickAction(Nothing))                                         // Design de text 

  var Phone_Input =  CardService.newDecoratedText()
      .setText(data[10])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://openclipart.org/image/2000px/262221"))

  var Company_Input = CardService.newDecoratedText()
      .setText(data[8])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://cdn-icons-png.flaticon.com/512/2083/2083337.png"))


  var Address_Input = CardService.newDecoratedText()
      .setText(data[9])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://www.pngfind.com/pngs/m/128-1288122_marker-circle-comments-address-icon-clipart-hd-png.png"))

  var Tax_Input = CardService.newDecoratedText()
      .setText(data[11])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://cdn-icons-png.flaticon.com/512/3408/3408755.png"))

  section.addWidget(Phone_Input)
  section.addWidget(Company_Input)
  section.addWidget(Address_Input)
  section.addWidget(Tax_Input)
//  section.addWidget(radioGroup)

  var action = CardService.newAction()
    .setFunctionName('Contact')
    .setParameters({ Name: obj.Name, email: obj.email})
    .setLoadIndicator(CardService.LoadIndicator.SPINNER)

  var textButton = CardService.newTextButton()
    .setText("Submit")
    .setOnClickAction(Nothing)
    .setBackgroundColor("#EBEBE4")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setDisabled(true)

  section.addWidget(textButton)

  section.addWidget(cardSection1Divider1)                             // Ce divider pour l'organisation et le design 
  var Nothing = CardService.newAction().setFunctionName('Nothing');      // Cet function ne fait rien 
  section.addWidget(CardService.newTextButton()
    .setText("Job Preference")
    .setOnClickAction(Nothing))                                         // Design de text 

 var decoratedTextid = CardService.newDecoratedText()
      .setText(obj.id)
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://e7.pngegg.com/pngimages/417/875/png-clipart-computer-icons-encapsulated-postscript-id-card-miscellaneous-text-thumbnail.png"))

  section.addWidget(decoratedTextid)

   var decoratedTextStatus = CardService.newDecoratedText()
      .setText(data[3])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://www.pngfind.com/pngs/m/75-756588_png-file-single-user-icon-png-transparent-png.png"))

  section.addWidget(decoratedTextStatus)
          
 var decoratedTextFolder = CardService.newDecoratedText()
      .setText(data[4])
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://pixlok.com/wp-content/uploads/2022/01/Folder-Icon-SVG-pmsnhdjei.png"))

  section.addWidget(decoratedTextFolder)
  }
  else{
    
  section.addWidget(cardSection1Divider1)                             // Ce divider pour l'organisation et le design 
  var Nothing = CardService.newAction().setFunctionName('Nothing');      // Cet function ne fait rien 
  section.addWidget(CardService.newTextButton()
    .setText("Input Details")
    .setOnClickAction(Nothing))                                         // Design de text 

  var payment = CardService.newSelectionInput()
  .setType(CardService.SelectionInputType.DROPDOWN)
  .setTitle("Herr")
  .setFieldName("Herr")
  .addItem("Herr", "Herr", true)
  .addItem("Frau", "Frau", false)
 // .setOnChangeAction(CardService.newAction().setFunctionName("onModeChange"));

  var Phone_Input = CardService.newTextInput()
    .setFieldName("Phone")
    .setTitle("Please enter your Phone Number")

  var Note_Input = CardService.newTextInput()
    .setFieldName("Note")
    .setTitle("Please enter your Note")

  var Company_Input = CardService.newTextInput()
    .setFieldName("Company")
    .setTitle("Please enter company")

  var Address_Input = CardService.newTextInput()
    .setFieldName("Address")
    .setTitle("Please enter Address")
    
  var Tax_Input = CardService.newTextInput()
    .setFieldName("Tax")
    .setTitle("Please enter Tax id")

  var radioGroup = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName("Types")
    .addItem("Quote", "Quote", true)
    .addItem("Ordred", "Ordred", false)

  section.addWidget(payment);
  section.addWidget(Phone_Input)
  section.addWidget(Note_Input)
  section.addWidget(Company_Input)
  section.addWidget(Address_Input)
  section.addWidget(Tax_Input)
  section.addWidget(radioGroup)

  var Sheet = CardService.newAction()
    .setFunctionName('Sheet')
    .setParameters({ Name: name, email: obj.email, id: obj.id, date: obj.date, attachements: obj.atts})
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var textButton = CardService.newTextButton()
    .setText("Submit")
    .setOnClickAction(Sheet)
    .setBackgroundColor("#482ff7")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)

  section.addWidget(textButton)   
  }
  card.addSection(section)
  return card.build(); 
  }else{   
      return Add();
  } 
}function Sheet(e) {
  var attachments = e.parameters.attachements
if(attachments != "No attachements" || attachments != undefined){
  var Date = e.parameters.date
  var folderName = name+'_'+Date
  var Status = e.formInput.Types
  createFolderIfNotExists(Status)
  var GDID = DriveApp.createFolder(Status).createFolder(folderName)
  var GDP = folderName
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var folderName = folderName; // Replace with the desired folder name
        var folderId = createFolderIfNotExists(folderName);
        var file = DriveApp.createFile(attachment, attachment);
        file.moveTo(DriveApp.getFolderById(folderId));
      }
}else{
    var Status = e.formInput.Types
    var GDID = "    "
    var GDP = "   "
}
  var name = e.parameters.Name
  var mail = e.parameters.Email
  var Id = e.parameters.id
  var Address = e.formInput.Address
  var Company = e.formInput.Company
  var Phone = e.formInput.Phone
  var Tax = e.formInput.Tax
  var Note= e.formInput.Note
  var Herr= e.formInput.Herr

  var Data = [Id,name,mail,Status,GDP,Note,GDID,Herr,Company,Address,Phone,Tax];
  var val = PropertiesService.getUserProperties().getProperty('Sheet')

  var ss = SpreadsheetApp.openById("1AtTPRZvdC_GF0yIk7jOLZf25YA3-oOkN-KNhJ1ngT5I"); 
  SpreadsheetApp.setActiveSpreadsheet(ss);
  sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  sheet.appendRow(Data);
  return Refrech();
}
function Nothing(e){                         // cette fontion ne fait rien mais utilisable dans le code 
  Logger.log("Nothing")
}
function createFolderIfNotExists(folderName) {
  var folder;
  var folderName = folderName
  var existingFolders = DriveApp.getFoldersByName(folderName);

  if (existingFolders.hasNext()) {
    folder = existingFolders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
 
  return folder.getId();
}
function attachements(attachments) {
  
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var folderName = "Your_Folder_Name"; // Replace with the desired folder name
        var folderId = createFolderIfNotExists(folderName);
        var file = DriveApp.createFile(attachment);
        file.moveTo(DriveApp.getFolderById(folderId));
      }
}
function Add() {
  var cardSection1Divider1 = CardService.newDivider();
  let card1 = CardService.newCardBuilder()
  var sec = CardService.newCardSection()

  var decoratedTextContact = CardService.newDecoratedText()
    .setText("Add your ID of google sheet")
    .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))

  sec.addWidget(decoratedTextContact)
  sec.addWidget(cardSection1Divider1)

  var API_Input = CardService.newTextInput()
    .setFieldName("SKey")
    .setTitle("Please enter your Token Key")

  sec.addWidget(API_Input)
  
  var action = CardService.newAction()
    .setFunctionName('Cache')
    .setLoadIndicator(CardService.LoadIndicator.SPINNER)

  var textButton = CardService.newTextButton()
    .setText("Submit")
    .setOnClickAction(action)
    .setBackgroundColor("#482ff7")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)

  sec.addWidget(textButton)
  card1.addSection(sec)
  return card1.build();
}
function Cache(e) {
   var cardx = CardService.newCardBuilder()
  var sect = CardService.newCardSection()

  var Nothing = CardService.newAction()
    .setFunctionName('Nothing');

  sect.addWidget(CardService.newTextButton()
    .setText('Congratulation your API was added correctly')
    .setOnClickAction(Nothing))

    var v = e.formInput.Skey
  
  // Store a value
    PropertiesService.getUserProperties().setProperty("Sheet", String(v));

  var actionx = CardService.newAction()
    .setFunctionName('onGmailMessage')

  var textButtonx = CardService.newTextButton()
    .setText("Back")
    .setOnClickAction(actionx)
    .setBackgroundColor("#482ff7")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)

  sect.addWidget(textButtonx)
  cardx.addSection(sect)
  return cardx.build();
}
function Refrech() {
  let card2 = CardService.newCardBuilder()
  var sec1 = CardService.newCardSection()

  var action1 = CardService.newAction()
    .setFunctionName('onGmailMessage')

  var textButton1 = CardService.newTextButton()
    .setText("Refrech")
    .setOnClickAction(action1)
    .setBackgroundColor("#482ff7")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)

  sec1.addWidget(textButton1)
  card2.addSection(sec1)
  return card2.build();
}
