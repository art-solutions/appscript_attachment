function loadAddOn(e) {
  var messageId = e.gmail.messageId;
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var message = GmailApp.getMessageById(messageId);
  var msgDate = message.getDate();
  var formattedDate = Utilities.formatDate(msgDate, "GMT+0", 'E MM/dd/yyyy');
  var rfcHeaderFrom = message.getFrom();
  var rfcHeaderFromMailRegex = /[^< ]+(?=>)/g;
  var obj = {};
  obj.message = message;
  obj.msgId = messageId;
  obj.id = message.getId();
  obj.Name = message.getFrom();
  obj.subj = message.getSubject();
  obj.Body = message.getPlainBody();
  obj.date = formattedDate;

  var attachments = message.getAttachments();
  if (attachments.length > 0) {
    obj.atts = attachments[0].getName();
  } else {
    obj.atts = "No attachments";
  }

  if (rfcHeaderFrom.match(rfcHeaderFromMailRegex) != null) {
    obj.email = rfcHeaderFrom.match(rfcHeaderFromMailRegex)[0];
  } else {
    obj.email = rfcHeaderFrom;
  }

  return DApp(obj);
}

function onHomepage(e) {
  var obj = {};
  return DApp(e);
}

function DApp(obj,action) {
  var value = PropertiesService.getUserProperties().getProperty("Sheet");

  if (value != null) {
    var ss = SpreadsheetApp.openById("148xKFGMVgAZgptLXU59nrn-0Iv9okWTVmQ87jVXhiL4");  // Replace with your actual Spreadsheet ID
    SpreadsheetApp.setActiveSpreadsheet(ss);
    var sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);

    if(action == "submit"){
      firstOccurrence = "submit"
      var name = obj[1]
      var email = obj[2]
    }
    else{
      var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
      var textFinder = range.createTextFinder(obj.msgId);
      var firstOccurrence = textFinder.findNext();
      var names = obj.Name                                               // Ce variable contient le Nom et le prenom et l'email exite dans le champ From
      var email = obj.email
      var reg = /(.+)( )/g;                                              // Ce variable est un regex qui regroupe une chaine de caractaire a deux groupes

      if(names.match(reg) != null){
          var name = names.match(reg)[0]                                    // Ce variable permet de separe le nom et le prenom dans un groupe 
      }else{
          var name = " "                                // Ce variable permet de separe le nom et le prenom dans un groupe 
      }
    }

    var cardSection1Divider1 = CardService.newDivider();
    var card = CardService.newCardBuilder();
    var section = CardService.newCardSection();
    var Nothing = CardService.newAction().setFunctionName('Nothing');

    var decoratedTextName = CardService.newDecoratedText()
      .setText(name)
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSUepaBdMZtoy5GmiKF_v1vkRbwo3MgxAiIwcaztDaqiYwLdV58jhq19hUX00btfdkBUF8&usqp=CAU"))

    section.addWidget(decoratedTextName)

    var decoratedTextEmail = CardService.newDecoratedText()
      .setText(email)
      .setStartIcon(CardService.newIconImage()
        .setIconUrl("https://cdn-icons-png.flaticon.com/512/2374/2374459.png"))

    section.addWidget(decoratedTextEmail)

    if(firstOccurrence != null){
      if(firstOccurrence != "submit"){
        var row = firstOccurrence.getRow();
        var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
        var data = rowData[0];
        var id = obj.id
      }
      else{
        var data = obj
        var id = data[12]
      }

      section = CardService.newCardSection();
      var cardSection1Divider1 = CardService.newDivider();

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
      //section.addWidget(radioGroup)

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
        .setText(id)
        .setStartIcon(CardService.newIconImage()
          .setIconUrl("https://e7.pngegg.com/pngimages/417/875/png-clipart-computer-icons-encapsulated-postscript-id-card-miscellaneous-text-thumbnail.png"))

      section.addWidget(decoratedTextid)

      var decoratedTextStatus = CardService.newDecoratedText()
          .setText(data[2])
          .setStartIcon(CardService.newIconImage()
            .setIconUrl("https://www.pngfind.com/pngs/m/75-756588_png-file-single-user-icon-png-transparent-png.png"))

      section.addWidget(decoratedTextStatus)
              
      var decoratedTextFolder = CardService.newDecoratedText()
          .setText(data[3])
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
      //.setOnChangeAction(CardService.newAction().setFunctionName("onModeChange"));

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

      Logger.log(section)

      var Sheet = CardService.newAction()
        .setFunctionName('Sheet')
        .setParameters({
            Name: name,
            email: obj.email,
            id: obj.id,
            date: obj.date,
            attachments: obj.atts,
        })
        .setLoadIndicator(CardService.LoadIndicator.SPINNER);


      var textButton = CardService.newTextButton()
        .setText("Submit")
        .setOnClickAction(Sheet)
        .setBackgroundColor("#482ff7")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

      section.addWidget(textButton)   
      }
      card.addSection(section)
      return card.build(); 
    }
  else{   
      return Add();
  } 
}

function Nothing(e){                         // cette fontion ne fait rien mais utilisable dans le code 
  Logger.log("Nothing")
}

function createFolderIfNotExists(folderName) {
  var folder;
  var existingFolders = DriveApp.getFoldersByName(folderName);

  if (existingFolders.hasNext()) {
    folder = existingFolders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
 
  return folder; // Return the ID of the folder for future reference
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

function Sheet(data, obj) {
  var ss = SpreadsheetApp.openById("148xKFGMVgAZgptLXU59nrn-0Iv9okWTVmQ87jVXhiL4"); // Replace with your actual Spreadsheet ID
  var sheet = ss.getSheets()[0]; // Assuming you want to work with the first sheet


  // Extract data from the input object
  var msg_id = data.messageMetadata.messageId || "";
  var name = data.parameters.Name || "";
  var email = data.parameters.email || "";
  var status = data.formInput.Types|| "";
  var notes = data.formInput.Note|| "";
  var herr = data.formInput.Herr || "";
  var company = data.formInput.Company || "";
  var address = data.formInput.Address || "";
  var phone = data.formInput.Phone || "";
  var tax_number = data.formInput.Tax || "";
  var id = data.parameters.id || "";

  // Create a folder in Google Drive to save attachments
  var folder = createFolderIfNotExists("order");
  var folderId = folder.getId();
  var folderUrl = folder.getUrl(); // Get the URL of the folder

  // Get the Gmail message associated with the messageId
  var messageId = data["messageMetadata"]["messageId"];
  var accessToken = data["messageMetadata"]["accessToken"];
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var message = GmailApp.getMessageById(messageId);
  
  // Get the email attachments
  var attachments = message.getAttachments();

  
  // Save attachments to the created folder with renamed filenames
  for (var k = 0; k < attachments.length; k++) {
    var attachment = attachments[k];
    var attachmentName = name + "_" + Utilities.formatDate(new Date(), "GMT+0", "dd_MM_yyyy") + "_" + attachment.getName();
    var file = DriveApp.createFile(attachmentName, attachment);
    file.moveTo(DriveApp.getFolderById(folderId));
  }


  result_array = [msg_id, name, email, status,folderUrl,notes, folderId, herr, company, address, phone, tax_number,id];

  // Append the data to the sheet
  sheet.appendRow(result_array);

  return DApp(result_array, "submit")
}
