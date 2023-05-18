//function run on sheet opening loads menu items
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Doc Generation')
      .addItem('Generate Proposal', 'genProposal')
      .addSubMenu(ui.createMenu('Generate SOW')
        .addItem('Geneate Managed Services or Agile SOW', 'genMSAgileSOW')
        .addItem('Geneate Ad Hoc Small SOW', 'genAdHocSmallSOW')
        .addItem('Geneate Ad Hoc Big SOW', 'genAdHocBigSOW')
        .addItem('Geneate Org Assessment SOW', 'genOrgAssessmentSOW'))
      .addToUi();
}

//function for menu item that runs the code to generate a proposal
function genProposal() {
  presentationid = mergeProposal('1HomquqMhAuR7eqe67ftyrMbufyUle5kjOG2l8mXQSXM');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  showAnchor('Open the Proposal Deck','https://docs.google.com/presentation/d/' + presentationid + '/edit#',"Proposal Generation");
}

//function for menu item that runs the code to generate an MSA/SOW     
function genMSAgileSOW() {
  docid = mergeMSA();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  showAnchor('Open the SOW Doc','https://docs.google.com/document/d/' + docid + '/edit#',"SOW Generation");
}

//fuction that generates an MSA/SOW
function mergeMSA() {
  
  const customerValuesRange = 'Customer!B1:B50'; //range for the account fields
  const msValuesRange = 'ProjectPricing!C20:C21'; //range for the ms quantity fields
  const msMilestonesRange = 'ManagedServicesMilestones!B2:B20'; //range for the ms milestones fields

    let dollarUSLocale = Intl.NumberFormat('en-US',{style: "currency",currency: "USD",minimumFractionDigits: 0,maximumFractionDigits: 0,});

  //id for the estimation sheet for the current project
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  dataSpreadsheetId = ss.getId();

  try {
      //get the account range in the estimate spreadsheet
      let customerValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(customerValuesRange).getValues();
      //create a new merged slide deck for the first line in the sheet
      
      const companyName = customerValues[3][0];
      const companyEntity = customerValues[9][0];
      const companyJurisdiction = customerValues[8][0];
      const effectiveDate = customerValues[22][0];
      const year = customerValues[23][0];
      const sowNumber = customerValues[20][0];
      const address = customerValues[7][0];
      const customerName = customerValues[4][0];
      const customerTitle = customerValues[5][0];
      const customerEmail = customerValues[6][0];
      const bwAccountManagerName = customerValues[26][0];
      const bwAccountManagerEmail = customerValues[27][0];
      const sowPrice = dollarUSLocale.format(customerValues[24][0]);
      const sowHalfPrice = dollarUSLocale.format(customerValues[25][0]);
      const agileProject = customerValues[16][0];
      const agileDuration = customerValues[15][0];

      if(agileProject){
        //use the agile SOW template
        templateMSAId = '11Gs70RUgdpQ7igFs-xGeRXarTZ_M8gF91td94ZNxf3I';
      } else {
        //use the managed services SOW template
        templateMSAId = '1QPc6y2f9q6vVx5bMDXveVPYNmpuChFDqHBoJoubfmEE';
      }
     
     //get the ms range in the estimate spreadsheet
      let msValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(msValuesRange).getValues();
      
      const msMonths = msValues[0][0];
      const msFTELevel = msValues[1][0];

      //get the account range in the estimate spreadsheet
      let msMilestones = SpreadsheetApp.openById(dataSpreadsheetId).getRange(msMilestonesRange).getValues();
      milestones = '';
      for(let i=0;i<msMilestones.length;i++){
          const msMilestonesRow = msMilestones[i]; //get the row
          if(msMilestonesRow[0]!=''){
            milestones += msMilestonesRow[0] + '\n';
          }
      }

      // Duplicate the template MSA using the Drive API.
      const copyTitle = companyName + ' MSA';
      let copyFile = {
        title: copyTitle,
        parents: [{id: 'root'}]
      };

      copyFile = Drive.Files.copy(copyFile, templateMSAId); //make a copy of the template
      const msaCopyId = copyFile.id;

      const newMSADoc = DocumentApp.openById(msaCopyId);
      const newMSABody = newMSADoc.getBody();
      const newMSAHeader = newMSADoc.getHeader();

      newMSABody.replaceText("[(][(]effective-date[)][)]",effectiveDate);
      newMSABody.replaceText("[(][(]year[)][)]",year);

      newMSABody.replaceText('[(][(]company-address[)][)]',address);
      newMSABody.replaceText('[(][(]company-name[)][)]',companyName);
      newMSABody.replaceText('[(][(]company-entity[)][)]',companyEntity);
      newMSABody.replaceText('[(][(]company-jurisdiction[)][)]',companyJurisdiction);

      newMSABody.replaceText('[(][(]customer-name[)][)]',customerName);
      newMSABody.replaceText('[(][(]customer-title[)][)]',customerTitle);
      newMSABody.replaceText('[(][(]customer-email[)][)]',customerEmail);
      newMSABody.replaceText('[(][(]sow-number[)][)]',sowNumber);

      newMSABody.replaceText('[(][(]bw-am-name[)][)]',bwAccountManagerName);
      newMSABody.replaceText('[(][(]bw-am-email[)][)]',bwAccountManagerEmail);

      newMSABody.replaceText('[(][(]ms-months[)][)]',msMonths);
      newMSABody.replaceText('[(][(]ms-fte-level[)][)]',msFTELevel);
      newMSABody.replaceText('[(][(]ms-milestones[)][)]',milestones);

      newMSABody.replaceText('[(][(]agile-sowprice[)][)]',sowPrice);
      newMSABody.replaceText('[(][(]agile-half-sowprice[)][)]',sowHalfPrice);
      newMSABody.replaceText('[(][(]agile-duration[)][)]',agileDuration);

      newMSAHeader.replaceText('[(][(]effective-date[)][)]',effectiveDate);

      newMSADoc.saveAndClose();

    return msaCopyId;

  } catch (err) {
      // TODO (Developer) - Handle exception
      console.log('Failed with error: %s', err.error);
    }
}

//function that generates a Proposal
function mergeProposal(templatePresentationId) {

   if(templatePresentationId == null) {
    templatePresentationId = '1HomquqMhAuR7eqe67ftyrMbufyUle5kjOG2l8mXQSXM';
  }

  let dollarUSLocale = Intl.NumberFormat('en-US',{style: "currency",currency: "USD",minimumFractionDigits: 0,maximumFractionDigits: 0,});

  let responses = [];

  const customerValuesRange = 'Customer!B1:B50'; //range for the account and proposal fields
  const featureSetValuesRange = "FeatureSets!A2:W13"; //range for the feature set fields
  const featureSetItemsValuesRange = "FeatureSetItems!A2:Y200"; //range for the feature set item fields

  //id for the estimation sheet for the current project
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  dataSpreadsheetId = ss.getId();

  //array for the clean up of all tags not used
  featureSetTags2 = [];
    for (let i = 0; i < 12; ++i) {
      //for numbers less than 10 we need to add in a 0
      if(i<9){
        fsNumber = "0" + (i+1).toString();
      } else {
        fsNumber = (i+1).toString();
      }
      //create JSON replace objects for each feature set level tag
      featureSetTags2.push("((FS" + fsNumber + "-number))");
      featureSetTags2.push("((FS" + fsNumber + "-name))");
      featureSetTags2.push("((FS" + fsNumber + "-description))");
      featureSetTags2.push("((FS" + fsNumber + "-percentage))");
      featureSetTags2.push("((FS" + fsNumber + "-sowprice))");

      for (let i = 0; i < 12; ++i) {
        //create JSON objects for 12 items for each of the feature set level tags
        featureSetTags2.push("((FS" + fsNumber + "-Item-" + (i+1).toString() + "-name))");
      }
    }

  try {
    //get the account range in the estimate spreadsheet
    let customerValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(customerValuesRange).getValues();

    //create a new merged slide deck for the first line in the sheet
    const companyName = customerValues[3][0]; // name in column 1
    const proposalName = customerValues[13][0]; // proposal name in column 2
    const proposalDate = Utilities.formatDate(customerValues[14][0], "GMT+1", "MM.dd.yyyy"); // date in column 3
    const sowprice =   dollarUSLocale.format(customerValues[24][0]);

    //craft JSON objects for account level replacements
    requestJSON = '{"replaceAllText":{"containsText":{"text":"((customer-name))","matchCase":true},"replaceText":"' + companyName + '"}},{"replaceAllText":{"containsText":{"text":"((proposal-name))","matchCase":true},"replaceText":"' + proposalName + '"}},{"replaceAllText":{"containsText":{"text":"((proposal-date))","matchCase":true},"replaceText":"' + proposalDate + '"}},{"replaceAllText":{"containsText":{"text":"((sowprice))","matchCase":true},"replaceText":"' + sowprice + '"}}';


    //get the feature set range in the estimate spreadsheet
    let featureSetValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(featureSetValuesRange).getValues();

    
    //loop through all the feature set rows
    for(let i=0;i<featureSetValues.length;i++){
      const featureSetRow = featureSetValues[i]; //get the row
      //only process for active rows
      if(featureSetRow[0] != "") {
        //create JSON replace objects and append to the full JSON
        requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-number))',featureSetRow[0]);
        requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-name))',featureSetRow[2]);
        requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-description))',featureSetRow[3]);
        requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-percentage))',Math.floor(featureSetRow[9] * 100) + "%");
        requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-sowprice))',dollarUSLocale.format(featureSetRow[8]));
      }
    }

    //get the feature set items range in the estimate spreadsheet
    let featureSetItemsValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(featureSetItemsValuesRange).getValues();
    for(let i=0;i<featureSetItemsValues.length;i++){
      const featureSetItemsRow = featureSetItemsValues[i]; //get the row
      //only process for active rows
      if(featureSetItemsRow[23] != "") {
        //create JSON replace objects and append to the full JSON
        requestJSON += ',' + jsonForReplace('((' + featureSetItemsRow[23] + '-Item-' + featureSetItemsRow[24] + '-name))',featureSetItemsRow[1]);
      }
    }

    //loop through all the tags to replace any not replaced above. this cleans up the document of any remaining tags
    for(let i=0;i<featureSetTags2.length;i++){
      tag = featureSetTags2[i];
      //create JSON replace objects and append to the full JSON
      requestJSON += ',' + jsonForReplace(tag,'');
    }

    //add brackets to JSON string
    requestJSON = "[" + requestJSON + "]";

    //console.log('requestJSON: ' + requestJSON);

    const requests = JSON.parse(requestJSON);
    
    // Duplicate the template presentation using the Drive API.
    const copyTitle = customerName + ' Salesforce Proposal';
    let copyFile = {
      title: copyTitle,
      parents: [{id: 'root'}]
    };

    copyFile = Drive.Files.copy(copyFile, templatePresentationId);
    const presentationCopyId = copyFile.id;


    /* load active slide deck */
      var deck = SlidesApp.openById(presentationCopyId);
      var slides = deck.getSlides();

    /* loop through slide deck slides and replace tokens with variable values */
      slides.forEach(function(slide){
        var shapes = (slide.getShapes());
        shapes.forEach(function(shape){
          shape.getText().replaceAllText('[(][(]customer-name[)][)]',customerName);
        }); 
      })
      deck.saveAndClose();

/*
    // Execute the replaces for this presentation.
    const result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);

    // Count the total number of replacements made.
    let numReplacements = 0;
    result.replies.forEach(function(reply) {
      numReplacements += reply.replaceAllText.occurrencesChanged;
    }); 

    console.log('Created presentation for %s with ID: %s', customerName, presentationCopyId);
    console.log('Replaced %s text instances', numReplacements);
*/
    return presentationCopyId;
  } catch (err) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', err.error);
  }  
}

//function that ceates a JSON object for the slide replace text
 function jsonForReplace(searchText,replaceText){
  return '{"replaceAllText":{"containsText":{"text":"' + searchText + '","matchCase":true},"replaceText":"' + replaceText + '"}}'
 }

//function that shows a URL in a dialog box
 function showAnchor(name,url,modalName) {
    var html = '<html><body style="font-family: Arial, Helvetica, sans-serif;"><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,modalName);
  }
