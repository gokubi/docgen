function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Doc Generation')
      .addItem('Generate Proposal', 'menuItem1')
      .addItem('Geneate SOW', 'menuItem2')
      .addToUi();
}

function menuItem1() {
  presentationid = mergeProposal('1HomquqMhAuR7eqe67ftyrMbufyUle5kjOG2l8mXQSXM');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  showAnchor('Open the Proposal Deck','https://docs.google.com/presentation/d/' + presentationid + '/edit#',"Proposal Generation");
}
     
function menuItem2() {
  docid = mergeMSA('1QPc6y2f9q6vVx5bMDXveVPYNmpuChFDqHBoJoubfmEE');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  showAnchor('Open the SOW Doc','https://docs.google.com/document/d/' + docid + '/edit#',"SOW Generation");
}

function mergeMSA(templateMSAId) {
  //if we are testing, supply the template doc id
  if(templateMSAId == null) {
    templateMSAId = '1QPc6y2f9q6vVx5bMDXveVPYNmpuChFDqHBoJoubfmEE';
  }

  const accountValuesRange = 'Customer!A2:P6'; //range for the account fields
  const msValuesRange = 'ProjectPricing!C20:C21'; //range for the ms quantity fields
  const msMilestonesRange = 'ProductsandQuantities!G3:G20'; //range for the ms milestones fields

  //id for the estimation sheet for the current project
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  dataSpreadsheetId = ss.getId();

  try {
      //get the account range in the estimate spreadsheet
      let accountValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(accountValuesRange).getValues();
      //create a new merged slide deck for the first line in the sheet
      
      const row = accountValues[0]; //get the row
      const companyName = row[0];
      const effectiveDate = row[9];
      const year = row[10];
      const address = row[13];
      const customerName = row[11];
      const customerTitle = row[12];
      const bwAccountManagerName = row[14];
      const bwAccountManagerEmail = row[15];
     
     //get the ms range in the estimate spreadsheet
      let msValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(msValuesRange).getValues();
      
      const firstMSRow = msValues[0]; //get the row
      const msMonths = firstMSRow[0];
      const secondMSRow = msValues[1]; //get the row
      const msFTELevel = secondMSRow[0];

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

      newMSABody.replaceText("[(][(]effective-date[)][)]",effectiveDate);
      newMSABody.replaceText("[(][(]year[)][)]",year);

      newMSABody.replaceText('[(][(]company-address[)][)]',address);
      newMSABody.replaceText('[(][(]company-name[)][)]',companyName);
      newMSABody.replaceText('[(][(]customer-name[)][)]',customerName);
      newMSABody.replaceText('[(][(]customer-title[)][)]',customerTitle);
      newMSABody.replaceText('[(][(]bw-am-name[)][)]',bwAccountManagerName);
      newMSABody.replaceText('[(][(]bw-am-email[)][)]',bwAccountManagerEmail);

      newMSABody.replaceText('[(][(]ms-months[)][)]',msMonths);
      newMSABody.replaceText('[(][(]ms-fte-level[)][)]',msFTELevel);
      newMSABody.replaceText('[(][(]ms-milestones[)][)]',milestones);

      newMSADoc.saveAndClose();

    return msaCopyId;

  } catch (err) {
      // TODO (Developer) - Handle exception
      console.log('Failed with error: %s', err.error);
    }
}

/**
 * Use the Sheets API to load data from a spreadsheet
 * @param {string} templatePresentationId
 * @param {string} dataSpreadsheetId
 * @returns {*[]}
 */
function mergeProposal(templatePresentationId) {

   if(templatePresentationId == null) {
    templatePresentationId = '1HomquqMhAuR7eqe67ftyrMbufyUle5kjOG2l8mXQSXM';
  }

  let dollarUSLocale = Intl.NumberFormat('en-US',{style: "currency",currency: "USD",minimumFractionDigits: 0,maximumFractionDigits: 0,});

  let responses = [];

  const accountValuesRange = 'Customer!A2:G6'; //range for the account and proposal fields
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
    let accountValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(accountValuesRange).getValues();

    //create a new merged slide deck for the first line in the sheet
    for (let i = 0; i < 1; ++i) {
      const row = accountValues[i]; //get the row
      const customerName = row[0]; // name in column 1
      const proposalName = row[1]; // proposal name in column 2
      const proposalDate = Utilities.formatDate(row[2], "GMT+1", "MM.dd.yyyy"); // date in column 3
      const sowprice =   dollarUSLocale.format(row[6]);

      //craft JSON objects for account level replacements
      requestJSON = '{"replaceAllText":{"containsText":{"text":"((customer-name))","matchCase":true},"replaceText":"' + customerName + '"}},{"replaceAllText":{"containsText":{"text":"((proposal-name))","matchCase":true},"replaceText":"' + proposalName + '"}},{"replaceAllText":{"containsText":{"text":"((proposal-date))","matchCase":true},"replaceText":"' + proposalDate + '"}},{"replaceAllText":{"containsText":{"text":"((sowprice))","matchCase":true},"replaceText":"' + sowprice + '"}}';


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
          requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-sowprice))',dollarUSLocale.format(featureSetRow[22]));
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

      console.log('requestJSON: ' + requestJSON);

      const requests = JSON.parse(requestJSON);
      
      // Duplicate the template presentation using the Drive API.
      const copyTitle = customerName + ' Salesforce Proposal';
      let copyFile = {
        title: copyTitle,
        parents: [{id: 'root'}]
      };

      copyFile = Drive.Files.copy(copyFile, templatePresentationId);
      const presentationCopyId = copyFile.id;

      // Execute the replaces for this presentation.
      const result = Slides.Presentations.batchUpdate({
        requests: requests
      }, presentationCopyId);

      // Count the total number of replacements made.
      let numReplacements = 0;
      result.replies.forEach(function(reply) {
        numReplacements += reply.replaceAllText.occurrencesChanged;
      }); 

/*
    // Fetch a list of all embedded charts in this
    // spreadsheet.
    var charts = [];
    var sheets = ss.getSheets();

    var position = {left: 40, top: 30};
    var size = {height: 510, width: 645};
    for (i = 0; i < sheets.length; i++) {
      charts = charts.concat(sheets[i].getCharts());
    }
    
    // If there aren't any charts, display a toast
    // message and return without doing anything
    // else.
  for (i = 0; i < charts.length; i++) {
      var slides = SlidesApp.openById(presentationCopyId);


     var newSlide = slides.appendSlide();
    newSlide.insertSheetsChart(
      charts[i],
      position.left,
      position.top,
      size.width,
      size.height); 
    }
*/

      console.log('Created presentation for %s with ID: %s', customerName, presentationCopyId);
      console.log('Replaced %s text instances', numReplacements);

      return presentationCopyId;
    }
  } catch (err) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', err.error);
  }

  
}

//function that ceates a JSON object for the slide replace text
 function jsonForReplace(searchText,replaceText){
  return '{"replaceAllText":{"containsText":{"text":"' + searchText + '","matchCase":true},"replaceText":"' + replaceText + '"}}'
 }

 function showAnchor(name,url,modalName) {
    var html = '<html><body style="font-family: Arial, Helvetica, sans-serif;"><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,modalName);
  }
