/*
This Apps Script is attached to an estimation spreadsheet template
and serves to make estimation and document generation easier and more consistent.
*/

//argument to work with Shared Drives in Google Workspace
const optionalArgs={supportsAllDrives: true};
//locale for formatting US dollars
const dollarUSLocale = Intl.NumberFormat('en-US',{style: "currency",currency: "USD",minimumFractionDigits: 0,maximumFractionDigits: 0,});

//load the UI for working with menu items and dialogs
var ui = SpreadsheetApp.getUi();

//Template Document Ids for generation of slides, spreadsheets, and documents
var standardProposalTemplateId = '';
var standardMSATemplateId = '';
var orgAssessmentSOWTemplateId ='';
var onDemandSupportSOWTemplateId ='';
var agileSOWTeamplateId ='';
var managedServicesSOWTemplateId ='';
var standardNDATemplateId = '';
var projectBATemplateId = '';

//spreasheet ranges in the estimation sheet
const msMilestonesRange = 'ManagedServicesMilestones!B2:B20'; //range for the ms milestones fields
const featureSetValuesRange = "FeatureSets!A2:W13"; //range for the feature set fields
const featureSetItemsValuesRange = "FeatureSetItems!A2:Y200"; //range for the feature set item fields
const projectBASOWRange = "ProjectPlan!B10:B13"; //range for the BA SOW details
const projectBARequirementsRange = "Requirements!F2:J200"; //range for the BA Requirements

//cell locations for pulling data
const oppIdCell = 'B1';
const companyNameCell = 'B6';
const customerNameCell = 'B7';
const customerTitleCell = 'B8';
const customerEmailCell = 'B9';
const addressCell = 'B10';
const companyJurisdictionCell = 'B11';
const companyEntityCell = 'B12';
const proposalNameCell = 'B16';
const proposalDateCell = 'B17';
const agileDurationCell = 'B18';  
const agileProjectCell = 'B19';
const projectTypeCell = 'B20';
const msaNeededCell = 'B22';
const sowNumberCell = 'B23';
const effectiveDateCell = 'B25';
const yearCell = 'B26';
const sowPriceCell = 'B27';
const sowHalfPriceCell = 'B28';
const bwAccountManagerNameCell = 'B35';
const bwAccountManagerEmailCell = 'B36';
const bwAccountManagerTitleCell = 'B37';
const msMonthsCell = 'B40';
const msLevelCell = 'B41';
const onDemandSupportSizeCell = 'B44';
const onDemandSupportPriceCell = 'B45';
const orgAssessmentPriceCell = 'B48';
const contractsFolderIdCell = 'B51'; 
const projectsFolderIdCell = 'B52'; 
const agileProjectSummaryCell = 'B53'; 

//id for the estimation sheet for the current project
var ss = SpreadsheetApp.getActiveSpreadsheet();
dataSpreadsheetId = ss.getId();

//key global variables
var oppId = '';
var contractsFolderId = '';
var projectsFolderId = '';

var ndaId = '';
var customerFolderExists = false;
var msaNeeded = false;

//function run on sheet opening loads menu items and global variables
function onOpen() {
  // create the menu items for the spreadsheet
  ui.createMenu('Solutions')
      
      .addSubMenu(ui.createMenu('Proposals')
        .addItem('Generate Proposal Slides', 'genProposal'))
      .addSubMenu(ui.createMenu('MSAs and SOWs')
        .addItem('Geneate Managed Services or Agile MSA/SOW', 'genMSAgileSOW')
        .addItem('Geneate On Demand Support MSA/SOW', 'genonDemandSupportSOW')
        .addItem('Geneate Org Assessment MSA/SOW', 'genOrgAssessmentSOW'))
      .addSubMenu(ui.createMenu('NDAs')
        .addItem('Generate NDA', 'genNDA'))
      .addSubMenu(ui.createMenu('Delivery')
        .addItem('Generate Project Requirements Sheet', 'genProjectBASheet'))
      .addItem('Clone Estimation Sheet for new Customer', 'genEstimationSheet')
      .addToUi();
  
  //get the ids for all templates from the sheet
  getTemplateIds();

  //get the customer tab range in the estimate spreadsheet
  var sheet = ss.getSheetByName('Customer');

  //if this isn't the default template file, but a copy of it, check for Opp Id and contracts folder id, and ask for them if they are missing
  if(!ss.getName().includes('Template')){
    //if opportunity id is blank, prompt for it
    if(sheet.getRange(oppIdCell).getValue() == ''){
      var response = ui.prompt('Enter the Id of the Opportunity you are estimating:');
      // Write Opp Id to field
      if (response.getSelectedButton() == ui.Button.OK) {
        sheet.getRange(oppIdCell).setValue(response.getResponseText());
      }
    }

    //if google folder id is blank, prompt for it
    if(sheet.getRange(contractsFolderIdCell).getValue() == ''){
      var response = ui.prompt('Enter the Folder Id of the customer contracts folder in Google Drive:');
      // Write folder Id to field
      if (response.getSelectedButton() == ui.Button.OK) {
        sheet.getRange(contractsFolderIdCell).setValue(response.getResponseText());
      }
    }
  }
  
  //set customerFolderExists if folder is there
  if(sheet.getRange(contractsFolderIdCell).getValue() != ''){
      customerFolderExists = true;
      contractsFolderId = sheet.getRange(contractsFolderIdCell).getValue();
  }

  //set MSA Needed if checked
  if(sheet.getRange(msaNeededCell).getValue() != ''){
      msaNeeded == true;
  }
}

//grab the template Ids from the constants tab
function getTemplateIds(){
  //get the constants tab to set template ids
  var constantsSheet = ss.getSheetByName('Constants');

  standardProposalTemplateId = constantsSheet.getRange('B10').getValue();
  standardMSATemplateId = constantsSheet.getRange('B11').getValue();
  orgAssessmentSOWTemplateId = constantsSheet.getRange('B12').getValue();
  onDemandSupportSOWTemplateId = constantsSheet.getRange('B13').getValue();
  agileSOWTeamplateId = constantsSheet.getRange('B14').getValue();
  managedServicesSOWTemplateId = constantsSheet.getRange('B15').getValue();
  standardNDATemplateId = constantsSheet.getRange('B16').getValue();
  projectBATemplateId = constantsSheet.getRange('B17').getValue();

}

//generate the project ba requirements sheet that will be used by delivery
function genProjectBASheet(){
  getTemplateIds();

  newProjectBASheetId = '';
  var customerSheet = ss.getSheetByName('Customer');
  const companyName = customerSheet.getRange(companyNameCell).getValue();
  projectsFolderId = customerSheet.getRange(projectsFolderIdCell).getValue();
  const sowPrice = dollarUSLocale.format(customerSheet.getRange(sowPriceCell).getValue());
  const projectType = customerSheet.getRange(projectTypeCell).getValue();
  const agileDuration = customerSheet.getRange(agileDurationCell).getValue();
  const msMonths = customerSheet.getRange(msMonthsCell).getValue();
  const oppId = customerSheet.getRange(oppIdCell).getValue();

  var durationText = '';
if(projectType == 'Agile Project'){
  durationText = agileDuration + ' weeks';
} else {
  durationText = msMonths + ' months';
}

  //if google folder id is blank, prompt for it
  if(projectsFolderId == ''){
    var response = ui.prompt('Enter the Folder Id of the customer projects folder in Google Drive:');
    // Write folder Id to field
    if (response.getSelectedButton() == ui.Button.OK) {
      projectsFolderId = response.getResponseText();
      customerSheet.getRange(projectsFolderIdCell).setValue(projectsFolderId);
    }
  }
  
  //make a copy of the projectBA template
  // Duplicate the template sheet using the Drive API.
  const copyTitle = companyName + ' Business Analyst Sheet';
  let copyFile = {
    title: copyTitle,
    parents: [{id: 'root'}]
  };

  //copy the standard MSA to a new file
  copyFile = Drive.Files.copy(copyFile,projectBATemplateId,optionalArgs);
  var newProjectBASheetId = copyFile.id;

  //move document to customer folder
  if(projectsFolderId!=''){
    var fileToMove = DriveApp.getFileById(newProjectBASheetId);
    var folder = DriveApp.getFolderById(projectsFolderId);
    fileToMove.moveTo(folder);
  }
  //get feature set items
  var data =[];
  let featureSetItemsValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(featureSetItemsValuesRange).getValues();

  //loop through all feature site items to get them into an array for saving to the ba sheet
  for(let i=0;i<featureSetItemsValues.length;i++){
    const featureSetItemsRow = featureSetItemsValues[i]; //get the row
    //only process for active rows
    if(featureSetItemsRow[23] != "" && featureSetItemsRow[4]) {
      //write them to an array
      data.push([featureSetItemsRow[0],featureSetItemsRow[2],featureSetItemsRow[1],'','Proposal/SOW','Proposed','','Sales Estimate: ' + featureSetItemsRow[3] + ' hours']);
    }
  }

  //get the location on the ba 
  let baRequirementsSheet = SpreadsheetApp.openById(newProjectBASheetId).getSheetByName('Requirements');
  var lastRow = baRequirementsSheet.getLastRow();

  //write array to ba sheet
  baRequirementsSheet.getRange(lastRow+1,6,data.length,8).setValues(data);

  //get SOW range
  let projectPlanSheet = SpreadsheetApp.openById(newProjectBASheetId).getSheetByName('ProjectPlan');
    projectPlanSheet.getRange('B10').setValue(sowPrice);
    projectPlanSheet.getRange('B11').setValue(projectType);
    projectPlanSheet.getRange('B12').setValue(durationText);
    projectPlanSheet.getRange('B13').setValue('https://bitwiseindustries.lightning.force.com/lightning/r/Opportunity/' + oppId + '/view');
    
  //share link to new sheet
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  if(newProjectBASheetId != null){
    showAnchor('Open the Project BA Sheet','https://docs.google.com/spreadsheets/d/' + newProjectBASheetId + '/edit#','The Project BA Sheet was saved to the customer contracts folder or your My Drive',"Project BA Sheet Generation");
  }
}

//make a copy of the estimation sheet template for a new prospect
function genEstimationSheet(){

  //get the new opp id
  var response = ui.prompt('Enter the Id of the Opportunity you are estimating:');
      // Write Opp Id to field
      if (response.getSelectedButton() == ui.Button.OK) {
        oppId = response.getResponseText();
      }

  //get the new google folder id
  var response = ui.prompt('Enter the Folder Id of the customer contracts folder in Google Drive:');
      // Write folder Id to field
      if (response.getSelectedButton() == ui.Button.OK) {
        contractsFolderId = response.getResponseText();
      }

  var companyName ='';
  //get the company name from the folder name
  if(contractsFolderId!=''){
    var folder = DriveApp.getFolderById(contractsFolderId);
    companyName = folder.getName();
  }

  //make a copy
  // Duplicate the template MSA using the Drive API.
  const copyTitle = companyName + ' Estimation Sheet';
  let copyFile = {
    title: copyTitle,
    parents: [{id: 'root'}]
  };

  //copy the standard MSA to a new file
  copyFile = Drive.Files.copy(copyFile,dataSpreadsheetId,optionalArgs);
  var newEstimationSheetId = copyFile.id;

  //move document to customer folder
  if(contractsFolderId!=''){
    customerFolderExists = true;
    var fileToMove = DriveApp.getFileById(newEstimationSheetId);
    var folder = DriveApp.getFolderById(contractsFolderId);
    fileToMove.moveTo(folder);
  }
  //update Opp id and google folder id
  //id for the estimation sheet for the current project
  var ss = SpreadsheetApp.openById(newEstimationSheetId);
  var customerSheet = ss.getSheetByName('Customer');

  customerSheet.getRange(oppIdCell).setValue(oppId);
  customerSheet.getRange(contractsFolderIdCell).setValue(contractsFolderId);
  
  //prompt with a link
  if(newEstimationSheetId != null){
    showAnchor('Open the Estimation Sheet','https://docs.google.com/spreadsheets/d/' + newEstimationSheetId + '/edit#','The cloned Estimation Sheet was created and saved!',"Estimation Sheet");
  }
}

//function for menu item that runs the code to generate a proposal
function genProposal() {
  getTemplateIds();
  presentationid = createProposal(standardProposalTemplateId);
  if(presentationid != null){
    showAnchor('Open the Proposal Deck','https://docs.google.com/presentation/d/' + presentationid + '/edit#','The proposal deck was saved to the customer folder.',"Proposal Generation");
  }
}

//function for menu item that runs the code to generate an Agile MSA/SOW     
function genMSAgileSOW() {
  getTemplateIds();
  docid = createMSASOW('MSAgile');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  let message = 'The ';
  if(msaNeeded){
      message += 'MSA/';
  }
  message += 'SOW was saved to ';

  if(customerFolderExists){
    message += 'the customer contracts folder.';
    
  } else {
    message += 'your My Drive.';
  }
    
  if(docid != null){
    
    showAnchor('Open the MSA/SOW Doc','https://docs.google.com/document/d/' + docid + '/edit#',message,"SOW Generation");
  }
}

//function for menu item that runs the code to generate a Managed Services MSA/SOW     
function genOrgAssessmentSOW() {
  getTemplateIds();
  docid = createMSASOW('OrgAssessment');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  if(docid != null){
    showAnchor('Open the MSA/SOW Doc','https://docs.google.com/document/d/' + docid + '/edit#','The MSA/SOW was saved to the customer contracts folder or your My Drive',"SOW Generation");
  }
}

//function for menu item that runs the code to generate an On Demand Support MSA/SOW     
function genonDemandSupportSOW() {
  getTemplateIds();
  docid = createMSASOW('onDemand');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  if(docid != null){
    showAnchor('Open the MSA/SOW Doc','https://docs.google.com/document/d/' + docid + '/edit#','The MSA/SOW was saved to the customer contracts folder or your My Drive',"SOW Generation");
  }
}

//function for menu item that runs the code to generate an NDA
function genNDA() {
  getTemplateIds();
  ndaId = createNDA(standardNDATemplateId);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  if(ndaId != null){
    showAnchor('Open the NDA','https://docs.google.com/document/d/' + ndaId + '/edit#','The NDA was saved to the customer contracts folder or your My Drive',"NDA Generation");
  }
}

//function to create an NDA
function createNDA(ndaTemplateId) {
  //get the customer tab range in the estimate spreadsheet
  var customerSheet = ss.getSheetByName('Customer');

  //get cells from the range
  const companyName = customerSheet.getRange(companyNameCell).getValue();
  const companyEntity = customerSheet.getRange(companyEntityCell).getValue();
  const companyJurisdiction = customerSheet.getRange(companyJurisdictionCell).getValue();
  const effectiveDate = customerSheet.getRange(effectiveDateCell).getValue();
  const year = customerSheet.getRange(yearCell).getValue();
  const address = customerSheet.getRange(addressCell).getValue();
  const customerName = customerSheet.getRange(customerNameCell).getValue();
  const customerTitle = customerSheet.getRange(customerTitleCell).getValue();
  const bwAccountManagerName = customerSheet.getRange(bwAccountManagerNameCell).getValue();
  const bwAccountManagerTitle = customerSheet.getRange(bwAccountManagerTitleCell).getValue();
  const contractsFolderId = customerSheet.getRange(contractsFolderIdCell).getValue();

  var promptText = 'Do you want to proceed creating an NDA for ' + companyName + '?';
  var response = ui.alert(promptText, ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (response == ui.Button.YES) {
    // Duplicate the template MSA using the Drive API.
    const copyTitle = companyName + ' NDA';
    let copyFile = {
      title: copyTitle,
      parents: [{id: 'root'}]
    };

    //copy the standard MSA to a new file
    copyFile = Drive.Files.copy(copyFile,ndaTemplateId,optionalArgs);
    ndaId = copyFile.id;

    //move document to customer folder
    if(contractsFolderId!=''){
      customerFolderExists = true;
      var fileToMove = DriveApp.getFileById(ndaId);
      var folder = DriveApp.getFolderById(contractsFolderId);
      fileToMove.moveTo(folder);
    }

    //get into the body and header of the doc
    const newNDADoc = DocumentApp.openById(ndaId);
    const newNDABody = newNDADoc.getBody();

    //replace merge fields in body
    newNDABody.replaceText("[(][(]effective-date[)][)]",effectiveDate);
    newNDABody.replaceText("[(][(]year[)][)]",year);
    newNDABody.replaceText('[(][(]company-address[)][)]',address);
    newNDABody.replaceText('[(][(]company-name[)][)]',companyName);
    newNDABody.replaceText('[(][(]company-entity[)][)]',companyEntity);
    newNDABody.replaceText('[(][(]company-jurisdiction[)][)]',companyJurisdiction);
    newNDABody.replaceText('[(][(]customer-name[)][)]',customerName);
    newNDABody.replaceText('[(][(]customer-title[)][)]',customerTitle);
    newNDABody.replaceText('[(][(]bw-am-name[)][)]',bwAccountManagerName);
    newNDABody.replaceText('[(][(]bw-am-title[)][)]',bwAccountManagerTitle);

    newNDADoc.saveAndClose();
  }
  return ndaId;
}

//function to create an MSA/SOW, accepts the type of SOW needed
function createMSASOW(sowType) {
  
  //get the customer tab range in the estimate spreadsheet
  var customerSheet = ss.getSheetByName('Customer');

  //get cells from the range
  const companyName = customerSheet.getRange(companyNameCell).getValue();
  const companyEntity = customerSheet.getRange(companyEntityCell).getValue();
  const companyJurisdiction = customerSheet.getRange(companyJurisdictionCell).getValue();
  const effectiveDate = customerSheet.getRange(effectiveDateCell).getValue();
  const year = customerSheet.getRange(yearCell).getValue();
  const address = customerSheet.getRange(addressCell).getValue();
  msaNeeded = customerSheet.getRange(msaNeededCell).getValue();
  const customerName = customerSheet.getRange(customerNameCell).getValue();
  const customerTitle = customerSheet.getRange(customerTitleCell).getValue();
  const customerEmail = customerSheet.getRange(customerEmailCell).getValue();
  const sowPrice = dollarUSLocale.format(customerSheet.getRange(sowPriceCell).getValue());
  const agileProject = customerSheet.getRange(agileProjectCell).getValue();
  const onDemandSupportSize = customerSheet.getRange(onDemandSupportSizeCell).getValue();
  const onDemandSupportPrice = dollarUSLocale.format(customerSheet.getRange(onDemandSupportPriceCell).getValue());
  const orgAssessmentPrice = dollarUSLocale.format(customerSheet.getRange(orgAssessmentPriceCell).getValue());
  const contractsFolderId = customerSheet.getRange(contractsFolderIdCell).getValue();

  var msaId = '';
  var sowId = '';
  var completeFileId = '';

  var promptPrice = 0;
  var promptText = 'Want to proceed with creating ';
  if (sowType == 'onDemand') {
    promptText += 'a ' + onDemandSupportSize + ' On Demand Support '
    promptPrice = onDemandSupportPrice;
  } else if (sowType == 'OrgAssessment') {
    promptText += 'an Org Assessment ';
    promptPrice = orgAssessmentPrice;
  } else {
    promptPrice = sowPrice;
    if (agileProject) {
      promptText += 'an Agile Project ';
    } else {
      promptText += 'a Managed Services Project ';
    }
  }
  
  if (msaNeeded){
    promptText += 'MSA and ';
  }
  promptText += 'SOW for ' + companyName + " for " + promptPrice + '?';

  var response = ui.alert(promptText, ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (response == ui.Button.YES) {

    //if the MSA Needed checkbox is set, create one
    if (msaNeeded){

      // Duplicate the template MSA using the Drive API.
        const copyTitle = companyName + ' MSA';
        let copyFile = {
          title: copyTitle,
          parents: [{id: 'root'}]
        };

        //copy the standard MSA to a new file
        copyFile = Drive.Files.copy(copyFile, standardMSATemplateId,optionalArgs);
        msaId = copyFile.id;

        //get into the body and header of the doc
        const newMSADoc = DocumentApp.openById(msaId);
        const newMSABody = newMSADoc.getBody();
        const newMSAHeader = newMSADoc.getHeader();

        //replace merge fields in body
        newMSABody.replaceText("[(][(]effective-date[)][)]",effectiveDate);
        newMSABody.replaceText("[(][(]year[)][)]",year);
        newMSABody.replaceText('[(][(]company-address[)][)]',address);
        newMSABody.replaceText('[(][(]company-name[)][)]',companyName);
        newMSABody.replaceText('[(][(]company-entity[)][)]',companyEntity);
        newMSABody.replaceText('[(][(]company-jurisdiction[)][)]',companyJurisdiction);
        newMSABody.replaceText('[(][(]customer-name[)][)]',customerName);
        newMSABody.replaceText('[(][(]customer-title[)][)]',customerTitle);
        newMSABody.replaceText('[(][(]customer-email[)][)]',customerEmail);

        //replace merge fields in header
        newMSAHeader.replaceText('[(][(]effective-date[)][)]',effectiveDate);

        newMSADoc.saveAndClose();
    } 
      
    //generate the SOW of the correct type
    if(sowType == 'MSAgile') {
      //check the spreadsheet for if Agile or MS
      if(agileProject){
        sowId = mergeSOW('agile');
      } else {
        sowId = mergeSOW('ms');
      }
    } else {
      sowId = mergeSOW(sowType);
    }

    //if we're creating an MSA, merge the new SOW into it, otherwise just use the SOW
    if(msaNeeded){
      //combine the files
      //var combinedTitle = "Combined Document Example";
      var combo = DocumentApp.openById(msaId);
      combo.setName(companyName + ' MSA and SOW');
      var comboBody = combo.getBody();
      
      //console.log('msaId & sowId',msaId + ' : ' + sowId);
      var docIDs = [sowId];
      for (var i = 0; i < docIDs.length; ++i ) {
        var otherBody = DocumentApp.openById(docIDs[i]).getActiveSection();    
        var totalElements = otherBody.getNumChildren();
        for( var j = 0; j < totalElements; ++j ) {
          var element = otherBody.getChild(j).copy();
          var type = element.getType();
          if( type == DocumentApp.ElementType.PARAGRAPH )
            comboBody.appendParagraph(element);
          else if( type == DocumentApp.ElementType.TABLE )
            comboBody.appendTable(element);
          else if( type == DocumentApp.ElementType.LIST_ITEM )
            comboBody.appendListItem(element);
          //else
          // throw new Error("Unknown element type: "+type);
        }
        DriveApp.getFileById(docIDs[i]).setTrashed(true);
      }
      combo.saveAndClose();
      completeFileId = msaId;

    } else {
      completeFileId = sowId;
    }

    //move document to customer folder
      if(contractsFolderId!=''){
        customerFolderExists = true;
        var fileToMove = DriveApp.getFileById(completeFileId);
        var folder = DriveApp.getFolderById(contractsFolderId);
        fileToMove.moveTo(folder);
      }

    return completeFileId;
  }
}

function mergeSOW(sowType) {

  //get the right SOW template id
  var sowTemplateId = '';
  if (sowType == 'onDemand') {
    sowTemplateId = onDemandSupportSOWTemplateId;
  } else if (sowType == 'OrgAssessment') {
    sowTemplateId = orgAssessmentSOWTemplateId;
  } else if (sowType == 'ms') {
    sowTemplateId = managedServicesSOWTemplateId;
  } else if (sowType == 'agile') {
    sowTemplateId = agileSOWTeamplateId;
  } 

  try {
      var customerSheet = ss.getSheetByName('Customer'); 

    //get cells from the range
      const companyName = customerSheet.getRange(companyNameCell).getValue();
      const companyEntity = customerSheet.getRange(companyEntityCell).getValue();
      const companyJurisdiction = customerSheet.getRange(companyJurisdictionCell).getValue();
      const effectiveDate = customerSheet.getRange(effectiveDateCell).getValue();
      const year = customerSheet.getRange(yearCell).getValue();
      const address = customerSheet.getRange(addressCell).getValue();
      const customerName = customerSheet.getRange(customerNameCell).getValue();
      const customerTitle = customerSheet.getRange(customerTitleCell).getValue();
      const customerEmail = customerSheet.getRange(customerEmailCell).getValue();
      const sowPrice = dollarUSLocale.format(customerSheet.getRange(sowPriceCell).getValue());
      const onDemandSupportPrice = dollarUSLocale.format(customerSheet.getRange(onDemandSupportPriceCell).getValue());
      const orgAssessmentPrice = dollarUSLocale.format(customerSheet.getRange(orgAssessmentPriceCell).getValue());
      const msMonths = customerSheet.getRange(msMonthsCell).getValue();
      const msFTELevel = customerSheet.getRange(msLevelCell).getValue();
      const bwAccountManagerName = customerSheet.getRange(bwAccountManagerNameCell).getValue();
      const bwAccountManagerEmail = customerSheet.getRange(bwAccountManagerEmailCell).getValue();
      const bwAccountManagerTitle = customerSheet.getRange(bwAccountManagerTitleCell).getValue();
      const sowNumber = customerSheet.getRange(sowNumberCell).getValue();
      const sowHalfPrice = dollarUSLocale.format(customerSheet.getRange(sowHalfPriceCell).getValue());
      const agileDuration = customerSheet.getRange(agileDurationCell).getValue(); 
      const agileProjectSummary = customerSheet.getRange(agileProjectSummaryCell).getValue(); 

      //get the account range in the estimate spreadsheet
      let msMilestones = SpreadsheetApp.openById(dataSpreadsheetId).getRange(msMilestonesRange).getValues();
      milestones = '';
      for(let i=0;i<msMilestones.length;i++){
          const msMilestonesRow = msMilestones[i]; //get the row
          if(msMilestonesRow[0]!=''){
            milestones += msMilestonesRow[0] + '\n';
          }
      }
      var msFeatureSetsLanguage = '';
      let msFeatureSets = SpreadsheetApp.openById(dataSpreadsheetId).getRange(featureSetValuesRange).getValues();
      for(let i=0;i<msFeatureSets.length;i++){
        const msFeatureSetRow = msFeatureSets[i]; //get the row
        if(msFeatureSetRow[5]){
            msFeatureSetsLanguage += msFeatureSetRow[2] + '\n';
          }
      }


      // Duplicate the template SOW using the Drive API.
      const copyTitle = companyName + ' SOW';
      let copyFile = {
        title: copyTitle,
        parents: [{id: 'root'}]
      };

      copyFile = Drive.Files.copy(copyFile, sowTemplateId,optionalArgs); //make a copy of the template
      const sowId = copyFile.id;

      const sowDoc = DocumentApp.openById(sowId);
      const sowBody = sowDoc.getBody();
      //const sowHeader = sowDoc.getHeader();

      sowBody.replaceText("[(][(]effective-date[)][)]",effectiveDate);
      sowBody.replaceText("[(][(]year[)][)]",year);
      sowBody.replaceText('[(][(]company-address[)][)]',address);
      sowBody.replaceText('[(][(]company-name[)][)]',companyName);
      sowBody.replaceText('[(][(]company-entity[)][)]',companyEntity);
      sowBody.replaceText('[(][(]company-jurisdiction[)][)]',companyJurisdiction);
      sowBody.replaceText('[(][(]customer-name[)][)]',customerName);
      sowBody.replaceText('[(][(]customer-title[)][)]',customerTitle);
      sowBody.replaceText('[(][(]customer-email[)][)]',customerEmail);
      sowBody.replaceText('[(][(]sow-number[)][)]',sowNumber);
      sowBody.replaceText('[(][(]bw-am-name[)][)]',bwAccountManagerName);
      sowBody.replaceText('[(][(]bw-am-email[)][)]',bwAccountManagerEmail);
      sowBody.replaceText('[(][(]bw-am-title[)][)]',bwAccountManagerTitle);
      sowBody.replaceText('[(][(]ms-months[)][)]',msMonths);
      sowBody.replaceText('[(][(]ms-fte-level[)][)]',msFTELevel);
      sowBody.replaceText('[(][(]ms-milestones[)][)]',milestones);
      sowBody.replaceText('[(][(]ms-sowprice[)][)]',sowPrice);
      sowBody.replaceText('[(][(]ms-featuresets[)][)]',msFeatureSetsLanguage);
      
      sowBody.replaceText('[(][(]agile-sowprice[)][)]',sowPrice);
      sowBody.replaceText('[(][(]agile-half-sowprice[)][)]',sowHalfPrice);
      sowBody.replaceText('[(][(]agile-duration[)][)]',agileDuration);
      sowBody.replaceText('[(][(]orgassessment-price[)][)]',orgAssessmentPrice);
      sowBody.replaceText('[(][(]ondemand-price[)][)]',onDemandSupportPrice);
      sowBody.replaceText('[(][(]agile-project-summary[)][)]',agileProjectSummary);

      sowDoc.saveAndClose();

      return sowId;
  } catch (err) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', err.error);
  }
}

//function that generates a Proposal
function createProposal(templatePresentationId) {

    let responses = [];

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
    featureSetTags2.push("((FS" + fsNumber + "-assumptions))");
    featureSetTags2.push("((FS" + fsNumber + "-percentage))");
    featureSetTags2.push("((FS" + fsNumber + "-sowprice))");
    
    for (let i = 0; i < 12; ++i) {
      //create JSON objects for 12 items for each of the feature set level tags
      featureSetTags2.push("((FS" + fsNumber + "-Item-" + (i+1).toString() + "-name))");
    }
  }

  try {
    //get the account range in the estimate spreadsheet
   // let customerValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(customerValuesRange).getValues();
    var customerSheet = ss.getSheetByName('Customer');

    //create a new merged slide deck for the first line in the sheet
    const companyName = customerSheet.getRange(companyNameCell).getValue(); // name
    const proposalName = customerSheet.getRange(proposalNameCell).getValue(); // proposal name
    const proposalDate = Utilities.formatDate(customerSheet.getRange(proposalDateCell).getValue(), "GMT+1", "MM.dd.yyyy"); // date
    const sowprice =   dollarUSLocale.format(customerSheet.getRange(sowPriceCell).getValue());
    const msLevel = customerSheet.getRange(msLevelCell).getValue(); // proposal name
    const msMonths = customerSheet.getRange(msMonthsCell).getValue(); // proposal name


    var response = ui.alert('Want to proceed with creating a proposal for ' + companyName, ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (response == ui.Button.YES) {

      //JSON for replacement of key fields
      requestJSON = jsonForReplace('((company-name))',companyName);
      requestJSON += ',' + jsonForReplace('((proposal-name))',proposalName);
      requestJSON += ',' + jsonForReplace('((proposal-date))',proposalDate);
      requestJSON += ',' + jsonForReplace('((sowprice))',sowprice);
      requestJSON += ',' + jsonForReplace('((ms-fte-level))',msLevel);
      requestJSON += ',' + jsonForReplace('((ms-months))',msMonths);


      //get the feature set range in the estimate spreadsheet
      let featureSetValues = SpreadsheetApp.openById(dataSpreadsheetId).getRange(featureSetValuesRange).getValues();

      //loop through all the feature set rows
      for(let i=0;i<featureSetValues.length;i++){
        const featureSetRow = featureSetValues[i]; //get the row
        //only process for active rows
        if(featureSetRow[0] != "") {

          var assumptionsText = '';
          if(featureSetRow[4]!=''){
            //construct the assumptions string, removing new lines
            assumptionsText = 'Assumptions: ' + featureSetRow[4].replace(/(\r\n|\n|\r)/gm, "");
          }

          //create JSON replace objects and append to the full JSON
          requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-number))',featureSetRow[0]);
          requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-name))',featureSetRow[2]);
          requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-description))',featureSetRow[3]);
          requestJSON += ',' + jsonForReplace('((' + featureSetRow[0] + '-assumptions))',assumptionsText);
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

      //parse the JSON for use in replace
      const requests = JSON.parse(requestJSON);

      // Duplicate the template presentation using the Drive API.
      const copyTitle = companyName + ' Salesforce Proposal';
      let copyFile = {
        title: copyTitle,
        parents: [{id: 'root'}]
      };

      copyFile = Drive.Files.copy(copyFile, templatePresentationId,optionalArgs);
      const presentationCopyId = copyFile.id;
      var sheet = ss.getSheetByName('Customer');

      //set customerFolderExists if folder is there
      if(sheet.getRange(contractsFolderIdCell).getValue() != ''){
          contractsFolderId = sheet.getRange(contractsFolderIdCell).getValue();
      }
      //move document to customer folder
      if(contractsFolderId!=''){
        customerFolderExists = true;
        var fileToMove = DriveApp.getFileById(presentationCopyId);
        var folder = DriveApp.getFolderById(contractsFolderId);
        fileToMove.moveTo(folder);
      } 

      // Execute the replaces for this presentation.
      const result = Slides.Presentations.batchUpdate({
        requests: requests
      }, presentationCopyId);
  
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

//function that shows a URL in a dialog box
 function showAnchor(name,url,text,modalName) {
    var html = '<html><body style="font-family: Arial, Helvetica, sans-serif;"><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a><p style="color:gray;">' + text + '</p></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,modalName);
  }
