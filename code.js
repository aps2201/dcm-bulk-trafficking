/***********************************************************************
Copyright 2018 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Note that these code samples being shared are not official Google
products and are not formally supported.
***********************************************************************/

/**
 * Setup custom menu for the sheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('DCM Functions')
      .addItem('Setup Sheets', 'setupTabs')
      .addSeparator()
      .addItem('List Sites', 'listSites')
      .addSeparator()
      .addItem('Bulk Create Campaigns', 'createCampaigns')
      .addItem('Bulk Create Placements', 'createPlacements')
      .addItem('Bulk Create Ads', 'createAds')
      .addItem('Bulk Create Tracking Creatives', 'createTrackingCreatives')
      .addItem('Bulk Create Landing Pages', 'createLandingPages')
      .addSeparator()
      .addItem('Upload Creatives', 'uploadCreatives')
      .addSeparator()
      .addItem('Campaign <-> Creative', 'createCampCreativeAss')
      .addItem('Ads <-> Creative', 'createAdsCreativeAss')
      .addToUi();
}

/**
 * Using DCM API list all the sites this profile has added
 * and print them out on the sheet.
 */
function listSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SITES_SHEET);
  var profileID = _fetchProfileId();
  initializeSheet_(SITES_SHEET, true);

  // setup header row
  sheet.getRange('A1')
      .setValue('Site Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('B1')
      .setValue('Directory Site ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');

  var sites = CampaignManager.Sites.list(profileID).sites;
  for (var i = 0; i < sites.length; i++) {
    var currentObject = sites[i];
    var rowNum = i+2;
    sheet.getRange('A' + rowNum)
        .setValue(currentObject.name)
        .setBackground('lightgray');
    sheet.getRange('B' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.directorySiteId)
        .setBackground('lightgray');
  }
}

/**
 * Read campaign information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createCampaigns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGNS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,6);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
      var newCampaign = _createOneCampaign(ss, values[i]);
      sheet.getRange('F' + rowNum)
          .setValue(newCampaign.id)
          .setBackground('lightgray');
    }
  }
  SpreadsheetApp.getUi().alert('Finished creating campaigns!');
}


/**
 * Read placement information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createPlacements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PLACEMENTS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // skip header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,10);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
      var newPlacement = _createOnePlacement(ss, values[i]);
      sheet.getRange('J' + rowNum)
          .setValue(newPlacement.id)
          .setBackground('lightgray');
    }
  }
  SpreadsheetApp.getUi().alert('Finished creating the placements!');
}


/**
 * Read campaign ads from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createAds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ADS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,9);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
    var newAd = _createOneAd(ss, values[i]);
    sheet.getRange('I' + rowNum)
        .setValue(newAd.id)
        .setBackground('lightgray');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the ads!');
}

/**
 * Read creatives information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */ 
function createCreatives() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,9);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
    var newCreative = _createOneCreative(ss, values[i]);
    sheet.getRange('I' + rowNum)
        .setValue(newCreative.id)
        .setBackground('lightgray');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the creatives!');
}

/**
 * Read landing pages information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createLandingPages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LANDING_PAGES_SHEET);
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,4);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
    var newLandingPage = _createOneLandingPage(ss, values[i]);
    var rowNum = i+1;
    sheet.getRange('D' + rowNum)
        .setValue(newLandingPage.id)
        .setBackground('lightgray');
    }
  }
  
  SpreadsheetApp.getUi().alert('Finished creating the landing pages!');
}

/**
 * A helper function which creates one campaign via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleCampaignArray An array containing campaign information
 * @return {object} Campaign object
 */
function _createOneCampaign(ss, singleCampaignArray){
  var profileID = _fetchProfileId();

  var advertiserId = singleCampaignArray[0];
  var name = singleCampaignArray[1];
  var defaultLandingPageId = singleCampaignArray[2];
  var startDate = Utilities.formatDate(
      singleCampaignArray[3], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var endDate = Utilities.formatDate(
      singleCampaignArray[4], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

  var campaignResource = {
    "kind": "dfareporting#campaign",
    "advertiserId": advertiserId,
    "name": name,
    "startDate": startDate,
    "endDate": endDate,
    "defaultLandingPageId":defaultLandingPageId
  };
  var newCampaign = CampaignManager.Campaigns
      .insert(campaignResource, profileID);
  return newCampaign;
}

/**
 * A helper function which creates one creative via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleCreativeArray An array containing creative information
 * @return {object} Creative object
 */
function _createOneCreative(ss, singleCreativeArray){
  var profileID = _fetchProfileId();

  var advertiserId = singleCreativeArray[0];
  var name = singleCreativeArray[1];
  var width = singleCreativeArray[2];
  var height = singleCreativeArray[3];
  var creativeType = singleCreativeArray[4];
  var assetType = singleCreativeArray[5];
  var assetName = singleCreativeArray[6];

  var creativeResource =  {
    "name": name,
    "advertiserId": advertiserId,
    "size": {
      "width": width,
      "height": height
    },
    "active": true,
    "type": creativeType,
    "creativeAssets": [
      {
        "assetIdentifier": {
          "type": assetType,
          "name": assetName
        }
      }
    ]
  };

  var newCreative = CampaignManager.Creatives
      .insert(creativeResource, profileID);
  return newCreative;

}

/**
 * A helper function which creates one ad via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleAdArray An array containing ad information
 * @return {object} Ad object
 */
function _createOneAd(ss, singleAdArray){
  var profileID = _fetchProfileId();

  var campaignId = singleAdArray[0];
  var name = singleAdArray[1];

  var startTime = Utilities.formatDate(
      singleAdArray[2], ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');

  var endTime = Utilities.formatDate(
      singleAdArray[3], ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');

  var impressionRatio = singleAdArray[4];
  var priority = singleAdArray[5];
  var type = singleAdArray[6];
  var placementId = singleAdArray[7];

  //https://developers.google.com/doubleclick-advertisers/v3.1/ads
  //priority requires double digit format even for values lower than 10
  //e.g. AD_PRIORITY_03
  if(priority<10){
    priority = "0"+priority;
  }

  var adResource = {
      "kind": "dfareporting#ad",
      "campaignId":campaignId,
      "name": name,
      "startTime": startTime ,
      "endTime": endTime,
      "deliverySchedule":{
        "impressionRatio":impressionRatio,
        "priority":"AD_PRIORITY_"+priority
      },
      "type":type
    };

  adResource.placementAssignments = [{"placementId":placementId}];

  var newAd = CampaignManager.Ads.insert(
    adResource, profileID);
  return newAd;
}

/**
 * A helper function which creates one placement via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing current active
 * spreadsheet
 * @param {Array} singlePlacementArray An array containing
 * placement information
 * @return {object} Placement object
 */
function _createOnePlacement(ss, singlePlacementArray) {
  var profileID = _fetchProfileId();

  var campaignID = singlePlacementArray[0];
  var name = singlePlacementArray[1];
  var siteId = singlePlacementArray[2];
  var paymentSource = 'PLACEMENT_AGENCY_PAID';
  var compatibility = (singlePlacementArray[3]).trim().toUpperCase();
  var size = singlePlacementArray[4];
  var sizeSplitted = size.split('x');

  var pricingScheduleStartDate = Utilities.formatDate(
      singlePlacementArray[5], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var pricingScheduleEndDate = Utilities.formatDate(
      singlePlacementArray[6], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var pricingSchedulePricingType = singlePlacementArray[7];
  var tagFormats = (singlePlacementArray[8]).split(',');
  for (var i = 0; i < tagFormats.length; i++) {
    tagFormats[i] = (tagFormats[i].trim()).replace(/\r?\n|\r/g, ', ');
  }

  var placementResource = {
    "kind": "dfareporting#placement",
    "campaignId": campaignID,
    "name": name,
    "directorySiteId": siteId,
    "paymentSource": paymentSource,
    "compatibility": compatibility,
    "size": {
      "width": sizeSplitted[0].trim(),
      "height": sizeSplitted[1].trim()
    },
    "pricingSchedule": {
      "startDate": pricingScheduleStartDate,
      "endDate": pricingScheduleEndDate,
      "pricingType": pricingSchedulePricingType
    },
    "tagFormats": tagFormats
  };

  var newPlacement = CampaignManager.Placements
      .insert(placementResource, profileID);
  return newPlacement;
}

/**
 * A helper function which creates one landing page via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing current active
 * spreadsheet
 * @param {Array} singleLandingPageArray An array containing
 * landing page information
 * @return {object} Landing Page object
 */
function _createOneLandingPage(ss, singleLandingPageArray) {
  var profileID = _fetchProfileId();
  
  var advertiserId = singleLandingPageArray[0];
  var name = singleLandingPageArray[1];
  var url = singleLandingPageArray[2];
  
  var landingPageResource = {
    "advertiserId": advertiserId,
    "kind": "dfareporting#landingPage",
    "name": name,
    "url": url
  }
  
  var newLandingPage = CampaignManager.AdvertiserLandingPages
      .insert(landingPageResource, profileID);
  return newLandingPage;
}

/** aps: I mess around with things down here */ 

/** 
 * aps: This snippet is taken from https://github.com/google/cm-creatives-drive-uploader and modified to the theme of the trafficking tool, it is limited to HTML5 creatives
 * Uploads any unprocessed creatives referenced in the associated Google Sheets
 * spreadsheet one-by-one, utilizing dedicated CM and Sheets API wrappers.
 *
 * @see {@link SheetsApi} and {@link CampaignManagerApi}. ++ add creative.id somewhere in the status
 */
function uploadCreatives() {
  var sheetsApi = new SheetsApi(SpreadsheetApp.getActiveSpreadsheet());
  var profileId = _fetchProfileId();
  var config = {data: {
    sheetName: CREATIVES_SHEET,
    startRow: 2,
    startCol: 1,
    advertiserId: 0,
    creativeName: 1,
    creativeDimensionsRaw: 2,
    creativeAssetName: 3,
    creativeAssetPath: 4,
    backupImageName: 5,
    backupImagePath: 6,
    backupImageClickThroughUrl: 7,
    },
  status: {column: 8, columnId: 'I', done: 'DONE'}
  }
  var cmApi = new CampaignManagerApi(profileId, config.data);

  var data = sheetsApi.getSheetData(
      config.data.sheetName, config.data.startRow, config.data.startCol);

  data.forEach(function(row, index) {
    if (row[config.data.advertiserId] && !row[config.status.column]) {
      try {
        var creativeResponse = cmApi.insertCreative(row);
        sheetsApi.setCellValue(
            config.data.sheetName,
            config.status.columnId + (config.data.startRow + index),
            creativeResponse.id); /** aps: replaced "done" status with creative ID to match overall theme */
        sheetsApi.setCellColourID(
          config.data.sheetName,
          config.status.columnId + (config.data.startRow + index) /** aps: add colour to creative ID to match overall theme */
        )
      } catch (e) {
        
        console.log(e);
        throw new Error(
            `Failed to upload Creative ${row[config.data.creativeName]} ` +
            `for Advertiser ${row[config.data.advertiserId]}! Check the logs ` +
            `at https://script.google.com/home/executions for more details.`);
        
      }
    } else SpreadsheetApp.getUi().alert('Finished creating the creatives!');
  });
}

/**
 * aps: Repurposed old creatiCreatives() for tracking creatives
 * Read creatives information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */ 
function createTrackingCreatives() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(T_CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,3);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
    var newCreative = _createOneTCreative(ss, values[i]);
    sheet.getRange('C' + rowNum)
        .setValue(newCreative.id)
        .setBackground('lightgray');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the tracking creatives!');
}

/**
 * A helper function which creates one creative via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleCreativeArray An array containing creative information
 * @return {object} Creative object
 */
function _createOneTCreative(ss, singleCreativeArray){
  var profileID = _fetchProfileId();

  var advertiserId = singleCreativeArray[0];
  var name = singleCreativeArray[1];

  var creativeResource =  {
    "name": name,
    "advertiserId": advertiserId,
    "active": true,
    "type": "TRACKING_TEXT"
  };

  var newCreative = CampaignManager.Creatives
      .insert(creativeResource, profileID);
  return newCreative;

}

/** 
 * 
 * Campaign - Creative Association Set
 * 
 * */
function createCampCreativeAss() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMP_CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,3);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
      _createOneCampCreativeAss(ss, values[i]);
      sheet.getRange('C' + rowNum)
          .setValue("DONE")
          .setBackground('lightgray');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the tracking creatives!');
}

function _createOneCampCreativeAss(ss,singleAssArray){
  var profileId = _fetchProfileId();

  var campaignId = singleAssArray[0];
  var creativeId = singleAssArray[1];

  var associationResource = {
    "creativeId": creativeId,
    "kind": "dfareporting#campaignCreativeAssociation"
  }
  var newCampCreAss = CampaignManager.CampaignCreativeAssociations
        .insert(associationResource,profileId,campaignId)
  return newCampCreAss
}


/**
 * 
 * Ads - Creatives Assignment set
 * 
 */
function createAdsCreativeAss() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ADS_CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();
  
  
  for (var i = 1; i < values.length; i++) { // exclude header row
    var rowNum = i+1;
    var status = sheet.getDataRange().getCell(rowNum,3);
    var first_col = sheet.getDataRange().getCell(rowNum,1);
    if (!first_col.isBlank() && status.isBlank()){
      _createOneAdsCreativeAss(ss, values[i]);
      sheet.getRange('C' + rowNum)
        .setValue("DONE")
        .setBackground('lightgray');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the tracking creatives!');
}

function _createOneAdsCreativeAss(ss,singleAdAssArray) {
  var profileId = _fetchProfileId();
  var adsId = singleAdAssArray[0];
  var creativeId = singleAdAssArray[1];

  var getAds = CampaignManager.Ads.get(profileId,adsId);
  
  /** assign creative */
  var nCA ={}
  if (getAds.creativeRotation.creativeAssignments == undefined) {
    var nCA = 0
    getAds.creativeRotation.creativeAssignments = []
  } else {
    var nCA = getAds.creativeRotation.creativeAssignments.length
  }
  getAds.creativeRotation.creativeAssignments[nCA] = {
  sslCompliant: true,
  creativeId: creativeId,
  active: false,
  clickThroughUrl: {
    defaultLandingPage: true 
    } 
  }
  var newAdCreAss = CampaignManager.Ads
          .update(getAds,profileId)

  return newAdCreAss
}


/** Listings */

/**
 * Using DCM API list all the advertisers this profile has added
 * and print them out on the sheet.
 */
function listAdvertisers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SITES_SHEET);
  var profileID = _fetchProfileId();

  // setup header row
  sheet.getRange('D1')
      .setValue('Advertiser Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('E1')
      .setValue('Advertiser ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');

  var advertisers = CampaignManager.Advertisers.list(profileID).advertisers;
  for (var i = 0; i < advertisers.length; i++) {
    var currentObject = advertisers[i];
    var rowNum = i+2;
    sheet.getRange('D' + rowNum)
        .setValue(currentObject.name)
        .setBackground('lightgray');
    sheet.getRange('E' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.id)
        .setBackground('lightgray');
  }
}

/** 
 * List Drive Folder file list to find creative file id
 *  
 * */

 function listCreativeFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SITES_SHEET);
  var folderId = _fetchFolderId()

  // setup header row
  sheet.getRange('G1')
      .setValue('File Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('H1')
      .setValue('File ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');

  var x = DriveApi.listFileIdsByFolder(folderId).items
  for (var i = 0; i < x.length; i++) {
    var currentObject = x[i];
    var rowNum = i+2
    sheet.getRange('G' + rowNum)
        .setValue(currentObject.title)
        .setBackground('lightblue')
    sheet.getRange('H' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.id)
        .setBackground('lemonchiffon')
  }
}

/** 
 * List creatives on a certain advertiser
 * 
 */
function listCreatives() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SITES_SHEET);
  var profileId = _fetchProfileId();

  // setup header row
  sheet.getRange('J1')
      .setValue('Creative Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('K1')
      .setValue('Creative ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('L1')
      .setValue('Advertiser ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');

  var creatives = CampaignManager.Creatives.list(profileId).creatives;
  for (var i = 0; i < creatives.length; i++) {
    var currentObject = creatives[i];
    var rowNum = i+2;
    sheet.getRange('J' + rowNum)
        .setValue(currentObject.name)
        .setBackground('lightblue');
    sheet.getRange('K' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.id)
        .setBackground('lemonchiffon');
    sheet.getRange('L' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.advertiserId)
        .setBackground('lavenderblush');
  }
}

