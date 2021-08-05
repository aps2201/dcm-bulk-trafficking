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

// Global variables/configurations
var DCMProfileID = 'DCMProfileID';
var AUTO_POP_HEADER_COLOR = '#a4c2f4';
var USER_INPUT_HEADER_COLOR = '#b6d7a8';
var AUTO_POP_CELL_COLOR = 'lightgray';
var AUTO_ATT_HEADER_COLOR = 'lightgreen'

// Data range values
var DCMUserProfileID = 'DCMUserProfileID';
var CreativeFolderID = 'CreativeFolderID';
var TimeZone = 'TimeZone'

// sheet names
var SETUP_SHEET = 'Setup';
var SITES_SHEET = 'Lists';
var CAMPAIGNS_SHEET = 'Campaigns';
var PLACEMENTS_SHEET = 'Placements';
var ADS_SHEET = 'Ads';
var CREATIVES_SHEET = 'Creatives';
var LANDING_PAGES_SHEET = 'LandingPages';

/**
 * Helper function to get DCM Profile ID.
 * @return {object} DCM Profile ID.
 */
function _fetchProfileId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName(DCMUserProfileID);
  return range.getValue();
}


/**
 * Find and clear, or create a new sheet named after the input argument.
 * @param {string} sheetName The name of the sheet which should be initialized.
 * @param {boolean} lock To lock the sheet after initialization or not
 * @return {object} A handle to a sheet.
 */
function initializeSheet_(sheetName, lock) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  if (lock) {
    sheet.protect().setWarningOnly(true);
  }
  return sheet;
}


/**
 * Initialize all tabs and their header rows
 */
function setupTabs() {
  _setupSetupSheet();
  _setupSitesSheet();
  _setupLandingPagesSheet();
  _setupCampaignsSheet();
  _setupPlacementsGroupsSheet();
  _setupCreativesSheet(); /** aps: original renamed as _setupCreativesSheet_OLD();*/
  _setupAdsSheet();
  
}

/**
 * Initialize the Setup sheet and its header row
 * @return {object} A handle to the sheet.
*/
function _setupSetupSheet() {  
  var sheet = initializeSheet_(SETUP_SHEET, false);
  var cell;
  
  sheet.getRange('B2').setValue("DCM Bulk Trafficking");
  sheet.getRange('B2:C2')
      .mergeAcross() /** aps: aesthetics... */
      .setFontWeight('bold')
      .setWrap(true)
      .setBackground(AUTO_POP_HEADER_COLOR)
      .setFontSize(12);

  sheet.getRange('B3')
      .setValue('For any questions contact panesh@ or jenicarawtani@ or toczos@'); /**original authors */
  sheet.getRange('B3:C3')
      .mergeAcross()
      .setValue('For any questions contact andaru.s@fivestones.net')      
      .setWrap(true);
  
  var instructions = [
    "Initial setup:",
    "# Make a copy of this template trix",
    "# In Menu, Go to [Tools] > [Script editor]",
    "# [New browser tab of the appscript] [Resources] > "+
    "[Cloud Platform Project...] If an appscript project "+
    "is linked, click on it and skip to enabling the API (Step 7)",
    "# [New browser tab of the appscript] [Resources] > "+
    "[Cloud Platform Project...] Click 'View API Console'",
    "# [New browser tab of cloud project] Create a new cloud "+
    "project and make a note of the project Id",
    "# [Browser tab of the appscript] [Resources] > "+
    "[Cloud Platform Project...] Click 'View API Console' and "+
    "enter the project ID and Click 'Change Project'",
    "# [New browser tab of cloud project] [Library] Search and "+
    "enable \"DCM/DFA Reporting And Trafficking API\"",
    "# [Browser tab of the appscript] [Resources] > "+
    "[Advanced Google Services]",
    "# [Advanced Google Services] Enable \"DCM/DFA Reporting And Trafficking API\"",
    "# Go back to appscript tab, select OK and close the [Advanced Google Services] window",
    null,
    null,
    "How to use:",
    "# Enter DCM Profile ID in C5 of this tab",
    "# [Sites tab] Retrieve the list of sites and IDs by [DCM Functions] > [List Sites]",
    "# [Campaigns tab] Bulk create Campaigns by [DCM Functions] > [Bulk Create Campaigns]",
    "# [Placements tab]  Bulk create Placements groups by [DCM Functions] > [Bulk Create Placements]",
    "# [Ads tab] Bulk create Ads by [DCM Functions] > [Bulk Create Ads]",
    "# [Creatives tab] Bulk create Creatives by [DCM Functions] > [Bulk Create Creatives]",
    "# [LandingPages tab] Bulk create Landing Pages by [DCM Functions] > [Bulk Create Landing Pages]"
  ]
  
  for(var i=0; i<instructions.length; i++) {
    cell = i+2
    var count = instructions[i] == null ? -1 : (i==0 ? 0 : count+1);
    var value = instructions[i] == null ? null : instructions[i].replace('#', count + ')');
    sheet.getRange('E' + cell).setValue(value);
    
    if (count == 0) {
      sheet.getRange('E' + cell + ':M' + cell)
        .setFontWeight("bold")
        .setWrap(true)
        .setBackground(AUTO_POP_HEADER_COLOR)
        .setFontSize(12);
    }
  }
    
  sheet.getRange('E' + (cell+3)).setValue("Legends")
      .setFontWeight("bold")
      .setFontSize(12);
  sheet.getRange('E' + (cell+4))
      .setValue("Green Cells / Columns are for input");
  sheet.getRange('E' + (cell+5))
      .setValue("Blue Cells /Columns are for the script to populate (do not edit)");
  
  sheet.getRange('E' + (cell+3) + ':M' + (cell+3))
      .setBackground("#f9cb9c");
  sheet.getRange('E' + (cell+4) + ':M' + (cell+4))
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E' + (cell+5) + ':M' + (cell+5))
      .setBackground(AUTO_POP_HEADER_COLOR);
  
  sheet.getRange('B5').setValue("User Profile ID")
                      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C5').setBackground(USER_INPUT_HEADER_COLOR);

  sheet.getRange("B5:C5").setFontWeight("bold").setWrap(true);
  
  /** aps: Creative Folder ID should go here */ 
  sheet.getRange('B7').setValue("Creative Folder ID")
                      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C7').setBackground(USER_INPUT_HEADER_COLOR);

  sheet.getRange("B7:C7").setFontWeight("bold").setWrap(true);
  
  /** aps: Time Zone should go here */ 
  sheet.getRange('B9').setValue("Time Zone")
                      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C9').setBackground(USER_INPUT_HEADER_COLOR);

  sheet.getRange("B9:C9").setFontWeight("bold").setWrap(true);

  var rule = SpreadsheetApp.newDataValidation().requireValueInList([
    'GMT+7',
    'GMT+8',
    'GMT+9',
    'GMT+10',
    'GMT+11',
    'GMT+12',
    'GMT-11',
    'GMT-10',
    'GMT-9',
    'GMT-8',
    'GMT-7',
    'GMT-6',
    'GMT-5',
    'GMT-4',
    'GMT-3',
    'GMT-2',
    'GMT-1',
    'GMT',
    'GMT+1',
    'GMT+2',
    'GMT+3',
    'GMT+4',
    'GMT+5',
    'GMT+6'
]).build();
  sheet.getRange('C9').setDataValidation(rule)
  return sheet;

}

/**
 * Initialize the Sites sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupSitesSheet() {
  var sheet = initializeSheet_(SITES_SHEET, false);

  sheet.getRange('A1')
      .setValue('Site Name')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Directory Site ID')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue('Advertiser Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('E1')
      .setValue('Advertiser ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('G1')
      .setValue('File Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('H1')
      .setValue('File ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
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
  sheet.getRange('N1')
      .setValue('Landing Page Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('O1')
      .setValue('Landing Page ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('P1')
      .setValue('Landing Page URL')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('Q1')
      .setValue('Advertiser ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  
  
  sheet.getRange('A1:Q1').setFontWeight('bold').setWrap(true);
  return sheet;
}

/**
 * Initialize the Campaigns sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupCampaignsSheet() {
  var sheet = initializeSheet_(CAMPAIGNS_SHEET, false);

  sheet.getRange('A1')
      .setValue('DCM Advertiser ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Campaign Name*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue('Landing Page ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue('Start Date*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1')
      .setValue('End Date*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1')
      .setValue('Campaign ID (auto-populated; do not edit)')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('A1:F1')
      .setFontWeight('bold')
      .setWrap(true);
  return sheet;

}

/**
 * Initialize the Placements sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupPlacementsGroupsSheet() {
  var sheet = initializeSheet_(PLACEMENTS_SHEET, false);

  sheet.getRange('A1')
      .setValue('Campaign ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Placement Name*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue('Site ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue('Compatibility*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1')
      .setValue('Size*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1')
      .setValue('Pricing Schedule Start Date*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1')
      .setValue('Pricing Schedule End Date*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1')
      .setValue('Pricing Schedule Pricing Type*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('I1')
      .setValue('Tag Formats*')
      .setBackground(USER_INPUT_HEADER_COLOR);

  sheet.getRange('J1').setValue('Placement ID (do not edit; auto-filling)')
       .setBackground(AUTO_POP_HEADER_COLOR);

  sheet.getRange('A1:J1').setFontWeight('bold').setWrap(true);

  /**
   * Set data validation rules
   */
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['APP','APP_INTERSTITIAL','DISPLAY','DISPLAY_INTERSTITIAL','IN_STREAM_AUDIO','IN_STREAM_VIDEO']).build();
  sheet.getRange('D2:D100').setDataValidation(rule)
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['PRICING_TYPE_CPM','PRICING_TYPE_CPA','PRICING_TYPE_CPC','PRICING_TYPE_CPM_ACTIVEVIEW','PRICING_TYPE_FLAT_RATE_CLICKS',
  'PRICING_TYPE_FLAT_RATE_IMPRESSIONS']).build();
  sheet.getRange('H2:H100').setDataValidation(rule)
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['DEFAULT','PLACEMENT_TAG_STANDARD','PLACEMENT_TAG_IFRAME_JAVASCRIPT','PLACEMENT_TAG_IFRAME_ILAYER','PLACEMENT_TAG_INTERNAL_REDIRECT','PLACEMENT_TAG_JAVASCRIPT','PLACEMENT_TAG_INTERSTITIAL_IFRAME_JAVASCRIPT','PLACEMENT_TAG_INTERSTITIAL_INTERNAL_REDIRECT','PLACEMENT_TAG_INTERSTITIAL_JAVASCRIPT','PLACEMENT_TAG_CLICK_COMMANDS','PLACEMENT_TAG_INSTREAM_VIDEO_PREFETCH','PLACEMENT_TAG_INSTREAM_VIDEO_PREFETCH_VAST_3','PLACEMENT_TAG_INSTREAM_VIDEO_PREFETCH_VAST_4','PLACEMENT_TAG_TRACKING','PLACEMENT_TAG_TRACKING_IFRAME','PLACEMENT_TAG_TRACKING_JAVASCRIPT']).build();
  sheet.getRange('I2:I100').setDataValidation(rule)

  return sheet;
}

/**
 * Initialize the Advertisers sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupAdsSheet() {
  var sheet = initializeSheet_(ADS_SHEET, false);

  sheet.getRange('A1')
      .setValue('Campaign ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Ad Name*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue('Start Date and Time*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue('End Date and Time*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1')
      .setValue('Impression Ratio*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1')
      .setValue('Priority*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1')
      .setValue('Type*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1')
      .setValue('Placement ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('I1')
      .setValue('Creative ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);   
  sheet.getRange('J1')
      .setValue('Landing Page URL (optional)')
      .setBackground(USER_INPUT_HEADER_COLOR);    
  sheet.getRange('K1').setValue('Ad ID (auto-populated; do not edit)')
       .setBackground(AUTO_POP_HEADER_COLOR);

  sheet.getRange('A1:K1').setFontWeight('bold').setWrap(true);
  
  /**
   * Set data validation rules
   */
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['AD_SERVING_STANDARD_AD','AD_SERVING_CLICK_TRACKER_DYNAMIC','AD_SERVING_TRACKING']).build();
  sheet.getRange('G2:G100').setDataValidation(rule)
  
  return sheet;

}

/**
 * Initialize the LandingPages sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupLandingPagesSheet() {
  var sheet = initializeSheet_(LANDING_PAGES_SHEET, false);

  sheet.getRange('A1')
      .setValue("Advertiser ID*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue("Landing Page Name*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue("Landing Page URL*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue("Landing Page ID (do not edit; auto-filling)")
      .setBackground(AUTO_POP_HEADER_COLOR);
      
  sheet.getRange("A1:H1").setFontWeight("bold").setWrap(true);
  return sheet;
}

/** aps: I mess around with things down here */ 

/**
 * Initialize the Creatives sheet and its header row
 * @return {object} A handle to the sheet. */
function _setupCreativesSheet() {
  var sheet = initializeSheet_(CREATIVES_SHEET, false);

  sheet.getRange('A1')
      .setValue('Advertiser ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Campaign ID*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue('Creative Type')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue('Creative Name*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1').setValue('Creative Size ("width"x"height")*').setBackground(
      USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1').setValue('Creative Asset Name*').setBackground(
      USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1')
      .setValue('Creative Asset Path (Drive ID)*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1')
      .setValue('Creative Backup Image Name*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('I1')
      .setValue('Creative Backup Image Path (Drive ID)*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('J1')
      .setValue('Creative Backup Image Custom Landing Page URL (optional)*')
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('K1')
      .setValue('Creative ID (auto-populated; do not edit)')
      .setBackground(AUTO_POP_HEADER_COLOR);
  
  sheet.getRange('A1:K1').setFontWeight('bold').setWrap(true);
  
  /**
   * Set data validation rules
   */
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['HTML', 'HTML_IMAGE','TRACKING_TEXT']).build();
  sheet.getRange('C2:C100').setDataValidation(rule)


  return sheet;
}

/**
 * Helper function to get Drive Folder ID containing creatives.
 * @return {object} Drive Folder ID.
 */
function _fetchFolderId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName(CreativeFolderID);
  return range.getValue();
}

/**
 * Helper function to get Time Zone.
 * @return {object} Time Zone.
 */
function _fetchTZ() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName(TimeZone);
  var gmt = range.getValue()
  if (gmt.includes('+')){
    return gmt.replace('+','-') /** aps: invert GMT to match the destination Time Zone */
  } else {
    return gmt.replace('+','-')
  }
  
  
}

