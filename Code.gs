
/** --- CONFIGURATION --- */

/*
// Code ans documentation at : 
// https://github.com/bactisme/appsscript-sellsy-to-gsheet

1/ Add the oAuth1 Library :  
- Check last version here : https://github.com/googleworkspace/apps-script-oauth1
- or add library 1CXDCY5sqT9ph64fFwSzVtXnbjpSfWdRymafDrtIZ7Z_hwysTY7IIhi7s

2/ Create an app here : https://www.sellsy.fr/developer/my-apps and past keys
// Variables to define 
var CONSUMER_KEY = 
var CONSUMER_TOKEN = 
var ACCESS_KEY = 
var ACCESS_TOKEN = 

3/ Customize in CUSTOM section

4/ Warnings
- The script only take 1000 first invoices (only first api page)
- See search param (date > 01-01-2021)
- docs : https://api.sellsy.fr/documentation/methodes

5/ WIP 
// configuration screen

*/

/** --- CUSTOM --- */

var PRINT_COLUMNS = "ident,displayedDate,formatted_created,formatted_payDateCustom,thirdname,subject,totalAmountTaxesFree,taxesAmountSum,totalAmount,smartTags,step_label";
var RESULTS_SHEET = "Factures 2021";
var FILTER_FUNCTION = "field_filter";
var START_ROW = 2;

function field_filter(column_name, data){
  if (column_name == "smartTags"){
    var s = [];
    for(var tag in data) {
      s.push(data[tag]['word']);
    }
    return s.join(',');
  }
  if (column_name == "subject"){
    var s = data.split('<br />');
    s = s[0].replace(/<[^>]+>/g, "");
    return s;
  }
  if (column_name.indexOf("Amount")> -1 || column_name.indexOf("amount") < -1){
    return parseFloat(data);
  }
  return data;
}

const DEFAULT_GET_LIST_SEARCH = {
  periodecreationDate_start: +((new Date(2021,0,1)).getTime()/1000),
  periodecreationDate_end: +((new Date(2021,11,31)).getTime()/1000)
}

const DEFAULT_GET_LIST_PAGINATION = {
  nbperpage: 1000,
  pagenum: 1
}

const DEFAULT_GET_LIST_ORDER = {
  direction: 'ASC',
  order: 'doc_displayedDate'
}

/** --- CODE --- */

const DEFAULT_ENDPOINT = 'https://apifeed.sellsy.com/0';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SellSy')
      .addItem('Mise à jour depuis SellSy', 'getInvoice')
      //.addSeparator()
      //.addItem('Paramétrage', 'setupScreen')
      .addToUi();
}

/**
 * Receive Webhook
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActive();
  Logger.log(e);
  getInvoice();
  return ContentService.createTextOutput(ss.getName());
}

function doPost(e) {
  Logger.log(e);
  return HtmlService.createHtmlOutput("post request received");
}

function setupScreen() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var start_date = documentProperties.getProperty('start_date');

  var ui = SpreadsheetApp.getUi();
  
  start_date = ui.prompt('Date de début', 'Sous la forme DD/MM/YYYY', ui.ButtonSet.OK);
}

function logCallbackUrl() {
  var service = getSellsyService();
  Logger.log(service.getCallbackUrl());
}

function getSellSyService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth1.createService('sellsy')
      // Set the endpoint URLs.
      .setAccessTokenUrl('https://apifeed.sellsy.com/0/access_token')
      .setRequestTokenUrl('https://apifeed.sellsy.com/0/request_token')
      .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
      // Set the consumer key and secret.
      .setConsumerKey(CONSUMER_KEY)
      .setConsumerSecret(CONSUMER_TOKEN)
      .setSignatureMethod('PLAINTEXT')
      //.setParamLocation('post-body')
      .setAccessToken(ACCESS_KEY, ACCESS_TOKEN);
}

function getInvoice(){
  var invoices = getDocument('invoice');
  var sheet = SpreadsheetApp.getActive().getSheetByName(RESULTS_SHEET);
  var columns = PRINT_COLUMNS.split(',');
  Logger.log(invoices);

  if (invoices){
    // print header lines
    for(var i = 0; i < columns.length; i++){
      sheet.getRange(START_ROW,i+1).setValue(columns[i]);
    }
    // print data
    var row = START_ROW+1;
    for(var line in invoices) {
      for(var d = 0; d < columns.length; d++){
        var column_name = columns[d];
        var data = invoices[line][column_name];
        if (FILTER_FUNCTION != ""){
          data = this[FILTER_FUNCTION](column_name, data);
        } 
        sheet.getRange(row,d+1).setValue(data);
      }
      row++;
    }
    sheet.getRange(1,1).setValue("Maj : "+Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm"));
  }
}

function getDocument(docType, search=DEFAULT_GET_LIST_SEARCH, pagination=DEFAULT_GET_LIST_PAGINATION, includePayments='N', order=DEFAULT_GET_LIST_ORDER){
  var params = {
      doctype: docType,
      pagination: pagination,
      search : search,
      order: order
      /*,
      includePayments*/
    };
  var d = makeRequest(
    "Document.getList", 
    params
  );
  if (d && d.hasOwnProperty('response') && d.response.result){
    return d.response.result;
  }
  return null;
}

function makeRequest(method, params, pagination){

  const postData = {
    request: 1,
    io_mode: 'json',
    do_in: JSON.stringify({
      method: method,
      params: params
    }),
  };

  var fetchParams = {
    method: 'post',
    payload: postData
  };

  var service = getSellSyService();
  var data = service.fetch(DEFAULT_ENDPOINT + "/", fetchParams);
  if (data.getResponseCode() == 200){
    return JSON.parse(data.getContentText());
  }
  return null;
}
