/*
// Code and documentation at : 
// https://github.com/bactisme/appsscript-sellsy-to-gsheet

1/ Add the oAuth1 Library :  
- Check last version here : https://github.com/googleworkspace/apps-script-oauth1
- or add library 1CXDCY5sqT9ph64fFwSzVtXnbjpSfWdRymafDrtIZ7Z_hwysTY7IIhi7s

3/ Run first time to make Configuration Page, then configure using column B

4/ 
- Search param not working ? (date > 01-01-2021)
- docs : https://api.sellsy.fr/documentation/methodes

*/

/**
 * Global Variable
 */
var CONSUMER_KEY = "";
var CONSUMER_TOKEN = "";
var ACCESS_KEY = "";
var ACCESS_TOKEN = "";
var RESULTS_SHEET = "";
var START_DATE = null;
var END_DATE = null;
var START_ROW = 2; // first row will be last update date and time
const DEFAULT_ENDPOINT = 'https://apifeed.sellsy.com/0';
var PRINT_COLUMNS = "ident,displayedDate,formatted_created,formatted_payDateCustom,thirdname,subject,totalAmountTaxesFree,taxesAmountSum,totalAmount,smartTags,step_label";

/**
 * Setup a configuration page in the current spreadsheet if not existing and read it to setup global variable
 */
function setupConfigurationSheet(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Configuration");
  if (!sheet){
    sheet = SpreadsheetApp.getActive().insertSheet();
    sheet.setName("Configuration");
  }
  sheet.getRange(1,1,8,1).setValues([["Configuration Page"],["CONSUMER_KEY"], ["CONSUMER_TOKEN"],["ACCESS_KEY"], ["ACCESS_TOKEN"], ["Sheet Name"], ["Start Date"], ["End Date"]]);

  var iconf = 2;
  CONSUMER_KEY = sheet.getRange(iconf++,2).getDisplayValue();
  CONSUMER_TOKEN = sheet.getRange(iconf++,2).getDisplayValue();
  ACCESS_KEY = sheet.getRange(iconf++,2).getDisplayValue();
  ACCESS_TOKEN = sheet.getRange(iconf++,2).getDisplayValue();
  RESULTS_SHEET =  sheet.getRange(iconf++,2).getDisplayValue();
  
  START_DATE =  sheet.getRange(iconf++,2).getValue();
  START_DATE = (new Date(START_DATE).getTime())/1000;

  END_DATE =  sheet.getRange(iconf++,2).getValue();  
  END_DATE = (new Date(END_DATE).getTime())/1000;

  DEFAULT_GET_LIST_SEARCH = {
    periodecreationDate_start: START_DATE,
    periodecreationDate_end: END_DATE,
  }
}

/**
 * Filter known problematic columns
 */
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

function sort_function(a, b){
  var createda = new Date(a.created);
  var createdb = new Date(b.created);

  if (createda.valueOf() > createdb.valueOf()){
    return 1;
  }else if (createda.valueOf() < createdb.valueOf()){
    return -1;
  }
  return 0;
}

var DEFAULT_GET_LIST_SEARCH = {
  periodecreationDate_start: +((new Date(2022,0,1)).getTime()/1000),
  periodecreationDate_end: +((new Date(2022,11,31)).getTime()/1000)
}

const DEFAULT_GET_LIST_PAGINATION = {
  nbperpage: 1000, //1000
  pagenum: 1
}

const DEFAULT_GET_LIST_ORDER = {
  direction: 'ASC',
  order: 'doc_displayedDate'
}

/**
 * Add a menu to force update
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SellSy')
      .addItem('Mise à jour depuis SellSy', 'getInvoiceAndCreditNote')
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
  getInvoice();
  return HtmlService.createHtmlOutput("post request received");
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

/**
 * Download Invoices, CréditNotes, merge them, sort them, filter fields and print it in the configured sheet
 */
function getInvoiceAndCreditNote(){
  setupConfigurationSheet();

  var invoices = getDocument('invoice'); //['invoice', 'creditnote']
  var creditnote = getDocument('creditnote'); //['invoice', 'creditnote']

  var alldocs = invoices.concat(creditnote);
  alldocs.sort(sort_function);

  var sheet = SpreadsheetApp.getActive().getSheetByName(RESULTS_SHEET);
  var columns = PRINT_COLUMNS.split(',');

  if (!sheet){
    sheet = SpreadsheetApp.getActive().insertSheet();
    sheet.setName(RESULTS_SHEET);
  }

  if (alldocs){
    // print header lines
    for(var i = 0; i < columns.length; i++){
      sheet.getRange(START_ROW,i+1).setValue(columns[i]);
    }

    // print data
    var row = START_ROW+1;

    var printed_values = null;
    printed_values = alldocs.map((obj) => {
      var field_value = [];
      for(var d = 0; d < columns.length; d++){
        var column_name = columns[d];
        field_value.push(field_filter(column_name, obj[column_name]));
      }
      return field_value;
    }); 

    sheet.getRange(row,1,printed_values.length, columns.length).setValues(printed_values);

    sheet.getRange(1,1).setValue("Maj : "+Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm"));
  }
}

/**
 * getDocument handle the loop arround the preparation of the request, and the loop arround API pagination, 
 */
function getDocument(docType, search=DEFAULT_GET_LIST_SEARCH, pagination=DEFAULT_GET_LIST_PAGINATION, includePayments='N', order=DEFAULT_GET_LIST_ORDER){
  var params = {
      doctype: docType,
      pagination: pagination,
      search : search,
      order: order
      /*,
      includePayments*/
    };

  var d = makeSellSyRequest(
    "Document.getList", 
    params
  );
  if (d && d.hasOwnProperty('response') && d.response.result){
    var results = Object.keys(d.response.result).map((k)=> {
      //d.response.result[k].iid = Number(d.response.result[k].id) // add a numeric id to sort it later
      return d.response.result[k];
    });

    // loop over remaining pages
    for(var i = 2; i <= d.response.infos.nbpages; i++){
      params.pagination.pagenum +=1; 

      var nd = makeSellSyRequest(
        "Document.getList", 
        params
      );
      if (nd && nd.hasOwnProperty('response') && nd.response.result){
        // aggregate results
        var new_results = Object.keys(nd.response.result).map((k)=> {
          //nd.response.result[k].iid = Number(nd.response.result[k].id)
          return nd.response.result[k];
        });
        results = results.concat(new_results);
      }else{
        break;
      }
    }

    return results;
  }
  return null;
}

/**
 * Do a request on SellSy API
 */
function makeSellSyRequest(method, params){

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
