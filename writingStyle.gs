/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Start', 'showSidebar')
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Writing Style');
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Scans the document for possible writing errors.
 */
function scanDocument() {
    var document = DocumentApp.getActiveDocument();
    // var header = document.getHeader(); // TODO
    var body = document.getBody();
    // var footer = document.getFooter(); // TODO
    // var footnotes = document.getFootnotes(); // TODO
    var paragraphs = body.getParagraphs();
    var bodyText = paragraphs[0].editAsText(); // Hardcoding paragraph zero for now.

    var recs = getRecommendationTest(document);

    return UpdateRecommendationsList(recs);
}


/**
 * Builds the html from results
 * @param data
 * @returns {string}
 * @constructor
 */
function UpdateRecommendationsList(data){
    var result = "";
    var paragraphs = {};

    var results = data[0].results;

    if(data === undefined){
        console.log("Data was not defined");
        return
    }
    results.forEach(function(rec) {
        var para_index = rec['paragraph_index'];
        console.log("Paragraph index: " + para_index);
        if(!(para_index in paragraphs)){
            paragraphs[para_index] = ""
        }
        if(para_index in paragraphs) {
            paragraphs[para_index] = paragraphs[para_index] + "<div id=" + rec['uuid'] + " class='recommendationCard'>\n" +
                " <div class=recHeader>\n" +
                "   <div class=recHeaderText>" +
                "     " + GetUserFriendlyType(rec['recommendation_type']) + "\n" +
                "     <div class=recSubTitle>" + rec['original_text'] + "</div>\n" +
                "   </div>\n" +
                "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_down' src=\"http://manicotti.se.rit.edu/thumbs-down.png\" alt=\"thumbs down\">\n" +
                "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_up' src=\"http://manicotti.se.rit.edu/thumbs-up.png\" alt=\"thumbs up\">\n" +
                "   </div>\n" +
                "   <div class=recText>" +GetRecString(rec['recommendation_type']) + rec['new_values'][0] + "</div>\n" +
                "</div>\n";
        }
    });
    Object.keys(paragraphs).sort().forEach(function(paragraphNum){
        result += "<div>\n" +
            "<div class='pargraphLabel'>Paragraph: " + paragraphNum + "</div>\n" +
            paragraphs[paragraphNum] +
            "</div>\n"
    });
    return result;
}

function GetUserFriendlyType(type){
    switch (type) {
        case "SimpleToCompound":
            return "Simple To Compound";
        case "PassiveToActive":
            return "Passive To Active";
        case "SentimentReversal":
            return "Sentiment Reversal";
    }
}

function GetRecString(type){
    switch (type) {
        case "SimpleToCompound":
            return "Consider changing to: ";
        case "PassiveToActive":
            return "Consider changing to: ";
        case "SentimentReversal":
            return "Consider changing to: ";
    }
}



function ThumbsUp(){
    Logger.log("Thumbs Up clicked");
    console.log("Thumbs Up clicked");
}

function ThumbsDown(){
    Logger.log("Thumbs Down clicked");
    console.log("Thumbs Down clicked");

}

/**
 * Test Function that retrieves a json file
 * @param document
 * @returns {*[]}
 */
function getRecommendationTest(document) {
    var results = [];
    var reqUrl = "http://bff53dca.ngrok.io/analyze";

    var payload = {
        "text": document.getBody().getText(),
        "paragraphs": [document.getBody().getText()]
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload" : JSON.stringify(payload)
    };

    var response = UrlFetchApp.fetch(reqUrl, options).getContentText();
    var responseObj = JSON.parse(response);
    results = results.concat(responseObj);

    return results;
}

/**
 * Send document text to the server & get json recommendations list
 * @param document
 */
function getRecommendation(document) {
    var text = DocumentApp.getActiveDocument().getBlob();


}

/**
 * User accepted recommendation, change the document text.
 * @param UID
 */
function implement_recommendation(UID) {
    //Find the rec in the dictionary of recs
    //Update the document text based on the rec data
}


/***
 * Devons old stuff, not sure what to do with it
 */
function simpleToComplexCheck(document) {
  var results = [];
  // var scriptProperties = PropertiesService.getScriptProperties();
  // var reqUrl = scriptProperties.getProperty("simpleToComplexEndpoint");
  // var similarityThreshold = scriptProperties.getProperty("similarityThreshold");

  var paragraphs = document.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    var startChar = document.newPosition(paragraph, 0).getOffset();
    var reqBody = {
      text: paragraph.getText(),
      paragraph_num: i,
      paragraph_start_char: startChar,
      threshold: similarityThreshold,
      gSuite: true
    };

    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(reqBody)
    };

    var response = UrlFetchApp.fetch(reqUrl, options).getContentText();
    var responseObj = JSON.parse(response);
    
    results = results.concat(responseObj);
  }
  
  Logger.log(results);
  return results;
}

/**
 * Performs a dummy check for passive voice in the text sample.
 * @param bodyText the Text instance for the writing sample
 */
function passiveVoiceCheck(document) {
  // var scriptProperties = PropertiesService.getScriptProperties();
  // var reqUrl = scriptProperties.getProperty("passiveVoiceEndpoint");
  var results = []

  var reqBody = {
    text: document.getBody().getText()
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(reqBody)
  };

  var response = UrlFetchApp.fetch(reqUrl, options).getContentText();
  var responseObj = JSON.parse(response);
  results = results.concat(responseObj['results']);
  
  return results;
}

/**
 * Performs the provided correction on the writing sample.
 * @param correctionObj object containing the paragraph, start and stop offsets,
 * and the correction to make
 */
function correctError(correctionObj) {
  var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
  var paragraph = paragraphs[correctionObj.paragraphNum];
  var paragraphText = paragraph.editAsText()
  paragraphText.deleteText(correctionObj.startOffset, correctionObj.endOffset);
  paragraphText.insertText(correctionObj.startOffset, correctionObj.replaceText);
}