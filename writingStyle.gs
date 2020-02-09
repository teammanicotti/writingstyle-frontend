var analyze_url_path = PropertiesService.getScriptProperties().getProperty("apiHost") + "/analyze";
var analytics_url_path = PropertiesService.getScriptProperties().getProperty("apiHost") + "/analytics";
var similarityThreshold = parseFloat(PropertiesService.getScriptProperties().getProperty("similarityThreshold"));
var PurgeInterval = 10;

var analyzationsSinceLastPurge = 0;
var activeFileID = DocumentApp.getActiveDocument().getId();

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
function scanDocument(hiddenItems) {
    var document = DocumentApp.getActiveDocument();
    var recs = getRecommendation(document);

    return UpdateRecommendationsList(recs, hiddenItems);
}

function getDocumentID() {
    return DocumentApp.getActiveDocument().getId();
}

/**
 * Builds the html from results
 * @param data
 * @param hiddenItems - stringified list of hidden element ids
 * @returns {string}
 * @constructor
 */
function UpdateRecommendationsList(data, hiddenItems){
    var html = "";
    var paragraphs = {};
    var mostRecentRecs = [];
    var results = data[0].results;

    if(data === undefined){
        console.log("Data was not defined");
        return
    }
    Logger.log(hiddenItems);
    results.forEach(function(rec) {
        mostRecentRecs.push(rec['uuid']);
        if(hiddenItems === null || hiddenItems.toString().indexOf(rec['uuid']) === -1){ //If the user has not already accepted/rejected it
            var para_index = rec['paragraph_index'];
            console.log("Paragraph index: " + para_index);
            if (!(para_index in paragraphs)) {
                paragraphs[para_index] = ""
            }
            if (para_index in paragraphs) {
                paragraphs[para_index] = paragraphs[para_index] + "<div id=" + rec['uuid'] + " class='recommendationCard'>\n" +
                    " <div class=recHeader>\n" +
                    "   <div class=recHeaderText>" +
                    "     " + GetUserFriendlyType(rec['recommendation_type']) + "\n" +
                    "     <div class=recSubTitle>" + rec['original_text'] + "</div>\n" +
                    "   </div>\n" +
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_down '" +  "data-recommendationType=\"" + rec['recommendation_type'] + "\" src=\"http://manicotti.se.rit.edu/thumbs-down.png\" alt=\"thumbs down\">\n" +
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_up ' " + "data-recommendationType=\"" + rec['recommendation_type'] + "\" src=\"http://manicotti.se.rit.edu/thumbs-up.png\" alt=\"thumbs up\">\n" +
                    "   </div>\n" +
                    "   <div class=recText>" + GetRecString(rec['recommendation_type']) + rec['new_values'][0] + "</div>\n" +
                    "</div>\n";
            }
        }
    });
    Object.keys(paragraphs).sort().forEach(function(paragraphNum){
        html += "<div>\n" +
            "<div class='pargraphLabel'>Paragraph: " + paragraphNum + "</div>\n" +
            paragraphs[paragraphNum] +
            "</div>\n"
    });

    //Find items currently in the doc and known to be hidden
    var newCache = calculate_new_hidden_cache(mostRecentRecs, hiddenItems);

    return [html, newCache];
}

function GetUserFriendlyType(type){
    switch (type) {
        case "SimpleToCompound":
            return "Simple To Compound";
        case "PassiveToActive":
            return "Passive To Active";
        case "SentimentReversal":
            return "Sentiment Reversal";
        case "Comparative":
            return "Comparative";
        case "Superlative":
            return "Superlative";
        case "DirectIndirectObjectChecking":
            return "Direct/Indirect Objects";
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
        case "Comparative":
            return "Consider changing to: ";
        case "Superlative":
            return "Consider changing to: ";
        case "DirectIndirectObjectChecking":
            return "Consider changing to: ";
    }
}


function ThumbsClicked(uuid, accepted, recommendationType){
    var userKey = Session.getTemporaryActiveUserKey();
    if(!userKey){
        userKey = 'unknown';
    }
    var payload = {
        "userKey": userKey, 
        "recommendationType": recommendationType,
        "accepted": accepted,
        "recId": uuid
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload" : JSON.stringify(payload)
    };

    Logger.log("thumbsClicked: " + payload);
    UrlFetchApp.fetch(analytics_url_path, options).getContentText();
}

/**
 * Test Function that retrieves a json file
 * @param document
 * @returns {*[]}
 */
function getRecommendation(document) {
    var results = [];
    var paragraphs = document.getBody().getParagraphs();
    var paragraph_text = [];
    for (var i = 0; i < paragraphs.length; i++) {
      paragraph_text.push(paragraphs[i].getText());
    }

    var payload = {
        "text": document.getBody().getText(),
        "paragraphs": paragraph_text,
        "similarityThreshold": similarityThreshold
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload" : JSON.stringify(payload)
    };

    Logger.log("recommendationRequest: " + JSON.stringify(payload));
    var response = UrlFetchApp.fetch(analyze_url_path, options).getContentText();
    var responseObj = JSON.parse(response);
    results = results.concat(responseObj);
    Logger.log("recommendationResponse: " + response);

    return results;
}

function calculate_new_hidden_cache(mostRecentRecs, oldRecs) {
    var newCache = [];
    if(oldRecs !== null) {
        var oldRecsStr = oldRecs.toString();
        mostRecentRecs.forEach(function (entry) {
            if (oldRecsStr.indexOf(entry) > -1) {
                newCache.push(entry);
            }
        });
    }
    return newCache;
}
