var reqUrl = "https://manicotti.se.rit.edu";
var ananlyze_url_path = "/analyze";
var rec_ack_path = "/recAck";
var PurgeInterval = 10;

var analyzationsSinceLastPurge = 0;
var activeFileID = DocumentApp.getActiveDocument().getId();
var cache = CacheService.getDocumentCache();
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

function onClose(e) {
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var rangeObj = body.findText(".*");

    if(rangeObj !== null){
        text.setBackgroundColor(rangeObj.getStartOffset(), rangeObj.getEndOffsetInclusive(), '#ffffff')
    }
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
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var rangeObj = body.findText(".*");

    if(rangeObj !== null){
        text.setBackgroundColor(rangeObj.getStartOffset(), rangeObj.getEndOffsetInclusive(), '#ffffff')
    }


    cache.put("current_recs", JSON.stringify(results));

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
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_down' src=\"http://manicotti.se.rit.edu/thumbs-down.png\" alt=\"thumbs down\">\n" +
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_up' src=\"http://manicotti.se.rit.edu/thumbs-up.png\" alt=\"thumbs up\">\n" +
                    "   </div>\n" +
                    "   <div class=recText>" + GetRecString(rec['recommendation_type']) + rec['new_values'][0] + "</div>\n" +
                    "</div>\n";

                HighlightText(rec['original_text'], '#f69e42')
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

function HighlightText(stringText, color) {
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var rangeObj = body.findText(stringText);

    if(rangeObj !== null){
        text.setBackgroundColor(rangeObj.getStartOffset(), rangeObj.getEndOffsetInclusive(), color)
    }
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

function DoSubstitution(recID){
    var currentRecommendations = JSON.parse(cache.get("current_recs"));

    if(currentRecommendations !== null) {
        currentRecommendations.forEach(function (rec) {
            if (recID === rec['uuid']) {
                var body = DocumentApp.getActiveDocument().getBody();
                HighlightText(rec['original_text'], '#ffffff')
                body.replaceText(rec['original_text'], rec['new_values'][0]);
                return;
            }
        });
    }
    else {
        Logger.log("Current recommendations list is null");
    }
}

function UndoHighlighting(recID) {
    var currentRecommendations = JSON.parse(cache.get("current_recs"));

    if(currentRecommendations !== null) {
        currentRecommendations.forEach(function (rec) {
            if (recID === rec['uuid']) {
                var body = DocumentApp.getActiveDocument().getBody();
                HighlightText(rec['original_text'], '#ffffff')
                return;
            }
        });
    }
    else {
        Logger.log("Current recommendations list is null");
    }
}


function ThumbsClicked(uuid, accepted){
    var options = {"method": "get"};
    //UrlFetchApp.fetch(reqUrl + rec_ack_path + (accepted ? "?accepted=true" : "?accepted=false"), options).getContentText();
}

/**
 * Test Function that retrieves a json file
 * @param document
 * @returns {*[]}
 */
function getRecommendation(document) {
    var results = [];

    var payload = {
        "text": document.getBody().getText(),
        "paragraphs": [document.getBody().getText()]
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload" : JSON.stringify(payload)
    };

    var response = UrlFetchApp.fetch(reqUrl + ananlyze_url_path, options).getContentText();
    var responseObj = JSON.parse(response);
    results = results.concat(responseObj);

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
