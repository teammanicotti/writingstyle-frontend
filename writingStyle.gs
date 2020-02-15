var apiHost = "https://20dfd657.ngrok.io";
//var analyze_url_path = PropertiesService.getScriptProperties().getProperty("apiHost") + "/analyze";
var analyze_url_path = apiHost + "/analyze";
//var rec_ack_path = PropertiesService.getScriptProperties().getProperty("apiHost") + "/recAck";
var rec_ack_path = apiHost + "/recAck";
var similarityThreshold = parseFloat(PropertiesService.getScriptProperties().getProperty("similarityThreshold"));
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
            if (!(para_index in paragraphs)) {
                paragraphs[para_index] = ""
            }
            if (para_index in paragraphs) {
                paragraphs[para_index] = paragraphs[para_index] + "<div id=" + rec['uuid'] + " class='recommendationCard'>\n" +
                    " <div class=recHeader>\n" +
                    "   <div class=recHeaderText>" +
                    "     " + GetUserFriendlyType(rec['recommendation_type']) + "\n" +
                    "     <div class=recSubTitle>" + rec['text_to_highlight'] + "</div>\n" +
                    "   </div>\n" +
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_down' src=\"http://manicotti.se.rit.edu/thumbs-down.png\" alt=\"thumbs down\">\n" +
                    "     <img id='" + rec['uuid'] + "' class='recIconThumb thumbs_up' src=\"http://manicotti.se.rit.edu/thumbs-up.png\" alt=\"thumbs up\">\n" +
                    "   </div>\n";
                if(rec['new_values'].length > 1){
                    var counter = 0;
                    paragraphs[para_index] += "<div class=recText id='newValueOptions_" + rec['uuid'] + "'>";
                    rec['new_values'].forEach(function (newVal) {
                        paragraphs[para_index] += "<input type='radio' class='radio_" + rec['uuid'] + "' id='" + counter + "'/>"; //'" +rec['new_values'][counter] + "</input>";
                        paragraphs[para_index] += "<span id='value_" + rec['uuid'] + "' class='" + counter + "'>" + rec['new_values'][counter] + "</span><br>";
                        counter++;
                    });
                    paragraphs[para_index] +="</div>";
                }
                else{
                    paragraphs[para_index] += "<div class=recText>" + GetRecString(rec['recommendation_type']) + rec['new_values'][0] + "</div>\n";
                }

                paragraphs[para_index] +="</div>\n";

                HighlightText(rec['original_text'], '#f69e42')
            }
        }
    });
    Object.keys(paragraphs).sort().forEach(function(paragraphNum){
        var num = (parseInt(paragraphNum, 10) + 1);
        if(paragraphNum > 0){
            html += "<div>" +
                        "<div class='pargraphLabel'>" +
                            "Paragraph: " + num + "" +
                            "<img id='paragraph_" + num + "' class='collapse' src='http://manicotti.se.rit.edu/plus.png' alt='plus'>" +
                        "</div>\n" +
                    "<div class='paragraph_recs' id='recommendations_" + num + "'>" + paragraphs[paragraphNum]  + "</div>";
                "</div>\n"
        }
        else{
            //First item shouldn't have a title because its baked in with the settings options
            html += "<div>" +
                        "<div class='paragraph_recs' id='recommendations_1'>" + paragraphs[paragraphNum]  + "</div>";
                    "</div>"
        }
    });

    //Find items currently in the doc and known to be hidden
    var newCache = calculate_new_hidden_cache(mostRecentRecs, hiddenItems);

    return [html, newCache];
}

function ShowErrorMultiSelect(count) {
    var ui = DocumentApp.getUi();
    if(count > 1){
        ui.alert('Please select only one option');
    }
    else {
        ui.alert('Please select an option');
    }

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

function DoSubstitution(recID, selected_index){
    var currentRecommendations = JSON.parse(cache.get("current_recs"));

    if(currentRecommendations !== null) {
        currentRecommendations.forEach(function (rec) {
            if (recID === rec['uuid'] && rec['is_replaceable']) {
                var body = DocumentApp.getActiveDocument().getBody();
                HighlightText(rec['original_text'], '#ffffff');
                //Logger.log("Replacing: '" + rec['original_text'] + "' with '" + rec['new_values'][selected_index] + "'");
                body.replaceText(rec['original_text'], rec['new_values'][selected_index]);
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
    UrlFetchApp.fetch(rec_ack_path + (accepted ? "?accepted=true" : "?accepted=false"), options).getContentText();

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
