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

function stuff() {
  return DocumentApp.getActiveDocument().getBody().editAsText().findText('was sad').getStartOffset();
}

function scanDocument() {
  var document = DocumentApp.getActiveDocument();
  // var header = document.getHeader(); // TODO
  var body = document.getBody();
  // var footer = document.getFooter(); // TODO
  // var footnotes = document.getFootnotes(); // TODO
  var paragraphs = body.getParagraphs();
  var bodyText = paragraphs[0].editAsText();
  var passiveVoiceResult = passiveVoiceCheck(bodyText);

  return passiveVoiceResult;
}

function passiveVoiceCheck(bodyText) {
  var result = bodyText.findText("was sad");
  if (result !== null) {
    return {
      paragraphNum: 0,
      elementText: result.getElement().asText().getText(),  
      startOffset: result.isPartial() ? result.getStartOffset() : 0,
      endOffset: result.isPartial() ? result.getEndOffsetInclusive() : result.getElement().asText().getText().length - 1
    };
  } else {
    return false;
  }
}

function correctError(correctionObj) {
  var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
  var paragraph = paragraphs[correctionObj.paragraphNum];
  var paragraphText = paragraph.editAsText()
  paragraphText.deleteText(correctionObj.startOffset, correctionObj.endOffset);
  paragraphText.insertText(correctionObj.startOffset, correctionObj.replaceText);
}