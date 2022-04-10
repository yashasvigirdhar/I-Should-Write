var LOG = false;

log = function (text) {
  if (LOG) {
    Logger.log(text);
  }
};

// Add a custom menu to the active document, including a separator and a sub-menu.
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Writing')
    .addItem('Show word count', 'showWordCountApi')
    .addToUi();
}

/**
 * Shows word count of the selected text in an alert.
 */
function showWordCountApi() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (selection) {
    var count = getWordCountForSelection(selection);
    DocumentApp.getUi().alert(count);
  } else {
    DocumentApp.getUi().alert('Nothing is selected.');
  }
}

function getWordCountForSelection(selection) {
  var count = 0;
  var elements = selection.getRangeElements();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i].getElement();
    if (element.editAsText) {
      var type = element.getType();
      log("element type: " + type);
      var curString = '';
      if (type == DocumentApp.ElementType.PARAGRAPH) {
        curString = element.asParagraph().getText();
      } else if (type == DocumentApp.ElementType.TEXT) {
        curString = element.asText().getText();
        // TEXT element can be partial too
        if (elements[i].isPartial) {
          log("Partial element, start:" + elements[i].getStartOffset() + " , end: " + elements[i].getEndOffsetInclusive() + ", **curString**: " + curString);
          curString = curString.substring(elements[i].getStartOffset(), elements[i].getEndOffsetInclusive() + 1);
        }
      }
      else {
        log("Unknown element type: " + type);
      }
      count += countWords(curString);
      log("Count reached: " + count);
    } else {
      log("Element can't be edited as text: " + element.getType());
    }
  }
  return count;
}

function countWords(str) {
  log("Splitting : " + str)
  const arr = str.split(' ');
  return arr.filter(word => word !== '').length;
}
