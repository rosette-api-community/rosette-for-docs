/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
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
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Highlight for Docs')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Gets the text the user has selected. If there is no selection,
 * this function displays an error message.
 *
 * @return {Array.<string>} The selected text.
 */
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText != '') {
            text.push(elementText);
          }
        }
      }
    }
    return text;
  }
  else
  {
    var body = DocumentApp.getActiveDocument().getBody();

    // Use editAsText to obtain a single text element containing
    // all the characters in the document.
  
    var text = body.editAsText();
    return text.getText();
  }
}

function runEntityExtractionGS(user_key) {
  if(user_key === ""){
  throw new Error ("Please include a Rosette API key");
  } else {
  var text = getSelectedText();

  var headers = {
           "accept": "application/json",
           "accept-encoding": "gzip",
           "content-type": "application/json",
           "X-RosetteAPI-Key" : user_key
         }
         
  var payload  = JSON.stringify({"content": text.toString()});
     
  var options = {
      "method" : "post",
      "headers" : headers,
      "payload" : payload,
  };

  var response  = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/entities', options);
  return response.toString();
  }
}

function runNameTranslationGS(mention, type, user_key) {
    if(user_key === ""){
  throw new Error ("Please include a Rosette API key");
  } else {
  var text = mention;

  var headers = {
           "accept": "application/json",
           "content-type": "application/json",
           "X-RosetteAPI-Key" : user_key
         }
         
  var payload  = JSON.stringify({"name": mention, "entityType": type, "targetLanguage": "eng"});
     
  var options = {
      "method" : "post",
      "headers" : headers,
      "payload" : payload,
  };

  var response  = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/translated-name', options);
  var responseString = response.toString();
  
  var responseToReturn = '{\"originalName\":\"' + mention + '\", ' + responseString.substr(1);

  return responseToReturn;
  }
}

// self-explanatory helper function
// returns rosapi response
var rosapiRequest = function(user_key, payload, endpoint, admOutput) {
    if(user_key === ""){
      throw new Error ("Please include a Rosette API key");
  } else {
  var url = 'https://api.rosette.com/rest/v1/' + endpoint;
  if (admOutput) {
    url = url + '?output=rosette';
  }
  var headers = {
    'X-RosetteAPI-Key': user_key,
    'Accept': 'application/json',
    'Content-Type': 'application/json'
  };
  
  var options = {
    'method': 'post',
    'headers': headers,
    'payload': payload,
    'muteHttpExceptions': true
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);
  return response;
  }
}

/**
 * Gets the stored user preferences for the origin and destination languages,
 * if they exist.
 *
 * @return {Object} The user's origin and destination language preferences, if
 *     they exist.
 */
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var languagePrefs = {
    key: userProperties.getProperty('userKey'),
    targetLanguage: userProperties.getProperty('targetLanguage'),
    sourceLanguage: userProperties.getProperty('sourceLanguage'),
    insertPerson: userProperties.getProperty('insertPerson'),
    insertLocation: userProperties.getProperty('insertLocation'),
    insertOrganisation: userProperties.getProperty('insertOrganisation')
  };
  return languagePrefs;
}

// Entity class
function Entity(name, entityType, startOffset, endOffset) {
  this.name = name;
  this.endOffset = endOffset;
  this.entityType = entityType;
  this.startOffset = startOffset;
}

// Select entity from adm object
function getEntity(adm, index) {
  if (index > adm["attributes"]["entityMentions"]["items"].len) {
    throw 'Index larger than result array';
  } else {
    var tmp = adm["attributes"]["entityMentions"]["items"][index];
    return new Entity(tmp["normalized"], tmp["entityType"], tmp["startOffset"], tmp["endOffset"]);
  }
}

/**
 * @param {string} key rosapi key
 * @param {string} responseData Raw rosapi entities/linked result
 * @param {string} sourceLang language
 * @param {string} targetLang language
 * @param {boolean} insertPer insert person entity types
 * @param {boolean} insertLoc insert location entity types
 * @param {boolean} savePrefs Save preferences for next run
 */
function nameInsertion(key, responseData, sourceLang, targetLang, insertPer, insertLoc, insertOrg, savePrefs) {

  if (savePrefs === true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('userKey', key);
    userProperties.setProperty('targetLanguage', targetLang);
    userProperties.setProperty('sourceLanguage', sourceLang);
    userProperties.setProperty('insertPerson', insertPer);
    userProperties.setProperty('insertLocation', insertLoc);
    userProperties.setProperty('insertOrganisation', insertOrg);
  }
  
  var data = JSON.parse(responseData);
 
  var offset = 0;
  for (var index = 0; index < data["attributes"]["entityMentions"]["items"].length; index++) {
    var entity = getEntity(data, index);
    if ((entity.entityType === "PERSON" && insertPer ) ||
        (entity.entityType === "LOCATION" && insertLoc) || 
         entity.entityType === "ORGANIZATION" && insertOrg) {
      

      if (sourceLang === "auto") {
        sourceLang = data["attributes"]["languageDetection"]["detectionResults"][0]["language"];
      }

      var jsonText = JSON.stringify({
        "name": entity.name,
        "sourceLanguageOfUse": sourceLang,
        "entityType": entity.entityType,
        "targetLanguage": targetLang
      });

      var response = rosapiRequest(key, jsonText, 'name-translation', false)

      var json = JSON.parse(response.getContentText());

      var newText = "[" + json["translation"] + "]";
      
      var text = DocumentApp.getActiveDocument().getBody().editAsText();  
      text.insertText(entity["endOffset"] + offset, newText);
      offset = offset + newText.length;
    }
  }
}


/**
 * Rosapi entities/linked call on the text of the current document
 *
 * @param {string} key rosApi key
 * @param {boolean} savePrefs Store key for future calls
 * @return {string} The adm object.
 */
function admResults(key, savePrefs) {
  if (savePrefs == true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('userKey', key);
  }
  var text = DocumentApp.getActiveDocument().getBody().editAsText().getText();  
  
  var jsonText = JSON.stringify({
    "content": text.toString()
  });
  
  var response = rosapiRequest(key, jsonText, 'entities/linked', true);
  var json = response.getContentText();
  
  return json.toString();
}


function colorText(mention, type) {
  var body = DocumentApp.getActiveDocument().getBody();  
  var colorVal;
  switch (type) 
  {
    case 'PERSON': colorVal = '#E93824';
      break;
    case 'ORGANIZATION': colorVal = '#FEC238';
      break;
    case 'LOCATION': colorVal = '#2CAAE2';
      break;
  }
  
  var foundElement = body.findText(mention);
  while (foundElement != null) {
    // Get the text object from the element
    var foundText = foundElement.getElement().asText();

    // Where in the Element is the found text?
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();

    // Change the background color to color associated with the type
    foundText.setBackgroundColor(start, end, colorVal);

    // Find the next match
    foundElement = body.findText(mention, foundElement);
  }
}

/**
 */
function getWikiData(user_key) {
  if(user_key === ""){
      throw new Error ("Please include a Rosette API key");
  } else {
  var text = getSelectedText();
  var headers = {
            "accept": "application/json",
            "accept-encoding": "gzip",
            "content-type": "application/json",
            "X-RosetteAPI-Key" : user_key
          }
  var payloadJSON = {"content": null};
  payloadJSON.content = text.toString();
  var payload  = JSON.stringify(payloadJSON);
      
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : payload,
  };
  var response  = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/entities/linked/', options);
   
  return parseEntities(response);
  }
}

function parseEntities(response) {
  if(JSON.parse(response)["entities"].length === 0){
    throw new Error( "No wiki data found");
  }
 var qidArray = [];
 var entitiesArray = JSON.parse(response).entities;
 var wikiLink = "https://en.wikipedia.org/wiki/";

 for(var i = 0; i< entitiesArray.length; i++){
   var linkJSON = UrlFetchApp.fetch("https://www.wikidata.org/w/api.php?action=wbgetentities&ids=" + entitiesArray[i].entityId + "&languages=en&props=sitelinks&sitefilter=enwiki&format=json");
   var str = entitiesArray[i].entityId.toString();
   var title = JSON.parse(linkJSON).entities[str].sitelinks.enwiki.title;
   title = title.replace(" ", "_");
   var finalLink = wikiLink.concat(title);
   var wikiLinksArray = getWikiContent(finalLink);
 }
  
 return wikiLinksArray;
}

function getWikiContent(link) {
  var response = UrlFetchApp.fetch(link);
  var content = response.getContentText();
  
  var title = /<h1 id="firstHeading"([\s\S]*?)<\/h1>/.exec(content);
  var image = /<a href="\/wiki\/File:([\s\S]*?)<\/a>/.exec(content);
  var paragraph = /<p>([\s\S]*?)<\/p>/.exec(content);
  var finalContent = [title, image, paragraph];
  
  return finalContent;
}
