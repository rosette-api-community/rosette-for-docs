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
      .setTitle('Rosette for Docs')
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
/*
* @param {String} user_key API key for Rosette
*
* @return {Object} JSON containing entities extracted from text with /entities
*/
function runEntityExtractionGS(user_key) {
  
  if(user_key === ""){
  throw new Error ("Please include a Rosette API key, 85");
  }
  var text = getSelectedText();

  var headers = {
           "accept": "application/json",
           "accept-encoding": "gzip",
           "content-type": "application/json",
           "X-RosetteAPI-App": "GoogleDocs",
           "X-RosetteAPI-Key" : user_key
         }
         
  var payload  = JSON.stringify({"content": text.toString()});
     
  var options = {
      "method" : "post",
      "headers" : headers,
      "payload" : payload,
  };
  var response  = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/entities', options);
  var json = JSON.parse(response);
  if(json["code"]){
     response = tooManyRequests('https://api.rosette.com/rest/v1/entities',options);
  }
  return response.toString();
}

/*
* @param {String} mention name to translate
* @param {String} type entity type
* @param {String} user_key API key for rosette
* @param {String} sourceLang code representing source language (optional, default auto)
* @param {String} targetLang code representing target language
*
* @return {String} translationPair string containing mention and its translation, wrapped in directional markers
*/
function runNameTranslationGS(mention, type, user_key, sourceLang, targetLang) {
  if(user_key === ""){
    throw new Error ("Please include a Rosette API key, 112");
  }
  var headers = {
           "accept": "application/json",
           "content-type": "application/json",
           "X-RosetteAPI-Key" : user_key
         };
    
  if(sourceLang == "auto"){
    sourceLang = getLanguage(user_key);
  }
  if( !useAPI(sourceLang,targetLang)){
    var response = "no_translation";
    if(sourceLang == "eng" || sourceLang == "spa" || sourceLang == "ara" || sourceLang == "zho"){
      var entityId = linkEntity(mention, user_key, sourceLang);
      if(entityId) var name = translateNameWiki(entityId, targetLang);
      if(name)  response = name;
    }
  }
  else{
    var response = translateNameRosette(mention, type, sourceLang, targetLang, user_key);  
  }
  var translation = addDirectionWrappers(response, targetLang);
  var translationPair = addDirectionWrappers(mention+" ("+translation+")", sourceLang);
  return translationPair
}

// self-explanatory helper function
// returns rosapi response
var rosapiRequest = function(user_key, payload, endpoint, admOutput) {
  if(user_key === ""){
      throw new Error ("Please include a Rosette API key, 143");
  } 
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
  var json = JSON.parse(response);
  if(json["code"]){
    response = tooManyRequests(url,options);
  }
  return response;
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

function setProperties(key, sourceLang, targetLang, insertPer, insertLoc, insertOrg){
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('userKey', key);
    userProperties.setProperty('targetLanguage', targetLang);
    userProperties.setProperty('sourceLanguage', sourceLang);
    userProperties.setProperty('insertPerson', insertPer);
    userProperties.setProperty('insertLocation', insertLoc);
    userProperties.setProperty('insertOrganisation', insertOrg);
}


// Entity class
function Entity(name, entityType, startOffset, endOffset, id) {
  this.name = name;
  this.endOffset = endOffset;
  this.entityType = entityType;
  this.startOffset = startOffset;
  this.entityId = id;
}

// Select entity from adm object
function getEntity(adm, index, linked) {
  if (index >= adm["attributes"]["entityMentions"]["items"].length) {
    throw 'Index larger than result array'; } 
  var entity = adm["attributes"]["entityMentions"]["items"][index];
  var other = null;
  if(linked) other = getResolvedEntity(entity,adm["attributes"]["resolvedEntities"]["items"]);
  if(other){
    return new Entity(entity["normalized"], entity["entityType"], entity["startOffset"], entity["endOffset"], other["entityId"]);
  }
  return new Entity(entity["normalized"],entity["entityType"],entity["startOffset"],entity["endOffset"],null);
}

//matches mentions to a resolved entity if possible
function getResolvedEntity(entity,data){
  var id = entity["coreferenceChainId"];
  for(var i = 0; i < data.length; i++){
    if(data[i]["coreferenceChainId"] == id) return data[i];
  }  
    return null;
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
function nameInsertion(key, responseData, sourceLang, targetLang, insertPer, insertLoc, insertOrg, savePrefs, linked) {
  if (savePrefs === true) {
    setProperties(key, sourceLang, targetLang, insertPer, insertLoc, insertOrg);
  }
  if(!insertPer && !insertLoc && !insertOrg) return;
  
  var offset = 0;
  var data = JSON.parse(responseData);
  if (sourceLang === "auto") {
        sourceLang = data["attributes"]["languageDetection"]["detectionResults"][0]["language"];
  }
  
  for (var index = 0; index < data["attributes"]["entityMentions"]["items"].length; index++) {
    var entity = getEntity(data, index, linked);
    if ((entity.entityType === "PERSON" && insertPer ) ||
        (entity.entityType === "LOCATION" && insertLoc) || 
         entity.entityType === "ORGANIZATION" && insertOrg) {
      
      if (sourceLang === "auto") {
        sourceLang = data["attributes"]["languageDetection"]["detectionResults"][0]["language"];
      }
      var newText = "no_translation";
      if(useAPI(sourceLang, targetLang)){
        var name = translateNameRosette(entity.name,entity.entityType,sourceLang,targetLang,key);
        if(name) newText =  name;
      }
      else if(linked){
        if(sourceLang == "eng"|| sourceLang == "spa" || sourceLang == "ara" || sourceLang == "zho"){
          var id = entity.entityId;
          if(id){
            var name = translateNameWiki(id, targetLang);
            if(name) newText = name;
          } 
        }
      }
      var text = DocumentApp.getActiveDocument().getBody().editAsText();
      var translation = addDirectionWrappers(newText,targetLang);
      Logger.log(entity.name + " " + entity["startOffset"] + " " + entity["endOffset"]);
      var translationPair = addDirectionWrappers(entity.name +" ["+translation+"] ", sourceLang);
      Logger.log(translationPair);
      //Logger.log(text.getText());
      text.deleteText(entity["startOffset"]+offset,entity["endOffset"]+offset);
      //Logger.log(text.getText());
      text.insertText(entity["startOffset"] + offset, translationPair);
      var colorVal;
      switch (entity.entityType) 
      {
        case 'PERSON': colorVal = '#E93824';
         break;
        case 'ORGANIZATION': colorVal = '#FEC238';
         break;
        case 'LOCATION': colorVal = '#2CAAE2';
         break;
      }
      text.setBackgroundColor(entity["startOffset"]+offset, entity["startOffset"]+offset+translationPair.length - 3, colorVal);
      //offset = offset + (translationPair.length - entity.name.length);
      offset = offset + (translationPair.length - Math.abs(entity["startOffset"] - entity["endOffset"])) -1;
    }    
  }
}

/*
* Certain languages are supported by /name-translation but not entity linking
* inserts these into the text
*
*/
function notLinkedInsertion(key, sourceLang, targetLang, insertPer, insertLoc, insertOrg, savePrefs) {
  var text = getSelectedText();
  var data = rosapiRequest(key, JSON.stringify({"content": text.toString()}), "entities", true);
  nameInsertion(key, data, sourceLang, targetLang, insertPer, insertLoc, insertOrg, savePrefs, false);
}

//inserts direction markers for overall document
function documentOrientation(sourceLang){
  var LRlang = ["eng","spa","zho","jpn","ita","nld","fra","deu", "ind","kor","por","rus"];
  var RLlang = ["ara","urd","heb","pus","fas"];
  var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
  var text = "";
  for(var index = 0; index < paragraphs.length; index++){
    if(LRlang.indexOf(sourceLang) != -1){
      text = paragraphs[index].editAsText();
      text.insertText(0, "\u202a")
      text.appendText("\u202c")
     paragraphs[index].setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    }
    else{
      text.insertText(0, "\u202b")
      text.appendText("\u202c")
      paragraphs[index].setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
}

/*
* @param {String} code 3 letter Rosette language code
* @return {String code wiki language code
*/ 
function changeLangCode(code){
  if(code == "ara") return "ar";
  if(code == "zho") return "zh-hans";
  if(code == "eng") return "en";
  if(code == "jpn") return "ja";
  if(code == "kos") return "ko";
  if(code == "pus") return "ps";
  if(code == "fas") return "fa";
  if(code == "rus") return "ru";
  if(code == "urd") return "ur";
  if(code == "spa") return "es";
  return code;
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
  //var jsonText = JSON.stringify({
  //  "content": text.toString(), "linkEntities":true
  //});
  var response = rosapiRequest(key, jsonText, 'entities/linked', true);
  //var response = rosapiRequest(key, jsonText,'entites',true)
  var json = response.getContentText();
  return json.toString();
}

//changes background color for entities
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
    foundElement = body.findText(mention, end+1);
  }
}

/**
*  gets linked entities from text to parse wikidata
*
* @param {String} user_key key for Rosette
* @param {sourceLang} optional workaround if rosette is identifying languages incorrectly
*
 */
function getWikiData(user_key,sourceLang) {
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
  if(sourceLang == "auto"){
     var payloadJSON = {"content": null};}
     //var payloadJSON = {"content": null,"linkedEntities":true}
    else{
      var payloadJSON = {"content":null,"language":sourceLang};

     //var payloadJSON = {"content": null,""language":sourceLang, linkedEntities":true}
      }
  payloadJSON.content = text.toString();
  var payload  = JSON.stringify(payloadJSON);
      
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : payload,
  };
    try{
  var response  = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/entities/linked/', options);
  //var response = UrlFetchApp.fetch('https://api.rosette.com/rest/v1/entities',options)
  var json = JSON.parse(response);
      if(json["code"]){
         response = tooManyRequests('https://api.rosette.com/rest/v1/entities/linked',options);
         //response = tooManyRequests('https://api.rosette.com/rest/v1/entities',options)}
    }
    }
  catch(err){
    throw "Unsupported Language. If you think this is a mistake, specify your language above";}
  return parseEntities(response,sourceLang);
  }
}

//extracts english and foreign (if applicable) wiki info using regex
function parseEntities(response,sourceLang) {
  if(JSON.parse(response)["entities"].length === 0){
    throw new Error( "No entities found");
  }
  
 var entitiesArray = JSON.parse(response).entities;
 var wikiLink = "https://en.wikipedia.org/wiki/";

 for(var i = 0; i< entitiesArray.length; i++){
   var linkJSON = UrlFetchApp.fetch("https://www.wikidata.org/w/api.php?action=wbgetentities&ids=" + entitiesArray[i].entityId + "&languages=en&props=sitelinks&sitefilter=enwiki&format=json");
   var str = entitiesArray[i].entityId.toString();
   var title = JSON.parse(linkJSON).entities[str].sitelinks.enwiki.title;
   title = title.replace(" ", "_");
   var finalLink = wikiLink.concat(title);
   var wikiLinksArray = getWikiContent(finalLink,"eng");
 }
  var wikiLinksArray2 = ['','',''];
  if(sourceLang != "eng" && sourceLang != "auto"){
    var langCode = changeLangCode(sourceLang);
    wikiLink = "https://"+langCode+".wikipedia.org/wiki/";

    for(var i = 0; i< entitiesArray.length; i++){
      linkJSON = UrlFetchApp.fetch("https://www.wikidata.org/w/api.php?action=wbgetentities&ids=" + entitiesArray[i].entityId + "&languages=" + langCode +"&props=sitelinks&sitefilter=" + langCode + "wiki&format=json");
      str = entitiesArray[i].entityId.toString();
      var json = JSON.parse(linkJSON);
      var wikiCode = Object.keys(json.entities[str].sitelinks)[0];
      title = json.entities[str].sitelinks[wikiCode].title;
      title = title.replace(" ", "_");
      finalLink = wikiLink.concat(title);
      wikiLinksArray2 = getWikiContent(finalLink,sourceLang);
    }
  }
  var finalWiki = wikiLinksArray.concat(wikiLinksArray2);
 return finalWiki;
}

function getWikiContent(link,sourceLang) {
  var response = UrlFetchApp.fetch(link);
  var content = response.getContentText();
  var ui = DocumentApp.getUi();
  var title = /<h1 id="firstHeading"(?:[\s\S]*?)<\/h1>/.exec(content);
  if(sourceLang == "spa"){
    var image = /<a href="\/wiki\/Archivo:(?:[\s\S]*?)<\/a>/.exec(content);
    var paragraphReg = /<p>(?:[\s\S]*?)<\/p>/g;
    var paragraph = paragraphReg.exec(content);
    while(paragraph[0].indexOf('<a href="/wiki/Archivo:') < 4 && paragraph.lastIndex < content.length){
       paragraph = paragraphReg.exec(content);}
  }
  else if(sourceLang == "ara"){
    var image = /<a href="\/wiki\/%D9%85%D9%84%D9%81:(?:[\s\S]*?)<\/a>/.exec(content);
      var paragraphReg = /<p>(?:[\s\S]*?)<\/p>/g;
    var paragraph = paragraphReg.exec(content);
    while(paragraph[0].indexOf('<a href="/wiki/%D9%85%D9%84%D9%81:') < 4 && paragraph.lastIndex < content.length){
       paragraph = paragraphReg.exec(content);}}
  else {var image = /<a href="\/wiki\/File:(?:[\s\S]*?)<\/a>/.exec(content);
        var paragraph = /<p>([\s\S]*?)<\/p>/.exec(content);}
  var finalContent = [title, image, paragraph];
  
  return finalContent;
}
/*
* @param {String} sourceLang 
* @param {String} targetLang
*
* @return {boolean} useApi
*/
function useAPI(sourceLang,targetLang){
   if(sourceLang == "eng"){
     var acceptedLanguages = ['ara','zho','eng','kor','pus','fas', 'rus'];    
     var pos = acceptedLanguages.indexOf(targetLang);
     if(pos > -1) return true;
   }
   else{
     var APIacceptedSources = ['ara','zho','jpn','kor','pus','fas','rus','urd'];
     var pos = APIacceptedSources.indexOf(sourceLang);
     if(pos > -1 && targetLang == "eng") return true;
   }
  return false;
}

//attempts to translate name using wikidata given entityID
function translateNameWiki(entityId, targetLanguage){
  var lang = changeLangCode(targetLanguage);
  var wikiData = UrlFetchApp.fetch("https://www.wikidata.org/w/api.php?action=wbgetentities&ids=" + entityId + "&props=labels&languages=" + lang +"&format=json");
  var json = JSON.parse(wikiData);
  var name = json.entities[entityId].labels[lang];
  if(name){
    return name.value;}
  return null;
}

//translates name using /name-translation
function translateNameRosette(mention, type, sourceLanguage, targetLanguage, key){
  var parameters = JSON.stringify({
        "name": mention,
        "sourceLanguageOfUse": sourceLanguage,
        "entityType": type,
        "targetLanguage": targetLanguage
      });
  var response = rosapiRequest(key, parameters, 'name-translation', false)
  var json = JSON.parse(response.getContentText());
  if(json["translation"]){
    return json["translation"]
  }
  else{
    return "no_translation"; }
}

//if sourceLanguage set as auto, calls /language
function getLanguage(key){
  var text = DocumentApp.getActiveDocument().getBody().editAsText().getText();  
  var jsonText = JSON.stringify({
    "content": text.toString()
  });
  //var jsonText = JSON.stringify({
  //"content": text.toString(),
  //"linkEntities": true
  //});
  var response = rosapiRequest(key, jsonText, 'language', false);
  var responseJSON = JSON.parse(response);
  return responseJSON["languageDetections"][0]["language"];
}
/*
* tries to link entities for wiki translation
*
*/
function linkEntity(entity, key, sourceLang){
  var jsonText = JSON.stringify({
    "content": entity,
    "language": sourceLang
  });
  var response = rosapiRequest(key, jsonText, 'entities/linked', false);
  //var response = rosapiRequest(key,jsonText,'entities',false);
    var responseJSON = JSON.parse(response); 
  try{  
    return responseJSON["entities"][0]["entityId"];
  }
  catch(err){
    return null;
  }
}
function tooManyRequests(url, options){
  var response = UrlFetchApp.fetch(url,options);
  var json = JSON.parse(response);
  while(json["code"]=="tooManyRequests"){
    Utilities.sleep(10);
    response = UrlFetchApp.fetch(url, options);
    json = JSON.parse(response);
  }
  return response;
}

function addDirectionWrappers(text, languageToWrap){
  var LRlang = ["eng","spa","zho","jpn","ita","nld","fra","deu", "ind","kor","por","rus"];
  var RLlang = ["ara","urd","heb","pus","fas"];
  if(LRlang.indexOf(languageToWrap) != -1){
    Logger.log(languageToWrap);
    text = "\u202a" + text + "\u202c";
  }
  else{
    text = "\u202b" + text + "\u202c";
  }
  return text;
}
