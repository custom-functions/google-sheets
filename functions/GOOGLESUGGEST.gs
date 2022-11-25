// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Google Suggest results for the given keyword.
 *
 * @param {string} query The query in the format "language:Query" ("de:Berlin") to get suggestions for.
 * @return {Array<string>} The list of suggestions.
 * @customfunction
 */
function GOOGLESUGGEST(query) {
  'use strict';
  if (!query) {
    return '';
  }
  var results = [];
  try {
    var language;
    var title;
    if (query.indexOf(':') !== -1) {
      language = query.split(/:(.+)?/)[0];
      title = query.split(/:(.+)?/)[1];
    } else {
      language = 'en';
      title = query;
    }
    if (!title) {
      return '';
    }
    var url = 'https://suggestqueries.google.com/complete/search' +
      '?output=toolbar' +
      '&hl=' + language +
      '&q=' + encodeURIComponent(title);
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChildren('CompleteSuggestion');
    for (var i = 0; i < entries.length; i++) {
      var text = entries[i].getChild('suggestion').getAttribute('data')
        .getValue();
      results[i] = text;
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
