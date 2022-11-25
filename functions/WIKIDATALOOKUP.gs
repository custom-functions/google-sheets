// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the Wikidata qid of the given identifier and property.
 * 
 * Internally, this function invokes a haswbstatement query against the  Wikidata API.
 *
 * @param {string} property The Wikidata property (such as "P298").
 * @param {string} identifier The identifier (value) to lookup (such as "AUT").
 * @return {string} The Wikidata qid.
 * @customfunction
 */
function WIKIDATALOOKUP(property, identifier) {
  'use strict';
  var results = [];
  try {
    var url = 'https://www.wikidata.org/w/api.php' +
      '?action=query' +
      '&format=json' +
      '&formatversion=2' +
      '&list=search' +
      '&ppprop=wikibase_item' +
      '&srsearch=haswbstatement:' + property + '=' + encodeURIComponent(identifier);
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    results[0] = json.query.search[0].title;
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
