// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Searches for Wikidata entities using Wikidata labels and aliases.
 *
 * @param {string} search The search string in the format "language:Query" ("de:Berlin") to get the Wikidata qid for.
 * @return {string} The Wikidata qid.
 * @customfunction
 */
function WIKIDATASEARCH(search) {
  'use strict';
  if (!search) {
    return '';
  }
  var results = [];
  try {
    var wbslanguage;
    var wbssearch;
    if (search.indexOf(':') !== -1) {
      wbslanguage = search.split(/:(.+)?/)[0];
      wbssearch = search.split(/:(.+)?/)[1];
    } else {
      wbslanguage = 'en';
      wbssearch = search;
    }
    if (!wbssearch) {
      return '';
    }
    var url = 'https://www.wikidata.org/w/api.php' +
      '?action=query' +
      '&list=wbsearch' +
      '&wbslanguage=' + wbslanguage +
      '&format=json' +
      '&wbssearch=' + encodeURIComponent(wbssearch);
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    results[0] = json.query.wbsearch[0].title;
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
