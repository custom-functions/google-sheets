// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the Wikidata qid of the corresponding Wikidata item for a Wikipedia article.
 *
 * @param {string} article The article in the format "language:Query" ("de:Berlin") to get the Wikidata qid for.
 * @return {string} The Wikidata qid.
 * @customfunction
 */
function WIKIDATAQID(article) {
  'use strict';
  if (!article) {
    return '';
  }
  var results = [];
  try {
    var language;
    var title;
    if (article.indexOf(':') !== -1) {
      language = article.split(/:(.+)?/)[0];
      title = article.split(/:(.+)?/)[1];
    } else {
      language = 'en';
      title = article;
    }
    if (!title) {
      return '';
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&format=json' +
      '&formatversion=2' +
      '&redirects=1' +
      '&prop=pageprops' +
      '&ppprop=wikibase_item' +
      '&titles=' + encodeURIComponent(title);
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    if (json.query.pages[0] && json.query.pages[0].pageprops.wikibase_item) {
      results[0] = json.query.pages[0].pageprops.wikibase_item;
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
