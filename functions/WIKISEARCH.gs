// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia article results for a query.
 *
 * @param {string} query The query in the format "language:Query" ("de:Berlin") to get search results for.
 * @param {boolean=} opt_didYouMean Whether to return a "did you mean" suggestion, defaults to false (optional).
 * @param {string=} opt_namespaces Only include pages in these namespaces (optional).
 * @return {Array<string>} The list of article results.
 * @customfunction
 */
function WIKISEARCH(query, opt_didYouMean, opt_namespaces) {
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
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&format=json' +
      '&list=search' +
      '&srinfo=suggestion' +
      '&srprop=' + // Empty on purpose
      '&srlimit=max' +
      '&srsearch=' + encodeURIComponent(title) +
      '&srnamespace=' + (opt_namespaces ?
        encodeURIComponent(opt_namespaces) : '0');
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    json.query.search.forEach(function (result, i) {
      result = result.title;
      if (opt_didYouMean) {
        if (i === 0) {
          results[i] = [
            result,
            json.query.searchinfo ? json.query.searchinfo.suggestion : title
          ];
        } else {
          results[i] = [result, ''];
        }
      } else {
        results[i] = result;
      }
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
