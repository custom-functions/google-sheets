// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the output of the Quarry (https://meta.wikimedia.org/wiki/Research:Quarry) query with the specified query ID.
 *
 * @param {number} queryId The query ID of the Quarry query to run.
 * @return {Array<string>} The list of query results, the first line represents the header.
 * @customfunction
 */
function WIKIQUARRY(queryId) {
  'use strict';
  if (!queryId) {
    return '';
  }
  var results = [];
  try {
    var url = 'https://quarry.wmflabs.org/query/' + queryId +
      '/result/latest/0/json';
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    results[0] = json.headers;
    results = results.concat(json.rows);
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
