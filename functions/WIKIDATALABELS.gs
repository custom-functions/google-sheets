// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the labels for a Wikidata item.
 *
 * @param {string} qid The Wikidata item's qid to get the labels for.
 * @param {Array<string>=} opt_targetLanguages The list of languages to limit the results to, or "all" (optional).
 * @return {Array<string>} The labels.
 * @customfunction
 */
function WIKIDATALABELS(qid, opt_targetLanguages) {
  'use strict';
  if (!qid) {
    return '';
  }
  var results = [];
  try {
    opt_targetLanguages = opt_targetLanguages || [];
    opt_targetLanguages = Array.isArray(opt_targetLanguages) ?
      opt_targetLanguages : [opt_targetLanguages];
    if (opt_targetLanguages.length === 0) {
      opt_targetLanguages = ['en'];
    }
    if (opt_targetLanguages.length === 1 && opt_targetLanguages[0] === 'all') {
      opt_targetLanguages = [];
    }
    var url = 'https://www.wikidata.org/w/api.php' +
      '?format=json' +
      '&action=wbgetentities' +
      '&props=labels' +
      '&ids=' + qid +
      (opt_targetLanguages.length ?
        '&languages=' + opt_targetLanguages.join('%7C') : '');
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    var labels = json.entities[qid].labels;
    var availableLanguages = Object.keys(labels).sort();
    availableLanguages.forEach(function (language) {
      var label = labels[language].value;
      results.push([language, label]);
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
