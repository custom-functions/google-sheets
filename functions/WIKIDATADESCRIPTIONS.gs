// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the descriptions for a Wikidata item.
 *
 * @param {string} qid The Wikidata item's qid to get the label for.
 * @param {Array<string>=} opt_targetLanguages The list of languages to limit the results to, or "all" (optional).
 * @return {Array<string>} The label.
 * @customfunction
 */
function WIKIDATADESCRIPTIONS(qid, opt_targetLanguages) {
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
      '&props=descriptions' +
      '&ids=' + qid +
      (opt_targetLanguages.length ?
        '&languages=' + opt_targetLanguages.join('%7C') : '');
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    var descriptions = json.entities[qid].descriptions;
    var availableLanguages = Object.keys(descriptions).sort();
    availableLanguages.forEach(function (language) {
      var description = descriptions[language].value;
      results.push([language, description]);
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
