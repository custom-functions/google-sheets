// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia translations (language links) for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") to get translations for.
 * @param {Array<string>=} opt_targetLanguages The list of languages to limit the results to (optional).
 * @param {boolean=} opt_skipHeader Whether to skip the header, defaults to false (optional).
 * @return {Array<string>} The list of translations.
 * @customfunction
 */
function WIKITRANSLATE(article, opt_targetLanguages, opt_skipHeader,
  _opt_returnAsObject) {
  'use strict';
  if (!article) {
    return '';
  }
  var results = {};
  opt_targetLanguages = opt_targetLanguages || [];
  opt_targetLanguages = Array.isArray(opt_targetLanguages) ?
    opt_targetLanguages : [opt_targetLanguages];
  var temp = {};
  opt_targetLanguages.forEach(function (lang) {
    temp[lang] = true;
  });
  opt_targetLanguages = Object.keys(temp);
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
    opt_targetLanguages.forEach(function (targetLanguage) {
      if (targetLanguage) {
        results[targetLanguage] = title.replace(/_/g, ' ');
      }
    });
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&prop=langlinks' +
      '&format=xml' +
      '&lllimit=max' +
      '&titles=' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query').getChild('pages')
      .getChild('page').getChild('langlinks').getChildren('ll');
    var targetLanguagesSet = opt_targetLanguages.length > 0;
    for (var i = 0; i < entries.length; i++) {
      var text = entries[i].getText();
      var lang = entries[i].getAttribute('lang').getValue();
      if ((targetLanguagesSet) && (opt_targetLanguages.indexOf(lang) === -1)) {
        continue;
      }
      results[lang] = text;
    }
    title = title.replace(/_/g, ' ');
    results[language] = title;
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  if (_opt_returnAsObject) {
    return results;
  }
  var arrayResults = [];
  for (var lang in results) {
    if (opt_skipHeader) {
      arrayResults.push(results[lang]);
    } else {
      arrayResults.push([lang, results[lang]]);
    }
  }
  return arrayResults.length > 0 ? arrayResults : '';
}
