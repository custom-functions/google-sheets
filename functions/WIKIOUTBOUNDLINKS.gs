// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia outbound links for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") to get outbound links for.
 * @param {string=} opt_namespaces Only include pages in these namespaces (optional).
 * @return {Array<string>} The list of outbound links.
 * @customfunction
 */
function WIKIOUTBOUNDLINKS(article, opt_namespaces) {
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
      '&prop=links' +
      '&plnamespace=' + (opt_namespaces ?
        encodeURIComponent(opt_namespaces) : '0') +
      '&format=xml' +
      '&pllimit=max' +
      '&titles=' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query').getChild('pages')
      .getChild('page').getChild('links').getChildren('pl');
    for (var i = 0; i < entries.length; i++) {
      var text = entries[i].getAttribute('title').getValue();
      results[i] = text;
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
