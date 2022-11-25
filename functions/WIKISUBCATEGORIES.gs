// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia subcategories for a Wikipedia category.
 *
 * @param {string} category The Wikipedia category in the format "language:Category_Title" ("en:Category:Visitor_attractions_in_Berlin") to get subcategories for.
 * @param {string=} opt_namespaces Only include pages in these namespaces (optional).
 * @return {Array<string>} The list of subcategories.
 * @customfunction
 */
function WIKISUBCATEGORIES(category, opt_namespaces) {
  'use strict';
  if (!category) {
    return '';
  }
  var results = [];
  try {
    var language;
    var title;
    if ((category.match(/:/g) || []).length > 1) {
      language = category.split(/:(.+)?/)[0];
      title = category.split(/:(.+)?/)[1];
    } else {
      language = 'en';
      title = category;
    }
    if (!title) {
      return '';
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&list=categorymembers' +
      '&cmlimit=max' +
      '&cmprop=title' +
      '&cmtype=subcat%7Cpage' +
      '&format=xml' +
      '&cmnamespace=' + (opt_namespaces ?
        encodeURIComponent(opt_namespaces) : '14') +
      '&cmtitle=' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query')
      .getChild('categorymembers').getChildren('cm');
    for (var i = 0; i < entries.length; i++) {
      var text = entries[i].getAttribute('title').getValue();
      results[i] = text;
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
