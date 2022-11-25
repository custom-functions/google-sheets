// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the Wikimedia Commons link for a file.
 *
 * @param {string} fileName The Wikimedia Commons file name in the format "language:File_Name" ("en:Flag of Berlin.svg") to get the link for.
 * @return {string} The link of the Wikimedia Commons file.
 * @customfunction
 */
function WIKICOMMONSLINK(fileName) {
  'use strict';
  if (!fileName) {
    return '';
  }
  var results = [];
  try {
    var language;
    var title;
    if (fileName.indexOf(':') !== -1) {
      language = fileName.split(/:(.+)?/)[0];
      title = fileName.split(/:(.+)?/)[1];
    } else {
      language = 'en';
      title = fileName;
    }
    if (!title) {
      return '';
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&prop=imageinfo' +
      '&iiprop=url' +
      '&format=xml' +
      '&titles=File:' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entry = document.getRootElement().getChild('query').getChild('pages')
      .getChild('page').getChild('imageinfo').getChild('ii');
    var fileUrl = entry.getAttribute('url').getValue();
    results[0] = fileUrl;
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}