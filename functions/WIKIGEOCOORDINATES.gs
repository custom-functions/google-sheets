// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia geocoordinates for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") to get geocoordinates for.
 * @return {Array<number>} The latitude and longitude.
 * @customfunction
 */
function WIKIGEOCOORDINATES(article) {
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
      '&prop=coordinates' +
      '&format=xml' +
      '&colimit=max' +
      '&coprimary=primary' +
      '&titles=' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var coordinates = document.getRootElement().getChild('query')
      .getChild('pages').getChild('page').getChild('coordinates')
      .getChild('co');
    var latitude = coordinates.getAttribute('lat').getValue();
    var longitude = coordinates.getAttribute('lon').getValue();
    results = [[latitude, longitude]];
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
