// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia articles around a Wikipedia article or around a point.
 *
 * @param {string} articleOrPoint The Wikipedia article in the format "language:Article_Title" ("de:Berlin") or the point in the format "language:Latitude,Longitude" ("en:37.786971,-122.399677") to get articles around for.
 * @param {number} radius The search radius in meters.
 * @param {boolean=} opt_includeDistance Whether to include the distance in the output, defaults to false (optional).
 * @param {string=} opt_namespaces Only include pages in these namespaces (optional).
 * @return {Array<string>} The list of articles around the given article or point.
 * @customfunction
 */
function WIKIARTICLESAROUND(articleOrPoint, radius, opt_includeDistance,
  opt_namespaces) {
  'use strict';
  if (!articleOrPoint) {
    return '';
  }
  var results = [];
  try {
    var language;
    var rest;
    var title;
    var latitude;
    var longitude;
    if (articleOrPoint.indexOf(':') !== -1) {
      language = articleOrPoint.split(/:(.+)?/)[0];
      rest = articleOrPoint.split(/:(.+)?/)[1];
      title = false;
      latitude = false;
      longitude = false;
      if (/^[-+]?([1-8]?\d(\.\d+)?|90(\.0+)?),\s*[-+]?(180(\.0+)?|((1[0-7]\d)|([1-9]?\d))(\.\d+)?)$/.test(rest)) {
        latitude = rest.split(',')[0];
        longitude = rest.split(',')[1];
      } else {
        title = rest;
      }
    } else {
      language = 'en';
      rest = articleOrPoint;
      title = false;
      latitude = false;
      longitude = false;
      if (/^[-+]?([1-8]?\d(\.\d+)?|90(\.0+)?),\s*[-+]?(180(\.0+)?|((1[0-7]\d)|([1-9]?\d))(\.\d+)?)$/.test(rest)) {
        latitude = rest.split(',')[0];
        longitude = rest.split(',')[1];
      } else {
        title = rest;
      }
    }
    if ((!title) && !(latitude && longitude)) {
      return;
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php';
    if (title) {
      url += '?action=query' +
        '&list=geosearch' +
        '&format=xml' +
        '&gslimit=max' +
        '&gsradius=' + radius +
        '&gspage=' + encodeURIComponent(title.replace(/\s/g, '_'));
    } else {
      url += '?action=query' +
        '&list=geosearch' +
        '&format=xml&gslimit=max' +
        '&gsradius=' + radius +
        '&gscoord=' + latitude + '%7C' + longitude;
    }
    url += '&gsnamespace=' + (opt_namespaces ?
      encodeURIComponent(opt_namespaces) : '0');
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query')
      .getChild('geosearch').getChildren('gs');
    for (var i = 0; i < entries.length; i++) {
      var title = entries[i].getAttribute('title').getValue();
      var lat = entries[i].getAttribute('lat').getValue();
      var lon = entries[i].getAttribute('lon').getValue();
      if (opt_includeDistance) {
        var dist = entries[i].getAttribute('dist').getValue();
        results[i] = [title, lat, lon, dist];
      } else {
        results[i] = [title, lat, lon];
      }
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
