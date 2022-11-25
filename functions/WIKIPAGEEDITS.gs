// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia pageedits statistics for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") to get pageedits statistics for.
 * @param {string=} opt_start The start date in the format "YYYYMMDD" ("2007-06-08") since when pageedits statistics should be retrieved from (optional).
 * @param {string=} opt_end The end date in the format "YYYYMMDD" ("2007-06-08") until when pageedits statistics should be retrieved to (optional).
 * @return {Array<number>} The list of pageedits between start and end and their deltas.
 * @customfunction
 */
function WIKIPAGEEDITS(article, opt_start, opt_end) {
  'use strict';

  var getIsoDate = function (date, time) {
    var date = new Date(date);
    var year = date.getFullYear();
    var month = (date.getMonth() + 1) < 10 ?
      '0' + (date.getMonth() + 1) :
      (date.getMonth() + 1).toString();
    var day = date.getDate() < 10 ?
      '0' + date.getDate() :
      date.getDate().toString();
    return year + '-' + month + '-' + day + time;
  };

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
    opt_start = opt_start || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    if (typeof opt_start === 'object') {
      opt_start = getIsoDate(opt_start, 'T00:00:00');
    }
    opt_end = opt_end || new Date();
    if (typeof opt_end === 'object') {
      opt_end = getIsoDate(opt_end, 'T23:59:59');
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&prop=revisions' +
      '&rvprop=size%7Ctimestamp' +
      '&rvlimit=max' +
      '&format=xml' +
      '&rvstart=' + opt_end + // Reversed on purpose due to confusing API name
      '&rvend=' + opt_start + // Reversed on purpose due to confusing API name
      '&titles=' + encodeURIComponent(title.replace(/\s/g, '_'));
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query').getChild('pages')
      .getChild('page').getChild('revisions').getChildren('rev');
    for (var i = 0; i < entries.length - 1 /* - 1 for the delta */; i++) {
      var timestamp = entries[i].getAttribute('timestamp').getValue().replace(
        /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$/,
        '$1-$2-$3-$4-$5-$6').split('-');
      timestamp = new Date(Date.UTC(
        parseInt(timestamp[0], 10), // Year
        parseInt(timestamp[1], 10) - 1, // Month
        parseInt(timestamp[2], 10), // Day
        parseInt(timestamp[3], 10), // Hour
        parseInt(timestamp[4], 10), // Minute
        parseInt(timestamp[5], 10))); // Second
      var delta = entries[i].getAttribute('size').getValue() -
        entries[i + 1].getAttribute('size').getValue();
      results.push([
        timestamp,
        delta
      ]);
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
