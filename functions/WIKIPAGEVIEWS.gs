// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia pageviews statistics for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") to get pageviews statistics for.
 * @param {string=} opt_start The start date in the format "YYYYMMDD" ("20070608") since when pageviews statistics should be retrieved from (optional).
 * @param {string=} opt_end The end date in the format "YYYYMMDD" ("20070608") until when pageviews statistics should be retrieved to (optional).
 * @param {boolean=} opt_sumOnly Whether to only return the sum of all pageviews in the requested period (optional).
 * @return {Array<number>} The list of pageviews between start and end per day.
 * @customfunction
 */
function WIKIPAGEVIEWS(article, opt_start, opt_end, opt_sumOnly) {
  'use strict';

  var getIsoDate = function (date) {
    var date = new Date(date);
    var year = date.getFullYear().toString();
    var month = (date.getMonth() + 1) < 10 ?
      '0' + (date.getMonth() + 1) :
      (date.getMonth() + 1).toString();
    var day = date.getDate() < 10 ?
      '0' + date.getDate() :
      date.getDate().toString();
    return year + month + day;
  };

  if (!article) {
    return '';
  }
  var results = [];
  var sum = 0;
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
      opt_start = getIsoDate(opt_start);
    }
    opt_end = opt_end || new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    if (typeof opt_end === 'object') {
      opt_end = getIsoDate(opt_end);
    }
    var url = 'https://wikimedia.org/api/rest_v1/metrics/pageviews/' +
      'per-article' +
      '/' + language + '.wikipedia' +
      '/all-access' +
      '/user' +
      '/' + encodeURIComponent(title.replace(/\s/g, '_')) +
      '/daily' +
      '/' + opt_start +
      '/' + opt_end;
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    json.items.forEach(function (item) {
      if (opt_sumOnly) {
        sum += item.views;
      } else {
        var timestamp = item.timestamp.replace(/^(\d{4})(\d{2})(\d{2})(\d{2})$/,
          '$1-$2-$3-$4').split('-');
        timestamp = new Date(Date.UTC(
          parseInt(timestamp[0], 10), // Year
          parseInt(timestamp[1], 10) - 1, // Month
          parseInt(timestamp[2], 10), // Day
          parseInt(timestamp[3], 10), // Hour
          0, // Minute
          0)); // Second))
        results.push([
          timestamp,
          item.views
        ]);
      }
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  if (opt_sumOnly) {
    return [sum];
  } else {
    results.reverse(); // Order from new to old
    return results.length > 0 ? results : '';
  }
}
