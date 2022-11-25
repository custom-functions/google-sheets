// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns most viewed pages for a project.
 *
 * @param {string} project The Wikimedia project to get pageviews statistics for.
 * @param {string=} opt_access The access method, defaults to "all-access" (optional).
 * @param {string=} opt_date The date in the format "YYYYMMDD" ("20070608") for which pageviews statistics should be retrieved (optional).
 * @return {Array<number>} The list of the most viewed pages.
 * @customfunction
 */
function WIKIPAGEVIEWSTOP(project, opt_access, opt_date) {
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

  if (!project) {
    return '';
  }
  var results = [];
  var sum = 0;
  try {
    opt_date = opt_date || new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    if (typeof opt_date === 'object') {
      opt_date = getIsoDate(opt_date);
    }
    var year = opt_date.substr(0, 4);
    var month = opt_date.substring(4, 6);
    var day = opt_date.substring(6, 8);
    var url = 'https://wikimedia.org/api/rest_v1/metrics/pageviews/' +
      'top' +
      '/' + project +
      '/' + (opt_access ? opt_access : 'all-access') +
      '/' + year +
      '/' + month +
      '/' + day;
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    json.items[0].articles.forEach(function (article) {
      results.push([
        article.article.replace(/_/g, ' '),
        article.views
      ]);
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
