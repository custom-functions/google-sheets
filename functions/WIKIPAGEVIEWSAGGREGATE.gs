// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns aggregated pageviews statistics for a project.
 *
 * @param {string} project The Wikimedia project to get pageviews statistics for.
 * @param {string=} opt_access The access method, defaults to "all-access" (optional).
 * @param {string=} opt_agent The agent, defaults to "all-agents" (optional).
 * @param {string=} opt_granularity The granularity, defaults to "daily" (optional).
 * @param {string=} opt_start The start date in the format "YYYYMMDDHH" ("2007060800") since when pageviews statistics should be retrieved from (optional).
 * @param {string=} opt_end The end date in the format "YYYYMMDDHH" ("2007060800") until when pageviews statistics should be retrieved to (optional).
 * @return {Array<number>} The list of aggregated pageviews between start and end.
 * @customfunction
 */
function WIKIPAGEVIEWSAGGREGATE(project, opt_access, opt_agent, opt_granularity,
  opt_start, opt_end) {
  'use strict';

  var getIsoDateWithHour = function (date) {
    var date = new Date(date);
    var year = date.getFullYear().toString();
    var month = (date.getMonth() + 1) < 10 ?
      '0' + (date.getMonth() + 1) :
      (date.getMonth() + 1).toString();
    var day = date.getDate() < 10 ?
      '0' + date.getDate() :
      date.getDate().toString();
    return year + month + day + '00';
  };

  if (!project) {
    return '';
  }
  var results = [];
  try {
    opt_start = opt_start || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    if (typeof opt_start === 'object') {
      opt_start = getIsoDateWithHour(opt_start);
    }
    opt_end = opt_end || new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    if (typeof opt_end === 'object') {
      opt_end = getIsoDateWithHour(opt_end);
    }
    var url = 'https://wikimedia.org/api/rest_v1/metrics/pageviews/' +
      'aggregate' +
      '/' + project +
      '/' + (opt_access ? opt_access : 'all-access') +
      '/' + (opt_agent ? opt_agent : 'all-agents') +
      '/' + (opt_granularity ? opt_granularity : 'daily') +
      '/' + opt_start +
      '/' + opt_end;
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    json.items.forEach(function (item) {
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
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  results.reverse(); // Order from new to old
  return results.length > 0 ? results : '';
}
