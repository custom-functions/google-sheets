// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns the number of unique devices for a project.
 *
 * @param {string} project The Wikimedia project to get pageviews statistics for.
 * @param {string=} opt_accessSite The accesses sites, defaults to "all-sites" (optional).
 * @param {string=} opt_granularity The granularity, defaults to "daily" (optional).
 * @param {string=} opt_start The start date in the format "YYYYMMDD" ("20070608") since when pageviews statistics should be retrieved from (optional).
 * @param {string=} opt_end The end date in the format "YYYYMMDD" ("20070608") until when pageviews statistics should be retrieved to (optional).
 * @return {Array<number>} The list of unique devices between start and end.
 * @customfunction
 */
function WIKIUNIQUEDEVICES(project, opt_accessSite, opt_granularity, opt_start,
  opt_end) {
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
    opt_start = opt_start || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    if (typeof opt_start === 'object') {
      opt_start = getIsoDate(opt_start);
    }
    opt_end = opt_end || new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    if (typeof opt_end === 'object') {
      opt_end = getIsoDate(opt_end);
    }
    var url = 'https://wikimedia.org/api/rest_v1/metrics/unique-devices' +
      '/' + project +
      '/' + (opt_accessSite ? opt_accessSite : 'all-sites') +
      '/' + (opt_granularity ? opt_granularity : 'daily') +
      '/' + opt_start +
      '/' + opt_end;
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    json.items.forEach(function (item) {
      var timestamp = item.timestamp.replace(/^(\d{4})(\d{2})(\d{2})$/,
        '$1-$2-$3').split('-');
      timestamp = new Date(Date.UTC(
        parseInt(timestamp[0], 10), // Year
        parseInt(timestamp[1], 10) - 1, // Month
        parseInt(timestamp[2], 10), // Day
        0, // Hour
        0, // Minute
        0)); // Second))
      results.push([
        timestamp,
        item.devices
      ]);
    });
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
