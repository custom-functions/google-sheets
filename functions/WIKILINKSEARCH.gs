// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikipedia articles that have a link that matches a given link pattern.
 *
 * @param {string} linkPattern The link pattern to search for in the format "language:example.com" or "language:*.example.com".
 * @param {string=} opt_protocol Protocol of the link, defaults to "http" (optional).
 * @param {string=} opt_namespaces Only include pages in these namespaces (optional).
 * @return {Array<string>} The list of articles that match the link pattern and the concrete link.
 * @customfunction
 */
function WIKILINKSEARCH(linkPattern, opt_protocol, opt_namespaces) {
  'use strict';
  if (!linkPattern) {
    return '';
  }
  var results = [];
  try {
    var language;
    var title;
    if (linkPattern.indexOf(':') !== -1) {
      language = linkPattern.split(/:(.+)?/)[0];
      title = linkPattern.split(/:(.+)?/)[1];
    } else {
      language = 'en';
      title = linkPattern;
    }
    if (!title) {
      return '';
    }
    var url = 'https://' + language + '.wikipedia.org/w/api.php' +
      '?action=query' +
      '&format=xml' +
      '&list=exturlusage' +
      '&eulimit=max' +
      '&euprop=title%7Curl' +
      '&euprotocol=' + (opt_protocol ? opt_protocol : 'http') +
      '&euquery=' + encodeURIComponent(title) +
      '&eunamespace=' + (opt_namespaces ?
        encodeURIComponent(opt_namespaces) : '0');
    var xml = UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText();
    var document = XmlService.parse(xml);
    var entries = document.getRootElement().getChild('query')
      .getChild('exturlusage').getChildren('eu');
    for (var i = 0; i < entries.length; i++) {
      var title = entries[i].getAttribute('title').getValue();
      var url = entries[i].getAttribute('url').getValue();
      results[i] = [title, url];
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
