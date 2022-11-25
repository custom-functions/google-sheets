// @author Thomas Steiner https://github.com/tomayac/wikipedia-tools-for-google-spreadsheets
/**
 * Returns Wikidata facts for a Wikipedia article.
 *
 * @param {string} article The Wikipedia article in the format "language:Article_Title" ("de:Berlin") or the Wikidata entity in the format "qid" ("Q42") to get Wikidata facts for.
 * @param {string=} opt_multiObjectMode Whether to return all object values (pass "all") or just the first (pass "first") when there are more than one object values (optional).
 * @param {Array<string>} opt_properties Limit the resulting facts to a list of properties (optional).
 * @return {Array<string>} The list of Wikidata facts.
 * @customfunction
 */
function WIKIDATAFACTS(article, opt_multiObjectMode, opt_properties) {
  'use strict';

  var simplifyClaims = function (claims) {
    var simpleClaims = {};
    for (var id in claims) {
      var claim = claims[id];
      simpleClaims[id] = simpifyClaim(claim);
    }
    return simpleClaims;
  };

  var simpifyClaim = function (claim) {
    var simplifiedClaim = [];
    var len = claim.length;
    for (var i = 0; i < len; i++) {
      var statement = claim[i];
      var simpifiedStatement = simpifyStatement(statement);
      if (simpifiedStatement !== null) {
        simplifiedClaim.push(simpifiedStatement);
      }
    }
    return simplifiedClaim;
  };

  var simpifyStatement = function (statement) {
    var mainsnak = statement.mainsnak;
    if (mainsnak === null) {
      return null;
    }
    var datatype = mainsnak.datatype;
    var datavalue = mainsnak.datavalue;
    if (datavalue === null || datavalue === undefined) {
      return null;
    }
    switch (datatype) {
      case 'string':
      case 'commonsMedia':
      case 'url':
      case 'math':
      case 'external-id':
        return datavalue.value;
      case 'monolingualtext':
        return datavalue.value.text;
      case 'wikibase-item':
        var qid = 'Q' + datavalue.value['numeric-id'];
        qids.push(qid);
        return qid;
      case 'time':
        return datavalue.value.time;
      case 'quantity':
        return datavalue.value.amount;
      default:
        return null;
    }
  };

  var getPropertyAndEntityLabels = function (propertiesAndEntities) {
    var labels = {};
    try {
      var size = 50;
      var j = propertiesAndEntities.length;
      for (var i = 0; i < j; i += size) {
        var chunk = propertiesAndEntities.slice(i, i + size);
        var url = 'https://www.wikidata.org/w/api.php' +
          '?action=wbgetentities' +
          '&languages=en' +
          '&format=json' +
          '&props=labels' +
          '&ids=' + chunk.join('%7C');
        var json = JSON.parse(UrlFetchApp.fetch(url, {
          headers: {
            'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
          }
        }).getContentText());
        var entities = json.entities;
        chunk.forEach(function (item) {
          if ((entities[item]) &&
            (entities[item].labels) &&
            (entities[item].labels.en) &&
            (entities[item].labels.en.value)) {
            labels[item] = entities[item].labels.en.value;
          } else {
            labels[item] = false;
          }
        });
      }
    } catch (e) {
      console.log(JSON.stringify(e));
    }
    return labels;
  };

  if (!article) {
    return '';
  }
  opt_properties = opt_properties || [];
  opt_properties = Array.isArray(opt_properties) ?
    opt_properties : [opt_properties];
  var temp = {};
  opt_properties.forEach(function (prop) {
    temp[prop] = true;
  });
  opt_properties = Object.keys(temp);
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
    var url;
    if (/^Q\d+$/.test(title)) {
      url = 'https://www.wikidata.org/w/api.php' +
        '?action=wbgetentities' +
        '&format=json' +
        '&props=claims' +
        '&ids=' + title;
    } else {
      url = 'https://wikidata.org/w/api.php' +
        '?action=wbgetentities' +
        '&sites=' + language + 'wiki' +
        '&format=json' +
        '&props=claims' +
        '&titles=' + encodeURIComponent(title.replace(/\s/g, '_'));
    }
    var json = JSON.parse(UrlFetchApp.fetch(url, {
      headers: {
        'X-User-Agent': 'Wikipedia Tools for Google Spreadsheets'
      }
    }).getContentText());
    var entity = Object.keys(json.entities)[0];
    var qids = [];
    var simplifiedClaims = simplifyClaims(json.entities[entity].claims);
    var properties = Object.keys(simplifiedClaims);
    if (opt_properties.length) {
      properties = properties.filter(function (property) {
        return opt_properties.indexOf(property) !== -1;
      });
    }
    var labels = getPropertyAndEntityLabels(properties.concat(qids));
    for (var claim in simplifiedClaims) {
      var claims = simplifiedClaims[claim].filter(function (value) {
        return value !== null;
      });
      // Only return single-object facts
      if (claims.length === 1) {
        var label = labels[claim];
        var value = /^Q\d+$/.test(claims[0]) ? labels[claims[0]] : claims[0];
        if (label && value) {
          results.push([label, value]);
        }
      }
      // Optionally return multi-object facts
      if ((
        (/^first$/i.test(opt_multiObjectMode)) ||
        (/^all$/i.test(opt_multiObjectMode))
      ) && (claims.length > 1)) {
        var label = labels[claim];
        claims.forEach(function (claimObject, i) {
          if (i > 0 && /^first$/i.test(opt_multiObjectMode)) {
            return;
          }
          var value = /^Q\d+$/.test(claimObject) ?
            labels[claimObject] : claimObject;
          if (label && value) {
            results.push([label, value]);
          }
        });
      }
    }
  } catch (e) {
    console.log(JSON.stringify(e));
  }
  return results.length > 0 ? results : '';
}
