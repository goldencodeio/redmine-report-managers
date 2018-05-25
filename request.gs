function APIRequest(reqUrl, options) {
  var url = 'http://redmine.zolotoykod.ru/' + reqUrl + '.json?key=' + OPTIONS.apiKey + '&limit=100';
  if (!options) options = {};
  if (options.query)
    options.query.forEach(function(item) {
      url += '&' + item.key + '=' + item.value;
    });

  var response = UrlFetchApp.fetch(encodeURI(url), options);

  // catch server errors
  if (response.getResponseCode() >= 400)
    throw response.getContentText();

  var result = JSON.parse(response.getContentText());

  if (result[reqUrl].length === 100) {
    var newOptions = Object.assign({query: []}, options);
    var offsetOptionsIndex = newOptions.query.findIndex(function(i) {return i.key === 'offset'});
    if (offsetOptionsIndex > -1)
      newOptions.query[offsetOptionsIndex].value += 100;
    else
      newOptions.query.push({key: 'offset', value: 100});

    var newResult = APIRequest(reqUrl, newOptions);
    result[reqUrl] = result[reqUrl].concat(newResult[reqUrl]);
  }

  return result;
}

function APIRequestById(reqUrl, id, options) {
  var url = 'http://redmine.zolotoykod.ru/' + reqUrl + '/' + id + '.json?key=' + OPTIONS.apiKey;
  if (!options) options = {};
  if (options.query)
    options.query.forEach(function(item) {
      url += '&' + item.key + '=' + item.value;
    });

  var response = UrlFetchApp.fetch(encodeURI(url), options);

  // catch server errors
  if (response.getResponseCode() >= 400)
    throw response.getContentText();

  var result = JSON.parse(response.getContentText());

  return result;
}
