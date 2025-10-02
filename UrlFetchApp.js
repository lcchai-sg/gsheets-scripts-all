// Documentation: https://developers.google.com/apps-script/reference/url-fetch/http-response

const list = [];

function fetch() {
  var response = UrlFetchApp.fetch('http://www.jyac.com/about-us/#our-team-section');
  raw_html = response.getContentText();
  var results = raw_html.match(/<div class="slide-desc"><p>.+/gi);
  for(var i=0; i<results.length; i++){
   list.push([results[i]]);
  }
 return list;
}

function append() {
  var array = fetch();
  var targetRange = SpreadsheetApp.getActiveSheet().getRange(4,1,array.length,array[0].length);
  console.log(array.length)
  console.log(targetRange.getValues().length)
  targetRange.setValues(array);
}
