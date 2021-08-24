
/**
 * Line Notify
 * @param {string} token 權杖
 * @param {string} message 訊息
 */
function lineNotify(token, message) {
  var options = {
    'method'  : 'post',
    'payload' : {'message' : message},
    'headers' : {'Authorization' : 'Bearer ' + token}
    };      
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify',options);
}
