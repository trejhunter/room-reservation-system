// Room Reservation System Email
// TJ Hunter, 2020

function makeEmail(request) {
  return (
    '<!DOCTYPE html><html><head><base target="_top"></head><body><div style="text-align: center;' +
    'font-family: Arial;"><div id="center" style="width:300px;border: 2px dotted grey;background:' +
    '#ececec; margin:25px;margin-left:auto; margin-right:auto;padding:15px;"><img src="https://upload.' +
    'wikimedia.org/wikipedia/commons/thumb/6/69/Calendar_font_awesome.svg/512px-Calendar_font_awesome' +
    '.svg.png"width="180" style="margin:10px 0px"><br /><div style=" border: 2px dotted grey;' +
    'background:white;margin-right:auto; margin-left:auto; padding:10px;"><h2>' +
    request.header +
    '</h2><h3>' +
    request.message +
    '<br /><br/>' +
    request.firstname +
    '&nbsp;' +
    request.lastname +
    '<br /><br />' +
    '<u>' +
    request.roomsRequested +
    '</u>' +
    '<br />' +
    request.room +
    '<br /><br />' +
    request.dateString +
    '<br />' +
    request.timeString.slice(0, request.timeString.length - 4) +
    '<br />' +
    request.endTimeString.slice(0, request.endTimeString.length - 4) +
    '<br /><br />' +
    request.reason +
    '<br /></h3><br />' +
    '<a href="' +
    request.buttonLink +
    '" class="btn" style="-webkit-border-radius: 28;' +
    '-moz-border-radius: 5;border-radius: 5px;font-family: Arial; color: #ffffff;font-size: 15px;' +
    'background: #ff7878;padding:8px 20px 8px 20px;text-decoration: none;">' +
    request.buttonText +
    '</a><br /><br /></div></div><div><p style="font-size:12px">' +
    '<a href="YOUR WEBSITE LINK/"> Andrews University Recreaction Center</a> </p></div></body></html>'
  );
}
