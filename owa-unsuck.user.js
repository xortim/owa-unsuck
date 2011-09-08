// ==UserScript==
// @name          OWA-unsuck
// @author        Tim Earle
// @namespace     org.endomusia.owa
// @description   Adds more functionality to outlook web access light
// @include       https://exchange.*.*/owa/*
// @include       https://webmail.*.*/owa/*
// @exclude       https://exchange.*.*/owa/auth/logoff.aspx?Cmd=logoff
// @exclude       https://exchange.*.*/exchweb/bin/auth/owalogon.asp*
// ==/UserScript==

window.doWork = function (junkBtn,navbar) {
  if(navbar) {
    var divider = document.createElement("td");
    divider.setAttribute("class", 'dv');
    divider.setAttribute("id","owa_unsuck_divider_1");
    divider.innerHTML = '<img alt="" src="8.1.340.0/themes/base/tbdv.gif"/></td>';
    navbar.insertBefore(divider, junkBtn.nextSibling);

    select_unread_btn = document.createElement("td");
    select_unread_btn.setAttribute("nowrap","");
    select_unread_btn.setAttribute("id","owa_unsuck_select_unread");
    select_unread_btn.innerHTML = '<a id="lnkHdrunread" title="Select Unread" class="btn" href="#"><img alt="" src="8.1.340.0/themes/base/msg-unrd.gif"/>Select Unread</a>';
    select_unread_btn.addEventListener("click",selectUnread, true);
    navbar.insertBefore(select_unread_btn,divider.nextSibling);

    select_read_btn = document.createElement("td");
    select_read_btn.setAttribute("nowrap","");
    select_read_btn.setAttribute("id","owa_unsuck_select_read");
    select_read_btn.innerHTML = '<a id="lnkHdrread" title="Select Read" class="btn" href="#"><img alt="" src="8.1.340.0/themes/base/msg-rd.gif"/>Select Read</a>';
    select_read_btn.addEventListener("click",selectRead, true);
    navbar.insertBefore(select_read_btn,divider.nextSibling);

    select_none_btn = document.createElement("td");
    select_none_btn.setAttribute("nowrap","");
    select_none_btn.setAttribute("id","owa_unsuck_select_none");
    select_none_btn.innerHTML = '<a id="lnkHdrread" title="Select None" class="btn" href="#">Select None</a>';
    select_none_btn.addEventListener("click",selectNone, true);
    navbar.insertBefore(select_none_btn,divider.nextSibling);
  }
}

window.selectNone = function (check) {
  if(check == null) check = false;
  var allIcons, thisIcon;
  allIcons = document.evaluate(
    "//img[@class='sI']", //just the read/unread icon
    document,
    null,
    XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE,
    null
  );
  for (var i = 0; i < allIcons.snapshotLength; i++) {
    thisIcon = allIcons.snapshotItem(i);
    var thisRow = thisIcon.parentNode.parentNode;
    var checkBox = thisRow.childNodes[3].childNodes[0];
    if(checkBox.checked == true) {
      checkBox.checked = false;
    }
  }
  return true;
}

window.selectUnread = function (check) {
  if(check == null) check = true;
  var allRows, cur;
  allRows = getAllRowsSnapshot();
  for (var i = 0; i < allRows.snapshotLength; i++) {
    cur = allRows.snapshotItem(i);
    var thisRow = cur.parentNode.parentNode;
    var checkBox = thisRow.childNodes[3].childNodes[0];
    if (thisRow.style.fontWeight == "bold") {
      if(checkBox.checked == false) {
        checkBox.checked = true;
      }
    } else {
        if(checkBox.checked == true) {
          checkBox.checked = false;
        }
    }
  }
  return true;
}

window.selectRead = function (check) {
  if(check == null) check = true;
  var allRows, cur;
  allRows = getAllRowsSnapshot();
  for (var i = 0; i < allRows.snapshotLength; i++) {
    cur = allRows.snapshotItem(i);
    var thisRow = cur.parentNode.parentNode;
    var checkBox = thisRow.childNodes[3].childNodes[0];
    if (thisRow.style.fontWeight != "bold") {
      if(checkBox.checked == false) {
        checkBox.checked = true;
      }
    } else {
      if(checkBox.checked == true) {
        checkBox.checked = false;
      }
    }
  }
  return true;
}

window.getAllRowsSnapshot = function() {
  var allIcons = document.evaluate(
    "//img[@class='sI']", //just the read/unread icon
    document,
    null,
    XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE,
    null
  );
  return allIcons;
}

window.timeoutRefresh = function() {
  setTimeout("window.location.reload()",300000);
  return;
}

// Add jQuery later on
/*
var GM_JQ = document.createElement('script');
GM_JQ.src = 'http://jquery.com/src/jquery-latest.js';
GM_JQ.type = 'text/javascript';
document.getElementsByTagName('head')[0].appendChild(GM_JQ);

// Check if jQuery's loaded
function GM_wait() {
	if(typeof unsafeWindow.jQuery == 'undefined') { window.setTimeout(GM_wait,100); }
	else { $ = unsafeWindow.jQuery; main(); }
}
GM_wait();
*/
main();

function main() {
  if(!document.getElementById('msgHd')) {
    var junkBtn, navbar;
    var messageList = document.evaluate(
      "//table[@class='lvw']", //the list view of mail items
      document,
      null,
      XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE,
      null
    );

    if(messageList.snapshotLength > 0) {
      if(document.getElementById('lnkHdrjunk')) {
        junkBtn = document.getElementById('lnkHdrjunk').parentNode;
      } else if (document.getElementById('lnkHdrnotjunk')) {
        junkBtn = document.getElementById('lnkHdrnotjunk').parentNode;
      } else { return; }
      navbar = junkBtn.parentNode;
      doWork(junkBtn,navbar);
      timeoutRefresh();
      updateTitlebar();
      return;
    } else { return; }
  }
}
