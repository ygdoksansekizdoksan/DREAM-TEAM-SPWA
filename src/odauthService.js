// instructions:
// - host a copy of callback.html and odauth.js on your domain.
// - embed odauth.js in your app like this:
//   <script id="odauth" src="odauth.js"
//           ></script>
// - define the onAuthenticated(token) function in your app to receive the auth token.
// - call odauth() to begin, as well as whenever you need an auth token
//   to make an API call. if you're making an api call in response to a user's
//   click action, call odAuth(true), otherwise just call odAuth(). the difference
//   is that sometimes odauth needs to pop up a window so the user can sign in,
//   grant your app permission, etc. the pop up can only be launched in response
//   to a user click, otherwise the browser's popup blocker will block it. when
//   odauth isn't called in click mode, it'll put a sign-in button at the top of
//   your page for the user to click. when it's done, it'll remove that button.
//
// how it all works:
// when you call odauth(), we first check if we have the user's auth token stored
// in a cookie. if so, we read it and immediately call your onAuthenticated() method.
// if we can't find the auth cookie, we need to pop up a window and send the user
// to Microsoft Account so that they can sign in or grant your app the permissions
// it needs. depending on whether or not odauth() was called in response to a user
// click, it will either pop up the auth window or display a sign-in button for
// the user to click (which results in the pop-up). when the user finishes the
// auth flow, the popup window redirects back to your hosted callback.html file,
// which calls the onAuthCallback() method below. it then sets the auth cookie
// and calls your app's onAuthenticated() function, passing in the optional 'window'
// argument for the popup window. your onAuthenticated function should close the
// popup window if it's passed in.
//
// subsequent calls to odauth() will usually complete immediately without the
// popup because the cookie is still fresh.


// for added security we require https
function ensureHttps() {
  if (window.location.protocol != "https:") {
    window.location.href = "https:" + window.location.href.substring(window.location.protocol.length);
  }
}

function onAuthCallback() {
  console.log('callback');
  var authInfo = getAuthInfoFromUrl();
  var token = authInfo["access_token"];
  window.opener.onAuthenticated(token, window);
}

function getAuthInfoFromUrl() {
  if (window.location.hash) {
    var authResponse = window.location.hash.substring(1);
    var authInfo = JSON.parse(
      '{"' + authResponse.replace(/&/g, '","').replace(/=/g, '":"') + '"}',
      function (key, value) { return key === "" ? value : decodeURIComponent(value); });
    return authInfo;
  }
  else {
    alert("failed to receive auth token");
  }
}



var storedAppInfo = null;

function provideAppInfo(appInfo) {

  if (!appInfo.hasOwnProperty("clientId")) {
    alert("clientId was is not defined");
    return;
  }
  if (!appInfo.hasOwnProperty("redirectUri")) {
    alert("redirectUri was is not defined");
    return;
  }
  if (!appInfo.hasOwnProperty("scopes")) {
    alert("scopes was is not defined");
    return;
  }
  if (!appInfo.hasOwnProperty("authServiceUri")) {
    alert("authServiceUri was is not defined");
    return;
  }

  storedAppInfo = appInfo;
}

function getAppInfo() {

  if (storedAppInfo) {
    return storedAppInfo;
  }

  alert("No AppInfo was provided, make sure provedAppInfo() was called");
}



function challengeForAuth() {
  var appInfo = getAppInfo();
  var url =
    appInfo.authServiceUri +
    "?client_id=" + appInfo.clientId +
    "&scope=" + encodeURIComponent(appInfo.scopes) +
    "&response_type=token" +
    "&redirect_uri=" + encodeURIComponent(appInfo.redirectUri);
  popup(url);
}

function popup(url) {
  var width = 525,
    height = 525,
    screenX = window.screenX,
    screenY = window.screenY,
    outerWidth = window.outerWidth,
    outerHeight = window.outerHeight;

  var left = screenX + Math.max(outerWidth - width, 0) / 2;
  var top = screenY + Math.max(outerHeight - height, 0) / 2;

  var features = [
    "width=" + width,
    "height=" + height,
    "top=" + top,
    "left=" + left,
    "status=no",
    "resizable=yes",
    "toolbar=no",
    "menubar=no",
    "scrollbars=yes"];
  var popup = window.open(url, "oauth", features.join(","));
  if (!popup) {
    alert("failed to pop up auth window");
  }

  popup.focus();
}