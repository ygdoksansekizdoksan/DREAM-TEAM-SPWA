
function oneDrive_login() {
    console.log('onedrive login');

    var appInfo = {
        "clientId": 'dabc0641-14b9-4c5f-8956-73693bbc3821',
        "redirectUri": "http://localhost:8080/callback.html",
        "scopes": "sites.read.all",
        "authServiceUri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    }

    provideAppInfo(appInfo);
    challengeForAuth();
}

function oneDrive_download(folder_path) {
    if(!is_authenticated()){
        alert("login into onedrive")
        return;
    }
    


}


window.onAuthenticated = function (token, authWindow) {
    if (token) {
        authWindow.close();
        this.console.log("token : ", token);
        store_session(token);
    }
}

function is_authenticated() {
    var expiresAt = JSON.parse(localStorage.getItem('oneDriveExpiresAt'));
    return new Date().getTime() < expiresAt;
}

function store_session(token) {
    var expiresAt = JSON.stringify(3600 * 1000 + new Date().getTime());
    localStorage.setItem("oneDriveExpiresAt", expiresAt);
    localStorage.setItem("oneDriveToken", token);
}