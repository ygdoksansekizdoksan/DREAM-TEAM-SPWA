
function oneDrive_login() {
    console.log('onedrive login');

    var appInfo = {
        "clientId": 'dabc0641-14b9-4c5f-8956-73693bbc3821',
        "redirectUri": "https://127.0.0.1:8080/callback.html",
        "scopes": "sites.read.all",
        "authServiceUri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    }

    provideAppInfo(appInfo);
    challengeForAuth();
}

function oneDrive_download(file_path) {
    if (!is_authenticated()) {
        alert("login into onedrive")
        return;
    }


    donwload_folder(localStorage.getItem("oneDriveToken"), file_path);
}


/*

Using 302 re-direct does not work a because of CORS.
Look at for https://github.com/microsoftgraph/microsoft-graph-docs/issues/43 more info.

Using the @microsoft.graph.downloadUrl property drive-item only works for personal accounts

*/

function donwload_folder(token, file_path) {
    var URI = "https://graph.microsoft.com/v1.0/me/drive/root:/" + file_path;
    console.log(URI);

    var xhttp = new XMLHttpRequest();
    
    xhttp.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
            var download_uri = JSON.parse(xhttp.responseText)["@microsoft.graph.downloadUrl"];
            var download_request = new XMLHttpRequest();
            
            download_request.onreadystatechange = function(){
                if(this.readyState == 4 && this.status == 200){
                    console.log(download_request.responseText);
                }
            }
            download_request.open("GET",download_uri,true);
            download_request.send();
        }
    };
    xhttp.open("GET",URI, true);
    xhttp.setRequestHeader("Authorization","bearer " + token);
    xhttp.send();
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