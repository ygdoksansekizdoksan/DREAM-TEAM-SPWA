
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


    download_folder(localStorage.getItem("oneDriveToken"), file_path).then(function(result){
        document.getElementById("step3").style.visibility = "visible";
        document.getElementById("file_header").innerHTML = "File Path : " + file_path;
        document.getElementById("file_contents").innerHTML = "File Contents : " + result[1];
    }).catch(function(error){
        if(error[0] == 401){
            alert("You are unauthorized, try logging in");
        }else if(error[0] == 404){
            alert("item not found, check path");
        }
        else{
            alert("You have a wierd error, check the console");
            console.log(error);
        }
    });
}


/*
    Using 302 re-direct does not work a because of CORS.
    Look at for https://github.com/microsoftgraph/microsoft-graph-docs/issues/43 more info.

    Make sure when using @microsoft.graph.downloadUrl property dont send authorization 
*/


function download_folder(token, file_path) {
    return new Promise(function (resolve, reject) {
        donwload_metadata(token, file_path).then(function (result) {
            var response = JSON.parse(result[1]);
            return response;
        }).then(function (result) {
            download_contents(result["@microsoft.graph.downloadUrl"]).then(function (result) {
                resolve(result);
            })
        }).catch(function (error) {
            reject(error);
        })
    });

}

function donwload_metadata(token, file_path) {
    return new Promise(function (resolve, reject) {
        var URI = "https://graph.microsoft.com/v1.0/me/drive/root:/" + file_path;

        var metaData_request = new XMLHttpRequest();

        metaData_request.onreadystatechange = function () {
            if (this.readyState == 4) {
                if (this.status == 200) {
                    resolve([this.status,metaData_request.responseText]);
                } else {
                    reject([this.status,metaData_request.responseText]);
                }
            }
        };
        metaData_request.open("GET", URI, true);
        metaData_request.setRequestHeader("Authorization", "bearer " + token);
        metaData_request.send();
    });

}

function download_contents(download_uri) {
    return new Promise(function (resolve, reject) {
        var download_request = new XMLHttpRequest();
        download_request.onreadystatechange = function () {
            if (this.readyState == 4) {
                if (this.status == 200) {
                    resolve([this.status,download_request.responseText]);
                } else {
                    reject([this.status,download_request.responseText]);
                }
            }
        }
        download_request.open("GET", download_uri, true);
        download_request.send();
    });
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


document.addEventListener( "DOMContentLoaded", function(){
    document.getElementById("step3").style.visibility = "hidden";
})