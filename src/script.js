

function login(){
   
    //OneDrive Application information, retrieved from Microsoft Graph API
    var appInfo = {
        "clientId": 'a60787c6-f851-4306-b7ec-4ec73c37e51d',
        "redirectUri": "https://onedrive.live.com",
        "scopes": "sites.read.all",
        "authServiceUri": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    }

    oneDrive_login(appInfo,function(){
        //ignore just for styling
        document.getElementById("step2").style.visibility = "visible";
    })
}

function logout(){
    oneDrive_logout();
}



// Downloads specific file from a OneDrive account
// file_path : absolute path to file, example : test/test/test.txt 
function download(file_path) {
    if (!is_authenticated()) {
        alert("login into onedrive")
        return;
    }

    oneDrive_download(file_path).then(function(result){
         // show contents of file
         document.getElementById("step3").style.visibility = "visible";
         document.getElementById("file_header").innerHTML = "File Path : " + file_path;
         document.getElementById("file_contents").innerHTML = "File Contents : " + result[1];
    }).catch(function(error){
        // Un-authorized
        if(error[0] == 401){
            alert("You are unauthorized, try logging in");
        }
        // File not found
        else if(error[0] == 404){
            alert("item not found, check path");
        }else if(error[0] == -1){
            alert("invalid file path");
        }
        else{
            alert("You have a weird error, check the console");
            console.log(error);
        }
    });
}



//ignore just for styling
document.addEventListener( "DOMContentLoaded", function(){
    document.getElementById("step3").style.visibility = "hidden";
    document.getElementById("step2").style.visibility = "hidden"
})