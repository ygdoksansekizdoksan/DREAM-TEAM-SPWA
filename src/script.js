
function OneDrive_Login(){
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