# little_onedrive_spwa
little_onedrive_spwa is a Single Page Web Application (SPWA) to connect to the [OneDrive REST API](https://docs.microsoft.com/en-gb/onedrive/developer/rest-api/?view=odsp-graph-online).

Check out the SPWA : https://aliceliveprojects.github.io/little_onedrive_spwa/src/index.html

The SPWA can do two things:

1. Reqest OneDrive authorisation token.
2. Download a specific file from a OneDrive account



## Setting Up Microsoft Developer Account

Before you can use the wonderful SPWA  you need to setup a Microsoft developer account.



### Creating Developer Account and Application

1. Navigate to Microsoft Graph :  <https://developer.microsoft.com/en-us/graph/>
2. Click the My Apps login, located on the top bar when not signed in.
3. Microsoft will prompt for a login
4. Once logged in, click the Add An App button
5. Name the application and create



---



### Filling Application Details

Within a specific application : 

1. Take note of the Application Id (also known as Client Id)
2. Under Platforms click the Add Platform button
3. Choose the Web Option
4. Check the Allow Implicit Flow checkbox
5. Under Redirect URLs enter the applications callback, example [application_domain]/callback.html
6. Under Microsoft Graph Permissions , remove any default permissions
7. Add the Sites.Read.All Application Permission (not Delegated Permissions)



## Filling SPWA Application Information

Within the script.js file locate ```appInfo``` on Line 7. Modify ```appInfo``` to the appropriate details, these details must match with the application details filled out earlier.

```clientId``` : application Id

```redirectUri``` : application redirect URL

```scopes```: application scopes delimited by a space and formatted to lowercase (example Sites.Read.All ===> sites.read.all)

```authServiceUri```: don't change (should be https://login.microsoftonline.com/common/oauth2/v2.0/authorize)



## Deploy application

Deploying the application is really easy.

1. Clone the repo
2. cd into the repo directory
3. run ``npm install``
4. cd into the src folder
5. run ```node server.js```



To run the SPWA on https read https://digitallabs.mmu.ac.uk/taming-the-urban-wild/#more-1657



## Notes on the OneDrive API

Some notes on the OneDrive API, which I painfully found so you don't have to.

The OneDrive REST API provides two ways to download the contents of a file : 

1. Using the /content endpoint

2. Using the '@microsoft.graph.downloadUrl' property


The /content endpoint returns a 302 response redirecting to a temporary pre-authenticated url to download the file contents. Requesting the /content endpoint from a sever works great, but, requesting from client-side the 302 causes a CORS error.



*For more information look at [Issues](https://github.com/microsoftgraph/microsoft-graph-docs/issues/43) or [Issues Archived](http://web.archive.org/web/20181205150708/https://github.com/microsoftgraph/microsoft-graph-docs/issues/43)*



---



To download the file the contents from client side the '@microsoft.graph.downloadUrl' property can be used. First you must request files meta-data which includes the '@microsoft.graph.downloadUrl' property. The '@microsoft.graph.downloadUrl' provides a temporary authenticated url to download file contents.



When requesting the '@microsoft.graph.downloadUrl' url DO NOT send any Authorisation headers. Sending Authorisation headers will cause a :

1. CORS error when requesting from client

2. 404 error when reqeusting from server


*For more information at [Docs](https://docs.microsoft.com/en-gb/onedrive/developer/rest-api/api/driveitem_get_content?view=odsp-graph-online) or [Docs Archived](http://web.archive.org/web/20181205150810/https://docs.microsoft.com/en-gb/onedrive/developer/rest-api/api/driveitem_get_content?view=odsp-graph-online)*
