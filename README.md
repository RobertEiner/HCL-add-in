# **Academic Trainee Assignment**

This is a project that was started during the HCL summer internship in the summer of 2023.

## **Assignment overview**

### **Description**

The idea for the assignment is to make it more seamless to download Outlook email attachments such as PDFs, images etc.

To do this, the task was to create a custom email notification pop-up that would appear each time the user received a new email. On the pop-up, there should be a clickable button that, when pressed, automatically downloads the attachment from the email and uploads it to a specific folder in the user's OneDrive account.

### **Structure**

The structure of the applicaiton consists of a few components:

#### **Microsoft Outlook add-in**

[MS add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) are basically embedded web applicaitons within Office applications that lets the developer cusotomize the appearance and functionality of them. However, they are quite limited in the ways you can do this.  

This component is based on a tutorial from the [Microsoft docs](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=yeomangenerator). The parts belonging to this component is the ***components*** and ***taskpane*** folders with their respective source code files. The most important being the taskpane.js. It's responsibilities:
*  Aquire an access token from MS Identity Platform to be able to authenticate to [MS Graph API](https://learn.microsoft.com/en-us/graph/use-the-api). Through Graph API, access would be granted to the user's OneDrive data, to be able to upload email attachments. 
* Pass the access token to the Express.js server, which is a necessary step towards accessing the Graph API.

#### **Microsoft Graph API**

Using the MS Graph API, a user's OneDrive data can be accessed. This is the recommended API from Microsoft for this purpose. This requires an authorization and authentication process. Information on how toauthorize to MS Graph API from an Outlook add-in can be found [here](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/authorize-to-microsoft-graph).

The tutorial used for implementing this is [this one](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs), it includes:
* Registering the application in the Azure Active Directory (requires admin account). An [MS Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) sandbox environment can be created if admin rights are missing.
* Cloning repository from GitHub

The application the tutorial creates only works on  Powerpoint, Excel and word, however, the authentication principles are the same for Outlook. 
The parts belonging to this component is in the ***server-helpers*** folder (contains files for facilitating the authentication process) and the ***routes*** folder (contains files for actually extracting data from OneDrive through Graph API). 

#### **Web notifications**

As previously mentioned, the MS add-ins are quite limited, it is not possible to modify the existing email pop-up in any way through the add-in. An alternative solution is to use web notifications with a NPM package called ***web-push***. This would make the add-in a web-only add-in, not available for Outlook on desktop, which is a limitation. 

In the ***client*** folder, the important files are:
<ul>
    <li>worker.js
        <ul>
            <li>Using the <a href="https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API">service worker API</a></li>
            <li>The service worker is like a proxy that sits between the web browser and the web applicaiton</li>
            <li>Helps out with managing the web notifcation</li>
        </ul> 
    </li>
    <li>client.js
    <ul>
        <li>Responsible for registering the service worker</li>
        <li>Makes http request to the server in the routes folder (index.js file)</li>
     </li>
    </ul>
</ul>


### **Issues**
There are a few issues that haven't been solved with this assignment, but it can most likely be finished. However, given the time constraint of the summer job, a decision was made to pause development and continue with another assignment.

#### ***Integration of web notification with Outlook add-in***
The web notification works as an isolated application, however it is not integrated with the add-in. How to implement this remains to be seen, as of now, a successful solution hasn't been found. 

#### ***Accessing OneDrive data through Microsoft Graph API***
The access token is successfully retrieved from the MS Identity Platform and sent to the server (getFilesRoute.js). When the access token is used to authenticate to MS Graph API, nothing happens, except an error is printed in the console of the browser: 

>AADSTS65001: The user or administrator has not consented to use the application with ID: '...' Send an interactive authorization request for this user and resource.

Admin consent has been given through the Azure Active Directory for the application and the permissions it needs. A way to send an interactive authorization request to the user has not yet been found. 

#### ***Triggering add-in when receiving new email***
Add-ins can be triggered by certain predefined [events](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch?tabs=xmlmanifest), such as writing a new email or changing the recipients to an email. But there is no such predefined event that can trigger the add-in when a new email is received. A possible workaround for this could be to listen for [change notifications](https://learn.microsoft.com/en-us/graph/api/resources/webhooks?view=graph-rest-1.0) through MS Graph API. This needs to be researched further, to see if it is feasible.  



## Test application at current state
To test this application at it's current state, follow these instructions:


### Prerequisites
* Code editor e.g. [VS Code](https://code.visualstudio.com/)
* [Node.js](https://nodejs.org/en) with NPM 
* Right now the app is registered on my account in a sandbox environment that will eventually expire. If development will continue using the code provided, the developer needs a Microsoft account with admin privileges to register the application in Azure Active Directory (If you don't have it, use [MS Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) to get sandbox environment with admin rights)

### Steps

**Clone repository**

>git clone https://github.com/RobertEiner/HCL-add-in.git

***Run MS Outlook add-in***
1. Navigate to root folder of the repo and run ***npm install***
2. Run ***npm start*** to start the webPack server that runs the add-in
3. Log in to your Outlook account with admin priviliges.
4. Open an email that contains an attachment in your inbox, press the three dots at the top right corner of the email. Scroll down the list until you find HCL-add-in and press "show taskpane". The taskpane should open.
5. Press the ***fetch*** button, and you will se the token that was aquired from the MS Identity Platform displayed in the taskpane, together with the attachment name of the email.
6. When this button is pressed, the idea is that the first ten folder/file names that is in the user's OneDrive account should display. However, the error message explained earlier regarding user consent appears in the bowser console. 

***Web notification***
1. Run server ***npm run app***
2. Open browser window and enter the URL: ***localhost:3001***
3. Press the button that says "Send"
4. Wait for push notification to arrive in the bottom right corner of the screen.












    

     