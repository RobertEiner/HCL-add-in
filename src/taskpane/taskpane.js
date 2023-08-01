/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import jwt_decode from "jwt-decode";
import {send} from "../client/client"

// This line sets up an event handler that is triggered when the
// office.js library is fully loaded and the host application (outlook in this case) is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    //check if host application is outlook. Office.HostType.Outlook is provided by the office.js lib
    // getUserData();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("notification").onclick = send;
  }
});

export async function getUserData() {
  try {
    let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
    let userToken = jwt_decode(userTokenEncoded); // Using the https://www.npmjs.com/package/jwt-decode library.
    console.log(userToken.name); // user name
    console.log(userToken.preferred_username); // email
    console.log("AUD: " + userToken.aud); // user id
    console.log("ID: " + userToken.oid);
    console.log("Access: " + userToken.scp);

    console.log("beforeeeee");
    const path = "/getuserfilenames";

    // Check if the browser supports notifications
    if ("Notification" in window) {
      // Request permission from the user
      Notification.requestPermission().then((permission) => {
        if (permission === "granted") {
          // Permission granted, show a notification
          const notification = new Notification("New Email", {
            body: "You've received a new email!",
            icon: "path/to/notification-icon.png",
          });

          // Handle click event on the notification (optional)
          notification.onclick = function () {
            // Handle the click event here (e.g., open your add-in)
          };
        }
      });
    }

    const response = await fetch(path, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + userTokenEncoded,
      },
    });
    console.log(response);

    document.getElementById("token").innerHTML = "Token bearer: " + userToken.name;
    document.getElementById("token-id").innerHTML = "Token ID: " + userToken.oid;
  } catch (exception) {
    document.getElementById("token").innerHTML = "Exception: " + exception.message;
    if (exception.code === 13003) {
      console.log(exception.message);
      // SSO is not supported for domain user accounts, only
      // Microsoft 365 Education or work account, or a Microsoft account.
    } else {
      // Handle error
      console.log(exception.message);
    }
  }
}

export async function run() {
  const item = Office.context.mailbox.item;
  const attachment = item.attachments[0];
  await getUserData();
  if (item.attachments.length > 0) {
    document.getElementById("item-subject").innerHTML = "Attachments: <br/>" + attachment.name;
    // item.getAttachmentContentAsync(item.attachments[0].id, handleAttachmentCb);
  } else {
    console.log("No attachments");
  }

  // }

  // item.getAttachmentsAsync(function(result){
  //   if(result.status !== Office.AsyncResultStatus.Succeeded) {
  //     document.getElementById("item-subject").innerHTML = "ERROR: " + result.error.message;
  //   } else {
  //     if (result.value.length > 0) {
  //       const attachment = result.value[0];
  //       attachmentString += "Name: " + attachment.name + " <br/> Size: " + attachment.size;
  //       document.getElementById("item-subject").innerHTML = "Attachments: <br/>" + attachmentString;

  //     } else {
  //       document.getElementById("item-subject").innerHTML = "Attachments: no attachments";
  //     }
  //   }
  // });

  // if(item.attachments.length > 0) {
  //   for(let i = 0; i < item.attachments.length; i++) {
  //     const attachment = item.attachments[i];

  //     attachmentString += "Name: " + attachment.name + "<br/>";
  //   }
  //   // attachmentIds += " ID: " + item.attachments[0].id + "<br/>";
  //   document.getElementById("item-subject").innerHTML = "Attachments: <br/>" + attachmentString;
  // } else {
  //   document.getElementById("item-subject").innerHTML = "No attachments available" + attachmentString;
  // }
}

export async function notify() {
  console.log("hellodgg");
  Notification.requestPermission().then((perm) => {
    console.log(perm);

    if (perm === "denied") {
      new Notification("example");
    } else {
      console.log("hhgg");
    }
  });
}

export async function handleAttachmentCb(result) {
  // if(result.status === Office.AsyncResultStatus.Succeeded) {
  // if (typeof result === "object") {
  document.getElementById("item-subject").innerHTML = "Attachments: <br/>" + result.value.name;
  document.getElementById("attachments-id").innerHTML = "ATT: ";
}

// }
// }

// document.getElementById("attachments-id").innerHTML = "ID: <br/>" + attachmentIds;
// SON.stringify(result.value.format)
