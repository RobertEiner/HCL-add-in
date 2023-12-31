/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import jwt_decode from "jwt-decode";
import { send } from "../client/client";

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
    let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      // forMSGraphAccess: true,
    });
    let userToken = jwt_decode(userTokenEncoded); // Using the https://www.npmjs.com/package/jwt-decode library.
    // console.log(userToken.name); // user name
    // console.log(userToken.preferred_username); // email
    // console.log("AUD: " + userToken.aud); // user id
    // console.log("ID: " + userToken.oid);
    // console.log("Access: " + userToken.scp);

    const path = "http://localhost:3001/getuserfilenames";
    console.log(userTokenEncoded)
  
    try {
      const response = await fetch(path, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + userTokenEncoded,
        },
      });
      // const data = await response.json();
      // console.log(data.message);
      const data = await response.json();
      console.log(data.errorDetails.errorMessage);
    } catch (err) {
      console.log(err);
    }
    //g
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
}

// export async function notify() {
//   console.log("hellodgg");
//   Notification.requestPermission().then((perm) => {
//     console.log(perm);

//     if (perm === "denied") {
//       new Notification("example");
//     } else {
//       console.log("hhgg");
//     }
//   });
// }

export async function handleAttachmentCb(result) {
  // if(result.status === Office.AsyncResultStatus.Succeeded) {
  // if (typeof result === "object") {
  document.getElementById("item-subject").innerHTML = "Attachments: <br/>" + result.value.name;
  document.getElementById("attachments-id").innerHTML = "ATT: ";
}
