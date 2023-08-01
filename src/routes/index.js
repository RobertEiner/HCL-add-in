// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file defines the home page request handling.

const express = require("express");
//router is like a mini-app that lives inside the app in app.js
const router = express.Router();
const webpush = require("web-push");
const path = require("path");

const publicKey = "BOUmVvzTvdtLU-j0kYSEV-UFtC5C8Ol1FUhMPm6kXwD3metSrG0S72O4cspVw1-6QBHUk1fzCvbyzKvLRC8GsEI";
const privateKey = "OhJmUzLVIkuZhrR75olGeVgxgAToaE7IN9w883C-Tik";

webpush.setVapidDetails("mailto:test@test.com", publicKey, privateKey);

// send the notification to the endpoint that the body of the incoming subscription object contains.
router.post("/subscribe", (req, res) => {
  const subscription = req.body;
  console.log("------")
  console.log("inside router post");
  // resource created
  res.status(201).json({});
  const payload = JSON.stringify({ title: "Robert Einer", body: "Hi, please download the attachments.." });
  console.log(`Inside post`);

  webpush.sendNotification(subscription, payload).catch((err) => console.error("ERROR:::: " + err));
});

module.exports = router;
