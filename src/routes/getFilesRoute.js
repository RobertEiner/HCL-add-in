// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require("express");

var router = express.Router();
const authHelper = require("../server-helpers/obo-auth-helper");
const getGraphData = require("../server-helpers/msgraph-helper");
const jwt = require("jsonwebtoken");
//
// TODO 9: Add route for /getuserfilenames REST API

router.get("/getuserfilenames", authHelper.validateJwt, async function (req, res) {
    console.log("gg");
    res.send("SERver says hi");
  // TODO 10: Exchange the access token for a Microsoft Graph token
  //          by using the OBO flow.
  try {
    const authHeader = req.headers.authorization;
    let oboRequest = {
      oboAssertion: authHeader.split(" ")[1],
      scopes: ["files.read"],
    };

    // The Scope claim tells you what permissions the client application has in the service.
    // In this case we look for a scope value of access_as_user, or full access to the service as the user.
    const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
    const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
    if (!accessAsUserScope) {
      res.status(401).send({ type: "Missing access_as_user" });
      return;
    }
    const cca = authHelper.getConfidentialClientApplication();
    const response = await cca.acquireTokenOnBehalfOf(oboRequest);
    // TODO 11: Call Microsoft Graph to get list of filenames.
  } catch (err) {
    // TODO 12: Handle any errors.
  }
});

module.exports = router;
