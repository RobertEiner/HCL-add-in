

const express = require("express");

var router = express.Router();
const authHelper = require("../server-helpers/obo-auth-helper");
const getGraphData = require("../server-helpers/msgraph-helper");
const jwt = require("jsonwebtoken");

router.get("/getuserfilenames", authHelper.validateJwt, async function (req, res) {
  // res.json({ message: "Hello from server" });
  console.log("before 4031");
  console.log(req.headers.authorization);

  // TODO 10: Exchange the access token for a Microsoft Graph token
  //          by using the OBO flow.

  try {
    const authHeader = req.headers.authorization;
    let oboRequest = {
      oboAssertion: authHeader.split(" ")[1],
      scopes: ["files.read", "profile", "userid", "mail.read"],
    };
    console.log("before 4032");
    console.log(authHeader)
    console.log("OBO: ")
    console.log(oboRequest.oboAssertion);
    // console.log(oboRequest.oboAssertion);

    // The Scope claim tells you what permissions the client application has in the service.
    // In this case we look for a scope value of access_as_user, or full access to the service as the user.
    const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
    const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
    if (!accessAsUserScope) {
      res.status(401).send({ type: "Missing access_as_user" });
      return;
    }
    
    const cca = authHelper.getConfidentialClientApplication();
    console.log("before 4033");
    console.log(cca);
    const response = await cca.acquireTokenOnBehalfOf(oboRequest);

    // TOD 11: Call Microsoft Graph to get list of filenames.

    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    console.log("before 4034");
    const graphData = await getGraphData(response.accessToken, rootUrl, params);
    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      console.log("code err");
      res.status(403).send({
        type: "Microsoft Graph",
        errorDetails: "An error occurred while calling the Microsoft Graph API.\n" + graphData,
      });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    // TODO 12: Check for expired token.
  } catch (err) {
    // TODO 12: Handle any errors.

    // On rare occasions the SSO access token is unexpired when Office validates it,
    // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Construct an error message to return to the client so it can refresh the SSO token.
    if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
      res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
    } else {
      res.status(403).send({ type: "Unknown", errorDetails: err });
    }
  }
});

router.get("/dialog.html", async function (req, res) {
  res.json({ message: "hello" });
});

module.exports = router;
// https://login.microsoftonline.com/5998e16b-cdb5-4590-9256-6f6d94dcbee1/adminconsent?client_id=7c5abd76-11c4-4993-9a55-ebc01c2d78d1