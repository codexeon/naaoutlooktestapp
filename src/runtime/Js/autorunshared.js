console.log("Start JS bundle");

import { PublicClientNext } from "@azure/msal-browser";

const clientId = "e3680097-6e43-43ab-8a38-23ed74b97839";
const config = {
  auth: {
    clientId,
    supportsNestedAppAuth: true,
    authority: "https://login.microsoftonline.com/common"
  },
  cache: {
    cacheLocation: "localStorage",
  }
};
let _pca = null;
async function ensureInitialized() {
  if (_pca == null) {
    _pca = await PublicClientNext.createPublicClientApplication(config);
  }
  return _pca;
}

function writeOutputToMail(output) {
  console.log(output);
  Office.context.mailbox.item.body.setAsync(
    output,
    {
      coercionType: "html",
    }, (result) => {
      console.log(result);
    }
  );
}

async function makeGraphCall(endpointUrl, accessToken) {
  let files = await fetch(endpointUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  });
  let responseJson = await files.json();
  return responseJson;
};

async function getTokenTest(eventObj) {
  const pca = await ensureInitialized();
  try {
    let result = await pca.ssoSilent({scopes: ["user.read"]});
    console.log("Received token response", result);
    console.log("Active account", pca.getActiveAccount());
    const graphResponse = await makeGraphCall("https://graph.microsoft.com/v1.0/me", result.accessToken);
    writeOutputToMail("Got token for account:" + result.account.username + " with length:" + result.accessToken.length + " with profileInfo:" + JSON.stringify(graphResponse));
  } catch(ex) {
    writeOutputToMail("Encountered an error" + ex);
  }
}

Office.actions.associate("getTokenTest", getTokenTest);
