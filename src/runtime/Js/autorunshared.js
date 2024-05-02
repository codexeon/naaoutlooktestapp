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
  Office.context.mailbox.item.body.prependAsync(
    output + "|||",
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

function wait(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

async function emptyPromise() {
  return Promise.resolve();
}

async function waitMultiplePromise() {
  for (let i = 0; i < 100; i++) {
    await emptyPromise();
  }
}

const loadTime = new Date().getTime();
async function getTokenTest(eventObj) {
  writeOutputToMail("Starting test iteration");
  const pca = await ensureInitialized();
  for (var i = 0; i < 2; i++) {
    const startTime = new Date().getTime();
    try {
      await waitMultiplePromise();
      // writeOutputToMail(`Waiting ${i*1000}ms`)
      // // await wait(i*1000);
      // writeOutputToMail("Done waiting");
      let result = await pca.ssoSilent({scopes: ["user.read"]});
      writeOutputToMail("Received token response", result);
      writeOutputToMail("Active account", pca.getActiveAccount());
      writeOutputToMail("Got token for account:" + result.account.username + " with length:" + result.accessToken.length);
      const endTime = new Date().getTime();
      writeOutputToMail("Time taken:" + (endTime - startTime) + "ms");
      // const graphResponse = await makeGraphCall("https://graph.microsoft.com/v1.0/me", result.accessToken);
      // writeOutputToMail(JSON.stringify(graphResponse));
    } catch(ex) {
      writeOutputToMail("Encountered an error" + ex);
    }
  }

  const nextIterationTime = new Date().getTime();
  writeOutputToMail(`Waiting 10 seconds for next iteration total time taken:${nextIterationTime - loadTime}ms`);
  setTimeout(() => {
    getTokenTest(eventObj);
  }, 10000);
}

Office.actions.associate("getTokenTest", getTokenTest);
