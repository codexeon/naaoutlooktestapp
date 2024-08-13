
import { createNestablePublicClientApplication } from "@azure/msal-browser";

if (eval('typeof Array.prototype.findLast === "undefined"')) {
  console.log("Running in Chakra");
} else {
  console.log("Running in V8");
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
  console.log("ensureInitialized called");
  if (_pca == null) {
    _pca = await createNestablePublicClientApplication(config);
  }
  console.log("ensureInitialized completed");
  return _pca;
}

async function getTokenTest(eventObj) {
  try {
    const pca = await ensureInitialized();
    let result = await pca.ssoSilent({scopes: ["user.read"]});
    writeOutputToMail("Got token for account:" + result.account.username + " with length:" + result.accessToken.length);

  }
  catch (ex) {
    writeOutputToMail("Encountered an error" + ex);
  }
}

Office.actions.associate("getTokenTest", getTokenTest);
