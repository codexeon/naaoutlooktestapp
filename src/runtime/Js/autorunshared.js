console.log("Start JS bundle");

// May be a bug this is needed, test to see if it can repro on othe rbuilds
setInterval(() => {
}, 1000);

function _generateUUID() {
    let uuid = '';
    const chars = '0123456789abcdef';
  
    for (let i = 0; i < 36; i++) {
      if (i === 8 || i === 13 || i === 18 || i === 23) {
        uuid += '-';
      } else if (i === 14) {
        uuid += '4';
      } else if (i === 19) {
        uuid += chars.charAt(Math.floor(Math.random() * chars.length));
      } else {
        uuid += chars.charAt(Math.floor(Math.random() * chars.length));
      }
    }
  
    return uuid;
}

function customGetRandomValues(inputArray) {
  const byteLength = inputArray.byteLength;
  const buffer = new Uint8Array(inputArray.buffer, inputArray.byteOffset, byteLength);
  for (let i = 0; i < byteLength; i++) {
    buffer[i] = Math.floor(Math.random() * 256);
  }
  return inputArray;
}

function base64Decode(input) {
    // Characters used in base64 encoding
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
    
    // Remove any characters not in the base64 characters list above and any trailing '='
    let str = input.replace(/[^A-Za-z0-9\+\/\=]/g, '');
    
    let output = '';
    
    for (let bc = 0, bs, buffer, idx = 0; buffer = str.charAt(idx++); ~buffer && (bs = bc % 4 ? bs * 64 + buffer : buffer, bc++ % 4) ? output += String.fromCharCode(255 & bs >> (-2 * bc & 6)) : 0) {
        buffer = chars.indexOf(buffer);
    }
    
    return output;
}

var crypto = {
    randomUUID: _generateUUID,
    getRandomValues: customGetRandomValues,
};

window.crypto = crypto;
window.atob = base64Decode;

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

function debugObject(obj, name) {
  console.log(name + " defined in environment", obj != null);
  var properties = Object.keys(window).join(",");
  console.log(properties);
}
async function getTokenTest(eventObj) {
  try {
    debugObject(window, "window");
    debugObject(self, "self");
    debugObject(globalThis, "globalThis");
    debugObject(global, "global");
  } catch (ex) {
    console.log("Encountered error", ex);
  }

  const pca = await ensureInitialized();
  try {
    let result = await pca.ssoSilent({scopes: ["user.read"]});
    console.log("Received token response", result);
    console.log("Active account", pca.getActiveAccount());
    writeOutputToMail("Got token for account:" + result.account.username + " with length:" + result.accessToken.length);
  } catch(ex) {
    writeOutputToMail("Encountered an error" + ex);
  }
}

Office.actions.associate("getTokenTest", getTokenTest);
