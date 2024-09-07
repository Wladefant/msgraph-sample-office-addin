// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* global authConfig, localStorage, Office */

// <ConsentJsSnippet>
'use strict';

// Ensure MSAL is loaded, or provide a meaningful error message if it's not
// @ts-ignore
var msal = msal || {
  PublicClientApplication: () => {
    throw new Error('MSAL library is not loaded. Make sure MSAL.js is correctly included.');
  },
};

try {
  // Initialize MSAL client
  const msalClient = new msal.PublicClientApplication({
    auth: {
      // Ensure authConfig is defined and clientId is provided
      clientId: authConfig?.clientId || (() => { throw new Error('authConfig.clientId is missing or undefined.'); })(),
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: true,
    },
  });

  // Define the MSAL request object
  const msalRequest = {
    scopes: ['https://graph.microsoft.com/.default'],
  };

  // Function that handles the response from MSAL redirect
  function handleResponse(response) {
    localStorage.removeItem('msalCallbackExpected');
    if (response !== null) {
      // Successful authentication
      localStorage.setItem('msalAccountId', response.account.homeId);
      Office.context.ui.messageParent(
        JSON.stringify({ status: 'success', result: response.accessToken })
      );
    } else {
      // No response, handle as needed
      console.warn('No response from MSAL handleRedirectPromise.');
    }
  }

  Office.onReady(() => {
    if (Office.context.ui.messageParent) {
      // Attempt to handle redirect response
      msalClient.handleRedirectPromise()
        .then(handleResponse)
        .catch((error) => {
          console.error('Error handling MSAL redirect:', error);
          Office.context.ui.messageParent(
            JSON.stringify({ status: 'failure', result: error })
          );
        });

      // Check if we are expecting a callback
      if (!localStorage.getItem('msalCallbackExpected')) {
        localStorage.setItem('msalCallbackExpected', 'yes');

        // Check if the user has previously signed in
        if (localStorage.getItem('msalAccountId')) {
          try {
            msalClient.acquireTokenRedirect(msalRequest);
          } catch (error) {
            console.error('Error acquiring token:', error);
          }
        } else {
          try {
            msalClient.loginRedirect(msalRequest);
          } catch (error) {
            console.error('Error during login redirect:', error);
          }
        }
      }
    } else {
      console.error('Office.context.ui.messageParent is not available.');
    }
  });
} catch (error) {
  console.error('Error initializing MSAL:', error);
  Office.context.ui.messageParent(
    JSON.stringify({ status: 'failure', result: error.message })
  );
}
// </ConsentJsSnippet>
