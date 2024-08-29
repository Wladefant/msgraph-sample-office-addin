// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import Router from 'express-promise-router';
import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import * as msal from '@azure/msal-node';

const authRouter = Router();

// <TokenExchangeSnippet>
// Initialize an MSAL confidential client
const msalClient = new msal.ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_APP_ID || '',
    clientSecret: process.env.AZURE_CLIENT_SECRET || ''
  }
});

const keyClient = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/common/discovery/v2.0/keys'
});

// Parses the JWT header and retrieves the appropriate public key
function getSigningKey(header: JwtHeader, callback: SigningKeyCallback): void {
  console.log('getSigningKey function started');
  if (header) {
    keyClient.getSigningKey(header.kid || '', (err, key) => {
      if (err) {
        callback(err, undefined);
      } else {
        callback(null, key?.getPublicKey());
      }
    });
  }
  console.log('getSigningKey function ended');
}

// Validates a JWT and returns it if valid
async function validateJwt(authHeader: string): Promise<string | null> {
  console.log('validateJwt function started');
  return new Promise((resolve) => {
    const token = authHeader.split(' ')[1];

    // Ensure that the audience matches the app ID
    // and the signature is valid
    const validationOptions = {
      audience: process.env.AZURE_APP_ID
    };

    jwt.verify(token, getSigningKey, validationOptions, (err) => {
      if (err) {
        console.log(`Verify error: ${JSON.stringify(err)}`);
        resolve(null);
      } else {
        resolve(token);
      }
    });
  });
  console.log('validateJwt function ended');
}

// Gets a Graph token from the API token contained in the
// auth header
export async function getTokenOnBehalfOf(authHeader: string): Promise<string | undefined> {
  console.log('getTokenOnBehalfOf function started');
  // Validate the supplied token if present
  const token = await validateJwt(authHeader);

  if (token) {
    const result = await msalClient.acquireTokenOnBehalfOf({
      oboAssertion: token,
      skipCache: true,
      scopes: ['https://graph.microsoft.com/.default']
    });

    return result?.accessToken;
  }
  console.log('getTokenOnBehalfOf function ended');
}
// </TokenExchangeSnippet>

// <GetAuthStatusSnippet>
// Checks if the add-in token can be silently exchanged
// for a Graph token. If it can, the user is considered
// authenticated. If not, then the add-in needs to do an
// interactive login so the user can consent.
authRouter.get('/status',
  async function(req, res) {
    console.log('authRouter.get("/status") function started');
    // Validate access token
    const authHeader = req.headers['authorization'];
    if (authHeader) {
      try {
        const graphToken = await getTokenOnBehalfOf(authHeader);

        // If a token was returned, consent is already
        // granted
        if (graphToken) {
          console.log(`Graph token: ${graphToken}`);
          res.status(200).json({
            status: 'authenticated'
          });
        } else {
          // Respond that consent is required
          res.status(200).json({
            status: 'consent_required'
          });
        }
      } catch (error) {
        // Respond that consent is required if the error indicates,
        // otherwise return the error.
        // @ts-ignore
        const payload = error.name === 'InteractionRequiredAuthError' ?
          { status: 'consent_required' } :
          { status: 'error', error: error};

        res.status(200).json(payload);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
    console.log('authRouter.get("/status") function ended');
  }
);
// </GetAuthStatusSnippet>

export default authRouter;
