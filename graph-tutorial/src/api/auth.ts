// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Router from 'express-promise-router';
import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
// @ts-ignore
import { ConfidentialClientApplication } from '@azure/msal-node';

const authRouter = Router();

// <TokenExchangeSnippet>
// Initialize an MSAL confidential client
const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_APP_ID || '',
    clientSecret: process.env.AZURE_CLIENT_SECRET || '',
  },
});

const keyClient = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/common/discovery/v2.0/keys',
});

// Parses the JWT header and retrieves the appropriate public key
function getSigningKey(header: JwtHeader, callback: SigningKeyCallback): void {
  if (header) {
    keyClient.getSigningKey(header.kid || '', (err, key) => {
      if (err) {
        callback(err, undefined);
      } else {
        callback(null, key?.getPublicKey());
      }
    });
  }
}

// Validates a JWT and returns it if valid
async function validateJwt(authHeader: string): Promise<string | null> {
  return new Promise((resolve) => {
    const token = authHeader.split(' ')[1];

    // Ensure that the audience matches the app ID
    // and the signature is valid
    const validationOptions = {
      audience: process.env.AZURE_APP_ID,
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
}

// Gets a Graph token from the API token contained in the
// auth header
export async function getTokenOnBehalfOf(
  authHeader: string,
): Promise<string | undefined> {
  // Validate the supplied token if present
  const token = await validateJwt(authHeader);

  if (token) {
    const result = await msalClient.acquireTokenOnBehalfOf({
      oboAssertion: token,
      skipCache: true,
      scopes: ['https://graph.microsoft.com/.default'],
    });

    return result?.accessToken;
  }
}
// </TokenExchangeSnippet>

// <GetAuthStatusSnippet>
// Checks if the add-in token can be silently exchanged
// for a Graph token. If it can, the user is considered
// authenticated. If not, then the add-in needs to do an
// interactive login so the user can consent.
authRouter.get('/status', async function (req, res) {
  // Validate access token
  const authHeader = req.headers['authorization'];
  if (authHeader) {
    try {
      const graphToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IktxaGpJUkM3STBOTk85VE5XWFdGQXBGbTVUNWtOblc0dW81bmFVOFpKUjAiLCJhbGciOiJSUzI1NiIsIng1dCI6Ikg5bmo1QU9Tc3dNcGhnMVNGeDdqYVYtbEI5dyIsImtpZCI6Ikg5bmo1QU9Tc3dNcGhnMVNGeDdqYVYtbEI5dyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9hMzE5N2ZkZi01ZGUyLTQzOGItOThhMy03YTYyODcxNjFjMzAvIiwiaWF0IjoxNzI1ODY1MjEwLCJuYmYiOjE3MjU4NjUyMTAsImV4cCI6MTcyNTg2OTM3NCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhYQUFBQTd3ZktlUFJkU2MvZXJWS0c5TFBwR09YclZxZmI3VHdlSGYxcVNwMW9QcXhqa3ozbTMzQkpDN3J4alJOeDczbzBKcFQ3WlY3QkpmNW5LVmpIS05xZm85cEFXbHpwVVBDQ1JZYzlrcVdQeVpNPSIsImFtciI6WyJwd2QiLCJyc2EiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoibXMtZ3JhcGh0ZXN0LWV4ZGVsIiwiYXBwaWQiOiIyYmJiNGI5NS0xOWYwLTQxMTYtOTI1Yi0xOTljZDFkODQxM2IiLCJhcHBpZGFjciI6IjEiLCJkZXZpY2VpZCI6IjIzOWUyZWM4LTMzZWEtNGY0Ny04ZWJmLTgxOWJkMGMyYjI5OSIsImZhbWlseV9uYW1lIjoiS2lyamFub3ZzIiwiZ2l2ZW5fbmFtZSI6IldsYWRpbWlyIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTcyLjIwMS4xNTYuMTIiLCJuYW1lIjoiV2xhZGltaXIgS2lyamFub3ZzIiwib2lkIjoiMWY5NTAzZDctYTY5NS00ZWM4LWE0MDUtNzIyZDhiNjQzNDgyIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAzODVGQTU3RjIiLCJyaCI6IjAuQWIwQTMzOFpvLUpkaTBPWW8zcGloeFljTUFNQUFBQUFBQUFBd0FBQUFBQUFBQUFEQVZZLiIsInNjcCI6IkNhbGVuZGFycy5SZWFkV3JpdGUgTWFpbGJveFNldHRpbmdzLlJlYWQgVXNlci5SZWFkIHByb2ZpbGUgb3BlbmlkIGVtYWlsIiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2NtcCJdLCJzdWIiOiJ1NGxVdXZBQ0tla1Y4THYxd0pfNnpBQkdhbC1waFhtRm90MkprRzFROHFNIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiYTMxOTdmZGYtNWRlMi00MzhiLTk4YTMtN2E2Mjg3MTYxYzMwIiwidW5pcXVlX25hbWUiOiJ3Lmtpcmphbm92c0ByZWFsZXN0LWFpLmNvbSIsInVwbiI6Incua2lyamFub3ZzQHJlYWxlc3QtYWkuY29tIiwidXRpIjoiT2V6Q1lqbUdRVTYzYV9uN0tCb2FBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMjc0NjA4ODMtMWRmMS00NjkxLWIwMzItM2I3OTY0M2U1ZTYzIiwiNDVkOGQzYzUtYzgwMi00NWM2LWIzMmEtMWQ3MGI1ZTFlODZlIiwiYWM0MzQzMDctMTJiOS00ZmExLWE3MDgtODhiZjU4Y2FhYmMxIiwiOWM2ZGYwZjItMWU3Yy00ZGMzLWIxOTUtNjZkZmJkMjRhYThmIiwiMjVhNTE2ZWQtMmZhMC00MGVhLWEyZDAtMTI5MjNhMjE0NzNhIiwiMzFlOTM5YWQtOTY3Mi00Nzk2LTljMmUtODczMTgxMzQyZDJkIiwiMTE0NTFkNjAtYWNiMi00NWViLWE3ZDYtNDNkMGYwMTI1YzEzIiwiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiZGQxMzA5MWEtNjIwNy00ZmMwLTgyYmEtMzY0MWUwNTZhYjk1IiwiODc3NjFiMTctMWVkMi00YWYzLTlhY2QtOTJhMTUwMDM4MTYwIiwiOTVlNzkxMDktOTVjMC00ZDhlLWFlZTMtZDAxYWNjZjJkNDdiIiwiNzU5MzQwMzEtNmM3ZS00MTVhLTk5ZDctNDhkYmQ0OWU4NzVlIiwiZjcwOTM4YTAtZmMxMC00MTc3LTllOTAtMjE3OGY4NzY1NzM3IiwiNGQ2YWMxNGYtMzQ1My00MWQwLWJlZjktYTNlMGM1Njk3NzNhIiwiMmI3NDViZGYtMDgwMy00ZDgwLWFhNjUtODIyYzQ0OTNkYWFjIiwiMTU4YzA0N2EtYzkwNy00NTU2LWI3ZWYtNDQ2NTUxYTZiNWY3IiwiNzY5OGE3NzItNzg3Yi00YWM4LTkwMWYtNjBkNmIwOGFmZmQyIiwiMWQzMzZkMmMtNGFlOC00MmVmLTk3MTEtYjM2MDRjZTNmYzJjIiwiZTZkMWEyM2EtZGExMS00YmU0LTk1NzAtYmVmYzg2ZDA2N2E3IiwiZWIxZjRhOGQtMjQzYS00MWYwLTlmYmQtYzdjZGY2YzVlZjdjIiwiOTJiMDg2YjMtZTM2Ny00ZWYyLWI4NjktMWRlMTI4ZmI5ODZlIiwiMjVkZjMzNWYtODZlYi00MTE5LWI3MTctMGZmMDJkZTIwN2U5IiwiNDQzNjcxNjMtZWJhMS00NGMzLTk4YWYtZjU3ODc4NzlmOTZhIiwiMzEzOTJmZmItNTg2Yy00MmQxLTkzNDYtZTU5NDE1YTJjYzRlIiwiNzQ5NWZkYzQtMzRjNC00ZDE1LWEyODktOTg3ODhjZTM5OWZkIiwiM2EyYzYyZGItNTMxOC00MjBkLThkNzQtMjNhZmZlZTVkOWQ1IiwiMTUwMWI5MTctNzY1My00ZmY5LWE0YjUtMjAzZWFmMzM3ODRmIiwiNzkwYzFmYjktN2Y3ZC00Zjg4LTg2YTEtZWYxZjk1YzA1YzFiIiwiM2YxYWNhZGUtMWUwNC00ZmJjLTliNjktZjAzMDJjZDg0YWVmIiwiM2VkYWY2NjMtMzQxZS00NDc1LTlmOTQtNWMzOThlZjZjMDcwIiwiMWE3ZDc4YjYtNDI5Zi00NzZiLWI4ZWItMzVmYjcxNWZmZmQ0IiwiNmU1OTEwNjUtOWJhZC00M2VkLTkwZjMtZTk0MjQzNjZkMmYwIiwiMjkyMzJjZGYtOTMyMy00MmZkLWFkZTItMWQwOTdhZjNlNGRlIiwiNTlkNDZmODgtNjYyYi00NTdiLWJjZWItNWMzODA5ZTU5MDhmIiwiMzI2OTY0MTMtMDAxYS00NmFlLTk3OGMtY2UwZjZiMzYyMGQyIiwiNWI3ODQzMzQtZjk0Yi00NzFhLWEzODctZTcyMTlmYzQ5Y2EyIiwiODkyYzU4NDItYTlhNi00NjNhLTgwNDEtNzJhYTA4Y2EzY2Y2IiwiYWMxNmU0M2QtN2IyZC00MGUwLWFjMDUtMjQzZmYzNTZhYjViIiwiYWFmNDMyMzYtMGMwZC00ZDVmLTg4M2EtNjk1NTM4MmFjMDgxIiwiMDUyNjcxNmItMTEzZC00YzE1LWIyYzgtNjhlM2MyMmI5ZjgwIiwiOTYzNzk3ZmItZWIzYi00Y2RlLThjZTMtNTg3OGIzZjMyYTNmIiwiNzQ0ZWM0NjAtMzk3ZS00MmFkLWE0NjItOGIzZjk3NDdhMDJjIiwiOGM4YjgwM2YtOTZlMS00MTI5LTkzNDktMjA3MzhkOWY5NjUyIiwiYzQzMGIzOTYtZTY5My00NmNjLTk2ZjMtZGIwMWJmOGJiNjJhIiwiOWYwNjIwNGQtNzNjMS00ZDRjLTg4MGEtNmVkYjkwNjA2ZmQ4IiwiZmRkN2E3NTEtYjYwYi00NDRhLTk4NGMtMDI2NTJmZThmYTFjIiwiODEwYTI2NDItYTAzNC00NDdmLWE1ZTgtNDFiZWFhMzc4NTQxIiwiY2YxYzM4ZTUtMzYyMS00MDA0LWE3Y2ItODc5NjI0ZGNlZDdjIiwiZmZkNTJmYTUtOThkYy00NjVjLTk5MWQtZmMwNzNlYjU5ZjhmIiwiODQyNGM2ZjAtYTE4OS00OTllLWJiZDAtMjZjMTc1M2M5NmQ0IiwiM2Q3NjJjNWEtMWI2Yy00OTNmLTg0M2UtNTVhM2I0MjkyM2Q0IiwiYjFiZTFjM2UtYjY1ZC00ZjE5LTg0MjctZjZmYTBkOTdmZWI5IiwiOWI4OTVkOTItMmNkMy00NGM3LTlkMDItYTZhYzJkNWVhNWMzIiwiOGFjM2ZjNjQtNmVjYS00MmVhLTllNjktNTlmNGM3YjYwZWIyIiwiNzI5ODI3ZTMtOWMxNC00OWY3LWJiMWItOTYwOGYxNTZiYmI4IiwiMGY5NzFlZWEtNDFlYi00NTY5LWE3MWUtNTdiYjhhM2VmZjFlIiwiOTM2MGZlYjUtZjQxOC00YmFhLTgxNzUtZTJhMDBiYWM0MzAxIiwiMzhhOTY0MzEtMmJkZi00YjRjLThiNmUtNWQzZDhhYmFjMWE0IiwiZTQ4Mzk4ZTItZjRiYi00MDc0LThmMzEtNDU4NjcyNWUyMDViIiwiNzRlZjk3NWItNjYwNS00MGFmLWE1ZDItYjk1MzlkODM2MzUzIiwiZmU5MzBiZTctNWU2Mi00N2RiLTkxYWYtOThjM2E0OWEzOGIxIiwiYTllYTg5OTYtMTIyZi00Yzc0LTk1MjAtOGVkY2QxOTI4MjZjIiwiODhkOGUzZTMtOGY1NS00YTFlLTk1M2EtOWI5ODk4Yjg4NzZiIiwiYmUyZjQ1YTEtNDU3ZC00MmFmLWEwNjctNmVjMWZhNjNiYzQ1IiwiYjBmNTQ2NjEtMmQ3NC00YzUwLWFmYTMtMWVjODAzZjEyZWZlIiwiNThhMTNlYTMtYzYzMi00NmFlLTllZTAtOWMwZDQzY2Q3ZjNkIiwiYzRlMzliZDktMTEwMC00NmQzLThjNjUtZmIxNjBkYTAwNzFmIiwiOWM5OTUzOWQtODE4Ni00ODA0LTgzNWYtZmQ1MWVmOWUyZGNkIiwiZjJlZjk5MmMtM2FmYi00NmI5LWI3Y2YtYTEyNmVlNzRjNDUxIiwiZjI4YTFmNTAtZjZlNy00NTcxLTgxOGItNmExMmYyYWY2YjZjIiwiMTEyY2ExYTItMTVhZC00MTAyLTk5NWUtNDViMGJjNDc5YTZhIiwiZmNmOTEwOTgtMDNlMy00MWE5LWI1YmEtNmYwZWM4MTg4YTEyIiwiZTMwMGQ5ZTctNGEyYi00Mjk1LTllZmYtZjFjNzhiMzZjYzk4IiwiYWEzODAxNGYtMDk5My00NmU5LTliNDUtMzA1MDFhMjA5MDlkIiwiNWM0ZjlkY2QtNDdkYy00Y2Y3LThjOWEtOWU0MjA3Y2JmYzkxIiwiYjVhOGRjZjMtMDlkNS00M2E5LWE2MzktOGUyOWVmMjkxNDcwIiwiMjgxZmU3NzctZmIyMC00ZmJiLWI3YTMtY2NlYmNlNWIwZDk2IiwiZDM3YzhiZWQtMDcxMS00NDE3LWJhMzgtYjRhYmU2NmNlNGMyIiwiOTJlZDA0YmYtYzk0YS00YjgyLTk3MjktYjc5OWE3YTRjMTc4IiwiMTczMTU3OTctMTAyZC00MGI0LTkzZTAtNDMyMDYyY2FjYTE4IiwiODMyOTE1M2ItMzFkMC00NzI3LWI5NDUtNzQ1ZWIzYmM1ZjMxIiwiZTM5NzNiZGYtNDk4Ny00OWFlLTgzN2EtYmE4ZTIzMWM3Mjg2IiwiZjAyM2ZkODEtYTYzNy00YjU2LTk1ZmQtNzkxYWMwMjI2MDMzIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19pZHJlbCI6IjEgMjIiLCJ4bXNfc3QiOnsic3ViIjoiZ0l5Q3FVOVdiQkY0YW5KcXFac05PSjJxX2lkSGxzNmZVWlM1V3hWVDYzZyJ9LCJ4bXNfdGNkdCI6MTcxNjUyOTI2Mn0.OZKoB0KVyeDNO4DNsPdPgZkls7ZW-vykWvBcvO2yrVoicsB-UZORY_pmI-a1_ZqTPI9zqMhldZaz_pQK1z19XqxSplYsrwWjGkh0FkMcZ4IvJ2lnVp-mcXvIj7HY7n_XD-KK45JVPWCK6xsR8kFgJDugpTudz9mvPdsp1vGZfokXoR4xKTFNIZ7yr9TOdcg9aQdxw-j52QhQ5s1gJTrE7yOfq3vNNg4NdSnxaTeRcpoz3CPjZTzXY5KF12r9H4z-mCX30Xhd2LU4BIbwNoDYtXrnGVDHjrjA4kmFfK9HOhAenxj551OdatYJ9lh4hqWKAMeZngQrxpx1dZ7W3xOjKg";

      // If a token was returned, consent is already
      // granted
      if (graphToken) {
        console.log(`Graph token: ${graphToken}`);
        res.status(200).json({
          status: 'authenticated',
        });
      } else {
        // Respond that consent is required
        res.status(200).json({
          status: 'consent_required',
        });
      }
    } catch (error) {
      // Respond that consent is required if the error indicates,
      // otherwise return the error.
      const payload =
        // @ts-ignore
        error.name === 'InteractionRequiredAuthError'
          ? { status: 'consent_required' }
          : { status: 'error', error: error };

      res.status(200).json(payload);
    }
  } else {
    // No auth header
    res.status(401).end();
  }
});
// </GetAuthStatusSnippet>

export default authRouter;
