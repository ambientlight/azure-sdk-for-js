// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Alias API usage. Simple CRUD operations are performed.
 */

import { getDefaultAzureCredential } from "@azure/identity";
import * as coreAuth from "@azure/core-auth";
import { OperationOptions } from '@azure/core-http';
import { CreatorClient } from "@azure/maps-creator";
import * as dotenv from "dotenv";
dotenv.config();

/**
 * Azure Maps supports two ways to authenticate requests:
 * - Shared Key authentication (subscription-key)
 * - Azure Active Directory (Azure AD) authentication
 * 
 * In this sample you can put MAPS_SUBSCRIPTION_KEY into .env file to use the first approach or populate
 * the three AZURE_CLIENT_ID, AZURE_CLIENT_SECRET & AZURE_TENANT_ID variables for trying out AAD auth.
 * 
 * More info is available at https://docs.microsoft.com/en-us/azure/azure-maps/azure-maps-authentication.
 */


/**
 * Empty token class definition. To be used with AzureKey credentials.
 */
class EmptyTokenCredential implements coreAuth.TokenCredential {
  async getToken(
    _scopes: string | string[],
    _options?: coreAuth.GetTokenOptions
  ): Promise<coreAuth.AccessToken | null> {
    return {
      token: "token",
      expiresOnTimestamp: Date.now() + 60 * 60 * 1000
    };
  }
}


async function main() {
  let credential: coreAuth.TokenCredential;
  let operationOptions: OperationOptions = {};

  if (process.env.MAPS_SUBSCRIPTION_KEY) {
    // Use subscription key authentication
    credential = new EmptyTokenCredential();
    operationOptions.requestOptions = { customHeaders: { "subscription-key": process.env.MAPS_SUBSCRIPTION_KEY } };
  }
  else {
    // Use Azure AD authentication
    credential = getDefaultAzureCredential();
  }

  const alias = new CreatorClient(credential).alias;

  console.log(" --- Create Alias:");
  const aliasCreateResponse = await alias.create(operationOptions);
  console.log(aliasCreateResponse);
  const aliasId = aliasCreateResponse.aliasId;

  // TO USE need to have some data uploaded already - please use env CREATOR_DWG_ZIP_UDID or CREATOR_GEOJSON_UDID
  const udid = process.env.CREATOR_DWG_ZIP_UDID;
  if ((typeof udid === "string") && (udid.length == 36)) {
    console.log(" --- Assign the aliasId to some Creator's udid:");
    console.log(await alias.assign(aliasId!, udid!, operationOptions));
  }

  console.log(" --- Get details about the created Alias:");
  console.log(await alias.getDetails(aliasId!, operationOptions));

  console.log(" --- Delete the created Alias:");
  await alias.delete(aliasId!, operationOptions);
  console.log("Done (no response body)");

  console.log(" --- List all the created Aliases:");
  for await (const aliasItem of alias.list(operationOptions)) {
    console.log(aliasItem);
  }

}

main();
