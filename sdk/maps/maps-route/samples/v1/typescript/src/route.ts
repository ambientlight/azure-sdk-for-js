// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Route API usage. Simple queries are performed.
 */

import fs from "fs";
import { getDefaultAzureCredential } from "@azure/identity";
import * as coreAuth from "@azure/core-auth";
import { OperationOptions } from '@azure/core-http';
import { RouteClient } from "@azure/maps-route";
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

  const route = new RouteClient(credential).route;

  const filePathForPostRouteDirections = "../../resources/route_directions_request_body.json";
  const filePathForPostRouteDirectionsBatch = "../../resources/route_directions_batch_request_body.json";
  const filePathForPostRouteMatrix = "../../resources/route_matrix_request_body.json";

  console.log(" --- Get route directions:");
  console.log(await route.getRouteDirections("json", "52.50931,13.42936:52.50274,13.43872", operationOptions));

  console.log(" --- Get route range:");
  const routeRangeOptions = { "timeBudgetInSec": 6000 };
  console.log(await route.getRouteRange("json", "50.97452,5.86605", { ...routeRangeOptions, ...operationOptions }));

  const postRouteDirectionsPayload = JSON.parse(fs.readFileSync(filePathForPostRouteDirections, "utf8"));
  console.log(" --- Post route directions:");
  console.log(await route.postRouteDirections("json", "52.50931,13.42936:52.50274,13.43872", postRouteDirectionsPayload, operationOptions));

  console.log(" --- Post route directions batch:");
  const postRouteDirectionsBatchPayload = JSON.parse(fs.readFileSync(filePathForPostRouteDirectionsBatch, "utf8"));
  console.log(await route.beginPostRouteDirectionsBatchAndWait("json", postRouteDirectionsBatchPayload, operationOptions));

  console.log(" --- Post route matrix:");
  const postRouteMatrixPayload = JSON.parse(fs.readFileSync(filePathForPostRouteMatrix, "utf8"));
  console.log(await route.beginPostRouteMatrixAndWait("json", postRouteMatrixPayload, operationOptions));

}

main();
