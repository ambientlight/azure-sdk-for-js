// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates WFS API usage. Simple queries are performed.
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

  const wfs = new CreatorClient(credential).wfs;

  // TO USE need to have some Dataset created already - please use env CREATOR_DATASET_ID
  const datasetId = process.env.CREATOR_DATASET_ID;
  if ((typeof datasetId !== "string") || (datasetId.length != 36)) {
    throw "This sample needs some Dataset created";
  }

  console.log(" --- Get conformance of the dataset:");
  console.log(await wfs.getConformance(datasetId, operationOptions));

  console.log(" --- Get landing page of the dataset:");
  console.log(await wfs.getLandingPage(datasetId, operationOptions));

  console.log(" --- Get collections:");
  console.log(await wfs.getCollections(datasetId, operationOptions));

  console.log(" --- Get collection:");
  console.log(await wfs.getCollection(datasetId, "facility", operationOptions));

  console.log(" --- Get collection definition:");
  console.log(await wfs.getCollectionDefinition(datasetId, "facility", operationOptions));

  console.log(" --- Get features of the collection:");
  const collectionFeatures = await wfs.getFeatures(datasetId, "facility", operationOptions);
  console.log(collectionFeatures);
  const featureId = collectionFeatures.features[0].id;

  console.log(" --- Get details of the feature of the collection:");
  console.log(await wfs.getFeature(datasetId, "facility", featureId!, operationOptions));

  /* This code works as expected however it would remove the feature and require to create the dataset once again.
    console.log(" --- Remove the feature:");
    await wfs.deleteFeature(datasetId, "facility", featureId!, operationOptions);
    console.log("Done (no response body)");*/

}

main();
