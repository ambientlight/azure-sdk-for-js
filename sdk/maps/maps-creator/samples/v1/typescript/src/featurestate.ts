// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Feature State API usage. Simple CRUD operations are performed.
 */

import fs from "fs";
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

  const featureState = new CreatorClient(credential).featureState;

  const filePathForCreate = "../../resources/featurestate_sample_create.json";
  const filePathForUpdate = "../../resources/featurestate_sample_put.json";
  const filePathForUpdateStates = "../../resources/featurestate_sample_update_states.json";

  // TO USE need to have some Dataset created already - please use env CREATOR_DATASET_ID
  const datasetId = process.env.CREATOR_DATASET_ID;
  if ((typeof datasetId !== "string") || (datasetId.length != 36)) {
    throw "This sample needs some Dataset created";
  }
  const featureId = "FCL13"; // this is taken from WFS GetFeatures call with "facility" as a collection

  const featureStateSetCreate = JSON.parse(fs.readFileSync(filePathForCreate, "utf8"));

  console.log(" --- Create Feature State Set:");
  const createResult = await featureState.createStateset(datasetId, featureStateSetCreate, operationOptions);
  console.log(createResult);
  const statesetId = createResult.statesetId;

  console.log(" --- Get Feature State Set:");
  console.log(await featureState.getStateset(statesetId!, operationOptions));

  const featureStateSetUpdateStates = JSON.parse(fs.readFileSync(filePathForUpdateStates, "utf8"));

  console.log(" --- Update states of the Feature State set:");
  console.log(await featureState.updateStates(statesetId!, featureId, featureStateSetUpdateStates, operationOptions));

  console.log(" --- Get states of the Feature State set:");
  console.log(await featureState.getStates(statesetId!, featureId, operationOptions));

  const featureStateSetUpdate = JSON.parse(fs.readFileSync(filePathForUpdate, "utf8"));

  console.log(" --- Update Feature State set:");
  console.log(await featureState.putStateset(statesetId!, featureStateSetUpdate, operationOptions));

  console.log(" --- Get states of the Feature State set:");
  console.log(await featureState.getStates(statesetId!, featureId, operationOptions));

  console.log(" --- Delete state of the Feature State set:");
  console.log(await featureState.deleteState(statesetId!, featureId, "s1", operationOptions));

  console.log(" --- Get states of the Feature State set:");
  console.log(await featureState.getStates(statesetId!, featureId, operationOptions));

  console.log(" --- Delete the created Feature State set:");
  await featureState.deleteStateset(statesetId!, operationOptions);
  console.log("Done (no response body)");

  console.log(" --- List all the Feature State sets:");
  for await (const featureStateSet of featureState.listStateset(operationOptions)) {
    console.log(featureStateSet);
  }
}

main();
