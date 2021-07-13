// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Render API usage. Simple queries are performed.
 */

import fs from "fs";
import { getDefaultAzureCredential } from "@azure/identity";
import * as coreAuth from "@azure/core-auth";
import { OperationOptions } from '@azure/core-http';
import { RenderClient } from "@azure/maps-render";
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

  let maps = new RenderClient(credential);
  const render = maps.render;
  const renderV2 = maps.renderV2;

  console.log(" --- Get copyright caption:");
  console.log(await render.getCopyrightCaption("json", operationOptions));

  console.log(" --- Get copyright for tile:");
  console.log(await render.getCopyrightForTile("json", 6, 9, 22, operationOptions));

  console.log(" --- Get copyright for world:");
  console.log(await render.getCopyrightForWorld("json", operationOptions));

  console.log(" --- Get copyright from bounding box:");
  console.log(await render.getCopyrightFromBoundingBox("json", "52.41064,4.84228", "52.41072,4.84239", operationOptions));

  if (!fs.existsSync("tmp"))
    fs.mkdirSync("tmp");

  console.log(" --- Get map imagery tile:");
  let result = await render.getMapImageryTile("png", "satellite", 6, 10, 22, operationOptions);
  // use result.blobBody for Browser, readableStreamBody for Node.js:
  result.readableStreamBody?.pipe(fs.createWriteStream("tmp/map_imagery_tile.png"));

  const statesetId = process.env.CREATOR_STATESET_ID;
  if ((typeof statesetId === "string") && (statesetId.length == 36)) {
    console.log(" --- Get map state tile:");
    result = await render.getMapStateTilePreview(6, 10, 22, statesetId, operationOptions);
    // use result.blobBody for Browser, readableStreamBody for Node.js:
    result.readableStreamBody?.pipe(fs.createWriteStream("tmp/state_tile.pbf"));
  }

  console.log(" --- Get map static image:");
  const mapStaticImageOptions = {
    "layer": "basic",
    "style": "dark",
    "zoom": 2,
    "bbox": "1.355233,42.982261,24.980233,56.526017"
  };
  result = await render.getMapStaticImage("png", { ...mapStaticImageOptions, ...operationOptions });
  // use result.blobBody for Browser, readableStreamBody for Node.js:
  result.readableStreamBody?.pipe(fs.createWriteStream("tmp/static_image.png"));

  console.log(" --- Get map tile:");
  const mapTileOptions = { "tileSize": "512" };
  result = await render.getMapTile("png", "basic", "main", 6, 10, 22, { ...mapTileOptions, ...operationOptions });
  // use result.blobBody for Browser, readableStreamBody for Node.js:
  result.readableStreamBody?.pipe(fs.createWriteStream("tmp/tile.png"));

  console.log(" --- Get map tile v2:");
  result = await renderV2.getMapTilePreview("microsoft.base", 6, 10, 22, { ...mapTileOptions, ...operationOptions });
  // use result.blobBody for Browser, readableStreamBody for Node.js:
  result.readableStreamBody?.pipe(fs.createWriteStream("tmp/tile_v2.vector.pbf"));

}

main();
