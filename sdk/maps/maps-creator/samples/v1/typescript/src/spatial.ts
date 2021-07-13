// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Spatial API usage. Simple queries are performed.
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

  const spatial = new CreatorClient(credential).spatial;

  const filePathForPostSpatialBuffer = "../../resources/spatial_buffer_request_body.json";
  const filePathForPostSpatialClosestPoint = "../../resources/spatial_closest_point_request_body.json";
  const filePathForPostSpatialGeofence = "../../resources/spatial_geofence_request_body.json";
  const filePathForPostSpatialPointInPolygon = "../../resources/spatial_point_in_polygon_request_body.json";

  // TO USE need to have some GeoJson Data uploaded already - please use env CREATOR_GEOJSON_UDID
  const udid = process.env.CREATOR_GEOJSON_UDID;
  if ((typeof udid !== "string") || (udid.length != 36)) {
    throw "This sample needs some GeoJson Data uploaded";
  }

  console.log(" --- Get buffer:");
  const res = await spatial.getBuffer("json", udid, "176.3", operationOptions);
  console.log(res);
  console.log(res.result?.features);

  console.log(" --- Get closest point:");
  console.log(await spatial.getClosestPoint("json", udid, 47.622942, -122.316456, operationOptions));

  console.log(" --- Get geofence:");
  const spatialGeofenceParams = {
    "searchBuffer": 50,
    "userTime": new Date("2017-07-21T17:32:28Z"),
    "mode": "EnterAndExit"
  };
  console.log(await spatial.getGeofence("json", "unique_device_name_under_account", udid, 48.36, -124.63, { ...spatialGeofenceParams, ...operationOptions }));

  console.log(" --- Get great circle distance:");
  console.log(await spatial.getGreatCircleDistance("json", "47.622942,-122.316456:47.610378,-122.200676", operationOptions));

  /* TODO: use udid with some Polygon as a "geometry" of the Feature
    console.log(" --- Get point in polygon:");
    console.log(await spatial.getPointInPolygon("json", udid, 47.622942, -122.316456, operationOptions));*/

  console.log(" --- Post buffer:");
  const postSpatialBufferPayload = JSON.parse(fs.readFileSync(filePathForPostSpatialBuffer, "utf8"));
  console.log(await spatial.postBuffer("json", postSpatialBufferPayload, operationOptions));

  console.log(" --- Post closest point:");
  const postSpatialClosestPointPayload = JSON.parse(fs.readFileSync(filePathForPostSpatialClosestPoint, "utf8"));
  console.log(await spatial.postClosestPoint("json", 47.622942, -122.316456, postSpatialClosestPointPayload, operationOptions));

  console.log(" --- Post geofence:");
  const postSpatialGeofencePayload = JSON.parse(fs.readFileSync(filePathForPostSpatialGeofence, "utf8"));
  console.log(await spatial.postGeofence("json", "unique_device_name_under_account", 48.36, -124.63, postSpatialGeofencePayload, { ...spatialGeofenceParams, ...operationOptions }));

  console.log(" --- Post point in polygon:");
  const postSpatialPointInPolygonPayload = JSON.parse(fs.readFileSync(filePathForPostSpatialPointInPolygon, "utf8"));
  console.log(await spatial.postPointInPolygon("json", 48.36, -124.63, postSpatialPointInPolygonPayload, operationOptions));

}

main();
