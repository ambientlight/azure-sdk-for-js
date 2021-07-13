// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Search API usage. Simple queries are performed.
 */

import fs from "fs";
import { getDefaultAzureCredential } from "@azure/identity";
import * as coreAuth from "@azure/core-auth";
import { OperationOptions } from '@azure/core-http';
import { SearchClient } from "@azure/maps-search";
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

  const search = new SearchClient(credential).search;

  const filePathForPostSearchAddressBatch = "../../resources/search_address_batch_request_body.json";
  const filePathForPostSearchAddressReverseBatch = "../../resources/search_address_reverse_batch_request_body.json"
  const filePathForPostSearchFuzzyBatch = "../../resources/search_fuzzy_batch_request_body.json";
  const filePathForPostSearchAlongRoute = "../../resources/search_along_route_request_body.json";
  const filePathForPostSearchInsideGeometry = "../../resources/search_inside_geometry_request_body.json";

  console.log(" --- Get search address:");
  console.log(await search.getSearchAddress("json", "15127 NE 24th Street, Redmond, WA 98052", operationOptions));

  console.log(" --- Get search address reverse:");
  console.log(await search.getSearchAddressReverse("json", "37.337,-121.89", operationOptions));

  console.log(" --- Get search address reverse cross street:");
  console.log(await search.getSearchAddressReverseCrossStreet("json", "37.337,-121.89", operationOptions));

  console.log(" --- Get search address structured:");
  const searchAddressStructuredOptions = {
    "countryCode": "US",
    "streetNumber": "15127",
    "streetName": "NE 24th Street",
    "municipality": "Redmond",
    "countrySubdivision": "WA",
    "postalCode": "98052"
  };
  console.log(await search.getSearchAddressStructured("json", { ...searchAddressStructuredOptions, ...operationOptions }));

  console.log(" --- Get search fuzzy:");
  const fuzzyResult = await search.getSearchFuzzy("json", "Seattle", operationOptions);
  console.log(fuzzyResult);

  // let's save geometry IDs from the fuzzy search for the getSearchPolygon example
  let geometries: string[] = [];
  fuzzyResult.results?.forEach((res) => geometries.push(res.dataSources?.geometry?.id!));

  console.log(" --- Get search nearby:");
  const searchNearbyOptions = { "radius": 8046 };
  console.log(await search.getSearchNearby("json", 40.706270, -74.011454, { ...searchNearbyOptions, ...operationOptions }));

  console.log(" --- Get search POI:");
  const searchPOIOptions = {
    "limit": 5,
    "lat": 47.606038,
    "lon": -122.333345,
    "radius": 8046
  };
  console.log(await search.getSearchPOI("json", "juice bars", { ...searchPOIOptions, ...operationOptions }));

  console.log(" --- Get search POI category:");
  const searchPOICategoryOptions = {
    "limit": 5,
    "lat": 47.606038,
    "lon": -122.333345,
    "radius": 8046
  };
  console.log(await search.getSearchPOICategory("json", "atm", operationOptions));

  console.log(" --- Get search POI category tree:");
  console.log(await search.getSearchPOICategoryTreePreview("json", operationOptions));

  console.log(" --- Get search polygon:");
  console.log(await search.getSearchPolygon("json", geometries, operationOptions));

  console.log(" --- Post search address batch:");
  const postSearchAddressBatchPayload = JSON.parse(fs.readFileSync(filePathForPostSearchAddressBatch, "utf8"));
  console.log(await search.beginPostSearchAddressBatchAndWait("json", postSearchAddressBatchPayload, operationOptions));

  console.log(" --- Post search address reverse batch:");
  const postSearchAddressReverseBatchPayload = JSON.parse(fs.readFileSync(filePathForPostSearchAddressReverseBatch, "utf8"));
  console.log(await search.beginPostSearchAddressReverseBatchAndWait("json", postSearchAddressReverseBatchPayload, operationOptions));

  console.log(" --- Post search fuzzy batch:");
  const postSearchFuzzyBatchPayload = JSON.parse(fs.readFileSync(filePathForPostSearchFuzzyBatch, "utf8"));
  console.log(await search.beginPostSearchFuzzyBatchAndWait("json", postSearchFuzzyBatchPayload, operationOptions));

  console.log(" --- Post search along route:");
  const searchAlongRouteOptions = { "limit": 2 };
  const postSearchAlongRoutePayload = JSON.parse(fs.readFileSync(filePathForPostSearchAlongRoute, "utf8"));
  console.log(await search.postSearchAlongRoute("json", "burger", 1000, postSearchAlongRoutePayload, { ...searchAlongRouteOptions, ...operationOptions }));

  console.log(" --- Post search inside geometry:");
  const searchInsideGeometryOptions = { "limit": 2 };
  const postSearchInsideGeometryPayload = JSON.parse(fs.readFileSync(filePathForPostSearchInsideGeometry, "utf8"));
  console.log(await search.postSearchInsideGeometry("json", "burger", postSearchInsideGeometryPayload, { ...searchInsideGeometryOptions, ...operationOptions }));

}

main();
