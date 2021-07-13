// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * @summary Demonstrates Weather API usage. Simple queries are performed.
 */

import { getDefaultAzureCredential } from "@azure/identity";
import * as coreAuth from "@azure/core-auth";
import { OperationOptions } from '@azure/core-http';
import { WeatherClient } from "@azure/maps-weather";
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

  const weather = new WeatherClient(credential).weather;

  console.log(" --- Get current weather conditions:");
  console.log(await weather.getCurrentConditions("json", "47.641268,-122.125679", operationOptions));

  console.log(" --- Get daily forecast:");
  const dailyForecastOptions = { "duration": 5 };
  console.log(await weather.getDailyForecast("json", "62.6490341,30.0734812", { ...dailyForecastOptions, ...operationOptions }));

  console.log(" --- Get daily indices:");
  const dailyIndicesOptions = { "indexGroupId": 11 };
  console.log(await weather.getDailyIndices("json", "43.84745,-79.37849", { ...dailyIndicesOptions, ...operationOptions }));

  console.log(" --- Get hourly forecast:");
  const hourlyForecastOptions = { "duration": 12 };
  console.log(await weather.getHourlyForecast("json", "47.632346,-122.138874", { ...hourlyForecastOptions, ...operationOptions }));

  console.log(" --- Get minute forecast:");
  const minuteForecastOptions = { "interval": 15 };
  console.log(await weather.getMinuteForecast("json", "47.632346,-122.138874", { ...minuteForecastOptions, ...operationOptions }));

  console.log(" --- Get quarter day forecast:");
  const quarterDayForecastOptions = { "duration": 1 };
  console.log(await weather.getQuarterDayForecast("json", "47.632346,-122.138874", { ...quarterDayForecastOptions, ...operationOptions }));

  console.log(" --- Get severe weather alerts:");
  console.log(await weather.getSevereWeatherAlerts("json", "48.057,-81.091", operationOptions));

  console.log(" --- Get weather along route:");
  console.log(await weather.getWeatherAlongRoute("json", "38.907,-77.037,0:38.907,-77.009,10:38.926,-76.928,20:39.033,-76.852,30:39.168,-76.732,40:39.269,-76.634,50:39.287,-76.612,60", operationOptions));

}

main();
