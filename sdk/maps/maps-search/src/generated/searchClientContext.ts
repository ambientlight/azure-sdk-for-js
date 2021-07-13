/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import * as coreHttp from "@azure/core-http";
import { Geography, SearchClientOptionalParams } from "./models";

const packageName = "@azure/maps-search";
const packageVersion = "1.0.0";

export class SearchClientContext extends coreHttp.ServiceClient {
  geography: Geography;
  xMsClientId?: string;
  apiVersion: string;

  /**
   * Initializes a new instance of the SearchClientContext class.
   * @param credentials Subscription credentials which uniquely identify client subscription.
   * @param options The parameter options
   */
  constructor(
    credentials: coreHttp.TokenCredential | coreHttp.ServiceClientCredentials,
    options?: SearchClientOptionalParams
  ) {
    if (credentials === undefined) {
      throw new Error("'credentials' cannot be null");
    }

    // Initializing default values for options
    if (!options) {
      options = {};
    }

    if (!options.userAgent) {
      const defaultUserAgent = coreHttp.getDefaultUserAgentValue();
      options.userAgent = `${packageName}/${packageVersion} ${defaultUserAgent}`;
    }

    if (!options.credentialScopes) {
      options.credentialScopes = ["https://atlas.microsoft.com/.default"];
    }

    super(credentials, options);

    this.requestContentType = "application/json; charset=utf-8";

    this.baseUri =
      options.endpoint || "https://{geography}.atlas.microsoft.com";

    // Assigning values to Constant parameters
    this.geography = options.geography || "us";
    this.apiVersion = options.apiVersion || "1.0";
  }
}
