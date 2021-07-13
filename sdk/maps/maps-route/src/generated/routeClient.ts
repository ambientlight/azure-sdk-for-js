/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import * as coreHttp from "@azure/core-http";
import { RouteImpl } from "./operations";
import { Route } from "./operationsInterfaces";
import { RouteClientContext } from "./routeClientContext";
import { RouteClientOptionalParams } from "./models";

export class RouteClient extends RouteClientContext {
  /**
   * Initializes a new instance of the RouteClient class.
   * @param credentials Subscription credentials which uniquely identify client subscription.
   * @param options The parameter options
   */
  constructor(
    credentials: coreHttp.TokenCredential | coreHttp.ServiceClientCredentials,
    options?: RouteClientOptionalParams
  ) {
    super(credentials, options);
    this.route = new RouteImpl(this);
  }

  route: Route;
}
