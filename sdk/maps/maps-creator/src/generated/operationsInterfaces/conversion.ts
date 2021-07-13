/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import "@azure/core-paging";
import { PagedAsyncIterableIterator } from "@azure/core-paging";
import * as coreHttp from "@azure/core-http";
import { PollerLike, PollOperationState } from "@azure/core-lro";
import {
  ConversionListDetailInfo,
  ConversionListNextOptionalParams,
  ConversionListOptionalParams,
  ConversionConvertOptionalParams,
  ConversionConvertResponse,
  ConversionGetOptionalParams,
  ConversionGetResponse,
  ConversionDeleteOptionalParams,
  ConversionGetOperationOptionalParams,
  ConversionGetOperationResponse
} from "../models";

/// <reference lib="esnext.asynciterable" />
/** Interface representing a Conversion. */
export interface Conversion {
  /**
   * **Applies to:** see pricing [tiers](https://aka.ms/AzureMapsPricingTier).
   *
   * Creator makes it possible to develop applications based on your private indoor map data using Azure
   * Maps API and SDK. [This](https://docs.microsoft.com/azure/azure-maps/creator-indoor-maps) article
   * introduces concepts and tools that apply to Azure Maps Creator.
   *
   * This API allows the caller to fetch a list of all successful data conversions submitted previously
   * using the [Conversion
   * API](https://docs.microsoft.com/en-us/rest/api/maps/conversion/convertpreview).
   *
   * ### Submit List Request
   *
   * To list all successful conversions you will issue a `GET` request with no additional parameters.
   *
   * ### List Data Response
   *
   * The Conversion List API returns the complete list of all conversion details in `json` format.<br>
   *
   * Here is a sample response returning the details of two successful conversion requests:
   *
   * <br>
   *
   * ```json
   * {
   *     "conversions":
   *     [
   *         {
   *             "conversionId": "54398242-ea6c-1f31-4fa6-79b1ae0fc24d",
   *             "udid": "31838736-8b84-11ea-bc55-0242ac130003",
   *             "created": "5/19/2020 9:00:00 AM +00:00",
   *             "description": "User provided description.",
   *             "featureCounts": {
   *                 "DIR": 1,
   *                 "LVL": 3,
   *                 "FCL": 1,
   *                 "UNIT": 150,
   *                 "CTG": 8,
   *                 "AEL": 0,
   *                 "OPN": 10
   *             }
   *         },
   *         {
   *             "conversionId": "2acf7d32-8b84-11ea-bc55-0242ac130003",
   *             "udid": "1214bc58-8b84-11ea-bc55-0242ac1300039",
   *             "created": "5/19/2020 9:00:00 AM +00:00",
   *             "description": "User provided description.",
   *             "featureCounts": {
   *                 "DIR": 1,
   *                 "LVL": 3,
   *                 "FCL": 1,
   *                 "UNIT": 150,
   *                 "CTG": 8,
   *                 "AEL": 0,
   *                 "OPN": 10
   *             }
   *         }
   *     ]
   * }
   * ```
   *
   * <br>
   * @param options The options parameters.
   */
  list(
    options?: ConversionListOptionalParams
  ): PagedAsyncIterableIterator<ConversionListDetailInfo>;
  /**
   * **Applies to:** see pricing [tiers](https://aka.ms/AzureMapsPricingTier).
   *
   * Creator makes it possible to develop applications based on your private indoor map data using Azure
   * Maps API and SDK. [This](https://docs.microsoft.com/azure/azure-maps/creator-indoor-maps) article
   * introduces concepts and tools that apply to Azure Maps Creator.
   *
   * The Conversion API lets the caller import a set of DWG design files as a zipped [Drawing
   * Package](https://aka.ms/am-drawing-package) into Azure Maps. The [Drawing
   * Package](https://aka.ms/am-drawing-package) should first be uploaded using the [Azure Maps Data
   * Service](https://docs.microsoft.com/rest/api/maps/data). Once uploaded, use the `udid` returned by
   * the [Data Upload API](https://docs.microsoft.com/rest/api/maps/data/uploadpreview) to call this
   * Conversion API.
   *
   * ## Convert DWG package
   *
   * The Conversion API performs a [long-running request](https://aka.ms/am-creator-lrt-v2).
   *
   * ## Debug DWG package issues
   *
   *
   * During the Conversion process, if there are any issues with the DWG package [errors and
   * warnings](https://aka.ms/am-conversion-errors) are provided in the response along with a *diagnostic
   * package* to visualize and diagnose these issues. In case any issues are encountered with your DWG
   * package, the Conversion operation status process as detailed
   * [here](https://aka.ms/am-creator-lrt-v2) returns the location of the *diagnostic package* that can
   * be downloaded by the caller to help them visualize and diagnose these issues. The *diagnostic
   * package* location can be found in the properties section of the conversion operation status response
   * and looks like the following:
   *
   * ```json
   * {
   *     "properties": {
   *         "diagnosticPackageLocation":
   * "https://us.atlas.microsoft.com/mapdata/{DiagnosticPackageId}?api-version=1.0"
   *     }
   * }
   * ```
   *
   * The *diagnostic package* can be downloaded by executing a `HTTP GET` request on the
   * `diagnosticPackageLocation`.
   * For more details on how to use the tool to visualize and diagnose all the errors and warnings see
   * [Drawing Error Visualizer](https://aka.ms/am-drawing-errors-visualizer). <br>
   *
   * A conversion operation will be marked as *success* if there are zero or more warnings but will be
   * marked as *failed* if any errors are encountered.
   * @param udid The unique data id for the content. The `udid` must have been obtained from a successful
   *             [Data Upload API](https://docs.microsoft.com/en-us/rest/api/maps/data%20v2/uploadpreview) call.
   * @param outputOntology Output ontology version. "facility-2.0" is the only supported value at this
   *                       time. Please refer to this
   *                       [article](https://docs.microsoft.com/en-us/azure/azure-maps/creator-facility-ontology) for more
   *                       information about Azure Maps Creator ontologies.
   * @param options The options parameters.
   */
  beginConvert(
    udid: string,
    outputOntology: string,
    options?: ConversionConvertOptionalParams
  ): Promise<
    PollerLike<
      PollOperationState<ConversionConvertResponse>,
      ConversionConvertResponse
    >
  >;
  /**
   * **Applies to:** see pricing [tiers](https://aka.ms/AzureMapsPricingTier).
   *
   * Creator makes it possible to develop applications based on your private indoor map data using Azure
   * Maps API and SDK. [This](https://docs.microsoft.com/azure/azure-maps/creator-indoor-maps) article
   * introduces concepts and tools that apply to Azure Maps Creator.
   *
   * The Conversion API lets the caller import a set of DWG design files as a zipped [Drawing
   * Package](https://aka.ms/am-drawing-package) into Azure Maps. The [Drawing
   * Package](https://aka.ms/am-drawing-package) should first be uploaded using the [Azure Maps Data
   * Service](https://docs.microsoft.com/rest/api/maps/data). Once uploaded, use the `udid` returned by
   * the [Data Upload API](https://docs.microsoft.com/rest/api/maps/data/uploadpreview) to call this
   * Conversion API.
   *
   * ## Convert DWG package
   *
   * The Conversion API performs a [long-running request](https://aka.ms/am-creator-lrt-v2).
   *
   * ## Debug DWG package issues
   *
   *
   * During the Conversion process, if there are any issues with the DWG package [errors and
   * warnings](https://aka.ms/am-conversion-errors) are provided in the response along with a *diagnostic
   * package* to visualize and diagnose these issues. In case any issues are encountered with your DWG
   * package, the Conversion operation status process as detailed
   * [here](https://aka.ms/am-creator-lrt-v2) returns the location of the *diagnostic package* that can
   * be downloaded by the caller to help them visualize and diagnose these issues. The *diagnostic
   * package* location can be found in the properties section of the conversion operation status response
   * and looks like the following:
   *
   * ```json
   * {
   *     "properties": {
   *         "diagnosticPackageLocation":
   * "https://us.atlas.microsoft.com/mapdata/{DiagnosticPackageId}?api-version=1.0"
   *     }
   * }
   * ```
   *
   * The *diagnostic package* can be downloaded by executing a `HTTP GET` request on the
   * `diagnosticPackageLocation`.
   * For more details on how to use the tool to visualize and diagnose all the errors and warnings see
   * [Drawing Error Visualizer](https://aka.ms/am-drawing-errors-visualizer). <br>
   *
   * A conversion operation will be marked as *success* if there are zero or more warnings but will be
   * marked as *failed* if any errors are encountered.
   * @param udid The unique data id for the content. The `udid` must have been obtained from a successful
   *             [Data Upload API](https://docs.microsoft.com/en-us/rest/api/maps/data%20v2/uploadpreview) call.
   * @param outputOntology Output ontology version. "facility-2.0" is the only supported value at this
   *                       time. Please refer to this
   *                       [article](https://docs.microsoft.com/en-us/azure/azure-maps/creator-facility-ontology) for more
   *                       information about Azure Maps Creator ontologies.
   * @param options The options parameters.
   */
  beginConvertAndWait(
    udid: string,
    outputOntology: string,
    options?: ConversionConvertOptionalParams
  ): Promise<ConversionConvertResponse>;
  /**
   * **Applies to:** see pricing [tiers](https://aka.ms/AzureMapsPricingTier).
   *
   * Creator makes it possible to develop applications based on your private indoor map data using Azure
   * Maps API and SDK. [This](https://docs.microsoft.com/azure/azure-maps/creator-indoor-maps) article
   * introduces concepts and tools that apply to Azure Maps Creator.
   *
   * This API allows the caller to fetch a successful data conversion submitted previously using the
   * [Conversion API](https://docs.microsoft.com/en-us/rest/api/maps/conversion/convertpreview).
   * @param conversionId The conversion id for the content. The `conversionId` must have been obtained
   *                     from a successful [Conversion
   *                     API](https://docs.microsoft.com/en-us/rest/api/maps/v2/conversion/convert) call.
   * @param options The options parameters.
   */
  get(
    conversionId: string,
    options?: ConversionGetOptionalParams
  ): Promise<ConversionGetResponse>;
  /**
   * **Applies to:** see pricing [tiers](https://aka.ms/AzureMapsPricingTier).
   *
   * Creator makes it possible to develop applications based on your private indoor map data using Azure
   * Maps API and SDK. [This](https://docs.microsoft.com/azure/azure-maps/creator-indoor-maps) article
   * introduces concepts and tools that apply to Azure Maps Creator.
   *
   * This API allows the caller to delete any data conversions created previously using the [Conversion
   * API](https://docs.microsoft.com/en-us/rest/api/maps/conversion/convertpreview).
   *
   * ### Submit Delete Request
   *
   * To delete your conversion data you will issue a `DELETE` request where the path will contain the
   * `conversionId` of the data to delete.
   *
   * ### Conversion Delete Response
   *
   * The Conversion Delete API returns a HTTP `204 No Content` response with an empty body, if the
   * converted data resources were deleted successfully.<br>
   * A HTTP `400 Bad Request` error response will be returned if no resource associated with the
   * passed-in `conversionId` is found.
   * @param conversionId The conversion id for the content. The `conversionId` must have been obtained
   *                     from a successful [Conversion
   *                     API](https://docs.microsoft.com/en-us/rest/api/maps/v2/conversion/convert) call.
   * @param options The options parameters.
   */
  delete(
    conversionId: string,
    options?: ConversionDeleteOptionalParams
  ): Promise<coreHttp.RestResponse>;
  /**
   * This path will be obtained from a call to POST /conversions.  While in progress, an http200 will be
   * returned with no extra headers -  followed by an http200 with Resource-Location header once
   * successfully completed.
   * @param operationId The ID to query the status for the dataset create/import request.
   * @param options The options parameters.
   */
  getOperation(
    operationId: string,
    options?: ConversionGetOperationOptionalParams
  ): Promise<ConversionGetOperationResponse>;
}
