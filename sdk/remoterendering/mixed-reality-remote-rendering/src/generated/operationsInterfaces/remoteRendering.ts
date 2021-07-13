/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import {
  CreateConversionSettings,
  RemoteRenderingCreateConversionOptionalParams,
  RemoteRenderingCreateConversionResponse,
  RemoteRenderingGetConversionOptionalParams,
  RemoteRenderingGetConversionResponse,
  RemoteRenderingListConversionsOptionalParams,
  RemoteRenderingListConversionsResponse,
  RenderingSessionSettings,
  RemoteRenderingCreateSessionOptionalParams,
  RemoteRenderingCreateSessionResponse,
  RemoteRenderingGetSessionOptionalParams,
  RemoteRenderingGetSessionResponse,
  UpdateSessionSettings,
  RemoteRenderingUpdateSessionOptionalParams,
  RemoteRenderingUpdateSessionResponse,
  RemoteRenderingStopSessionOptionalParams,
  RemoteRenderingStopSessionResponse,
  RemoteRenderingListSessionsOptionalParams,
  RemoteRenderingListSessionsResponse,
  RemoteRenderingListConversionsNextOptionalParams,
  RemoteRenderingListConversionsNextResponse,
  RemoteRenderingListSessionsNextOptionalParams,
  RemoteRenderingListSessionsNextResponse
} from "../models";

/** Interface representing a RemoteRendering. */
export interface RemoteRendering {
  /**
   * Creates a conversion using an asset stored in an Azure Blob Storage account.
   * @param accountId The Azure Remote Rendering account ID.
   * @param conversionId An ID uniquely identifying the conversion for the given account. The ID is case
   *                     sensitive, can contain any combination of alphanumeric characters including hyphens and underscores,
   *                     and cannot contain more than 256 characters.
   * @param body Request body configuring the settings for an asset conversion.
   * @param options The options parameters.
   */
  createConversion(
    accountId: string,
    conversionId: string,
    body: CreateConversionSettings,
    options?: RemoteRenderingCreateConversionOptionalParams
  ): Promise<RemoteRenderingCreateConversionResponse>;
  /**
   * Gets the status of a particular conversion.
   * @param accountId The Azure Remote Rendering account ID.
   * @param conversionId An ID uniquely identifying the conversion for the given account. The ID is case
   *                     sensitive, can contain any combination of alphanumeric characters including hyphens and underscores,
   *                     and cannot contain more than 256 characters.
   * @param options The options parameters.
   */
  getConversion(
    accountId: string,
    conversionId: string,
    options?: RemoteRenderingGetConversionOptionalParams
  ): Promise<RemoteRenderingGetConversionResponse>;
  /**
   * Gets a list of all conversions.
   * @param accountId The Azure Remote Rendering account ID.
   * @param options The options parameters.
   */
  listConversions(
    accountId: string,
    options?: RemoteRenderingListConversionsOptionalParams
  ): Promise<RemoteRenderingListConversionsResponse>;
  /**
   * Creates a new rendering session.
   * @param accountId The Azure Remote Rendering account ID.
   * @param sessionId An ID uniquely identifying the rendering session for the given account. The ID is
   *                  case sensitive, can contain any combination of alphanumeric characters including hyphens and
   *                  underscores, and cannot contain more than 256 characters.
   * @param body Settings of the session to be created.
   * @param options The options parameters.
   */
  createSession(
    accountId: string,
    sessionId: string,
    body: RenderingSessionSettings,
    options?: RemoteRenderingCreateSessionOptionalParams
  ): Promise<RemoteRenderingCreateSessionResponse>;
  /**
   * Gets the properties of a particular rendering session.
   * @param accountId The Azure Remote Rendering account ID.
   * @param sessionId An ID uniquely identifying the rendering session for the given account. The ID is
   *                  case sensitive, can contain any combination of alphanumeric characters including hyphens and
   *                  underscores, and cannot contain more than 256 characters.
   * @param options The options parameters.
   */
  getSession(
    accountId: string,
    sessionId: string,
    options?: RemoteRenderingGetSessionOptionalParams
  ): Promise<RemoteRenderingGetSessionResponse>;
  /**
   * Updates the max lease time of a particular rendering session.
   * @param accountId The Azure Remote Rendering account ID.
   * @param sessionId An ID uniquely identifying the rendering session for the given account. The ID is
   *                  case sensitive, can contain any combination of alphanumeric characters including hyphens and
   *                  underscores, and cannot contain more than 256 characters.
   * @param body Settings used to update the session.
   * @param options The options parameters.
   */
  updateSession(
    accountId: string,
    sessionId: string,
    body: UpdateSessionSettings,
    options?: RemoteRenderingUpdateSessionOptionalParams
  ): Promise<RemoteRenderingUpdateSessionResponse>;
  /**
   * Stops a particular rendering session.
   * @param accountId The Azure Remote Rendering account ID.
   * @param sessionId An ID uniquely identifying the rendering session for the given account. The ID is
   *                  case sensitive, can contain any combination of alphanumeric characters including hyphens and
   *                  underscores, and cannot contain more than 256 characters.
   * @param options The options parameters.
   */
  stopSession(
    accountId: string,
    sessionId: string,
    options?: RemoteRenderingStopSessionOptionalParams
  ): Promise<RemoteRenderingStopSessionResponse>;
  /**
   * Gets a list of all rendering sessions.
   * @param accountId The Azure Remote Rendering account ID.
   * @param options The options parameters.
   */
  listSessions(
    accountId: string,
    options?: RemoteRenderingListSessionsOptionalParams
  ): Promise<RemoteRenderingListSessionsResponse>;
  /**
   * ListConversionsNext
   * @param accountId The Azure Remote Rendering account ID.
   * @param nextLink The nextLink from the previous successful call to the ListConversions method.
   * @param options The options parameters.
   */
  listConversionsNext(
    accountId: string,
    nextLink: string,
    options?: RemoteRenderingListConversionsNextOptionalParams
  ): Promise<RemoteRenderingListConversionsNextResponse>;
  /**
   * ListSessionsNext
   * @param accountId The Azure Remote Rendering account ID.
   * @param nextLink The nextLink from the previous successful call to the ListSessions method.
   * @param options The options parameters.
   */
  listSessionsNext(
    accountId: string,
    nextLink: string,
    options?: RemoteRenderingListSessionsNextOptionalParams
  ): Promise<RemoteRenderingListSessionsNextResponse>;
}
