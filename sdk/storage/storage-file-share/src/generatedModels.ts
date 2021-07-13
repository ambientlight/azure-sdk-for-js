// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// explicitly exporting types that we need.

export {
  CopyStatusType,
  DeleteSnapshotsOptionType,
  DirectoryCreateResponse,
  DirectoryDeleteResponse,
  DirectoryGetPropertiesResponse,
  DirectoryItem,
  DirectoryListFilesAndDirectoriesSegmentResponse,
  DirectoryListHandlesResponse,
  DirectorySetMetadataResponse,
  DirectorySetPropertiesResponse,
  FileAbortCopyResponse,
  FileCreateResponse,
  FileDeleteResponse,
  FileDownloadHeaders,
  FileDownloadOptionalParams,
  FileDownloadResponse as FileDownloadResponseModel,
  FileGetPropertiesResponse,
  FileGetRangeListHeaders,
  FileItem,
  FileListHandlesResponse,
  FileServiceProperties,
  FileSetMetadataResponse,
  FileStartCopyResponse,
  FileUploadRangeFromURLOptionalParams,
  FileUploadRangeFromURLResponse,
  FileUploadRangeResponse,
  HandleItem,
  ListSharesIncludeType,
  FileRange as RangeModel,
  ServiceGetPropertiesResponse,
  ServiceSetPropertiesResponse,
  ShareCreatePermissionResponse,
  ShareCreateResponse,
  ShareCreateSnapshotResponse,
  ShareDeleteResponse,
  ShareGetAccessPolicyHeaders,
  ShareGetPermissionResponse,
  ShareGetPropertiesResponse as ShareGetPropertiesResponseModel,
  ShareGetStatisticsResponse as ShareGetStatisticsResponseModel,
  ShareItemInternal,
  ShareSetAccessPolicyResponse,
  ShareSetMetadataResponse,
  ShareSetPropertiesResponse,
  SignedIdentifier as SignedIdentifierModel,
  SourceModifiedAccessConditions,
  FileForceCloseHandlesHeaders,
  DirectoryForceCloseHandlesHeaders,
  DirectoryCreateHeaders,
  DirectoryDeleteHeaders,
  DirectoryGetPropertiesHeaders,
  FilesAndDirectoriesListSegment,
  ListFilesAndDirectoriesSegmentResponse,
  DirectoryListFilesAndDirectoriesSegmentHeaders,
  ListHandlesResponse,
  DirectoryListHandlesHeaders,
  DirectorySetMetadataHeaders,
  DirectorySetPropertiesHeaders,
  FileAbortCopyHeaders,
  FileCreateHeaders,
  FileDeleteHeaders,
  FileGetPropertiesHeaders,
  FileProperty,
  FileListHandlesHeaders,
  CorsRule,
  Metrics,
  FileSetHttpHeadersHeaders as FileSetHTTPHeadersHeaders,
  FileSetHttpHeadersResponse as FileSetHTTPHeadersResponse,
  FileSetMetadataHeaders,
  FileStartCopyHeaders,
  FileUploadRangeFromURLHeaders,
  FileUploadRangeHeaders,
  FileType,
  ServiceGetPropertiesHeaders,
  ListSharesResponse as ListSharesResponseModel,
  RetentionPolicy,
  ServiceListSharesSegmentHeaders,
  ServiceSetPropertiesHeaders,
  ShareCreatePermissionHeaders,
  ShareCreateHeaders,
  ShareCreateSnapshotHeaders,
  ShareDeleteHeaders,
  SharePermission,
  ShareGetPermissionHeaders,
  ShareGetPropertiesHeaders,
  ShareStats,
  ShareGetStatisticsHeaders,
  SharePropertiesInternal,
  ShareSetAccessPolicyHeaders,
  ShareSetMetadataHeaders,
  ShareSetPropertiesHeaders,
  AccessPolicy,
  LeaseAccessConditions,
  LeaseDurationType,
  LeaseStateType,
  LeaseStatusType,
  CopyFileSmbInfo,
  ShareProtocolSettings,
  ShareSmbSettings,
  SmbMultichannel,
  FileGetRangeListResponse as FileGetRangeListDiffResponse,
  ShareFileRangeList,
  ClearRange,
  ShareAccessTier,
  ShareRootSquash
} from "./generated/src/models";

import { ShareSetPropertiesResponse, ShareSetPropertiesHeaders } from "./generated/src/models";

/**
 * Contains response data for the setQuota operation.
 */
export type ShareSetQuotaResponse = ShareSetPropertiesResponse;

/**
 * Defines headers for setQuota operation.
 */
export type ShareSetQuotaHeaders = ShareSetPropertiesHeaders;
