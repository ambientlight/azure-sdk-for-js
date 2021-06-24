// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { FullOperationResponse } from "@azure/core-client";
import { failureStates, LROState, successStates } from "./stateMachine";

function getProvisioningState(rawResponse: FullOperationResponse): string {
  const { properties, provisioningState } =
    rawResponse.parsedBody ??
    (rawResponse.bodyAsText ? JSON.parse(rawResponse.bodyAsText) : {});
  const state: string | undefined =
    properties?.provisioningState ?? provisioningState;
  return state?.toLowerCase() ?? "succeeded";
}

export function isBodyPollingDone(rawResponse: FullOperationResponse) {
  const state = getProvisioningState(rawResponse);
  if (failureStates.includes(state)) {
    throw new Error(`Provisioning state: ${state}`);
  }
  return successStates.includes(state);
}

/**
 * Creates a polling strategy based on BodyPolling which uses the provisioning state
 * from the result to determine the current operation state
 */
export function processBodyPollingOperationResult<TResult>(
  rawResponse: FullOperationResponse,
  flatResponse: TResult
): LROState<TResult> {
  return {
    rawResponse,
    flatResponse,
    done: isBodyPollingDone(rawResponse)
  };
}
