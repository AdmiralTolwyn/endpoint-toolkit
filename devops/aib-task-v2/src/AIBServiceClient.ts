"use strict";
// v1-patched: import path changed from azure-arm-rest-v2 to azure-arm-rest (v3)
import { ServiceClient } from 'azure-pipelines-tasks-azure-arm-rest/AzureServiceClient';
import webClient = require("azure-pipelines-tasks-azure-arm-rest/webClient");


// AIBServiceClient extends ServiceClient from azure-arm-rest, with some modifications
export default class AIBServiceClient extends ServiceClient {
    // getLongRunningOperationResult extends the same-name function in azure-arm-rest, with retry on http timeout
    public async getLongRunningOperationResult(response: webClient.WebResponse, timeoutInMinutes?: number): Promise<webClient.WebResponse> {
        timeoutInMinutes = timeoutInMinutes || this.longRunningOperationRetryTimeout;
        var timeout = new Date().getTime() + timeoutInMinutes * 60 * 1000;
        var waitIndefinitely = timeoutInMinutes == 0;
        while (true) {
            try {
                // when http timeout, this will directly return and throw exception
                var operationResponse = await super.getLongRunningOperationResult(response, timeoutInMinutes);
            }
            catch(exception) {
                let exceptionString: string = exception.toString();

                // if http timeout and need retry
                if (exceptionString.indexOf('Request timeout') != -1 && (waitIndefinitely || timeout > new Date().getTime())) {
                    console.log('Encountered request timeout issue. Will retry');
                    continue;
                }

                throw exception;
            }

            return operationResponse;
        }
    }
}
