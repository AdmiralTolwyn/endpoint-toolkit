"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// v1-patched: import path changed from azure-arm-rest-v2 to azure-arm-rest (v3)
const AzureServiceClientBase_1 = require("azure-pipelines-tasks-azure-arm-rest/AzureServiceClientBase");
const webClient = require("azure-pipelines-tasks-azure-arm-rest/webClient");
const AIBServiceClient_1 = __importDefault(require("./AIBServiceClient"));
const constants_1 = require("./constants");
const Utils_1 = __importDefault(require("./Utils"));
class ImageBuilderClient {
    _client;
    _subscriptionId;
    _taskParameters;
    constructor(credentials, subscriptionId, taskParameters) {
        credentials.baseUrl = "https://management.azure.com";
        this._client = new AIBServiceClient_1.default(credentials, subscriptionId);
        this._taskParameters = taskParameters;
        this._subscriptionId = subscriptionId;
    }
    async getTemplateId(templateName) {
        var httpRequest = new webClient.WebRequest();
        httpRequest.method = 'GET';
        httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName }, [], constants_1.AIB_API_VERSION);
        var resourceId = "";
        try {
            var response = await this._client.beginRequest(httpRequest);
            if (response.statusCode != 200 || response.body.status == "Failed")
                throw (0, AzureServiceClientBase_1.ToError)(response);
            if (response.statusCode == 200 && response.body.id)
                resourceId = response.body.id;
        }
        catch (error) {
            throw Error(`get template call failed for template ${templateName} with error: ${this._client.getFormattedError(error)}`);
        }
        return resourceId;
    }
    async putImageTemplate(template, templateName) {
        console.log("starting put template...");
        var httpRequest = new webClient.WebRequest();
        httpRequest.method = 'PUT';
        httpRequest.body = template;
        httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName }, [], constants_1.AIB_API_VERSION);
        try {
            var response = await this._client.beginRequest(httpRequest);
            if (response.statusCode == 201) {
                response = await this._client.getLongRunningOperationResult(response);
            }
            if (response.statusCode != 200 || response.body.status == "Failed") {
                throw (0, AzureServiceClientBase_1.ToError)(response);
            }
            if (response.statusCode == 200 && response.body && response.body.status == "Succeeded") {
                console.log("put template: ", response.body.status);
            }
        }
        catch (error) {
            throw Error(`put template call failed for template ${templateName} with error: ${this._client.getFormattedError(error)}`);
        }
    }
    // v1-patched: runTemplate with progress polling (replaces blind getLongRunningOperationResult)
    async runTemplate(templateName, buildStart) {
        try {
            console.log("starting run template...");
            var httpRequest = new webClient.WebRequest();
            httpRequest.method = 'POST';
            httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}/run`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName }, [], constants_1.AIB_API_VERSION);
            var response = await this._client.beginRequest(httpRequest);
            if (response.statusCode == 202) {
                // v1-patched: use custom progress polling instead of getLongRunningOperationResult
                var asyncUrl = response.headers["azure-asyncoperation"] || response.headers["location"];
                if (!asyncUrl)
                    throw new Error("No async operation URL returned from POST /run");
                await this._pollBuildWithProgress(asyncUrl, templateName, buildStart || Date.now());
            }
            else if (response.statusCode != 200 || (response.body && response.body.status == "Failed")) {
                throw (0, AzureServiceClientBase_1.ToError)(response);
            }
            else if (response.statusCode == 200 && response.body && response.body.status == "Succeeded") {
                console.log("run template: ", response.body.status);
            }
        }
        catch (error) {
            throw Error(`post template call failed for template ${templateName} with error: ${this._client.getFormattedError(error)}`);
        }
    }
    // v1-patched: custom polling loop with phase transitions, elapsed time, and heartbeats
    async _pollBuildWithProgress(asyncUrl, templateName, buildStart) {
        var POLL_INTERVAL_SECS = 30;
        var STATUS_CHECK_INTERVAL = 2; // check template status every Nth poll
        var lastPhase = "";
        var pollCount = 0;
        var phaseTimestamps = [];
        var formatElapsed = (fromMs) => {
            var totalSec = Math.floor((Date.now() - fromMs) / 1000);
            var h = Math.floor(totalSec / 3600);
            var m = Math.floor((totalSec % 3600) / 60);
            var s = totalSec % 60;
            if (h > 0)
                return `${h}h ${m}m ${s}s`;
            if (m > 0)
                return `${m}m ${s}s`;
            return `${s}s`;
        };
        var printPhaseChange = (newPhase) => {
            if (newPhase && newPhase !== lastPhase) {
                var arrow = lastPhase ? ` → ${newPhase}` : newPhase;
                var prefix = lastPhase ? "  ◆" : "  ▶";
                console.log(`${prefix} Phase: ${arrow}  [${formatElapsed(buildStart)}]`);
                phaseTimestamps.push({ phase: newPhase, time: Date.now() });
                lastPhase = newPhase;
            }
        };
        // Initial phase
        printPhaseChange("Queued");
        while (true) {
            await Utils_1.default.sleep(POLL_INTERVAL_SECS * 1000);
            pollCount++;
            // Check async operation status
            var opRequest = new webClient.WebRequest();
            opRequest.method = "GET";
            opRequest.uri = asyncUrl;
            var opResponse;
            try {
                opResponse = await this._client.beginRequest(opRequest);
            }
            catch (error) {
                var errStr = String(error);
                if (errStr.indexOf('Request timeout') != -1) {
                    console.log(`  ⏳ Poll timeout — retrying...  [${formatElapsed(buildStart)}]`);
                    continue;
                }
                throw error;
            }
            var opStatus = opResponse.body && opResponse.body.status;
            // Check template status periodically for richer phase info
            if (pollCount % STATUS_CHECK_INTERVAL === 0) {
                try {
                    var status = await this._getLastRunStatus(templateName);
                    if (status) {
                        printPhaseChange(status.runSubState || status.runState);
                        if (status.message && status.message !== lastPhase) {
                            var msg = status.message;
                            if (msg.length > 0 && msg.length < 500 && msg.indexOf("InProgress") === -1)
                                console.log(`  ℹ ${msg}`);
                        }
                    }
                }
                catch (e) {
                    // Template status check is best-effort; don't fail the build
                }
            }
            // Async operation final states
            if (opStatus === "Succeeded") {
                printPhaseChange("Completed");
                console.log(`  ✔ Build succeeded  [${formatElapsed(buildStart)}]`);
                // Print phase timeline
                if (phaseTimestamps.length > 1) {
                    console.log("");
                    console.log("  Phase Timeline:");
                    for (var i = 0; i < phaseTimestamps.length; i++) {
                        var duration = i < phaseTimestamps.length - 1
                            ? Math.floor((phaseTimestamps[i + 1].time - phaseTimestamps[i].time) / 1000)
                            : 0;
                        var durationStr = duration > 0
                            ? ` (${Math.floor(duration / 60)}m ${duration % 60}s)`
                            : "";
                        console.log(`    ${i + 1}. ${phaseTimestamps[i].phase}${durationStr}`);
                    }
                    console.log("");
                }
                return;
            }
            if (opStatus === "Failed" || opStatus === "Canceled") {
                // Fetch final template status for error details
                try {
                    var status = await this._getLastRunStatus(templateName);
                    if (status && status.message)
                        console.log(`  ✘ Build ${opStatus.toLowerCase()}: ${status.message}`);
                }
                catch (e) {
                    // ignore
                }
                throw new Error(`Image build ${opStatus.toLowerCase()}. Check Azure portal for details.`);
            }
            // Still in progress — print heartbeat every 2 minutes
            if (pollCount % 4 === 0) {
                var phase = lastPhase || "Running";
                console.log(`  ⟳ ${phase} ...  [${formatElapsed(buildStart)}]`);
            }
        }
    }
    // v1-patched: fetch template's lastRunStatus for phase info
    async _getLastRunStatus(templateName) {
        var httpRequest = new webClient.WebRequest();
        httpRequest.method = 'GET';
        httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName }, [], constants_1.AIB_API_VERSION);
        var response = await this._client.beginRequest(httpRequest);
        if (response.statusCode === 200 && response.body && response.body.properties && response.body.properties.lastRunStatus)
            return response.body.properties.lastRunStatus;
        return null;
    }
    async deleteTemplate(templateName) {
        try {
            console.log(`deleting template ${templateName}...`);
            var httpRequest = new webClient.WebRequest();
            httpRequest.method = 'DELETE';
            httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName }, [], constants_1.AIB_API_VERSION);
            var response = await this._client.beginRequest(httpRequest);
            if (response.statusCode == 202) {
                response = await this._client.getLongRunningOperationResult(response);
            }
            if (response.statusCode != 200 || response.body.status == "Failed") {
                throw (0, AzureServiceClientBase_1.ToError)(response);
            }
            if (response.statusCode == 200 && response.body && response.body.status == "Succeeded") {
                console.log("delete template: ", response.body.status);
            }
        }
        catch (error) {
            throw Error(`delete template call failed for template ${templateName} with error: ${this._client.getFormattedError(error)}`);
        }
    }
    async getRunOutput(templateName, runOutput) {
        console.log("getting runOutput for ", runOutput);
        var httpRequest = new webClient.WebRequest();
        httpRequest.method = 'GET';
        httpRequest.uri = this._client.getRequestUri(`/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/Microsoft.VirtualMachineImages/imagetemplates/{imageTemplateName}/runOutputs/{runOutput}`, { '{subscriptionId}': this._subscriptionId, '{resourceGroupName}': this._taskParameters.ibResourceGroup, '{imageTemplateName}': templateName, '{runOutput}': runOutput }, [], constants_1.AIB_API_VERSION);
        var output = "";
        try {
            var response = await this._client.beginRequest(httpRequest);
            if (response.statusCode != 200 || response.body.status == "Failed")
                throw (0, AzureServiceClientBase_1.ToError)(response);
            if (response.statusCode == 200 && response.body) {
                if (response.body && response.body.properties.artifactId)
                    output = response.body.properties.artifactId;
                else if (response.body && response.body.properties.artifactUri)
                    output = response.body.properties.artifactUri;
                else
                    console.log(`Error to parse response.body -- ${response.body}.`);
            }
        }
        catch (error) {
            throw Error(`get runOutput call failed for template ${templateName} for ${runOutput} with error: ${this._client.getFormattedError(error)}`);
        }
        return output;
    }
}
exports.default = ImageBuilderClient;
