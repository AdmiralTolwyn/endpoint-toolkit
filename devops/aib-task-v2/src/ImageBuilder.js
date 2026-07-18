"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const path = require("path");
const tl = require("azure-pipelines-task-lib/task");
// v1-patched: import path changed from azure-arm-rest-v2 to azure-arm-rest (v3)
const armStorage = require("azure-pipelines-tasks-azure-arm-rest/azure-arm-storage");
const azure_arm_endpoint_1 = require("azure-pipelines-tasks-azure-arm-rest/azure-arm-endpoint");
const util = require("util");
const fs = require("fs");
// v1-patched: @azure/storage-blob replaces deprecated azure-storage SDK
const storage_blob_1 = require("@azure/storage-blob");
var archiver = require('archiver');
const Utils_1 = __importDefault(require("./Utils"));
const TaskParameters_1 = __importDefault(require("./TaskParameters"));
const BuildTemplate_1 = __importDefault(require("./BuildTemplate"));
const AzureImageBuilderClient_1 = __importDefault(require("./AzureImageBuilderClient"));
var containerName = 'imagebuilder-vststask';
class ImageBuilder {
    _taskParameters;
    _aibClient;
    _buildTemplate;
    // v1-patched: @azure/storage-blob types replace azure-storage's blobService
    _blobServiceClient;
    _storageCredential;
    constructor() {
        try {
            this._taskParameters = new TaskParameters_1.default();
            this._taskParameters.printSummary();
            this._buildTemplate = new BuildTemplate_1.default(this._taskParameters);
        }
        catch (error) {
            throw (`error happened while initializing Image builder: ${error}`);
        }
    }
    async execute() {
        // v1-patched: elapsed time tracking
        var buildStart = Date.now();
        var elapsed = () => {
            var s = Math.floor((Date.now() - buildStart) / 1000);
            var m = Math.floor(s / 60);
            var sec = s % 60;
            return m > 0 ? `${m}m ${sec}s` : `${sec}s`;
        };
        try {
            var creds = await this._getARMCredentials(this._taskParameters.ibSubscriptionId);
            const subscriptionId = tl.getEndpointDataParameter(this._taskParameters.ibSubscriptionId, "subscriptionId", false);
            this._aibClient = new AzureImageBuilderClient_1.default(creds, subscriptionId, this._taskParameters);
            await this._initBlobService();
            var blobName = this._taskParameters.buildFolder + "/" + this._getBlobsPrefixPath() + this._taskParameters.buildFolder + `_${Date.now()}`;
            if (Utils_1.default.IsEqual(this._taskParameters.provisioner, "powershell"))
                blobName = blobName + '.zip';
            else
                blobName = blobName + '.tar.gz';
            var blobUrl = await this._uploadPackage(containerName, blobName);
            console.log(`[${elapsed()}] Upload complete.`);
            var templateJson = this._buildTemplate.getTemplate(blobUrl);
            var templateName = this.getTemplateName();
            console.log("template name: ", templateName);
            var runOutputName = templateJson.properties.distribute[0].runOutputName;
            var isVhdDistribute = templateJson.properties.distribute[0].type == "VHD";
            var templateStr = JSON.stringify(templateJson);
            await this._aibClient.putImageTemplate(templateStr, templateName);
            console.log(`[${elapsed()}] Template deployed.`);
            await this._aibClient.runTemplate(templateName, buildStart);
            var out = await this._aibClient.getRunOutput(templateName, runOutputName);
            tl.setVariable('templateName', templateName);
            var id = await this._aibClient.getTemplateId(templateName);
            tl.setVariable('templateId', id);
            if (out) {
                tl.setVariable('imageUri', out);
            }
            if (Utils_1.default.IsEqual(templateJson.properties.source.type, "PlatformImage")) {
                tl.setVariable('pirPublisher', templateJson.properties.source.publisher);
                tl.setVariable('pirOffer', templateJson.properties.source.offer);
                tl.setVariable('pirSku', templateJson.properties.source.sku);
                tl.setVariable('pirVersion', templateJson.properties.source.version);
            }
            console.log("==============================================================================");
            console.log("## task output variables ##");
            console.log("$(imageUri) = ", tl.getVariable('imageUri'));
            if (isVhdDistribute) {
                console.log("$(templateName) = ", tl.getVariable('templateName'));
                console.log("$(templateId) = ", tl.getVariable('templateId'));
            }
            console.log("Total elapsed: %s", elapsed());
            console.log("==============================================================================");
            this.cleanup(isVhdDistribute, templateName, containerName, blobName);
        }
        catch (error) {
            throw error;
        }
    }
    // v1-patched: @azure/storage-blob initialization
    async _initBlobService() {
        var storageDetails = await this._getStorageAccountDetails();
        this._storageCredential = new storage_blob_1.StorageSharedKeyCredential(storageDetails.name, storageDetails.primaryAccessKey);
        this._blobServiceClient = new storage_blob_1.BlobServiceClient(storageDetails.primaryBlobUrl, this._storageCredential);
    }
    // v1-patched: modern blob SDK upload + HTTPS-only SAS with timeout-based expiry
    async _uploadPackage(containerName, blobName) {
        var archivedWebPackage;
        try {
            // use zip for Windows and tar for Linux
            if (Utils_1.default.IsEqual(this._taskParameters.provisioner, "powershell"))
                archivedWebPackage = await this.createArchive(this._taskParameters.buildPath, this._generateTemporaryFile(tl.getVariable('AGENT.TEMPDIRECTORY'), `.zip`));
            else
                archivedWebPackage = await this.createArchive(this._taskParameters.buildPath, this._generateTemporaryFile(tl.getVariable('AGENT.TEMPDIRECTORY'), `.tar.gz`));
        }
        catch (error) {
            throw Error(`unable to create archive build: ${error}`);
        }
        console.log(`created archive ${archivedWebPackage}`);
        // Ensure container exists
        var containerClient = this._blobServiceClient.getContainerClient(containerName);
        await containerClient.createIfNotExists();
        // Upload blob
        console.log(`uploading to ${containerName}/${blobName}...`);
        var blockBlobClient = containerClient.getBlockBlobClient(blobName);
        await blockBlobClient.uploadFile(archivedWebPackage);
        // SAS expiry = build timeout + 1h buffer (minimum 4h, cap 25h)
        // buildTimeoutInMinutes=0 means "wait indefinitely"; AIB default is 240 min
        var timeoutMin = this._taskParameters.buildTimeoutInMinutes || 240;
        var sasHours = Math.max(4, Math.min(25, Math.ceil(timeoutMin / 60) + 1));
        var startsOn = new Date();
        startsOn.setMinutes(startsOn.getMinutes() - 5);
        var expiresOn = new Date();
        expiresOn.setHours(expiresOn.getHours() + sasHours);
        var sasToken = (0, storage_blob_1.generateBlobSASQueryParameters)({
            containerName,
            blobName,
            permissions: storage_blob_1.BlobSASPermissions.parse("r"),
            startsOn,
            expiresOn,
            protocol: storage_blob_1.SASProtocol.Https
        }, this._storageCredential).toString();
        return `${blockBlobClient.url}?${sasToken}`;
    }
    async createArchive(folderPath, targetPath) {
        return new Promise((resolve, reject) => {
            console.log('Archiving ' + folderPath + ' to ' + targetPath);
            var output = fs.createWriteStream(targetPath);
            var archive;
            if (targetPath.endsWith(".zip")) {
                archive = archiver('zip');
            }
            else {
                archive = archiver('tar', {
                    gzip: true,
                    gzipOptions: {
                        level: 1
                    }
                });
            }
            output.on('close', function () {
                console.log(archive.pointer() + ' total bytes');
                tl.debug('Successfully created archive ' + targetPath);
                resolve(targetPath);
            });
            output.on('error', function (error) {
                reject(error);
            });
            archive.on('error', function (error) {
                reject(error);
            });
            archive.glob("**", {
                cwd: folderPath,
                dot: true
            });
            archive.pipe(output);
            archive.finalize();
        });
    }
    _generateTemporaryFile(folderPath, extension) {
        var randomString = Math.random().toString().split('.')[1];
        var tempPath = path.join(folderPath, 'temp_web_package_' + randomString + extension);
        if (tl.exist(tempPath)) {
            return this._generateTemporaryFile(folderPath, extension);
        }
        return tempPath;
    }
    async _getStorageAccountDetails() {
        console.log(`getting storage account details for ${this._taskParameters.storageAccountName}`);
        const subscriptionId = tl.getEndpointDataParameter(this._taskParameters.ibSubscriptionId, "subscriptionId", false);
        const credentials = await this._getARMCredentials(this._taskParameters.ibSubscriptionId);
        const storageArmClient = new armStorage.StorageManagementClient(credentials, subscriptionId);
        const storageAccount = await this._getStorageAccount(storageArmClient);
        const storageAccountResourceGroupName = armStorage.StorageAccounts.getResourceGroupNameFromUri(storageAccount.id);
        const accessKeys = await storageArmClient.storageAccounts.listKeys(storageAccountResourceGroupName, this._taskParameters.storageAccountName, null, storageAccount.type);
        return {
            name: this._taskParameters.storageAccountName,
            primaryBlobUrl: storageAccount.properties.primaryEndpoints.blob,
            resourceGroupName: storageAccountResourceGroupName,
            primaryAccessKey: accessKeys[0]
        };
    }
    async _getStorageAccount(storageArmClient) {
        const storageAccounts = await storageArmClient.storageAccounts.listClassicAndRMAccounts(null);
        return await storageArmClient.storageAccounts.get(this._taskParameters.storageAccountName);
    }
    async _getARMCredentials(connectedServiceName) {
        var endpoint = await new azure_arm_endpoint_1.AzureRMEndpoint(connectedServiceName).getEndpoint();
        return endpoint.applicationTokenCredentials;
    }
    _getBlobsPrefixPath() {
        var uniqueValue = Date.now().toString();
        var releaseId = tl.getVariable("release.releaseid");
        var releaseAttempt = tl.getVariable("release.attemptnumber");
        var prefixFolderPath = "";
        if (!!releaseId && !!releaseAttempt) {
            prefixFolderPath = util.format("%s-%s/", releaseId, releaseAttempt);
        }
        else {
            prefixFolderPath = util.format("%s-%s/", tl.getVariable("build.buildid"), uniqueValue);
        }
        return prefixFolderPath;
    }
    getTemplateName() {
        // Date.now() method returns the number of milliseconds elapsed since January 1, 1970
        return "t_" + Date.now().toString();
    }
    async cleanup(isVhdDistribute, templateName, containerName, blobName) {
        try {
            if (!isVhdDistribute) {
                await Promise.all([this._aibClient.deleteTemplate(templateName), this._deleteBlob(containerName, blobName)]);
            }
            else {
                console.log(`template ${templateName} has vhd distribute so skipping delete template.`);
                await this._deleteBlob(containerName, blobName);
            }
        }
        catch (error) {
            console.log(`Error in cleanup: `, error);
        }
    }
    // v1-patched: @azure/storage-blob delete
    async _deleteBlob(containerName, blobName) {
        console.log(`deleting storage blob ${containerName}/${blobName}`);
        try {
            var containerClient = this._blobServiceClient.getContainerClient(containerName);
            await containerClient.getBlockBlobClient(blobName).delete();
            console.log(`blob ${containerName}/${blobName} is deleted`);
        }
        catch (error) {
            console.log(`unable to delete blob ${containerName}/${blobName}: ${error}`);
        }
    }
}
exports.default = ImageBuilder;
