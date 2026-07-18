import path = require("path");

import tl = require("azure-pipelines-task-lib/task");
// v1-patched: import path changed from azure-arm-rest-v2 to azure-arm-rest (v3)
import armStorage = require('azure-pipelines-tasks-azure-arm-rest/azure-arm-storage');
import msRestAzure = require("azure-pipelines-tasks-azure-arm-rest/azure-arm-common");
import { AzureRMEndpoint } from 'azure-pipelines-tasks-azure-arm-rest/azure-arm-endpoint';
import azureModel = require('azure-pipelines-tasks-azure-arm-rest/azureModels');
import util = require('util');
import fs = require('fs');
// v1-patched: @azure/storage-blob replaces deprecated azure-storage SDK
import { BlobServiceClient, StorageSharedKeyCredential, generateBlobSASQueryParameters, BlobSASPermissions, SASProtocol } from "@azure/storage-blob";
var archiver = require('archiver');
import Utils from "./Utils";

import TaskParameters from "./TaskParameters";
import BuildTemplate from "./BuildTemplate";
import ImageBuilderClient from "./AzureImageBuilderClient"

var containerName = 'imagebuilder-vststask';

export default class ImageBuilder {

    private _taskParameters: TaskParameters;
    private _aibClient: ImageBuilderClient;
    private _buildTemplate: BuildTemplate;
    // v1-patched: @azure/storage-blob types replace azure-storage's blobService
    private _blobServiceClient: BlobServiceClient;
    private _storageCredential: StorageSharedKeyCredential;

    constructor()
    {
        try{
            this._taskParameters = new TaskParameters();
            this._taskParameters.printSummary();
            this._buildTemplate = new BuildTemplate(this._taskParameters);
        }
        catch (error) {
            throw (`error happened while initializing Image builder: ${error}`);
        }
    }

    public async execute(): Promise<void> {
        // v1-patched: elapsed time tracking
        var buildStart = Date.now();
        var elapsed = () => {
            var s = Math.floor((Date.now() - buildStart) / 1000);
            var m = Math.floor(s / 60);
            var sec = s % 60;
            return m > 0 ? `${m}m ${sec}s` : `${sec}s`;
        };

        try{
            var creds = await this._getARMCredentials(this._taskParameters.ibSubscriptionId);
            const subscriptionId: string = tl.getEndpointDataParameter(this._taskParameters.ibSubscriptionId, "subscriptionId", false);
            this._aibClient = new ImageBuilderClient(creds, subscriptionId, this._taskParameters);
            await this._initBlobService();

            var blobName: string = this._taskParameters.buildFolder + "/" + this._getBlobsPrefixPath() + this._taskParameters.buildFolder + `_${Date.now()}`;
            if (Utils.IsEqual(this._taskParameters.provisioner, "powershell"))
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

            if(Utils.IsEqual(templateJson.properties.source.type, "PlatformImage"))
            {
                tl.setVariable('pirPublisher', templateJson.properties.source.publisher);
                tl.setVariable('pirOffer', templateJson.properties.source.offer);
                tl.setVariable('pirSku', templateJson.properties.source.sku);
                tl.setVariable('pirVersion', templateJson.properties.source.version);
            }

            console.log("==============================================================================")
            console.log("## task output variables ##");
            console.log("$(imageUri) = ", tl.getVariable('imageUri'));
            if(isVhdDistribute)
            {
                console.log("$(templateName) = ", tl.getVariable('templateName'));
                console.log("$(templateId) = ", tl.getVariable('templateId'));
            }
            console.log("Total elapsed: %s", elapsed());
            console.log("==============================================================================")

            this.cleanup(isVhdDistribute, templateName, containerName, blobName);
        }
        catch(error)
        {
            throw error;
        }
    }

    // v1-patched: @azure/storage-blob initialization
    private async _initBlobService(): Promise<void>
    {
        var storageDetails: StorageAccountInfo = await this._getStorageAccountDetails();
        this._storageCredential = new StorageSharedKeyCredential(storageDetails.name, storageDetails.primaryAccessKey);
        this._blobServiceClient = new BlobServiceClient(storageDetails.primaryBlobUrl, this._storageCredential);
    }

    // v1-patched: modern blob SDK upload + HTTPS-only SAS with timeout-based expiry
    private async _uploadPackage(containerName: string, blobName: string) : Promise<string> {
        var archivedWebPackage: string;

        try {
            // use zip for Windows and tar for Linux
            if (Utils.IsEqual(this._taskParameters.provisioner, "powershell"))
                archivedWebPackage = await this.createArchive(this._taskParameters.buildPath,
                                                                    this._generateTemporaryFile(tl.getVariable('AGENT.TEMPDIRECTORY'), `.zip`));
            else
                archivedWebPackage = await this.createArchive(this._taskParameters.buildPath,
                                                                 this._generateTemporaryFile(tl.getVariable('AGENT.TEMPDIRECTORY'), `.tar.gz`));
        }
        catch(error) {
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

        var sasToken = generateBlobSASQueryParameters(
            {
                containerName,
                blobName,
                permissions: BlobSASPermissions.parse("r"),
                startsOn,
                expiresOn,
                protocol: SASProtocol.Https
            },
            this._storageCredential
        ).toString();

        return `${blockBlobClient.url}?${sasToken}`;
    }

    private async createArchive(folderPath: string, targetPath: string) : Promise<string> {
        return new Promise<string>((resolve, reject) => {
            console.log('Archiving ' + folderPath + ' to ' + targetPath);
            var output = fs.createWriteStream(targetPath);

            var archive: any;
            if(targetPath.endsWith(".zip")){
                archive = archiver('zip');
            } else {
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

            output.on('error', function(error) {
                reject(error);
            });

            archive.on('error', function(error) {
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

    private _generateTemporaryFile(folderPath: string, extension: string) {
        var randomString = Math.random().toString().split('.')[1];
        var tempPath = path.join(folderPath, 'temp_web_package_' + randomString + extension);
        if(tl.exist(tempPath)) {
            return this._generateTemporaryFile(folderPath, extension);
        }
        return tempPath;
    }

    private async _getStorageAccountDetails(): Promise<StorageAccountInfo> {
        console.log(`getting storage account details for ${this._taskParameters.storageAccountName}`);
        const subscriptionId: string = tl.getEndpointDataParameter(this._taskParameters.ibSubscriptionId, "subscriptionId", false);
        const credentials = await this._getARMCredentials(this._taskParameters.ibSubscriptionId);
        const storageArmClient = new armStorage.StorageManagementClient(credentials, subscriptionId);
        const storageAccount: azureModel.StorageAccount = await this._getStorageAccount(storageArmClient);

        const storageAccountResourceGroupName = armStorage.StorageAccounts.getResourceGroupNameFromUri(storageAccount.id);

        const accessKeys = await storageArmClient.storageAccounts.listKeys(storageAccountResourceGroupName, this._taskParameters.storageAccountName, null, storageAccount.type);

        return <StorageAccountInfo>{
            name: this._taskParameters.storageAccountName,
            primaryBlobUrl: storageAccount.properties.primaryEndpoints.blob,
            resourceGroupName: storageAccountResourceGroupName,
            primaryAccessKey: accessKeys[0]
        }
    }

    private async _getStorageAccount(storageArmClient: armStorage.StorageManagementClient): Promise<azureModel.StorageAccount> {
        const storageAccounts = await storageArmClient.storageAccounts.listClassicAndRMAccounts(null);
        return await storageArmClient.storageAccounts.get(this._taskParameters.storageAccountName);
      }

    private async _getARMCredentials(connectedServiceName: string): Promise<msRestAzure.ApplicationTokenCredentials> {
        var endpoint = await new AzureRMEndpoint(connectedServiceName).getEndpoint();
        return endpoint.applicationTokenCredentials;
    }

    private _getBlobsPrefixPath(): string {
        var uniqueValue = Date.now().toString();
        var releaseId = tl.getVariable("release.releaseid");
        var releaseAttempt = tl.getVariable("release.attemptnumber");
        var prefixFolderPath: string = "";
        if (!!releaseId && !!releaseAttempt) {
            prefixFolderPath = util.format("%s-%s/", releaseId, releaseAttempt);
        } else {
            prefixFolderPath = util.format("%s-%s/", tl.getVariable("build.buildid"), uniqueValue);
        }
        return prefixFolderPath;
    }

    private getTemplateName(): string {
        // Date.now() method returns the number of milliseconds elapsed since January 1, 1970
        return "t_" + Date.now().toString();
    }
    
    private async cleanup(isVhdDistribute: boolean, templateName: string, containerName: string, blobName: string)
    {
        try{
            if(!isVhdDistribute)
            {
                await Promise.all([this._aibClient.deleteTemplate(templateName), this._deleteBlob(containerName, blobName)]);
            }
            else
            {
                console.log(`template ${templateName} has vhd distribute so skipping delete template.`);
                await this._deleteBlob(containerName, blobName);
            }
        }
        catch(error)
        {
            console.log(`Error in cleanup: `, error);
        }
    }

    // v1-patched: @azure/storage-blob delete
    private async _deleteBlob(containerName: string, blobName: string)
    {
        console.log(`deleting storage blob ${containerName}/${blobName}`);
        try {
            var containerClient = this._blobServiceClient.getContainerClient(containerName);
            await containerClient.getBlockBlobClient(blobName).delete();
            console.log(`blob ${containerName}/${blobName} is deleted`);
        } catch(error) {
            console.log(`unable to delete blob ${containerName}/${blobName}: ${error}`);
        }
    }
}

interface StorageAccountInfo {
    name: string;
    resourceGroupName: string;
    primaryAccessKey: string;
    primaryBlobUrl: string;
}
