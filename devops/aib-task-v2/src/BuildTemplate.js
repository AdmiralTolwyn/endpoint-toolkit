"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const Utils_1 = __importDefault(require("./Utils"));
// v1-patched: programmatic object construction replaces string-template approach
// to support nested objects (validate, optimize, errorHandling, targetRegions, etc.)
var templateSource = new Map([
    ["managedimage", { type: "ManagedImage", imageId: "" }],
    ["gallery", { type: "SharedImageVersion", imageVersionId: "" }],
    ["marketplace", { type: "PlatformImage", publisher: "", offer: "", sku: "", version: "latest" }]
]);
class BuildTemplate {
    _taskParameters;
    constructor(taskParameters) {
        this._taskParameters = taskParameters;
    }
    getTemplate(blobUrl) {
        var template = {
            location: this._taskParameters.ibLocation,
            identity: {
                type: "UserAssigned",
                userAssignedIdentities: {
                    [this._taskParameters.identity]: {}
                }
            },
            properties: {
                source: this._buildSource(),
                customize: this._buildCustomizers(blobUrl),
                distribute: [this._buildDistribute()],
                vmProfile: this._buildVmProfile(),
                buildTimeoutInMinutes: this._taskParameters.buildTimeoutInMinutes,
                errorHandling: {
                    onCustomizerError: this._taskParameters.errorHandlingOnCustomizerError,
                    onValidationError: this._taskParameters.errorHandlingOnValidationError
                }
            }
        };
        // optimize — vmBoot
        if (this._taskParameters.optimizeVmBoot) {
            template.properties.optimize = { vmBoot: { state: "Enabled" } };
        }
        // validate
        if (this._taskParameters.enableValidation) {
            template.properties.validate = this._buildValidate();
        }
        // autoRun
        if (this._taskParameters.autoRun) {
            template.properties.autoRun = { state: "Enabled" };
        }
        // stagingResourceGroup
        if (this._taskParameters.stagingResourceGroup) {
            template.properties.stagingResourceGroup = this._taskParameters.stagingResourceGroup;
        }
        // managedResourceTags
        if (Object.keys(this._taskParameters.managedResourceTags).length > 0) {
            template.properties.managedResourceTags = this._taskParameters.managedResourceTags;
        }
        return template;
    }
    _buildSource() {
        var source = { ...templateSource.get(this._taskParameters.imageSource) };
        if (Utils_1.default.IsEqual(source.type, "PlatformImage")) {
            source.publisher = this._taskParameters.imagePublisher;
            source.offer = this._taskParameters.imageOffer;
            source.sku = this._taskParameters.imageSku;
            source.version = this._taskParameters.baseImageVersion || "latest";
            if (this._taskParameters.planName && this._taskParameters.planProduct && this._taskParameters.planPublisher) {
                source.planInfo = {
                    planName: this._taskParameters.planName,
                    planProduct: this._taskParameters.planProduct,
                    planPublisher: this._taskParameters.planPublisher
                };
            }
        }
        else if (Utils_1.default.IsEqual(source.type, "ManagedImage"))
            source.imageId = this._taskParameters.sourceResourceId;
        else
            source.imageVersionId = this._taskParameters.imageVersionId;
        console.log("Source for image: ", JSON.stringify(source, null, 2));
        return source;
    }
    _buildCustomizers(blobUrl) {
        var customizers = [];
        if (Utils_1.default.IsEqual(this._taskParameters.provisioner, "shell")) {
            var packageName = `/tmp/${this._taskParameters.buildFolder}`;
            customizers.push({
                type: "File",
                name: "vststask_file_copy",
                sourceUri: blobUrl,
                destination: `${packageName}.tar.gz`
            });
            var inline = "#\n";
            inline += `mkdir -p ${packageName}\n`;
            inline += `sudo tar -xzvf ${packageName}.tar.gz -C ${packageName}\n`;
            if (this._taskParameters.inlineScript)
                inline += `${this._taskParameters.inlineScript}\n`;
            customizers.push({
                type: "Shell",
                name: "vststask_inline",
                inline: inline.split("\n")
            });
        }
        else if (Utils_1.default.IsEqual(this._taskParameters.provisioner, "powershell")) {
            var packageName = "c:\\buildartifacts\\" + this._taskParameters.buildFolder;
            var inline = `New-Item -Path c:\\buildartifacts -ItemType Directory -Force\n`;
            inline += `$ProgressPreference = 'SilentlyContinue'\n`;
            inline += `Invoke-WebRequest -Uri '${blobUrl}' -OutFile '${packageName}.zip' -UseBasicParsing\n`;
            inline += `Expand-Archive -Path '${packageName}.zip' -DestinationPath '${packageName}' -Force\n`;
            inline += `Remove-Item '${packageName}.zip' -Force\n`;
            if (this._taskParameters.inlineScript)
                inline += `${this._taskParameters.inlineScript}\n`;
            var psCustomizer = {
                type: "PowerShell",
                name: "vststask_inline",
                inline: inline.split("\n")
            };
            if (this._taskParameters.runElevated) {
                psCustomizer.runElevated = true;
                if (this._taskParameters.runAsSystem)
                    psCustomizer.runAsSystem = true;
            }
            customizers.push(psCustomizer);
        }
        // Windows Update customizer (optional, PowerShell only)
        if (Utils_1.default.IsEqual(this._taskParameters.provisioner, "powershell") && this._taskParameters.windowsUpdateProvisioner) {
            customizers.push({
                type: "PowerShell",
                name: "PreWindowsUpdateWait",
                runElevated: true,
                inline: ["Start-Sleep -Seconds 300"]
            });
            customizers.push({
                type: "WindowsUpdate",
                searchCriteria: "IsInstalled=0",
                filters: [
                    "exclude:$_.Title -like '*Preview*'",
                    "include:$true"
                ]
            });
        }
        return customizers;
    }
    _buildValidate() {
        var validate = {
            continueDistributeOnFailure: this._taskParameters.continueDistributeOnFailure,
            sourceValidationOnly: this._taskParameters.sourceValidationOnly,
            inVMValidations: []
        };
        if (this._taskParameters.validationInlineScript) {
            var validatorType = Utils_1.default.IsEqual(this._taskParameters.provisioner, "shell") ? "Shell" : "PowerShell";
            var validator = {
                type: validatorType,
                name: "vststask_validation",
                inline: this._taskParameters.validationInlineScript.split("\n")
            };
            if (validatorType === "PowerShell" && this._taskParameters.runElevated)
                validator.runElevated = true;
            validate.inVMValidations.push(validator);
        }
        return validate;
    }
    _buildDistribute() {
        var tp = this._taskParameters;
        var distribute;
        if (Utils_1.default.IsEqual(tp.distributeType, "managedimage")) {
            distribute = {
                type: "ManagedImage",
                imageId: tp.imageIdForDistribute,
                location: tp.managedImageLocation,
                runOutputName: "ManagedImage_distribute"
            };
        }
        else if (Utils_1.default.IsEqual(tp.distributeType, "gallery") || Utils_1.default.IsEqual(tp.distributeType, "sig")) {
            // targetRegions replaces deprecated replicationRegions
            var regionInput = tp.targetRegions || tp.replicationRegions || "";
            var regions = regionInput.split(",").map(r => r.trim()).filter(r => r.length > 0);
            var targetRegions = regions.map(name => {
                var region = { name };
                if (tp.distributeStorageAccountType)
                    region.storageAccountType = tp.distributeStorageAccountType;
                return region;
            });
            distribute = {
                type: "SharedImage",
                galleryImageId: tp.galleryImageId,
                targetRegions: targetRegions,
                excludeFromLatest: tp.excludeFromLatest,
                runOutputName: "SharedImage_distribute"
            };
            // Versioning
            if (tp.distributeVersioningScheme) {
                distribute.versioning = { scheme: tp.distributeVersioningScheme };
                if (Utils_1.default.IsEqual(tp.distributeVersioningScheme, "Latest") && tp.distributeVersioningMajor)
                    distribute.versioning.major = Number(tp.distributeVersioningMajor);
            }
        }
        else {
            // VHD
            distribute = {
                type: "VHD",
                runOutputName: "VHD_distribute"
            };
            if (tp.vhdUri)
                distribute.uri = tp.vhdUri;
        }
        return distribute;
    }
    _buildVmProfile() {
        var vmProfile = {
            vmSize: this._taskParameters.vmSize
        };
        if (this._taskParameters.osDiskSizeGB > 0)
            vmProfile.osDiskSizeGB = this._taskParameters.osDiskSizeGB;
        if (this._taskParameters.vnetSubnetId) {
            vmProfile.vnetConfig = { subnetId: this._taskParameters.vnetSubnetId };
            if (this._taskParameters.containerInstanceSubnetId)
                vmProfile.vnetConfig.containerInstanceSubnetId = this._taskParameters.containerInstanceSubnetId;
            if (this._taskParameters.proxyVmSize && !this._taskParameters.containerInstanceSubnetId)
                vmProfile.vnetConfig.proxyVmSize = this._taskParameters.proxyVmSize;
        }
        if (this._taskParameters.buildVmIdentities.length > 0)
            vmProfile.userAssignedIdentities = this._taskParameters.buildVmIdentities;
        return vmProfile;
    }
}
exports.default = BuildTemplate;
