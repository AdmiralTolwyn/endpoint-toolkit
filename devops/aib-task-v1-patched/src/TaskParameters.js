"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const tl = __importStar(require("azure-pipelines-task-lib/task"));
const constants = __importStar(require("./constants"));
const Utils_1 = __importDefault(require("./Utils"));
var repoRoot = tl.getVariable('System.DefaultWorkingDirectory');
var defaultPublisher = "Publisher:Offer:Sku";
var defaultPurchasePLan = "PlanName:PlanProduct:PlanPublisher";
class TaskParameters {
    // image builder inputs
    ibSubscriptionId;
    ibResourceGroup;
    ibLocation;
    // source
    imageSource;
    sourceResourceId;
    imageVersionId;
    baseImageVersion;
    imagePublisher;
    imageOffer;
    imageSku;
    planName;
    planProduct;
    planPublisher;
    //customize
    buildPath;
    buildFolder;
    blobName = "";
    inlineScript;
    provisioner;
    windowsUpdateProvisioner;
    storageAccountName;
    buildTimeoutInMinutes;
    vmSize;
    osDiskSizeGB;
    identity;
    buildVmIdentities;
    vnetSubnetId;
    containerInstanceSubnetId;
    proxyVmSize;
    runElevated;
    runAsSystem;
    // validate (v1-patched)
    enableValidation;
    sourceValidationOnly;
    continueDistributeOnFailure;
    validationInlineScript;
    //distribute
    distributeType;
    imageIdForDistribute;
    replicationRegions;
    targetRegions;
    managedImageLocation;
    galleryImageId;
    distributeStorageAccountType;
    excludeFromLatest;
    distributeVersioningScheme;
    distributeVersioningMajor;
    vhdUri;
    // optional / advanced (v1-patched)
    optimizeVmBoot;
    autoRun;
    stagingResourceGroup;
    managedResourceTags;
    errorHandlingOnCustomizerError;
    errorHandlingOnValidationError;
    constructor() {
        // ib inputs
        console.log("start reading task parameters...");
        this.ibSubscriptionId = tl.getInput(constants.IbSubscription, true);
        if (!this.ibSubscriptionId)
            throw Error("subscription is required");
        this.ibResourceGroup = tl.getInput(constants.IbAzureResourceGroup, true);
        if (!this.ibResourceGroup)
            throw Error("resource group is required");
        this.ibLocation = tl.getInput(constants.IbLocation, true);
        if (!this.ibLocation)
            throw Error("template location is required");
        this.imageSource = tl.getInput(constants.ImageSource, true);
        if (!this.imageSource)
            throw Error("image source is required");
        this.sourceResourceId = tl.getInput(constants.ImageResourceId, false);
        this.imageVersionId = tl.getInput(constants.ImageVersionId, false);
        this.baseImageVersion = tl.getInput(constants.BaseImageVersion, false);
        this.storageAccountName = tl.getInput(constants.StorageAccountName, true);
        if (!this.storageAccountName)
            throw Error("storage account is required");
        if (this.imageSource == "marketplace") {
            var img = tl.getPathInput(constants.BaseImagePubOfferSku, false);
            if (!img || Utils_1.default.IsEqual(img, defaultPublisher))
                img = tl.getInput(constants.BaseImage, false);
            this._extractImageDetails(img);
            var pp = tl.getPathInput(constants.PlatformImagePurchasePlan, false);
            if (pp && !Utils_1.default.IsEqual(pp, defaultPurchasePLan))
                this._extractPurchasePlan(pp);
        }
        var bp = tl.getPathInput(constants.BuildFolder, true, false);
        if (!bp)
            throw Error("build folder is required");
        var x = bp.split(path.sep);
        this.buildFolder = x[x.length - 1];
        this.buildPath = this._makeAbsolute(path.normalize(bp.trim()));
        console.log("found build at: ", this.buildPath);
        this.inlineScript = tl.getInput(constants.InlineScript, false);
        this.provisioner = tl.getInput(constants.Provisioner, false);
        this.windowsUpdateProvisioner = tl.getBoolInput(constants.WindowsUpdateProvisioner, false);
        this.runElevated = tl.getBoolInput(constants.RunElevated, false);
        this.runAsSystem = tl.getBoolInput(constants.RunAsSystem, false);
        // validate (v1-patched)
        this.enableValidation = tl.getBoolInput(constants.EnableValidation, false);
        this.sourceValidationOnly = tl.getBoolInput(constants.SourceValidationOnly, false);
        this.continueDistributeOnFailure = tl.getBoolInput(constants.ContinueDistributeOnFailure, false);
        this.validationInlineScript = tl.getInput(constants.ValidationInlineScript, false) || "";
        this.distributeType = tl.getInput(constants.DistributeType, true);
        if (!this.distributeType)
            throw Error("distribute type is required");
        this.imageIdForDistribute = tl.getInput(constants.ImageIdForDistribute, false);
        this.galleryImageId = tl.getInput(constants.GalleryImageId, false);
        this.replicationRegions = tl.getInput(constants.ReplicationRegions, false);
        this.targetRegions = tl.getInput(constants.TargetRegions, false) || "";
        this.managedImageLocation = tl.getInput(constants.ManagedImageLocation, false);
        this.distributeStorageAccountType = tl.getInput(constants.DistributeStorageAccountType, false) || "";
        this.excludeFromLatest = tl.getBoolInput(constants.ExcludeFromLatest, false);
        this.distributeVersioningScheme = tl.getInput(constants.DistributeVersioningScheme, false) || "";
        this.distributeVersioningMajor = tl.getInput(constants.DistributeVersioningMajor, false) || "";
        this.vhdUri = tl.getInput(constants.VhdUri, false) || "";
        // build timeout in minutes
        this.buildTimeoutInMinutes = Number(tl.getInput(constants.BuildTimeoutInMinutes, false));
        if (isNaN(this.buildTimeoutInMinutes))
            throw Error("build timeout in minutes should be integer value");
        if (this.buildTimeoutInMinutes < 0)
            throw Error("build timeout in minutes should not be negative");
        // vm size
        this.vmSize = tl.getInput(constants.VMSize, false);
        // os disk size (v1-patched)
        var osDisk = tl.getInput(constants.OsDiskSizeGB, false) || "0";
        this.osDiskSizeGB = Number(osDisk);
        if (isNaN(this.osDiskSizeGB) || this.osDiskSizeGB < 0)
            throw Error("osDiskSizeGB must be a non-negative integer");
        // identity
        this.identity = tl.getInput(constants.ManagedIdentity, true);
        var buildVmIdInput = tl.getInput(constants.BuildVmIdentities, false) || "";
        this.buildVmIdentities = buildVmIdInput.split("\n").map(s => s.trim()).filter(s => s.length > 0);
        // Vnet 
        this.vnetSubnetId = tl.getInput(constants.VnetSubnetId, false);
        this.containerInstanceSubnetId = tl.getInput(constants.ContainerInstanceSubnetId, false) || "";
        this.proxyVmSize = tl.getInput(constants.ProxyVmSize, false) || "";
        // Advanced / optional settings (v1-patched)
        this.optimizeVmBoot = tl.getBoolInput(constants.OptimizeVmBoot, false);
        this.autoRun = tl.getBoolInput(constants.AutoRun, false);
        this.stagingResourceGroup = tl.getInput(constants.StagingResourceGroup, false) || "";
        this.errorHandlingOnCustomizerError = tl.getInput(constants.ErrorHandlingOnCustomizerError, false) || "cleanup";
        this.errorHandlingOnValidationError = tl.getInput(constants.ErrorHandlingOnValidationError, false) || "cleanup";
        var tagsInput = tl.getInput(constants.ManagedResourceTags, false) || "";
        if (tagsInput) {
            try {
                this.managedResourceTags = JSON.parse(tagsInput);
            }
            catch {
                throw Error("managedResourceTags must be valid JSON, e.g. {\"key\": \"value\"}");
            }
        }
        else {
            this.managedResourceTags = {};
        }
        console.log("end reading parameters");
    }
    // ── v1-patched: configuration summary ──
    printSummary() {
        var line = "──────────────────────────────────────────────────────────────────────────────";
        console.log("");
        console.log(line);
        console.log("  CONFIGURATION SUMMARY");
        console.log(line);
        console.log("  Resource Group:        %s", this.ibResourceGroup);
        console.log("  Location:              %s", this.ibLocation);
        console.log("  Storage Account:       %s", this.storageAccountName);
        var idParts = this.identity.split("/");
        var shortId = idParts[idParts.length - 1] || this.identity;
        console.log("  Managed Identity:      %s", shortId);
        console.log(line);
        console.log("  SOURCE");
        console.log("  Type:                  %s", this.imageSource);
        if (this.imageSource === "marketplace") {
            console.log("  Publisher:             %s", this.imagePublisher);
            console.log("  Offer:                 %s", this.imageOffer);
            console.log("  SKU:                   %s", this.imageSku);
            console.log("  Version:               %s", this.baseImageVersion || "latest");
            if (this.planName)
                console.log("  Purchase Plan:         %s / %s / %s", this.planName, this.planProduct, this.planPublisher);
        }
        else if (this.imageSource === "managedimage") {
            console.log("  Image ID:              %s", this.sourceResourceId);
        }
        else {
            console.log("  Image Version ID:      %s", this.imageVersionId);
        }
        console.log(line);
        console.log("  VM PROFILE");
        console.log("  Size:                  %s", this.vmSize);
        if (this.osDiskSizeGB > 0)
            console.log("  OS Disk:               %d GB", this.osDiskSizeGB);
        if (this.vnetSubnetId) {
            console.log("  VNet Subnet:           %s", this.vnetSubnetId.split("/").pop());
            if (this.containerInstanceSubnetId)
                console.log("  Container Subnet:      %s", this.containerInstanceSubnetId.split("/").pop());
            if (this.proxyVmSize)
                console.log("  Proxy VM Size:         %s", this.proxyVmSize);
        }
        if (this.buildVmIdentities.length > 0)
            console.log("  Build VM Identities:   %d configured", this.buildVmIdentities.length);
        console.log(line);
        console.log("  CUSTOMIZATION");
        console.log("  Provisioner:           %s", this.provisioner);
        console.log("  Build Path:            %s", this.buildPath);
        if (this.runElevated)
            console.log("  Run Elevated:          yes" + (this.runAsSystem ? " (as SYSTEM)" : ""));
        if (this.windowsUpdateProvisioner)
            console.log("  Windows Update:        enabled");
        if (this.inlineScript) {
            var scriptLines = this.inlineScript.split("\n").filter(l => l.trim().length > 0).length;
            console.log("  Inline Script:         %d line(s)", scriptLines);
        }
        console.log(line);
        console.log("  DISTRIBUTION");
        console.log("  Type:                  %s", this.distributeType);
        if (Utils_1.default.IsEqual(this.distributeType, "gallery") || Utils_1.default.IsEqual(this.distributeType, "sig")) {
            console.log("  Gallery Image ID:      %s", this.galleryImageId);
            var regionStr = this.targetRegions || this.replicationRegions;
            if (regionStr)
                console.log("  Target Regions:        %s", regionStr);
            if (this.distributeStorageAccountType)
                console.log("  Storage Account Type:  %s", this.distributeStorageAccountType);
            console.log("  Exclude from Latest:   %s", this.excludeFromLatest ? "yes" : "no");
            if (this.distributeVersioningScheme)
                console.log("  Versioning:            %s%s", this.distributeVersioningScheme, this.distributeVersioningMajor ? ` (major: ${this.distributeVersioningMajor})` : "");
        }
        else if (Utils_1.default.IsEqual(this.distributeType, "managedimage")) {
            console.log("  Image ID:              %s", this.imageIdForDistribute);
            console.log("  Location:              %s", this.managedImageLocation);
        }
        else if (this.vhdUri) {
            console.log("  VHD URI:               %s", this.vhdUri);
        }
        if (this.enableValidation) {
            console.log(line);
            console.log("  VALIDATION");
            console.log("  Source Only:            %s", this.sourceValidationOnly ? "yes" : "no");
            console.log("  Continue on Failure:   %s", this.continueDistributeOnFailure ? "yes" : "no");
            if (this.validationInlineScript)
                console.log("  Validation Script:     %d line(s)", this.validationInlineScript.split("\n").filter(l => l.trim()).length);
        }
        var advancedLines = [];
        if (this.buildTimeoutInMinutes > 0)
            advancedLines.push(`Timeout: ${this.buildTimeoutInMinutes} min`);
        if (this.optimizeVmBoot)
            advancedLines.push("VM Boot Optimization: enabled");
        if (this.autoRun)
            advancedLines.push("Auto Run: enabled");
        if (this.stagingResourceGroup)
            advancedLines.push(`Staging RG: ${this.stagingResourceGroup}`);
        if (this.errorHandlingOnCustomizerError !== "cleanup")
            advancedLines.push(`On Customizer Error: ${this.errorHandlingOnCustomizerError}`);
        if (this.errorHandlingOnValidationError !== "cleanup")
            advancedLines.push(`On Validation Error: ${this.errorHandlingOnValidationError}`);
        if (Object.keys(this.managedResourceTags).length > 0)
            advancedLines.push(`Resource Tags: ${Object.keys(this.managedResourceTags).length} tag(s)`);
        if (advancedLines.length > 0) {
            console.log(line);
            console.log("  ADVANCED");
            advancedLines.forEach(l => console.log("  %s", l));
        }
        console.log(line);
        console.log("");
    }
    _extractImageDetails(img) {
        this.imagePublisher = "";
        this.imageOffer = "";
        this.imageSku = "";
        var parts = img.split(':');
        if (parts.length != 3)
            throw Error("Platform Base Image should have '{publisher}:{offer}:{sku}'. All fields are required.");
        this.imagePublisher = parts[0];
        this.imageOffer = parts[1];
        this.imageSku = parts[2];
    }
    _extractPurchasePlan(pp) {
        this.planName = "";
        this.planProduct = "";
        this.planPublisher = "";
        var parts = pp.split(':');
        if (parts.length != 3)
            throw Error("Purchase plan should have '{planName}:{planProduct}:{planPublisher}'. All fields are required.");
        this.planName = parts[0];
        this.planProduct = parts[1];
        this.planPublisher = parts[2];
    }
    _makeAbsolute(normalizedPath) {
        var result = normalizedPath;
        if (!path.isAbsolute(normalizedPath)) {
            result = path.join(repoRoot, normalizedPath);
        }
        return result;
    }
}
exports.default = TaskParameters;
