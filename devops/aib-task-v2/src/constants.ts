export var IbSubscription = "ibSubscription";
export var IbLocation = "ibLocation";
export var IbAzureResourceGroup = "ibAzureResourceGroup";

export var ImageSource = "imageSource";
export var ImageResourceId = "customImageResourceId";
export var BaseImage = "baseImage";
export var BaseImageVersion = "baseImageVersion";
export var BaseImagePubOfferSku = "baseImagePubOfferSku"
export var PlatformImagePurchasePlan = "platformImagePurchasePlan"
export var ImageVersionId = "imageVersionId";
export var ManagedIdentity = "managedIdentity";
export var VnetSubnetId = "vnetSubnetId"

export var BuildFolder = "packagePath";
export var InlineScript = "inlineScript";
export var ConnectedServiceName = "connectedServiceName"; // subscription Id
export var StorageAccountName = "storageAccountName";
export var Provisioner = "provisioner";
export var WindowsUpdateProvisioner = "windowsUpdateProvisioner";
export var BuildTimeoutInMinutes = "buildTimeoutInMinutes"
export var VMSize = "vmSize";
export var RunElevated = "runElevated";
export var RunAsSystem = "runAsSystem";

export var GalleryImageId = "galleryImageId";
export var DistributeType = "distributeType";
export var ImageIdForDistribute = "imageIdForDistribute";
export var ReplicationRegions = "replicationRegions";
export var ManagedImageLocation = "managedImageLocation";

// ── v1-patched additions (API 2024-02-01) ────────────────────────────────────

export var AIB_API_VERSION = "2024-02-01";

export var BuildVmIdentities = "buildVmIdentities";
export var ContainerInstanceSubnetId = "containerInstanceSubnetId";
export var ProxyVmSize = "proxyVmSize";
export var OsDiskSizeGB = "osDiskSizeGB";

export var EnableValidation = "enableValidation";
export var SourceValidationOnly = "sourceValidationOnly";
export var ContinueDistributeOnFailure = "continueDistributeOnFailure";
export var ValidationInlineScript = "validationInlineScript";

export var TargetRegions = "targetRegions";
export var DistributeStorageAccountType = "distributeStorageAccountType";
export var ExcludeFromLatest = "excludeFromLatest";
export var DistributeVersioningScheme = "distributeVersioningScheme";
export var DistributeVersioningMajor = "distributeVersioningMajor";
export var VhdUri = "vhdUri";

export var OptimizeVmBoot = "optimizeVmBoot";
export var AutoRun = "autoRun";
export var StagingResourceGroup = "stagingResourceGroup";
export var ManagedResourceTags = "managedResourceTags";
export var ErrorHandlingOnCustomizerError = "errorHandlingOnCustomizerError";
export var ErrorHandlingOnValidationError = "errorHandlingOnValidationError";
