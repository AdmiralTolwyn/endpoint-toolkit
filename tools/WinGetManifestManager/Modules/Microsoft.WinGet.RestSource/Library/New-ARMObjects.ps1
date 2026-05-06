# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
function New-ARMObjects {
    <#
    .SYNOPSIS
    Creates the Azure Resources to stand-up a Windows Package Manager REST Source.

    .DESCRIPTION
    Uses the custom PowerShell object provided by the "New-ARMParameterObjects" cmdlet to create Azure resources, and will 
    create the the Key Vault secrets and publish the Windows Package Manager REST source REST apis to the Azure Function.

    .PARAMETER ARMObjects
    Object returned from the "New-ARMParameterObjects" providing the paths to the ARM Parameters and Template files.

    .PARAMETER RestSourcePath
    Path to the compiled Function ZIP containing the REST APIs

    .PARAMETER ResourceGroup
    Resource Group that will be used to create the ARM Objects in.

    .EXAMPLE
    New-ARMObjects -ARMObjects $ARMObjects -RestSourcePath "C:\WinGet-CLI-RestSource\WinGet.RestSource.Functions.zip" -ResourceGroup "WinGet"

    Parses through the $ARMObjects variable, creating all identified Azure Resources following the provided ARM Parameters and Template information.
    #>
    param(
        [Parameter(Position = 0, Mandatory = $true)] [array] [ref] $ARMObjects,
        [Parameter(Position = 1, Mandatory = $true)] [string] $RestSourcePath,
        [Parameter(Position = 2, Mandatory = $true)] [string] $ResourceGroup
    )

    # Function to create a new Function App key
    function New-FunctionAppKey {
        $private:characters = 'abcdefghiklmnoprstuvwxyzABCDEFGHIJKLMENOPTSTUVWXYZ'
        $private:randomChars = 1..64 | ForEach-Object { Get-Random -Maximum $characters.length }
    
        # Set the output field separator to empty instead of space
        $private:ofs = ''
        return [String]$characters[$randomChars]
    }

    # Function to ensure Azure role assignment
    function Set-RoleAssignment {
        param(
            [string] $PrincipalId,
            [string] $RoleName,
            [string] $ResourceGroup,
            [string] $ResourceName,
            [string] $ResourceType
        )
        
        $GetAssignment = Get-AzRoleAssignment -ObjectId $PrincipalId -RoleDefinitionName $RoleName -ResourceGroupName $ResourceGroup -ResourceName $ResourceName -ResourceType $ResourceType
        if (!$GetAssignment) {
            Write-Verbose "Creating role assignment. Role: $RoleName"
            $Result = New-AzRoleAssignment -ObjectId $PrincipalId -RoleDefinitionName $RoleName -ResourceGroupName $ResourceGroup -ResourceName $ResourceName -ResourceType $ResourceType -ErrorVariable ErrorNew
            if ($ErrorNew) {
                Write-Error "Failed to set Azure role. Role: $RoleName Error: $ErrorNew"
                return $false
            }
        }

        return $true
    }

    function Get-SecureString {
        param(
            [string] $InputString
        )

        $Result = New-Object SecureString
        foreach ($char in $InputString.ToCharArray()) {
            $Result.AppendChar($char)
        }

        return $Result
    }

    ## TODO: Consider multiple instances of same Azure Resource in the future
    ## Azure resource names retrieved from the Parameter files.
    $StorageAccountName = $ARMObjects.Where({ $_.ObjectType -eq 'StorageAccount' }).Parameters.Parameters.storageAccountName.value
    $KeyVaultName = $ARMObjects.Where({ $_.ObjectType -eq 'Keyvault' }).Parameters.Parameters.name.value
    $CosmosAccountName = $ARMObjects.Where({ $_.ObjectType -eq 'CosmosDBAccount' }).Parameters.Parameters.name.value
    $AppConfigName = $ARMObjects.Where({ $_.ObjectType -eq 'AppConfig' }).Parameters.Parameters.appConfigName.value
    $FunctionName = $ARMObjects.Where({ $_.ObjectType -eq 'Function' }).Parameters.Parameters.functionName.value

    ## Azure Keyvault Secret Names - Do not change values (Must match with values in the Template files)
    $CosmosAccountEndpointKeyName = 'CosmosAccountEndpoint'
    $AzureFunctionHostKeyName = 'AzureFunctionHostKey'
    $AppConfigPrimaryEndpointName = 'AppConfigPrimaryEndpoint'
    $AppConfigSecondaryEndpointName = 'AppConfigSecondaryEndpoint'

    ## Creates the Azure Resources following the ARM template / parameters
    Write-Information 'Creating Azure Resources following ARM Templates.'

    ## This is order specific, please ensure you used the New-ARMParameterObjects function to create this object in the pre-determined order.
    foreach ($Object in $ARMObjects) {
        if ($Object.DeploymentSuccess) {
            Write-Verbose "Skipped the Azure Object - $($Object.ObjectType). Deployment success detected."
            continue
        }

        Write-Information "Creating the Azure Object - $($Object.ObjectType)"

        ## Pre ARM deployment operations
        if ($Object.ObjectType -eq 'Function') {
            $CosmosAccountEndpointValue = Get-SecureString($(Get-AzCosmosDBAccount -Name $CosmosAccountName -ResourceGroupName $ResourceGroup).DocumentEndpoint)
            Write-Verbose 'Creating Keyvault Secret for Azure CosmosDB endpoint.'
            $Result = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $CosmosAccountEndpointKeyName -SecretValue $CosmosAccountEndpointValue -ErrorVariable ErrorSet
            if ($ErrorSet) {
                Write-Error "Failed to set keyvault secret. Name: $CosmosAccountEndpointKeyName Error: $ErrorSet"
                return $false
            }

            $AppConfigEndpointValue = Get-SecureString($(Get-AzAppConfigurationStore -Name $AppConfigName -ResourceGroupName $ResourceGroup).Endpoint)
            Write-Verbose 'Creating Keyvault Secret for Azure App Config endpoint.'
            $Result = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $AppConfigPrimaryEndpointName -SecretValue $AppConfigEndpointValue -ErrorVariable ErrorSet
            if ($ErrorSet) {
                Write-Error "Failed to set keyvault secret. Name: $AppConfigPrimaryEndpointName Error: $ErrorSet"
                return $false
            }
            $Result = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $AppConfigSecondaryEndpointName -SecretValue $AppConfigEndpointValue -ErrorVariable ErrorSet
            if ($ErrorSet) {
                Write-Error "Failed to set keyvault secret. Name: $AppConfigSecondaryEndpointName Error: $ErrorSet"
                return $false
            }
        } elseif ($Object.ObjectType -eq 'ApiManagement') {
            ## Create instance manually if not exist
            $ApiManagementParameters = $Object.Parameters.Parameters
            $ApiManagement = Get-AzApiManagement -ResourceGroupName $ResourceGroup -Name $ApiManagementParameters.serviceName.value -ErrorVariable ErrorGet -ErrorAction SilentlyContinue

            ## V2 SKUs (BasicV2, StandardV2) are not accepted by New-AzApiManagement -Sku
            ## (the cmdlet's [PsApiManagementSku] enum only knows the legacy SKUs:
            ## Developer/Standard/Premium/Basic/Consumption). For V2 we skip the manual
            ## cmdlet creation and let the subsequent New-AzResourceGroupDeployment do it
            ## via the ARM template (which accepts the SKU as a free-form string parameter).
            $SkuValue = "$($ApiManagementParameters.sku.value)"
            $IsV2Sku  = $SkuValue -match 'v2$'

            if (!$ApiManagement -and -not $IsV2Sku) {
                Write-Warning "Creating new Api Aanagement service. Name: $($ApiManagementParameters.serviceName.value)"
                Write-Warning 'This is a long-running action. It can take between 30 and 40 minutes to create and activate an API Management service.'
                $ApiManagement = New-AzApiManagement -ResourceGroupName $ResourceGroup -Name $ApiManagementParameters.serviceName.value -Location $ApiManagementParameters.location.value -Organization $ApiManagementParameters.publisherName.value -AdminEmail $ApiManagementParameters.publisherEmail.value -Sku $ApiManagementParameters.sku.value -SystemAssignedIdentity -ErrorVariable DeployError
                if ($DeployError) {
                    Write-Error "Failed to create Api Aanagement service. Error: $DeployError"
                    return $false
                }
            } elseif (!$ApiManagement -and $IsV2Sku) {
                Write-Verbose "API Management SKU '$SkuValue' detected — deferring creation to ARM template (cmdlet enum does not support V2 SKUs)."
                Write-Verbose 'This is a long-running action. It can take between 30 and 40 minutes to create and activate an API Management service.'
            }

            ## Set secret get permission for Api Management service (skipped for V2 — done post-deploy below)
            if ($ApiManagement) {
                Write-Verbose 'Set keyvault secret access for Api Management service'
                $Result = Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ResourceGroupName $ResourceGroup -ObjectId $ApiManagement.Identity.PrincipalId -PermissionsToSecrets Get -BypassObjectIdValidation -ErrorVariable ErrorSet
                if ($ErrorSet) {
                    Write-Error "Failed to set keyvault secret access for Api Management Service. Error: $ErrorSet"
                    return $false
                }
            } elseif ($IsV2Sku) {
                ## V2 fallback: the Az.ApiManagement cmdlet sometimes returns $null for V2 SKUs.
                ## If APIM actually exists in ARM, fetch its identity via REST and apply the
                ## keyvault access policy now (BEFORE the ARM deploy needs it for KV references).
                ## Without this the ARM deployment fails with: 'does not have secrets get permission
                ## on key vault' (Code:ValidationError) when resolving named values backed by KV refs.
                $apimSubId = (Get-AzContext).Subscription.Id
                $apimResUri = "/subscriptions/$apimSubId/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$($ApiManagementParameters.serviceName.value)`?api-version=2023-05-01-preview"
                $apimGet = Invoke-AzRestMethod -Path $apimResUri -Method GET -ErrorAction SilentlyContinue
                if ($apimGet -and $apimGet.StatusCode -ge 200 -and $apimGet.StatusCode -lt 300) {
                    $apimObj = $apimGet.Content | ConvertFrom-Json -Depth 20
                    if ($apimObj.identity -and $apimObj.identity.principalId) {
                        Write-Verbose "Set keyvault secret access for Api Management service (V2 REST fallback, principalId=$($apimObj.identity.principalId))"
                        $Result = Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ResourceGroupName $ResourceGroup -ObjectId $apimObj.identity.principalId -PermissionsToSecrets Get -BypassObjectIdValidation -ErrorVariable ErrorSet
                        if ($ErrorSet) {
                            Write-Error "Failed to set keyvault secret access for Api Management Service (V2 REST fallback). Error: $ErrorSet"
                            return $false
                        }
                    } else {
                        Write-Verbose 'V2 APIM has no system-assigned identity yet — keyvault access policy will be set post-deploy.'
                    }
                } else {
                    Write-Verbose 'V2 APIM does not exist yet — keyvault access policy will be set post-deploy after ARM creates it.'
                }
            }

            ## Update backend urls and re-create parameters file if needed
            $FunctionApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroup -Name $FunctionName
            $FunctionAppUrl = "https://$($FunctionApp.DefaultHostName)/api"
            if ($ApiManagementParameters.backendUrls.value.Where({ $_ -eq $FunctionAppUrl }).Count -eq 0) {
                $ApiManagementParameters.backendUrls.value += $FunctionAppUrl
                Write-Verbose -Message "Re-creating the Parameter file for $($Object.ObjectType) in the following location: $($Object.ParameterPath)"
                $ParameterFile = $Object.Parameters | ConvertTo-Json -Depth 8
                $ParameterFile | Out-File -FilePath $Object.ParameterPath -Force
            }
        }

        ## ARM deployment operations
        ## Creates the Azure Resource
        ## -ErrorAction SilentlyContinue + -ErrorVariable suppresses the noisy auto-rendered
        ## cmdlet error block (line numbers, giant correlation IDs); the caller logs a clean
        ## summary instead and the retry loop in Deploy-WinGetSource.ps1 handles recovery.
        $DeployResult = New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroup -TemplateFile $Object.TemplatePath -TemplateParameterFile $Object.ParameterPath -Mode Incremental -ErrorVariable DeployError -ErrorAction SilentlyContinue

        ## Verifies that no error occured when creating the Azure resource
        if ($DeployError -or ($DeployResult.ProvisioningState -ne 'Succeeded' -and $DeployResult.ProvisioningState -ne 'Created')) {
            ## Extract just the Status Message from the ARM error (drops correlationId, timestamps,
            ## and "Showing 1 out of 1 error(s)" boilerplate). The full error is still on $DeployError
            ## for callers that want it via -ErrorVariable.
            $errMsg = "$DeployError"
            if ($errMsg -match 'Status Message:\s*(.+?)(?:\s*Please provide correlationId|\s*\(Code:|\r|\n|$)') {
                $errMsg = $Matches[1].Trim()
            }
            $errCode = if ("$DeployError" -match '\(Code:([^)]+)\)') { $Matches[1] } else { 'DeploymentFailed' }
            $ErrReturnObject = @{
                DeployError  = $DeployError
                DeployResult = $DeployResult
            }

            Write-Verbose "ARM deploy failed: [$errCode] $errMsg"
            Write-Verbose "Full error: $DeployError"
            $null = $ErrReturnObject  # retained for diagnostic capture
            return $false
        }

        Write-Information -MessageData "$($Object.ObjectType) was successfully created."

        ## Post ARM deployment operations
        if ($Object.ObjectType -eq 'Function') {
            ## Publish GitHub Functions to newly created Azure Function
            $FunctionName = $Object.Parameters.Parameters.functionName.value
            $FunctionApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroup -Name $FunctionName
            $FunctionAppId = $FunctionApp.Id

            ## Assign necessary Azure roles
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Account Contributor' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Blob Data Owner' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Table Data Contributor' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Queue Data Contributor' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Queue Data Message Processor' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'Storage Queue Data Message Sender' -ResourceGroup $ResourceGroup -ResourceName $StorageAccountName -ResourceType 'Microsoft.Storage/storageAccounts')) { return $false }
            if (!$(Set-RoleAssignment -PrincipalId $FunctionApp.IdentityPrincipalId -RoleName 'App Configuration Data Reader' -ResourceGroup $ResourceGroup -ResourceName $AppConfigName -ResourceType 'Microsoft.AppConfiguration/configurationStores')) { return $false }

            ## Set keyvault secrets get permission
            Write-Verbose 'Set keyvault secret access for Azure Function'
            $Result = Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ResourceGroupName $ResourceGroup -ObjectId $FunctionApp.IdentityPrincipalId -PermissionsToSecrets Get -BypassObjectIdValidation -ErrorVariable ErrorSet
            if ($ErrorSet) {
                Write-Error "Failed to set keyvault secret access for Azure Function. Error: $ErrorSet"
                return $false
            }

            ## Assign cosmos db roles
            $CosmosAccount = Get-AzCosmosDBAccount -ResourceGroupName $ResourceGroup -Name $CosmosAccountName
            $RoleId = '00000000-0000-0000-0000-000000000002' ## Built in contributor role
            $RoleDefinitionId = "$($CosmosAccount.Id)/sqlRoleAssignments/$RoleId"
            if ((Get-AzCosmosDBSqlRoleAssignment -ResourceGroupName $ResourceGroup -AccountName $CosmosAccountName).Where({ $_.PrincipalId -eq $FunctionApp.IdentityPrincipalId -and $_.RoleDefinitionId -eq $RoleDefinitionId }).Count -eq 0) {
                Write-Verbose 'Assigning Cosmos DB Account contributor role'
                $Result = New-AzCosmosDBSqlRoleAssignment -AccountName $CosmosAccountName -ResourceGroupName $ResourceGroup -RoleDefinitionId $RoleId -Scope '/' -PrincipalId $FunctionApp.IdentityPrincipalId -ErrorVariable ErrorNew
                if ($ErrorNew) {
                    Write-Error "Failed to assign Azure Function with Cosmos DB contributor role. Error: $ErrorNew"
                    return $false
                }
            }

            ## Create Function app key and also add to keyvault
            $NewFunctionKeyValue = New-FunctionAppKey
            $Result = Invoke-AzRestMethod -Path "$FunctionAppId/host/default/functionKeys/WinGetRestSourceAccess?api-version=2024-04-01" -Method PUT -Payload (@{properties = @{name = 'WinGetRestSourceAccess'; value = $NewFunctionKeyValue } } | ConvertTo-Json -Depth 8)
            if ($Result.StatusCode -ne 200 -and $Result.StatusCode -ne 201) {
                Write-Error "Failed to create Azure Function key. $($Result.Content)"
                return $false
            }

            Write-Information -MessageData 'Add Function App host key to keyvault.'
            $Result = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $AzureFunctionHostKeyName -SecretValue (Get-SecureString($NewFunctionKeyValue)) -ErrorVariable ErrorSet
            if ($ErrorSet) {
                Write-Error "Failed to set keyvault secret. Name: $AzureFunctionHostKeyName Error: $ErrorSet"
                return $false
            }

            Write-Information -MessageData 'Publishing function files to the Azure Function.'
            $DeployResult = Publish-AzWebApp -ArchivePath $RestSourcePath -ResourceGroupName $ResourceGroup -Name $FunctionName -Force -ErrorVariable DeployError

            ## Verifies that no error occured when publishing the Function App
            if ($DeployError -or !$DeployResult) {
                $ErrReturnObject = @{
                    DeployError  = $DeployError
                    DeployResult = $DeployResult
                }

                Write-Error "Failed to publishing the Function App. Error: $DeployError" -TargetObject $ErrReturnObject
                return $false
            }

            ## Restart the Function App
            Write-Verbose 'Restarting Azure Function.'
            if (!$(Restart-AzFunctionApp -Name $FunctionName -ResourceGroupName $ResourceGroup -Force -PassThru)) {
                Write-Error "Failed to restart Function App. Name: $FunctionName"
                return $false
            }
        } elseif ($Object.ObjectType -eq 'ApiManagement') {
            ## V2 SKU post-deploy: ARM created APIM (cmdlet path was skipped above), so set
            ## the keyvault secret access policy now using the freshly-created APIM identity.
            $ApiManagementParameters = $Object.Parameters.Parameters
            $SkuValue = "$($ApiManagementParameters.sku.value)"
            if ($SkuValue -match 'v2$') {
                $apimPrincipalId = $null

                ## Try the cmdlet first (works for some V2 SKUs depending on Az.ApiManagement version)
                $ApiManagement = Get-AzApiManagement -ResourceGroupName $ResourceGroup -Name $ApiManagementParameters.serviceName.value -ErrorAction SilentlyContinue
                if ($ApiManagement -and $ApiManagement.Identity -and $ApiManagement.Identity.PrincipalId) {
                    $apimPrincipalId = $ApiManagement.Identity.PrincipalId
                } else {
                    ## Cmdlet failed or returned no identity — fall back to ARM REST.
                    ## Common for V2 SKUs because Get-AzApiManagement can't deserialize the SKU enum.
                    $apimSubId = (Get-AzContext).Subscription.Id
                    $apimResUri = "/subscriptions/$apimSubId/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$($ApiManagementParameters.serviceName.value)`?api-version=2023-05-01-preview"
                    $apimGet = Invoke-AzRestMethod -Path $apimResUri -Method GET -ErrorAction SilentlyContinue
                    if ($apimGet -and $apimGet.StatusCode -ge 200 -and $apimGet.StatusCode -lt 300) {
                        $apimObj = $apimGet.Content | ConvertFrom-Json -Depth 20
                        if ($apimObj.identity -and $apimObj.identity.principalId) {
                            $apimPrincipalId = $apimObj.identity.principalId
                        }
                    }
                }

                if ($apimPrincipalId) {
                    Write-Verbose "Set keyvault secret access for Api Management service (V2 post-deploy, principalId=$apimPrincipalId)"
                    $Result = Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ResourceGroupName $ResourceGroup -ObjectId $apimPrincipalId -PermissionsToSecrets Get -BypassObjectIdValidation -ErrorVariable ErrorSet
                    if ($ErrorSet) {
                        Write-Error "Failed to set keyvault secret access for Api Management Service (V2 post-deploy). Error: $ErrorSet"
                        return $false
                    }
                } else {
                    ## Pre-deploy V2 fallback already sets this in the wrapper script before ARM runs,
                    ## so a missing identity here is usually harmless. Demote from Warning to Verbose.
                    Write-Verbose 'V2 APIM identity not retrievable post-deploy via cmdlet or REST — pre-deploy KV access policy should already be in place.'
                }
            }
        }

        $Object.DeploymentSuccess = $true
    }

    return $true
}
