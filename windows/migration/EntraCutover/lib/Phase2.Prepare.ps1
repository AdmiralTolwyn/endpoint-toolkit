# Phase 2 - Prepare: reversible staging before the Teardown point-of-no-return.
# Dot-sourced by Invoke-EntraCutover.ps1; contract in that file's header block.
# $PSScriptRoot here is THIS file's own folder (lib\) even under dot-sourcing,
# so capture the tool root (repo checkout / staged Bin) once, at file scope.
$Script:ECToolRoot = Split-Path $PSScriptRoot -Parent

# --------------------------------------------------------------------------
# Private helpers (<Verb>-EC<Noun>)
# --------------------------------------------------------------------------

function Get-ECPropertyValue {
    <#
        Strict-mode-safe optional read. State.Device/Options fields are added
        incrementally across phases and round-trip through JSON (hashtable ->
        PSCustomObject) between runs, so a missing key must return $Default
        instead of throwing PropertyNotFoundStrict under Set-StrictMode -Version 2.
    #>
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [hashtable]) {
        if ($Object.ContainsKey($Name)) { return $Object[$Name] }
        return $Default
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $Default }
    return $prop.Value
}

function Set-ECStateValue {
    <#
        State sub-objects are hashtables before the first Save/reload and
        PSCustomObjects after (ConvertFrom-Json). Setting an existing
        PSCustomObject property works via dot-assignment; setting a NEW one
        (e.g. Device starts as {}) does not and needs Add-Member -Force.
    #>
    param([Parameter(Mandatory)]$Target, [Parameter(Mandatory)][string]$Name, $Value)
    if ($Target -is [hashtable]) {
        $Target[$Name] = $Value
        return
    }
    try { $Target.$Name = $Value }
    catch { $Target | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

function New-ECRandomPassword {
    <#
        24-char password from a CSPRNG (RNGCryptoServiceProvider), rerolled
        until all four complexity classes are present. Ambiguous glyphs
        (0/O/1/I/l) excluded so it can be transcribed by a human operator.
    #>
    param([int]$Length = 24)
    $upper  = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lower  = 'abcdefghijkmnpqrstuvwxyz'
    $digit  = '23456789'
    $symbol = '!@#$%^&*-_=+?'
    $all = $upper + $lower + $digit + $symbol
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    try {
        while ($true) {
            $bytes = New-Object byte[] $Length
            $rng.GetBytes($bytes)
            $chars = New-Object char[] $Length
            for ($i = 0; $i -lt $Length; $i++) { $chars[$i] = $all[$bytes[$i] % $all.Length] }
            $candidate = -join $chars
            if (($candidate -cmatch '[A-Z]') -and ($candidate -cmatch '[a-z]') -and
                ($candidate -match '[0-9]') -and ($candidate -match '[^A-Za-z0-9]')) {
                return $candidate
            }
        }
    }
    finally { $rng.Dispose() }
}

# --------------------------------------------------------------------------
# Phase entry point
# --------------------------------------------------------------------------

function Invoke-PhasePrepare {
    param([hashtable]$Ctx)

    Invoke-Step -Ctx $Ctx -Name 'Prepare.StageBin' -Action {
        foreach ($d in @($Ctx.Paths.Root, $Ctx.Paths.Bin, $Ctx.Paths.Backup, $Ctx.Paths.Rollback)) {
            if ($d -and -not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        }
        # Guard against self-copy when re-invoked from the already-staged Bin (Resume).
        if ($Script:ECToolRoot -ne $Ctx.Paths.Bin) {
            Copy-Item -Path (Join-Path $Script:ECToolRoot '*') -Destination $Ctx.Paths.Bin -Recurse -Force
        }

        $ppkgStaged = $false
        $srcPpkg = $Ctx.Options.PpkgPath
        if ($srcPpkg) {
            $destPpkg = Join-Path $Ctx.Paths.Root 'join.ppkg'
            Copy-Item -Path $srcPpkg -Destination $destPpkg -Force
            Set-ECStateValue -Target $Ctx.State.Options -Name 'PpkgPath' -Value $destPpkg
            Save-CutoverState -State $Ctx.State
            $ppkgStaged = $true
        }
        return @{ BinPath = $Ctx.Paths.Bin; PpkgStaged = $ppkgStaged }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.KfmGate' -Action {
        $kfm = Get-ECKfmStatus
        $healthy = [bool](Get-ECPropertyValue -Object $kfm -Name 'Healthy' -Default $false)
        if (-not $healthy) {
            if ($Ctx.Options.SkipKfmGate) {
                Write-Log 'KFM gate skipped by operator' 'WARN'
            }
            else {
                throw 'KFM not verified healthy - enable KFM and confirm sync, or use -SkipKfmGate to accept data risk'
            }
        }
        return $kfm
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.BreakGlass' -Action {
        $userName = $Script:BreakGlassUser
        $existing = Get-LocalUser -Name $userName -ErrorAction SilentlyContinue

        # Operator-supplied password (fleet/unattended): use it verbatim - nothing
        # is generated, echoed, or DPAPI-stored, because the operator already holds
        # it. Otherwise (attended) generate a random one and show it once.
        $bgCred = $null
        if ($Ctx.ContainsKey('BreakGlassCredential')) { $bgCred = $Ctx.BreakGlassCredential }
        $operatorSupplied = [bool]$bgCred

        $plain  = $null
        if ($operatorSupplied) {
            $secure = $bgCred.Password
        }
        else {
            $plain  = New-ECRandomPassword -Length 24
            $secure = ConvertTo-SecureString -String $plain -AsPlainText -Force
        }

        if (-not $existing) {
            New-LocalUser -Name $userName -Password $secure -FullName 'EntraCutover Break-Glass' `
                -Description 'Temporary local admin created by EntraCutover Prepare phase.' `
                -PasswordNeverExpires -AccountNeverExpires | Out-Null
            # SID form is locale-safe; the localized "Administrators" name is not.
            $adminGroup = Get-LocalGroup -SID 'S-1-5-32-544'
            Add-LocalGroupMember -Group $adminGroup -Member $userName -ErrorAction SilentlyContinue
            Write-Log "break-glass account '$userName' created and added to Administrators."
        }
        else {
            Set-LocalUser -Name $userName -Password $secure -PasswordNeverExpires $true
            Write-Log "break-glass account '$userName' already exists - password set (resume)."
        }

        if ($operatorSupplied) {
            # Do NOT persist an operator-held secret; record only the source.
            Set-ECStateValue -Target $Ctx.State.Device -Name 'BreakGlassSecret' -Value $null
            Set-ECStateValue -Target $Ctx.State.Device -Name 'BreakGlassSource' -Value 'Operator'
            Save-CutoverState -State $Ctx.State
            Write-Log "break-glass password set from -BreakGlassCredential (operator-held; not stored or displayed)."
            return @{ Account = $userName; Source = 'Operator' }
        }

        # Attended: DPAPI-store (machine scope, on-box recovery) and show once.
        $secret = Protect-Secret -Plain $plain
        Set-ECStateValue -Target $Ctx.State.Device -Name 'BreakGlassSecret' -Value $secret
        Set-ECStateValue -Target $Ctx.State.Device -Name 'BreakGlassSource' -Value 'Generated'
        Save-CutoverState -State $Ctx.State

        $bar = '*' * 70
        Write-Host ''
        Write-Host "  $bar" -ForegroundColor Yellow
        Write-Host '  BREAK-GLASS LOCAL ADMINISTRATOR - RECORD THIS NOW, SHOWN ONCE' -ForegroundColor Yellow
        Write-Host "  Account : $userName" -ForegroundColor Yellow
        Write-Host "  Password: $plain" -ForegroundColor Yellow
        Write-Host '  Recoverable on THIS box only (DPAPI, admin/SYSTEM). Store it securely now.' -ForegroundColor Yellow
        Write-Host "  $bar" -ForegroundColor Yellow
        Write-Host ''
        $plain = $null

        return @{ Account = $userName; Source = 'Generated' }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.RollbackBlob' -Action {
        $blobPath = Join-Path $Ctx.Paths.Rollback 'odj.blob'
        $domain   = Get-ECPropertyValue -Object $Ctx.State.Device -Name 'Domain'
        $computer = $env:COMPUTERNAME
        $djoin    = "$env:windir\System32\djoin.exe"
        $method   = 'None'

        if (-not $domain) {
            Write-Log 'no domain recorded in state - skipping djoin rollback blob capture.' 'WARN'
        }
        else {
            $djoinArgs = @('/provision', '/domain', $domain, '/machine', $computer, '/reuse', '/savefile', $blobPath)

            if ($Ctx.DomainCredential) {
                $method = 'Credential'
                $outFile = Join-Path $env:TEMP 'ec-djoin-out.txt'
                $errFile = Join-Path $env:TEMP 'ec-djoin-err.txt'
                Remove-Item -Path $outFile, $errFile -Force -ErrorAction SilentlyContinue
                try {
                    # Start-Process -Credential does NOT elevate; djoin /provision needs
                    # AD rights on the computer object (not local admin), so this is fine.
                    $proc = Start-Process -FilePath $djoin -ArgumentList $djoinArgs -Credential $Ctx.DomainCredential `
                        -RedirectStandardOutput $outFile -RedirectStandardError $errFile -NoNewWindow -Wait -PassThru -ErrorAction Stop
                    $stdout = if (Test-Path $outFile) { Get-Content -Path $outFile -Raw -ErrorAction SilentlyContinue } else { '' }
                    $stderr = if (Test-Path $errFile) { Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue } else { '' }
                    Write-Log ("djoin (credential) exit {0}: {1}" -f $proc.ExitCode, ("$stdout $stderr").Trim())
                }
                catch { Write-Log "djoin (credential) attempt failed: $($_.Exception.Message)" 'WARN' }
                finally { Remove-Item -Path $outFile, $errFile -Force -ErrorAction SilentlyContinue }
            }

            if (-not (Test-Path $blobPath)) {
                $method = 'CurrentContext'
                try {
                    $result = Invoke-Exe -Path $djoin -Arguments $djoinArgs
                    if ($result.ExitCode -ne 0) { Write-Log "djoin (current context) exit $($result.ExitCode)" 'WARN' }
                }
                catch { Write-Log "djoin (current context) attempt failed: $($_.Exception.Message)" 'WARN' }
            }
        }

        $present = [bool](Test-Path $blobPath)
        if ($present) {
            $method = if ($method -eq 'None') { 'CurrentContext' } else { $method }
            Write-Log '/reuse has RESET the AD computer account password - if you abort before Teardown, run -Mode Rollback to restore machine trust from the blob.' 'WARN'
        }
        else {
            $method = 'None'
            Write-Log 'no rollback blob: post-unjoin rollback will require manual domain rejoin with credentials' 'WARN'
        }

        return @{ BlobPresent = $present; Method = $method }
    } -AllowFail

    Invoke-Step -Ctx $Ctx -Name 'Prepare.RegistryBackups' -Action {
        $keys = @(
            'HKLM\SOFTWARE\Microsoft\Enrollments',
            'HKLM\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked',
            'HKLM\SOFTWARE\Microsoft\PolicyManager',
            'HKLM\SOFTWARE\Microsoft\Provisioning\OMADM',
            'HKLM\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin',
            'HKLM\SYSTEM\CurrentControlSet\Control\CloudDomainJoin'
        )
        $exported = 0
        foreach ($k in $keys) {
            if (New-RegistryBackup -KeyPath $k) { $exported++ }
        }
        return @{ Exported = $exported }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.BlockHybridRejoin' -Action {
        $wjKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin'
        $prevBlock = 'absent'
        $prevAuto  = 'absent'
        if (Test-Path $wjKey) {
            $item = Get-ItemProperty -Path $wjKey -ErrorAction SilentlyContinue
            if ($item -and ($item.PSObject.Properties.Name -contains 'BlockAADWorkplaceJoin')) { $prevBlock = [string]$item.BlockAADWorkplaceJoin }
            if ($item -and ($item.PSObject.Properties.Name -contains 'autoWorkplaceJoin')) { $prevAuto = [string]$item.autoWorkplaceJoin }
        }

        $taskPath = '\Microsoft\Windows\Workplace Join\'
        $taskName = 'Automatic-Device-Join'
        $taskExisted = $false
        $taskWasEnabled = $false
        $task = Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task) {
            $taskExisted = $true
            $taskWasEnabled = [bool]($task.State -ne 'Disabled')
            Disable-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction SilentlyContinue | Out-Null
        }

        if (-not (Test-Path $wjKey)) { New-Item -Path $wjKey -Force | Out-Null }
        Set-ItemProperty -Path $wjKey -Name 'BlockAADWorkplaceJoin' -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $wjKey -Name 'autoWorkplaceJoin' -Value 0 -Type DWord -Force

        return @{
            PrevBlockAADWorkplaceJoin = $prevBlock
            PrevAutoWorkplaceJoin     = $prevAuto
            TaskExisted               = $taskExisted
            TaskWasEnabled            = $taskWasEnabled
        }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.SuspendBitLocker' -Action {
        $wasProtected = $false
        try {
            $vol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
            if ($vol -and $vol.ProtectionStatus -eq 'On') {
                $wasProtected = $true
                Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 3 | Out-Null
                Write-Log "BitLocker suspended on $env:SystemDrive for up to 3 reboots."
            }
        }
        catch { Write-Log "BitLocker not available/queryable: $($_.Exception.Message)" 'WARN' }
        return @{ WasProtected = $wasProtected }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.LegalNotice' -Action {
        $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        $caption = ''
        $text = ''
        if (Test-Path $key) {
            $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
            if ($item -and ($item.PSObject.Properties.Name -contains 'legalnoticecaption') -and $item.legalnoticecaption) { $caption = $item.legalnoticecaption }
            if ($item -and ($item.PSObject.Properties.Name -contains 'legalnoticetext') -and $item.legalnoticetext) { $text = $item.legalnoticetext }
        }
        Set-MigrationNotice -Caption 'Device migration in progress' `
            -Text 'This device is being migrated to cloud management. Do NOT sign in until this notice is gone. Phase: Prepare complete.'
        return @{ OriginalCaption = $caption; OriginalText = $text }
    }

    Invoke-Step -Ctx $Ctx -Name 'Prepare.ResumeTask' -Action {
        Register-ResumeTask
        return $null
    }
}
