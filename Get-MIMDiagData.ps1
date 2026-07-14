<#
.SYNOPSIS
    MIM / FIM diagnostic data collector.

.VERSION
    Version：Get-MIMDiagData_v1_19.ps1
    Updated: v1.14 stores all collection output under the -Logpath output root, uses MIMLOG_<timestamp> as the root folder name, and avoids creating temporary working folders under Windows Temp. v1.13 fixes Metaverse extension DataTable handling and Windows PowerShell Out-File compatibility. v1.12 fixes Metaverse rules extension collection when running from the FIMService server by avoiding SQL double-hop authentication. v1.11 added Metaverse rules extension configuration/DLL collection and includes it in the HTML report. Previous updates: v1.10 removes MA XML parse-check output and fixes UTF-8 reading for MA XML Japanese OU names; quiet mode by default, service/portal build collection, SharePoint version collection, name resolution/DC inventory, robust MA XML summary parsing, optional PCNS collection, improved best-effort metaverse object extraction for OBJ mode, skips CONFIG/EVENTLOG folders in OBJ mode, and creates an HTML summary report.

.USAGE
    Note: MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。

    ALL mode:
      .\Get-MIMDiagData.ps1 -Logpath C:\Temp\MIMDiag -SyncServer WIN-U74H5QBQ28C -MimServiceUri http://WIN-M6F0SOQJ62G.contoso.com:5725/ResourceManagementService [-PcnSServer <DCName>]

    OBJ mode:
      .\Get-MIMDiagData.ps1 -Logpath C:\Temp\MIMDiag -SyncServer <SyncServer> -GetObjDomainName contoso.com -GetObjADdn "CN=user1,OU=SyncUsers,DC=contoso,DC=com" -DomainAdminName "CONTOSO\Administrator"

.NOTES
    - Run in elevated Windows PowerShell.
    - ALL mode collects data marked as *ALL in the design note.
    - OBJ mode collects data marked as *OBJ in the design note.
    - LithnetMIISAutomation is required on the Sync server for MA export and best-effort CS/MV object extraction.
    - FIMAutomation snap-in is required on the execution server for MIM Service resource export.
    - PCNS collection is optional. If PCNS is not installed, omit -PcnSServer or specify -SkipPCNS.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Logpath,

    [string]$SyncServer = $env:COMPUTERNAME,

    [string]$MimServiceUri,

    [string]$MimServiceServer,

    [string]$PcnSServer,

    [string]$GetObjDomainName,
    [string]$GetObjADdn,
    [string]$DomainAdminName,

    [int]$EventLogDays = 7,

    # Run only FIMService-side diagnostics. Use this on the FIMService server when WinRM to the FIMSynchronizationService server is not available.
    [switch]$FIMServiceOnly,

    # Run only FIMSynchronizationService-side diagnostics. Use this on the FIMSynchronizationService server when WinRM from the FIMService server is not available.
    [switch]$FIMSyncOnly,

    # Default behavior is quiet: errors are recorded to files but not printed in red.
    [switch]$ShowErrors,

    # Disable transcript when the console must be as quiet as possible.
    [switch]$NoTranscript,

    # Skip PCNS collection when Password Change Notification Service is not installed or not in scope.
    [switch]$SkipPCNS
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$ScriptStartTime = Get-Date
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path -LiteralPath $Logpath)) {
    New-Item -Path $Logpath -ItemType Directory -Force | Out-Null
}

$ObjectMode = -not [string]::IsNullOrWhiteSpace($GetObjADdn)

if ($FIMServiceOnly -and $FIMSyncOnly) {
    throw "-FIMServiceOnly and -FIMSyncOnly cannot be specified together."
}

$RootPrefix = "MIMLOG"
if ($FIMServiceOnly) { $RootPrefix = "MIMLOG_FIMSERVICE" }
elseif ($FIMSyncOnly) { $RootPrefix = "MIMLOG_FIMSYNC" }

$Root        = Join-Path $Logpath "${RootPrefix}_$TimeStamp"
$SyncDataDir = Join-Path $Root "SYNC DATA"
$ConfigDir   = Join-Path $Root "CONFIG"
$EventLogDir = Join-Path $Root "EVENTLOG"
$DiagDir     = Join-Path $Root "DIAGNOSTIC"

# OBJ mode collects only object-focused diagnostics.
# Do not create CONFIG / EVENTLOG folders in OBJ mode because they are not used.
$null = New-Item -Path $SyncDataDir -ItemType Directory -Force
$null = New-Item -Path $DiagDir     -ItemType Directory -Force
if (-not $ObjectMode) {
    $null = New-Item -Path $ConfigDir   -ItemType Directory -Force
    $null = New-Item -Path $EventLogDir -ItemType Directory -Force
}

$Global:DiagErrorCsv = Join-Path $DiagDir "Get-MIMDiagData_Errors.csv"
$Global:DiagLogTxt   = Join-Path $DiagDir "Get-MIMDiagData_Summary.txt"
$Global:DiagErrorDetailsTxt = Join-Path $DiagDir "Get-MIMDiagData_ErrorDetails.txt"

if ($ObjectMode) {
    $missing = @()
    if ([string]::IsNullOrWhiteSpace($GetObjDomainName)) { $missing += "GetObjDomainName" }
    if ([string]::IsNullOrWhiteSpace($DomainAdminName)) { $missing += "DomainAdminName" }
    if ($missing.Count -gt 0) { throw "OBJ mode requires: $($missing -join ', ')" }
}

# OBJ mode and FIMSyncOnly do not require FIMService ResourceManagementService URI.
if (-not $MimServiceUri -and -not $ObjectMode -and -not $FIMSyncOnly) {
    if ($MimServiceServer) {
        $MimServiceUri = "http://$MimServiceServer`:5725/ResourceManagementService"
    }
    else {
        $MimServiceUri = "http://$env:COMPUTERNAME`:5725/ResourceManagementService"
    }
}

$TranscriptPath = Join-Path $Root "Get-MIMDiagData_Transcript.txt"
$script:TranscriptStopped = $false
if (-not $NoTranscript) {
    try { Start-Transcript -Path $TranscriptPath -Force | Out-Null } catch {}
}
else {
    "Transcript disabled by -NoTranscript." | Out-File -FilePath $TranscriptPath -Encoding UTF8
    $script:TranscriptStopped = $true
}

function Write-DiagStatus {
    param([string]$Message,[ConsoleColor]$Color = [ConsoleColor]::Gray)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line -ForegroundColor $Color
    Add-Content -Path $Global:DiagLogTxt -Value $line -Encoding UTF8
}

function Limit-Text {
    param(
        [string]$Text,
        [int]$MaxLength = 1200
    )
    if ($null -eq $Text) { return $null }
    if ($Text.Length -le $MaxLength) { return $Text }
    return ($Text.Substring(0, $MaxLength) + " ... [truncated; see raw output files if available]")
}


function Read-TextFileBestEffort {
    param([Parameter(Mandatory=$true)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }

    # MA export XML files can contain Japanese OU names. Prefer UTF-8 because
    # Windows PowerShell 5.1 Get-Content defaults to ANSI and can cause mojibake.
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8Strict.GetString($bytes)
    }
    catch {
        return [System.Text.Encoding]::Default.GetString($bytes)
    }
}

function Write-DiagError {
    param([string]$Stage,[string]$Target,[System.Management.Automation.ErrorRecord]$ErrorRecord)

    $fullMessage = if ($ErrorRecord -and $ErrorRecord.Exception) { $ErrorRecord.Exception.Message } else { 'Unknown error' }
    $shortMessage = Limit-Text -Text $fullMessage -MaxLength 500
    $detail = @"
===== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====
Stage : $Stage
Target: $Target
ExceptionType: $(if ($ErrorRecord -and $ErrorRecord.Exception) { $ErrorRecord.Exception.GetType().FullName } else { '' })
Category: $(if ($ErrorRecord) { $ErrorRecord.CategoryInfo.ToString() } else { '' })
Message:
$fullMessage
ScriptStack:
$(if ($ErrorRecord) { $ErrorRecord.ScriptStackTrace } else { '' })

"@
    Add-Content -Path $Global:DiagErrorDetailsTxt -Value $detail -Encoding UTF8

    $row = [pscustomobject]@{
        TimeCreated   = Get-Date
        Stage         = $Stage
        Target        = $Target
        Message       = $shortMessage
        ExceptionType = if ($ErrorRecord -and $ErrorRecord.Exception) { $ErrorRecord.Exception.GetType().FullName } else { $null }
        Category      = if ($ErrorRecord) { $ErrorRecord.CategoryInfo.ToString() } else { $null }
        ScriptStack   = if ($ErrorRecord) { Limit-Text -Text $ErrorRecord.ScriptStackTrace -MaxLength 500 } else { $null }
    }
    if (-not (Test-Path $Global:DiagErrorCsv)) {
        $row | Export-Csv -Path $Global:DiagErrorCsv -NoTypeInformation -Encoding UTF8
    }
    else {
        $row | Export-Csv -Path $Global:DiagErrorCsv -NoTypeInformation -Encoding UTF8 -Append
    }

    $summaryLine = "[{0}] WARN recorded [{1}][{2}]. See {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Stage, $Target, $Global:DiagErrorCsv
    Add-Content -Path $Global:DiagLogTxt -Value $summaryLine -Encoding UTF8

    if ($ShowErrors) {
        Write-Host $summaryLine -ForegroundColor DarkYellow
    }
}

function New-SafeFileName {
    param([Parameter(Mandatory=$true)][string]$Name)
    $safe = $Name -replace '[\\/:*?"<>|\s]+','_'
    $safe = $safe.Trim('_')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "NoName" }
    return $safe
}

function Invoke-ExportFimConfigSafe {
    param(
        [Parameter(Mandatory=$true)][string]$CustomConfig,
        [switch]$OnlyBaseResources
    )
    $attempts = 3
    $lastError = $null
    for ($i = 1; $i -le $attempts; $i++) {
        try {
            Write-DiagStatus "Export-FIMConfig attempt $i/$attempts : $CustomConfig" DarkGray
            if ($OnlyBaseResources) {
                return @(Export-FIMConfig -Uri $MimServiceUri -OnlyBaseResources -CustomConfig $CustomConfig -ErrorAction Stop)
            }
            else {
                return @(Export-FIMConfig -Uri $MimServiceUri -CustomConfig $CustomConfig -ErrorAction Stop)
            }
        }
        catch {
            $lastError = $_
            Write-DiagError -Stage 'Invoke-ExportFimConfigSafe' -Target $CustomConfig -ErrorRecord $_
            if ($i -lt $attempts) { Start-Sleep -Seconds 2 }
        }
    }
    throw $lastError
}

function Test-IsLocalComputer {
    param([string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $true }
    $names = @('.', 'localhost', $env:COMPUTERNAME)
    try { $names += [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch {}
    return (($names | ForEach-Object { $_.ToLowerInvariant() }) -contains $ComputerName.ToLowerInvariant())
}

function Test-TcpPortQuiet {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][int]$Port,
        [int]$TimeoutMs = 3000
    )
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $ok) {
            $client.Close()
            return $false
        }
        $client.EndConnect($iar)
        $client.Close()
        return $true
    }
    catch {
        return $false
    }
}

function Get-SyncManagementAgentMap {
    param([string]$ComputerName)

    $scriptBlock = {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        $rows = @()
        try {
            Import-Module LithnetMIISAutomation -Force -ErrorAction Stop
            $rows = @(Get-ManagementAgent -ErrorAction Stop | ForEach-Object {
                $guidText = $null
                foreach ($propName in @('Guid','ID','Id','ObjectID')) {
                    if ($_.PSObject.Properties[$propName] -and $_.$propName) {
                        $guidText = $_.$propName.ToString().Trim('{}').ToUpperInvariant()
                        break
                    }
                }
                [pscustomobject]@{
                    Name     = $_.Name
                    Type     = $_.Type
                    Guid     = $guidText
                    GuidText = $guidText
                    Source   = 'LithnetMIISAutomation'
                }
            })
        }
        catch {
            try {
                $rows = @(Get-WmiObject -Namespace 'root\MicrosoftIdentityIntegrationServer' -Class 'MIIS_ManagementAgent' -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{
                        Name     = $_.Name
                        Type     = $_.Type
                        Guid     = $_.Guid
                        GuidText = $_.Guid.ToString().Trim('{}').ToUpperInvariant()
                        Source   = 'WMI local on Sync server'
                    }
                })
            }
            catch {
                [pscustomobject]@{
                    Name     = $null
                    Type     = $null
                    Guid     = $null
                    GuidText = $null
                    Source   = 'Failed'
                    Error    = $_.Exception.Message
                }
            }
        }
        return $rows
    }

    $result = @(Invoke-OnComputer -ComputerName $ComputerName -ScriptBlock $scriptBlock)
    return @($result | Where-Object { $_.Name })
}

function Invoke-OnComputer {
    param([string]$ComputerName,[scriptblock]$ScriptBlock,[object[]]$ArgumentList = @())
    if (Test-IsLocalComputer -ComputerName $ComputerName) { & $ScriptBlock @ArgumentList }
    else { Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction Stop -WarningAction SilentlyContinue }
}

function Copy-RemoteFolderToLocal {
    param([string]$ComputerName,[string]$RemotePath,[string]$LocalPath)
    New-Item -Path $LocalPath -ItemType Directory -Force | Out-Null
    if (Test-IsLocalComputer -ComputerName $ComputerName) {
        Copy-Item -Path (Join-Path $RemotePath '*') -Destination $LocalPath -Recurse -Force -ErrorAction Stop
        return
    }
    $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop -WarningAction SilentlyContinue
    try { Copy-Item -FromSession $session -Path (Join-Path $RemotePath '*') -Destination $LocalPath -Recurse -Force -ErrorAction Stop }
    finally { if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue } }
}



function New-RemoteWorkPath {
    param([Parameter(Mandatory=$true)][string]$Name)

    # Do not use Windows Temp for intermediate collection files.
    # Stage remote/local working files under the same -Logpath output root,
    # then copy the collected files to the final folder and clean up the stage folder.
    $safeName = New-SafeFileName $Name
    $workRoot = Join-Path $Root '_REMOTE_WORK'
    return (Join-Path $workRoot ("{0}_{1}" -f $safeName, $TimeStamp))
}

function Remove-RemoteTempFolder {
    param(
        [string]$ComputerName,
        [string]$RemotePath,
        [string]$Stage = 'CleanupRemoteTemp'
    )

    if ([string]::IsNullOrWhiteSpace($RemotePath)) { return }

    try {
        Invoke-OnComputer -ComputerName $ComputerName -ScriptBlock {
            param([string]$Path)
            if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            }
        } -ArgumentList @($RemotePath)
    }
    catch {
        Write-DiagError -Stage $Stage -Target "$ComputerName $RemotePath" -ErrorRecord $_
    }
}

function Get-FimAttrValues {
    param($Resource,[string]$Name)
    $attr = $Resource.ResourceManagementAttributes | Where-Object { $_.AttributeName -eq $Name } | Select-Object -First 1
    if ($null -eq $attr) { return @() }
    if ($attr.IsMultiValue) { return @($attr.Values | ForEach-Object { $_.ToString() }) }
    if ($null -ne $attr.Value) { return @($attr.Value.ToString()) }
    return @()
}

function Get-FimAttrValue {
    param($Resource,[string]$Name)
    $values = Get-FimAttrValues -Resource $Resource -Name $Name
    if ($values.Count -eq 0) { return $null }
    return ($values -join '; ')
}

function Convert-FimExportToObject {
    param($ExportObject)
    $rmo = $ExportObject.ResourceManagementObject
    $hash = [ordered]@{
        ObjectID    = $rmo.ObjectIdentifier
        ObjectGuid  = ($rmo.ObjectIdentifier -replace '^urn:uuid:','')
        ObjectType  = Get-FimAttrValue $rmo 'ObjectType'
        DisplayName = Get-FimAttrValue $rmo 'DisplayName'
    }
    foreach ($attr in $rmo.ResourceManagementAttributes) {
        if (-not $hash.Contains($attr.AttributeName)) { $hash[$attr.AttributeName] = Get-FimAttrValue $rmo $attr.AttributeName }
    }
    [pscustomobject]$hash
}

function Get-FimAttributeValueKind {
    param([string]$AttributeName,[string]$ValueText)
    if ([string]::IsNullOrWhiteSpace($ValueText)) { return 'Empty' }
    if ($AttributeName -in @('InitialFlow','PersistentFlow')) { return 'SynchronizationRuleFlowXml' }
    if ($ValueText -match '^urn:uuid:') { return 'MIMReferenceObjectID' }
    if ($ValueText -match '^[{][0-9a-fA-F-]{36}[}]$') { return 'GUID' }
    if ($ValueText -match '^<.+>$') { return 'XmlFragment' }
    return 'TextOrScalar'
}

function Get-FimAttributeFriendlyDescription {
    param([string]$AttributeName,[string]$ValueText)
    switch -Regex ($AttributeName) {
        '^IsMultiValue$' { return 'Internal field. True means this attribute can contain multiple values.' }
        '^HasReference$' { return 'Internal field. True means this value references another MIM Service resource.' }
        '^InitialFlow$' { return 'Initial attribute flow. Evaluated mainly when the connector space object is initially created/provisioned.' }
        '^PersistentFlow$' { return 'Persistent attribute flow. Evaluated during synchronization after the object exists.' }
        '^FlowType$' { return 'Synchronization rule direction/type. Commonly 1=Inbound, 2=Outbound.' }
        '^ConnectedSystem$' { return 'Connected Management Agent identifier. Use ManagementAgent_List_From_SyncServer.csv to resolve the MA name.' }
        '^ManagementAgentID$' { return 'Reference to the Management Agent resource in MIM Service.' }
        '^CreateConnectedSystemObject$' { return 'Whether the rule can create a connector space / connected directory object.' }
        '^CreateILMObject$' { return 'Whether the rule can create a metaverse object.' }
        '^DisconnectConnectedSystemObject$' { return 'Whether the connected system object is disconnected when the rule no longer applies.' }
        '^RelationshipCriteria$' { return 'Join / relationship criteria between metaverse and connector space attributes.' }
        '^msidmOutboundIsFilterBased$' { return 'Whether outbound synchronization rule scoping is filter based.' }
        default { return $null }
    }
}

function Convert-FimExportToKeyValueRows {
    param([string]$ResourceType,$ExportObject)
    $rmo = $ExportObject.ResourceManagementObject
    $displayName = Get-FimAttrValue $rmo 'DisplayName'
    foreach ($attr in $rmo.ResourceManagementAttributes) {
        $valueText = if ($attr.IsMultiValue) { ($attr.Values | ForEach-Object { $_.ToString() }) -join '; ' } else { if ($null -ne $attr.Value) { $attr.Value.ToString() } else { '' } }
        $isMultiMeaning = if ($attr.IsMultiValue) { 'True: Multiple values can be stored in this attribute.' } else { 'False: This is a single-value attribute.' }
        $hasReferenceMeaning = if ($attr.HasReference) { 'True: The value is a reference to another MIM Service resource.' } else { 'False: The value is not a MIM Service resource reference.' }
        [pscustomobject]@{
            ResourceType         = $ResourceType
            ObjectID             = $rmo.ObjectIdentifier
            DisplayName          = $displayName
            AttributeName        = $attr.AttributeName
            IsMultiValue         = $attr.IsMultiValue
            IsMultiValueMeaning  = $isMultiMeaning
            HasReference         = $attr.HasReference
            HasReferenceMeaning  = $hasReferenceMeaning
            ValueKind            = Get-FimAttributeValueKind -AttributeName $attr.AttributeName -ValueText $valueText
            AttributeDescription = Get-FimAttributeFriendlyDescription -AttributeName $attr.AttributeName -ValueText $valueText
            Value                = $valueText
        }
    }
}

function Normalize-GuidText { param([string]$Value) if ([string]::IsNullOrWhiteSpace($Value)) { return $null } return $Value.Trim().Trim('{').Trim('}').ToUpperInvariant() }

function Get-FlowTypeName {
    param($FlowType)
    switch ("$FlowType") { '0' { 'Inbound' } '1' { 'Outbound' } '2' { 'Inbound and Outbound' } default { "Unknown / $FlowType" } }
}

function Convert-FlowXmlToObject {
    param([string]$SyncRuleDisplayName,[string]$ManagementAgentName,[string]$ManagementAgentGuid,[string]$FlowSet,[int]$No,[string]$XmlText)
    $direction='Unknown'; $source=''; $destination=''; $allowsNull=''; $functionId=''; $functionExpression=''; $scoping=''
    try {
        [xml]$xml = $XmlText
        $root = $xml.DocumentElement
        if ($root.Name -eq 'export-flow') { $direction = 'Outbound / export-flow' }
        elseif ($root.Name -eq 'import-flow') { $direction = 'Inbound / import-flow' }
        $allowsNull = $root.GetAttribute('allows-null')
        $destNode = $root.SelectSingleNode('dest'); if ($destNode) { $destination = $destNode.InnerText }
        $scopeNode = $root.SelectSingleNode('scoping'); if ($scopeNode) { $scoping = $scopeNode.InnerXml }
        $srcNode = $root.SelectSingleNode('src')
        $srcItems = @()
        if ($srcNode) {
            $attrNodes = $srcNode.SelectNodes('attr')
            if ($attrNodes.Count -gt 0) { foreach ($attrNode in $attrNodes) { $srcItems += $attrNode.InnerText } }
            elseif (-not [string]::IsNullOrWhiteSpace($srcNode.InnerText)) { $srcItems += $srcNode.InnerText }
        }
        $source = $srcItems -join ', '
        $fnNode = $root.SelectSingleNode('fn')
        if ($fnNode) {
            $functionId = $fnNode.GetAttribute('id')
            $args = @(); foreach ($arg in $fnNode.SelectNodes('arg')) { $args += $arg.InnerText }
            $functionExpression = $args -join ' + '
        }
    }
    catch { $direction='ParseError' }
    [pscustomobject]@{
        SyncRuleDisplayName=$SyncRuleDisplayName; ManagementAgentName=$ManagementAgentName; ManagementAgentGuid=$ManagementAgentGuid
        FlowSet=$FlowSet; No=$No; Direction=$direction; AllowsNull=$allowsNull; Source=$source; Destination=$destination
        FunctionId=$functionId; FunctionExpression=$functionExpression; Scoping=$scoping; RawXml=$XmlText
    }
}

function Find-ReferenceInObject {
    param([pscustomobject]$Object,[string[]]$Needles)
    $matched = @()
    foreach ($prop in $Object.PSObject.Properties) {
        $value = [string]$prop.Value
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        foreach ($needle in $Needles) {
            if ([string]::IsNullOrWhiteSpace($needle)) { continue }
            if ($value -like "*$needle*") { $matched += $prop.Name; break }
        }
    }
    return ($matched | Sort-Object -Unique)
}

function Get-ObjectDisplayNameFromRef {
    param([string]$ReferenceValue,[hashtable]$Dictionary)
    if ([string]::IsNullOrWhiteSpace($ReferenceValue)) { return $null }
    $refs = $ReferenceValue -split '; ' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $names = foreach ($ref in $refs) {
        $guid = $ref -replace '^urn:uuid:',''
        if ($Dictionary.ContainsKey($ref)) { $Dictionary[$ref].DisplayName }
        elseif ($Dictionary.ContainsKey($guid)) { $Dictionary[$guid].DisplayName }
        else { $ref }
    }
    return ($names -join '; ')
}

function Get-ChildText {
    param([System.Xml.XmlNode]$Node,[string]$XPath)
    $n = $Node.SelectSingleNode($XPath)
    if ($null -eq $n) { return $null }
    return ($n.InnerText -replace "`r|`n", ' ').Trim()
}

function Convert-CsExportToSummaryCsv {
    param([string]$XmlPath,[string]$CsvPath,[string]$Type,[string]$MAName)
    if (-not (Test-Path $XmlPath)) { return }
    try { [xml]$doc = Read-TextFileBestEffort -Path $XmlPath } catch { Write-DiagError -Stage 'Convert-CsExportToSummaryCsv' -Target $XmlPath -ErrorRecord $_; return }
    $rows = foreach ($obj in $doc.SelectNodes('//cs-object')) {
        $delta = $obj.SelectSingleNode('unapplied-export/delta')
        $err = $obj.SelectSingleNode('export-errordetail/export-error')
        $changedAttrs = @()
        if ($delta) {
            foreach ($attr in $delta.SelectNodes('attr')) {
                $attrName = $attr.GetAttribute('name'); $attrOp = $attr.GetAttribute('operation')
                $values = @($attr.SelectNodes('value') | ForEach-Object { ($_.InnerText -replace "`r|`n", ' ').Trim() })
                $changedAttrs += "$attrName[$attrOp]=$($values -join ';')"
            }
        }
        [pscustomobject]@{
            Type=$Type; MAName=$MAName; CsDn=$obj.GetAttribute('cs-dn'); ObjectType=$obj.GetAttribute('object-type'); ObjectId=$obj.GetAttribute('id')
            AccountName=Get-ChildText -Node $obj -XPath 'account-name'; UserPrincipalName=Get-ChildText -Node $obj -XPath 'user-principal-name'
            DomainName=Get-ChildText -Node $obj -XPath 'domain-name'; PartitionName=Get-ChildText -Node $obj -XPath 'partition-name'
            PendingDn=if($delta){$delta.GetAttribute('dn')}else{$null}; PendingOperation=if($delta){$delta.GetAttribute('operation')}else{$null}
            PrimaryObjectType=Get-ChildText -Node $obj -XPath 'unapplied-export/delta/primary-objectclass'; ChangedAttributes=$changedAttrs -join ' | '
            ErrorType=if($err){$err.GetAttribute('error-type')}else{$null}; CdError=if($err){$err.GetAttribute('cd-error')}else{$null}
            ErrorCode=if($err){$err.GetAttribute('error-code')}else{$null}; ErrorLiteral=if($err){$err.GetAttribute('error-literal')}else{$null}
            ServerErrorDetail=if($err){$err.GetAttribute('server-error-detail')}else{$null}; FirstOccurred=if($err){$err.GetAttribute('first-occurred')}else{$null}
            DateOccurred=if($err){$err.GetAttribute('date-occurred')}else{$null}; RetryCount=if($err){$err.GetAttribute('retry-count')}else{$null}
        }
    }
    $rows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
}

function Export-SyncDataAll {
    param([string]$OutDir)
    Write-DiagStatus "Collecting SYNC DATA from $SyncServer" Cyan
    $localOut = Join-Path $OutDir 'CSExport'
    New-Item -Path $localOut -ItemType Directory -Force | Out-Null
    $remoteTemp = New-RemoteWorkPath -Name "CSExport"
    $scriptBlock = {
        param($RemoteOut)
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        $ErrorActionPreference = 'Stop'
        $WarningPreference = 'SilentlyContinue'
        $ProgressPreference = 'SilentlyContinue'
        New-Item -Path $RemoteOut -ItemType Directory -Force | Out-Null
        $CandidateBinDirs = @('C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Bin','C:\Program Files\Microsoft Azure AD Sync\Bin')
        $BinDir = $CandidateBinDirs | Where-Object { Test-Path (Join-Path $_ 'csexport.exe') } | Select-Object -First 1
        if (-not $BinDir) { throw 'csexport.exe was not found.' }
        $CsExport = Join-Path $BinDir 'csexport.exe'
        $Analyzer = Join-Path $BinDir 'CSExportAnalyzer.exe'
        try {
            Import-Module LithnetMIISAutomation -Force -ErrorAction Stop
            $mas = @(Get-ManagementAgent -ErrorAction Stop | Select-Object Name,Type)
        }
        catch {
            try {
                $mas = @(Get-WmiObject -Namespace 'root\MicrosoftIdentityIntegrationServer' -Class 'MIIS_ManagementAgent' -ErrorAction Stop | Select-Object Name,Type)
            }
            catch {
                throw "Failed to enumerate Management Agents by Lithnet and WMI. $($_.Exception.Message)"
            }
        }
        $filters = @(
            [pscustomobject]@{Name='PendingExport';Filter='/f:x'},
            [pscustomobject]@{Name='ExportError';Filter='/f:e'},
            [pscustomobject]@{Name='PendingImport';Filter='/f:i'},
            [pscustomobject]@{Name='ImportError';Filter='/f:m'}
        )
        $summary = foreach ($ma in $mas) {
            $safeMa = ($ma.Name -replace '[\\/:*?"<>|\s]+','_').Trim('_')
            $maDir = Join-Path $RemoteOut $safeMa
            New-Item -Path $maDir -ItemType Directory -Force | Out-Null
            foreach ($filter in $filters) {
                $xml = Join-Path $maDir "$safeMa`_$($filter.Name).xml"
                $csv = Join-Path $maDir "$safeMa`_$($filter.Name)_Analyzer.csv"
                $stdout = Join-Path $maDir "$safeMa`_$($filter.Name)_stdout.txt"
                $stderr = Join-Path $maDir "$safeMa`_$($filter.Name)_stderr.txt"
                $status='Unknown'; $message=$null
                try {
                    $p = Start-Process -FilePath $CsExport -ArgumentList @($ma.Name,$xml,$filter.Filter) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
                    if ($p.ExitCode -eq 0) { $status='Success'; if (Test-Path $Analyzer) { & $Analyzer $xml | Out-File -FilePath $csv -Encoding UTF8 } }
                    else { $status='Failed'; $message="csexport exit code: $($p.ExitCode)" }
                }
                catch { $status='Failed'; $message=$_.Exception.Message }
                [pscustomobject]@{MAName=$ma.Name;MAType=$ma.Type;ExportType=$filter.Name;Filter=$filter.Filter;Status=$status;Message=$message;XmlPath=$xml;AnalyzerCsv=$csv}
            }
        }
        $summary | Export-Csv -Path (Join-Path $RemoteOut 'CSExport_RunSummary.csv') -NoTypeInformation -Encoding UTF8
    }
    try {
        Invoke-OnComputer -ComputerName $SyncServer -ScriptBlock $scriptBlock -ArgumentList @($remoteTemp)
        Copy-RemoteFolderToLocal -ComputerName $SyncServer -RemotePath $remoteTemp -LocalPath $localOut
        Get-ChildItem -Path $localOut -Filter '*.xml' -Recurse | ForEach-Object {
            $type = if ($_.BaseName -match 'PendingExport') {'PendingExport'} elseif ($_.BaseName -match 'ExportError') {'ExportError'} elseif ($_.BaseName -match 'PendingImport') {'PendingImport'} elseif ($_.BaseName -match 'ImportError') {'ImportError'} else {'Unknown'}
            $maName = Split-Path $_.DirectoryName -Leaf
            Convert-CsExportToSummaryCsv -XmlPath $_.FullName -CsvPath (Join-Path $_.DirectoryName "$($_.BaseName)_Summary.csv") -Type $type -MAName $maName
        }
    }
    catch { Write-DiagError -Stage 'Export-SyncDataAll' -Target $SyncServer -ErrorRecord $_ }
    finally { Remove-RemoteTempFolder -ComputerName $SyncServer -RemotePath $remoteTemp -Stage 'Export-SyncDataAll-CleanupRemoteTemp' }
}

function Export-ManagementAgentConfigs {
    param([string]$OutDir)
    Write-DiagStatus "Exporting MA configurations from $SyncServer" Cyan
    $localOut = Join-Path $OutDir 'ManagementAgents'
    New-Item -Path $localOut -ItemType Directory -Force | Out-Null
    $remoteTemp = New-RemoteWorkPath -Name "MAConfig"
    $scriptBlock = {
        param($RemoteOut)
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        $ErrorActionPreference = 'Stop'
        $WarningPreference = 'SilentlyContinue'
        $ProgressPreference = 'SilentlyContinue'
        Import-Module LithnetMIISAutomation -Force
        New-Item -Path $RemoteOut -ItemType Directory -Force | Out-Null
        Get-ChildItem -Path $RemoteOut -File -Include '*.xml','MA_Export_Summary.csv','MA_Export_Errors.csv' -Recurse:$false -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        $used = @{}
        $errors = @()
        $result = foreach ($MA in Get-ManagementAgent) {
            $safe = ($MA.Name -replace '[\/:*?"<>|\s]+','_').Trim('_')
            if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'NoName' }
            if ($used.ContainsKey($safe)) { $safe = "{0}_{1}" -f $safe,([guid]::NewGuid().ToString('N').Substring(0,8)) }
            $used[$safe] = $true
            $xml = Join-Path $RemoteOut "$safe.xml"
            try {
                Export-ManagementAgent -MA $MA.Name -File $xml 2>$null
                [pscustomobject]@{MAName=$MA.Name;MAType=$MA.Type;XmlPath=$xml;Status='Success';Message=$null}
            }
            catch {
                $msg = $_.Exception.Message
                if ($msg -and $msg.Length -gt 1000) { $msg = $msg.Substring(0,1000) + ' ... [truncated]' }
                $errors += [pscustomobject]@{MAName=$MA.Name;MAType=$MA.Type;XmlPath=$xml;Status='Failed';Message=$msg}
                [pscustomobject]@{MAName=$MA.Name;MAType=$MA.Type;XmlPath=$xml;Status='Failed';Message='Export failed. See MA_Export_Errors.csv.'}
            }
        }
        $result | Export-Csv -Path (Join-Path $RemoteOut 'MA_Export_Summary.csv') -NoTypeInformation -Encoding UTF8
        $errors | Export-Csv -Path (Join-Path $RemoteOut 'MA_Export_Errors.csv') -NoTypeInformation -Encoding UTF8
    }
    try {
        Invoke-OnComputer -ComputerName $SyncServer -ScriptBlock $scriptBlock -ArgumentList @($remoteTemp)
        Copy-RemoteFolderToLocal -ComputerName $SyncServer -RemotePath $remoteTemp -LocalPath $localOut

        # Some exported MA files can contain XML fragments that fail strict [xml] parsing.
        # Therefore, this script does not create MA_XML_Parse_Check.csv anymore.
        # Summary extraction below uses UTF-8-aware text reading + regex and continues per file.
        $ouRows = foreach ($file in Get-ChildItem -Path $localOut -Filter '*.xml') {
            try {
                $raw = Read-TextFileBestEffort -Path $file.FullName
                $maName = ([regex]::Match($raw,'(?s)<ma-data>.*?<name>(.*?)</name>').Groups[1].Value)
                $category = ([regex]::Match($raw,'(?s)<category>(.*?)</category>').Groups[1].Value)
                if ($category -eq 'AD') {
                    foreach ($m in [regex]::Matches($raw,'(?s)<inclusion>(.*?)</inclusion>')) {
                        [pscustomobject]@{MAName=$maName;Category=$category;ScopeType='Inclusion';DistinguishedName=[System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim());XmlFile=$file.FullName}
                    }
                    foreach ($m in [regex]::Matches($raw,'(?s)<exclusion>(.*?)</exclusion>')) {
                        [pscustomobject]@{MAName=$maName;Category=$category;ScopeType='Exclusion';DistinguishedName=[System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim());XmlFile=$file.FullName}
                    }
                }
            } catch {
                Write-DiagError -Stage 'Export-ManagementAgentConfigs-OURegex' -Target $file.FullName -ErrorRecord $_
            }
        }
        $ouRows | Export-Csv -Path (Join-Path $localOut 'ADMA_OU_Scope.csv') -NoTypeInformation -Encoding UTF8

        $connectionRows = foreach ($file in Get-ChildItem -Path $localOut -Filter '*.xml') {
            try {
                $raw = Read-TextFileBestEffort -Path $file.FullName
                [pscustomobject]@{
                    MAName        = ([regex]::Match($raw,'(?s)<ma-data>.*?<name>(.*?)</name>').Groups[1].Value)
                    Category      = ([regex]::Match($raw,'(?s)<category>(.*?)</category>').Groups[1].Value)
                    CsvSampleFile = ([regex]::Match($raw,'(?s)<sample_file>(.*?)</sample_file>').Groups[1].Value)
                    SqlServer     = ([regex]::Match($raw,'(?s)<server>(.*?)</server>').Groups[1].Value)
                    SqlDatabase   = ([regex]::Match($raw,'(?s)<databasename>(.*?)</databasename>').Groups[1].Value)
                    SqlTable      = ([regex]::Match($raw,'(?s)<tablename>(.*?)</tablename>').Groups[1].Value)
                    XmlFile       = $file.FullName
                }
            } catch {
                Write-DiagError -Stage 'Export-ManagementAgentConfigs-ConnectionRegex' -Target $file.FullName -ErrorRecord $_
            }
        }
        $connectionRows | Export-Csv -Path (Join-Path $localOut 'MA_Connection_Summary.csv') -NoTypeInformation -Encoding UTF8
    }
    catch { Write-DiagError -Stage 'Export-ManagementAgentConfigs' -Target $SyncServer -ErrorRecord $_ }
    finally { Remove-RemoteTempFolder -ComputerName $SyncServer -RemotePath $remoteTemp -Stage 'Export-ManagementAgentConfigs-CleanupRemoteTemp' }
}

function Export-MimResourceCollection {
    param([string]$ResourceType,[string]$OutPrefix,[string]$OutDir)
    Write-DiagStatus "Exporting MIM Service resource: $ResourceType" Cyan
    $items = @(Invoke-ExportFimConfigSafe -CustomConfig "/$ResourceType" -OnlyBaseResources)
    $flat = foreach ($item in $items) { Convert-FimExportToObject $item }
    $kv = foreach ($item in $items) { Convert-FimExportToKeyValueRows -ResourceType $ResourceType -ExportObject $item }
    $flat | Sort-Object DisplayName | Export-Csv -Path (Join-Path $OutDir "$OutPrefix`_Summary.csv") -NoTypeInformation -Encoding UTF8
    $kv | Sort-Object DisplayName,AttributeName | Export-Csv -Path (Join-Path $OutDir "$OutPrefix`_AllAttributes.csv") -NoTypeInformation -Encoding UTF8
    return $flat
}

function Export-SynchronizationRuleDetails {
    param([string]$OutDir,$MaMap)
    $srDir = Join-Path $OutDir 'SynchronizationRules'
    New-Item -Path $srDir -ItemType Directory -Force | Out-Null
    Write-DiagStatus 'Exporting SynchronizationRule details' Cyan

    try {
        $allRules = @(Invoke-ExportFimConfigSafe -CustomConfig '/SynchronizationRule' -OnlyBaseResources)
    }
    catch {
        $_ | Out-File -FilePath (Join-Path $srDir 'SynchronizationRules_ExportError.txt') -Encoding UTF8
        Write-DiagError -Stage 'Export-SynchronizationRuleDetails' -Target '/SynchronizationRule' -ErrorRecord $_
        return @()
    }

    [pscustomobject]@{
        MimServiceUri = $MimServiceUri
        ExportedRuleCount = $allRules.Count
        ExportTime = Get-Date
    } | Export-Csv -Path (Join-Path $srDir 'SynchronizationRules_ExportSummary.csv') -NoTypeInformation -Encoding UTF8

    if ($allRules.Count -eq 0) {
        'SynchronizationRule export succeeded, but zero rules were returned.' | Out-File -FilePath (Join-Path $srDir 'SynchronizationRules_ExportResult.txt') -Encoding UTF8
        return @()
    }

    @(
        [pscustomobject]@{FieldName='IsMultiValue';FalseMeaning='Single-value attribute';TrueMeaning='Multi-value attribute';Note='This describes the MIM Service schema property of the attribute, not whether the current value is true or false.'},
        [pscustomobject]@{FieldName='HasReference';FalseMeaning='Normal scalar/text value';TrueMeaning='Reference to another MIM Service resource';Note='When true, the Value usually contains an urn:uuid or object reference.'},
        [pscustomobject]@{FieldName='ValueKind';FalseMeaning='N/A';TrueMeaning='N/A';Note='Friendly classification added by this script, such as TextOrScalar, MIMReferenceObjectID, XmlFragment, or SynchronizationRuleFlowXml.'},
        [pscustomobject]@{FieldName='InitialFlow';FalseMeaning='N/A';TrueMeaning='N/A';Note='Initial attribute flows. Usually evaluated on initial provisioning/creation.'},
        [pscustomobject]@{FieldName='PersistentFlow';FalseMeaning='N/A';TrueMeaning='N/A';Note='Persistent attribute flows. Usually evaluated during ongoing synchronization.'}
    ) | Export-Csv -Path (Join-Path $srDir 'SynchronizationRule_Attributes_FieldGuide.csv') -NoTypeInformation -Encoding UTF8

    $ruleList = foreach ($rule in $allRules) {
        $rmo = $rule.ResourceManagementObject
        $csRaw = Get-FimAttrValue $rmo 'ConnectedSystem'; $csGuid = Normalize-GuidText $csRaw
        $ma = $MaMap | Where-Object { $_.GuidText -eq $csGuid } | Select-Object -First 1
        $flowTypeRaw = Get-FimAttrValue $rmo 'FlowType'
        [pscustomobject]@{
            SyncRuleDisplayName=Get-FimAttrValue $rmo 'DisplayName'; SyncRuleObjectID=$rmo.ObjectIdentifier
            ManagementAgentName=$ma.Name; ManagementAgentType=$ma.Type; ManagementAgentGuid=$ma.GuidText
            ConnectedSystemRaw=$csRaw; ConnectedSystemGuid=$csGuid; ConnectedObjectType=Get-FimAttrValue $rmo 'ConnectedObjectType'
            ILMObjectType=Get-FimAttrValue $rmo 'ILMObjectType'; FlowTypeRaw=$flowTypeRaw; FlowTypeName=Get-FlowTypeName $flowTypeRaw
            CreateConnectedSystemObject=Get-FimAttrValue $rmo 'CreateConnectedSystemObject'; CreateILMObject=Get-FimAttrValue $rmo 'CreateILMObject'
            DisconnectConnectedSystemObject=Get-FimAttrValue $rmo 'DisconnectConnectedSystemObject'; Precedence=Get-FimAttrValue $rmo 'Precedence'
            msidmOutboundIsFilterBased=Get-FimAttrValue $rmo 'msidmOutboundIsFilterBased'
        }
    }
    $ruleList | Sort-Object ManagementAgentName,SyncRuleDisplayName | Export-Csv -Path (Join-Path $srDir 'All_SynchronizationRules.csv') -NoTypeInformation -Encoding UTF8
    $allFlowRows = @()
    foreach ($rule in ($allRules | Sort-Object { Get-FimAttrValue $_.ResourceManagementObject 'DisplayName' })) {
        $rmo = $rule.ResourceManagementObject
        $displayName = Get-FimAttrValue $rmo 'DisplayName'
        $safe = New-SafeFileName $displayName
        $csRaw = Get-FimAttrValue $rmo 'ConnectedSystem'; $csGuid = Normalize-GuidText $csRaw
        $ma = $MaMap | Where-Object { $_.GuidText -eq $csGuid } | Select-Object -First 1
        Convert-FimExportToKeyValueRows -ResourceType 'SynchronizationRule' -ExportObject $rule | Export-Csv -Path (Join-Path $srDir "$safe`_Attributes.csv") -NoTypeInformation -Encoding UTF8
        $flowRows = @()
        foreach ($attrName in @('InitialFlow','PersistentFlow')) {
            $attr = $rmo.ResourceManagementAttributes | Where-Object { $_.AttributeName -eq $attrName } | Select-Object -First 1
            if ($null -eq $attr -or $null -eq $attr.Values) { continue }
            $i=0
            foreach ($flow in $attr.Values) {
                if ($null -eq $flow) { continue }
                $i++
                $fo = Convert-FlowXmlToObject -SyncRuleDisplayName $displayName -ManagementAgentName $ma.Name -ManagementAgentGuid $ma.GuidText -FlowSet $attrName -No $i -XmlText $flow.ToString()
                $flowRows += $fo; $allFlowRows += $fo
            }
        }
        $flowRows | Export-Csv -Path (Join-Path $srDir "$safe`_AttributeFlows.csv") -NoTypeInformation -Encoding UTF8
    }
    $allFlowRows | Export-Csv -Path (Join-Path $srDir 'All_SynchronizationRule_AttributeFlows.csv') -NoTypeInformation -Encoding UTF8
    return $ruleList
}

function Export-SyncRuleWorkflowMprSetMap {
    param([string]$OutDir,$SyncRules,$Workflows,$Mprs,$Sets)
    Write-DiagStatus 'Creating SynchronizationRule / Workflow / MPR / Set map' Cyan
    $setDict = @{}
    foreach ($set in $Sets) { $setDict[$set.ObjectID]=$set; $setDict[$set.ObjectGuid]=$set }
    $map = foreach ($sr in ($SyncRules | Sort-Object DisplayName)) {
        $srNeedles = @($sr.ObjectID,$sr.ObjectGuid,$sr.DisplayName) | Where-Object { $_ }
        $relatedWorkflows = foreach ($wf in $Workflows) {
            $matched = Find-ReferenceInObject -Object $wf -Needles $srNeedles
            if ($matched.Count -gt 0) { [pscustomobject]@{WorkflowObject=$wf;WorkflowMatchedAttrs=($matched -join '; ')} }
        }
        if (-not $relatedWorkflows -or $relatedWorkflows.Count -eq 0) {
            [pscustomobject]@{SyncRuleDisplayName=$sr.DisplayName;SyncRuleObjectID=$sr.ObjectID;SyncRuleFlowType=$sr.FlowType;SyncRuleConnectedSystem=$sr.ConnectedSystem;WorkflowDisplayName=$null;WorkflowObjectID=$null;WorkflowMatchedAttrs=$null;MPRDisplayName=$null;MPRObjectID=$null;MPRDisabled=$null;MPRActionType=$null;MPRMatchedAttrs=$null;PrincipalSetDisplayName=$null;PrincipalSetObjectID=$null;ResourceCurrentSetName=$null;ResourceCurrentSetID=$null;ResourceFinalSetName=$null;ResourceFinalSetID=$null}
            continue
        }
        foreach ($rwf in $relatedWorkflows) {
            $wf = $rwf.WorkflowObject
            $wfNeedles = @($wf.ObjectID,$wf.ObjectGuid,$wf.DisplayName) | Where-Object { $_ }
            $relatedMprs = foreach ($mpr in $Mprs) {
                $m = Find-ReferenceInObject -Object $mpr -Needles $wfNeedles
                if ($m.Count -gt 0) { [pscustomobject]@{MPRObject=$mpr;MPRMatchedAttrs=($m -join '; ')} }
            }
            if (-not $relatedMprs -or $relatedMprs.Count -eq 0) {
                [pscustomobject]@{SyncRuleDisplayName=$sr.DisplayName;SyncRuleObjectID=$sr.ObjectID;SyncRuleFlowType=$sr.FlowType;SyncRuleConnectedSystem=$sr.ConnectedSystem;WorkflowDisplayName=$wf.DisplayName;WorkflowObjectID=$wf.ObjectID;WorkflowMatchedAttrs=$rwf.WorkflowMatchedAttrs;MPRDisplayName=$null;MPRObjectID=$null;MPRDisabled=$null;MPRActionType=$null;MPRMatchedAttrs=$null;PrincipalSetDisplayName=$null;PrincipalSetObjectID=$null;ResourceCurrentSetName=$null;ResourceCurrentSetID=$null;ResourceFinalSetName=$null;ResourceFinalSetID=$null}
                continue
            }
            foreach ($rmpr in $relatedMprs) {
                $mpr = $rmpr.MPRObject
                [pscustomobject]@{
                    SyncRuleDisplayName=$sr.DisplayName; SyncRuleObjectID=$sr.ObjectID; SyncRuleFlowType=$sr.FlowType; SyncRuleConnectedSystem=$sr.ConnectedSystem
                    WorkflowDisplayName=$wf.DisplayName; WorkflowObjectID=$wf.ObjectID; WorkflowMatchedAttrs=$rwf.WorkflowMatchedAttrs
                    MPRDisplayName=$mpr.DisplayName; MPRObjectID=$mpr.ObjectID; MPRDisabled=$mpr.Disabled; MPRActionType=$mpr.ActionType; MPRMatchedAttrs=$rmpr.MPRMatchedAttrs
                    PrincipalSetDisplayName=Get-ObjectDisplayNameFromRef -ReferenceValue $mpr.PrincipalSet -Dictionary $setDict; PrincipalSetObjectID=$mpr.PrincipalSet
                    ResourceCurrentSetName=Get-ObjectDisplayNameFromRef -ReferenceValue $mpr.ResourceCurrentSet -Dictionary $setDict; ResourceCurrentSetID=$mpr.ResourceCurrentSet
                    ResourceFinalSetName=Get-ObjectDisplayNameFromRef -ReferenceValue $mpr.ResourceFinalSet -Dictionary $setDict; ResourceFinalSetID=$mpr.ResourceFinalSet
                }
            }
        }
    }
    $sortedMap = @($map | Sort-Object SyncRuleDisplayName,WorkflowDisplayName,MPRDisplayName)
    $rootMapPath = Join-Path $OutDir 'All_SyncRule_Workflow_MPR_Set_Map.csv'
    $sortedMap | Export-Csv -Path $rootMapPath -NoTypeInformation -Encoding UTF8

    $srDir = Join-Path $OutDir 'SynchronizationRules'
    New-Item -Path $srDir -ItemType Directory -Force | Out-Null
    $sortedMap | Export-Csv -Path (Join-Path $srDir 'All_SyncRule_Workflow_MPR_Set_Map.csv') -NoTypeInformation -Encoding UTF8

    [pscustomobject]@{
        MapCsvRoot = $rootMapPath
        MapCsvSynchronizationRulesFolder = (Join-Path $srDir 'All_SyncRule_Workflow_MPR_Set_Map.csv')
        RowCount = $sortedMap.Count
        ExportTime = Get-Date
    } | Export-Csv -Path (Join-Path $srDir 'SyncRule_Workflow_MPR_Set_Map_Status.csv') -NoTypeInformation -Encoding UTF8

    return $sortedMap
}

function Export-MimServiceConfig {
    param([string]$OutDir)
    Write-DiagStatus "Collecting CONFIG from MIM Service: $MimServiceUri" Cyan
    Add-PSSnapin FIMAutomation -ErrorAction SilentlyContinue
    $uriObj = [System.Uri]$MimServiceUri
    $tcpOk = Test-TcpPortQuiet -ComputerName $uriObj.Host -Port $uriObj.Port
    [pscustomobject]@{MimServiceUri=$MimServiceUri;Host=$uriObj.Host;Port=$uriObj.Port;TcpTestSucceeded=$tcpOk} | Export-Csv -Path (Join-Path $OutDir 'MIMService_ConnectionCheck.csv') -NoTypeInformation -Encoding UTF8
    if (-not $tcpOk) {
        Write-DiagStatus "MIM Service TCP check failed: $($uriObj.Host):$($uriObj.Port). MIM Service config export may fail." Yellow
    }
    $maMap = @()
    try {
        $maMap = @(Get-SyncManagementAgentMap -ComputerName $SyncServer)
        $maMap | Sort-Object Name | Export-Csv -Path (Join-Path $OutDir 'ManagementAgent_List_From_SyncServer.csv') -NoTypeInformation -Encoding UTF8
    } catch { Write-DiagError -Stage 'Export-MimServiceConfig' -Target 'MA map from Sync server' -ErrorRecord $_ }
    $ruleList = Export-SynchronizationRuleDetails -OutDir $OutDir -MaMap $maMap
    $workflows = Export-MimResourceCollection -ResourceType 'WorkflowDefinition' -OutPrefix 'WorkflowDefinitions' -OutDir $OutDir
    $mprs = Export-MimResourceCollection -ResourceType 'ManagementPolicyRule' -OutPrefix 'ManagementPolicyRules' -OutDir $OutDir
    $sets = Export-MimResourceCollection -ResourceType 'Set' -OutPrefix 'Sets' -OutDir $OutDir
    try {
        $attrs = @(Invoke-ExportFimConfigSafe -CustomConfig '/AttributeTypeDescription' -OnlyBaseResources)
        $result = foreach ($item in $attrs) {
            $rmo = $item.ResourceManagementObject
            $validationAttrs = $rmo.ResourceManagementAttributes | Where-Object { $_.AttributeName -match 'Regex|Regular|Validation|Pattern|Minimum|Maximum|Length' } | Sort-Object AttributeName
            $validationSummary = foreach ($vAttr in $validationAttrs) {
                $txt = if ($vAttr.IsMultiValue) { ($vAttr.Values | ForEach-Object { $_.ToString() }) -join '; ' } else { if ($null -ne $vAttr.Value) { $vAttr.Value.ToString() } else { '' } }
                if (-not [string]::IsNullOrWhiteSpace($txt)) { "$($vAttr.AttributeName)=$txt" }
            }
            [pscustomobject]@{Name=Get-FimAttrValue $rmo 'Name';DisplayName=Get-FimAttrValue $rmo 'DisplayName';DataType=Get-FimAttrValue $rmo 'DataType';Validation=($validationSummary -join ' | ');ObjectID=$rmo.ObjectIdentifier}
        }
        $result | Sort-Object DisplayName,Name | Export-Csv -Path (Join-Path $OutDir 'MIM_AllAttributes_DataType_Validation.csv') -NoTypeInformation -Encoding UTF8
    } catch { Write-DiagError -Stage 'Export-MimServiceConfig' -Target 'AttributeTypeDescription' -ErrorRecord $_ }
    try {
        $syncRulesFlat = @(Invoke-ExportFimConfigSafe -CustomConfig '/SynchronizationRule' -OnlyBaseResources | ForEach-Object { Convert-FimExportToObject $_ })
        Export-SyncRuleWorkflowMprSetMap -OutDir $OutDir -SyncRules $syncRulesFlat -Workflows $workflows -Mprs $mprs -Sets $sets | Out-Null
    } catch { Write-DiagError -Stage 'Export-MimServiceConfig' -Target 'SyncRule Workflow MPR Set map' -ErrorRecord $_ }
}

function Export-PCNSConfig {
    param([string]$OutDir)
    $pcnsDir = Join-Path $OutDir 'PCNS'
    New-Item -Path $pcnsDir -ItemType Directory -Force | Out-Null

    if ($SkipPCNS -or [string]::IsNullOrWhiteSpace($PcnSServer)) {
        Write-DiagStatus 'Skipping PCNS configuration collection because -PcnSServer was not specified or -SkipPCNS was used.' Yellow
        [pscustomobject]@{
            Status     = 'Skipped'
            Reason     = 'PcnSServer not specified or SkipPCNS specified'
            PcnSServer = $PcnSServer
            TimeCreated = Get-Date
        } | Export-Csv -Path (Join-Path $pcnsDir 'PCNS_CollectionStatus.csv') -NoTypeInformation -Encoding UTF8
        'PCNS collection was skipped. This is expected when PCNS is not installed or not in scope.' | Out-File -FilePath (Join-Path $pcnsDir 'PCNS_CollectionStatus.txt') -Encoding UTF8
        return
    }

    Write-DiagStatus "Collecting PCNS configuration from $PcnSServer" Cyan
    $remoteTemp = New-RemoteWorkPath -Name "PCNS"
    $scriptBlock = {
        param($RemoteOut)
        New-Item -Path $RemoteOut -ItemType Directory -Force | Out-Null
        hostname | Out-File -FilePath (Join-Path $RemoteOut 'Hostname.txt') -Encoding UTF8
        $relatedServices = @(Get-Service NTDS,DNS,KDC,Netlogon,PCNSSVC -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType)
        $relatedServices | Export-Csv -Path (Join-Path $RemoteOut 'PCNS_RelatedServices.csv') -NoTypeInformation -Encoding UTF8
        $pcnsService = $relatedServices | Where-Object { $_.Name -eq 'PCNSSVC' } | Select-Object -First 1
        if ($null -eq $pcnsService) {
            [pscustomobject]@{
                Status = 'PCNS_NotInstalled'
                Message = 'PCNSSVC service was not found on this server. PCNS-specific configuration files are not expected.'
                ComputerName = $env:COMPUTERNAME
                TimeCreated = Get-Date
            } | Export-Csv -Path (Join-Path $RemoteOut 'PCNS_CollectionStatus.csv') -NoTypeInformation -Encoding UTF8
        }
        $pcnsFolder = 'C:\Program Files\Microsoft Password Change Notification'
        $pcnscfg = Join-Path $pcnsFolder 'pcnscfg.exe'
        if (Test-Path $pcnscfg) {
            Push-Location $pcnsFolder
            try { $pcnsConfig = .\pcnscfg.exe list; $pcnsConfig | Out-File -FilePath (Join-Path $RemoteOut 'pcnscfg_list.txt') -Encoding UTF8 } finally { Pop-Location }
            $includeGroup = ($pcnsConfig | Where-Object { $_ -match 'Inclusion Group Name\.\.:' }) -replace '^.*Inclusion Group Name\.\.:\s*',''
            $excludeGroup = ($pcnsConfig | Where-Object { $_ -match 'Exclusion Group Name\.\.:' }) -replace '^.*Exclusion Group Name\.\.:\s*',''
            $targets = @($pcnsConfig | Where-Object { $_ -match 'Server FQDN or Address:' } | ForEach-Object { ($_ -replace '^.*Server FQDN or Address:\s*','').Trim() }) | Where-Object { $_ }
            $targets | ForEach-Object { $t=$_; foreach($p in @(135,445)){ [pscustomobject]@{Target=$t;Port=$p;TcpTestSucceeded=Test-NetConnection -ComputerName $t -Port $p -InformationLevel Quiet -WarningAction SilentlyContinue} } } | Export-Csv -Path (Join-Path $RemoteOut 'PCNS_TargetNetworkTests.csv') -NoTypeInformation -Encoding UTF8
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                foreach ($pair in @(@('Inclusion',$includeGroup),@('Exclusion',$excludeGroup))) {
                    $label=$pair[0]; $groupName=$pair[1]; $out=Join-Path $RemoteOut "PCNS_$label`GroupMembers.csv"
                    if ([string]::IsNullOrWhiteSpace($groupName)) { [pscustomobject]@{Label=$label;Status='No group configured';GroupName=$groupName} | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8; continue }
                    $sam = ($groupName -split '\\')[-1]
                    $rows=@()
                    $g = Get-ADGroup -Identity $sam -Properties DistinguishedName,GroupScope,GroupCategory -ErrorAction SilentlyContinue
                    if ($g) {
                        $rows += [pscustomobject]@{Label=$label;MemberType='Group';Name=$g.Name;SamAccountName=$g.SamAccountName;ObjectClass='group';DistinguishedName=$g.DistinguishedName}
                        Get-ADGroupMember -Identity $sam -Recursive -ErrorAction SilentlyContinue | ForEach-Object { $rows += [pscustomobject]@{Label=$label;MemberType='RecursiveMember';Name=$_.Name;SamAccountName=$_.SamAccountName;ObjectClass=$_.ObjectClass;DistinguishedName=$_.DistinguishedName} }
                    } else { $rows += [pscustomobject]@{Label=$label;Status='Group not found';GroupName=$groupName;SamAccountName=$sam} }
                    $rows | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8
                }
                setspn -Q 'PCNSCLNT/*' | Out-File -FilePath (Join-Path $RemoteOut 'PCNS_SPN_Query.txt') -Encoding UTF8
            } catch { $_ | Out-File -FilePath (Join-Path $RemoteOut 'PCNS_ADModule_Error.txt') -Encoding UTF8 }
        } else {
            "pcnscfg.exe not found: $pcnscfg" | Out-File -FilePath (Join-Path $RemoteOut 'pcnscfg_list.txt') -Encoding UTF8
        }

        $statusPath = Join-Path $RemoteOut 'PCNS_CollectionStatus.csv'
        if (-not (Test-Path -LiteralPath $statusPath)) {
            $status = if ($null -eq $pcnsService) { 'PCNS_NotInstalled' } elseif (Test-Path -LiteralPath $pcnscfg) { 'Success' } else { 'PCNS_ConfigToolNotFound' }
            $message = switch ($status) {
                'Success' { 'PCNS service and pcnscfg.exe were found. PCNS configuration collection completed.' }
                'PCNS_NotInstalled' { 'PCNSSVC service was not found on this server.' }
                'PCNS_ConfigToolNotFound' { 'PCNSSVC service was found, but pcnscfg.exe was not found in the expected folder.' }
                default { 'PCNS collection completed with status information.' }
            }
            [pscustomobject]@{
                Status        = $status
                Message       = $message
                ComputerName  = $env:COMPUTERNAME
                PcnSServer    = $env:COMPUTERNAME
                PCNSSVCStatus = if ($pcnsService) { $pcnsService.Status } else { $null }
                PcnscfgPath   = $pcnscfg
                PcnscfgExists = Test-Path -LiteralPath $pcnscfg
                TimeCreated   = Get-Date
            } | Export-Csv -Path $statusPath -NoTypeInformation -Encoding UTF8
        }

        $statusObj = Import-Csv -Path $statusPath -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($statusObj) {
            @(
                "Status: $($statusObj.Status)",
                "Message: $($statusObj.Message)",
                "ComputerName: $($statusObj.ComputerName)",
                "PCNSSVCStatus: $($statusObj.PCNSSVCStatus)",
                "PcnscfgPath: $($statusObj.PcnscfgPath)",
                "PcnscfgExists: $($statusObj.PcnscfgExists)",
                "TimeCreated: $($statusObj.TimeCreated)"
            ) | Out-File -FilePath (Join-Path $RemoteOut 'PCNS_CollectionStatus.txt') -Encoding UTF8
        }
    }
    try { Invoke-OnComputer -ComputerName $PcnSServer -ScriptBlock $scriptBlock -ArgumentList @($remoteTemp); Copy-RemoteFolderToLocal -ComputerName $PcnSServer -RemotePath $remoteTemp -LocalPath $pcnsDir }
    catch { Write-DiagError -Stage 'Export-PCNSConfig' -Target $PcnSServer -ErrorRecord $_ }
    finally { Remove-RemoteTempFolder -ComputerName $PcnSServer -RemotePath $remoteTemp -Stage 'Export-PCNSConfig-CleanupRemoteTemp' }
}

function Export-EventLogs {
    param([string]$OutDir)
    Write-DiagStatus 'Exporting Event Logs from execution server' Cyan

    $statusPath = Join-Path $OutDir 'EventLog_ExportStatus.txt'
    $logsToExport = @(
        @{LogName='Application'; FileName='Application.evtx'},
        @{LogName='System'; FileName='System.evtx'},
        @{LogName='Security'; FileName='Security.evtx'},
        @{LogName='Forefront Identity Manager'; FileName='Forefront_Identity_Manager.evtx'},
        @{LogName='Forefront Identity Manager Management Agent'; FileName='Forefront_Identity_Manager_Management_Agent.evtx'}
    )

    $statusLines = New-Object System.Collections.Generic.List[string]
    foreach ($item in $logsToExport) {
        $logName = $item.LogName
        $fileName = $item.FileName
        try {
            $logInfo = Get-WinEvent -ListLog $logName -ErrorAction Stop
            if (-not $logInfo) {
                $statusLines.Add("Skipped`t$logName`tLog not found")
                continue
            }

            $evtx = Join-Path $OutDir $fileName
            Write-DiagStatus "Exporting Event Log: $logName" Cyan
            & wevtutil epl $logName $evtx /ow:true
            if (Test-Path -LiteralPath $evtx) {
                $statusLines.Add("Exported`t$logName`t$evtx")
            }
            else {
                $statusLines.Add("Failed`t$logName`tEVTX file was not created")
            }
        }
        catch {
            $msg = Limit-Text -Text $_.Exception.Message -MaxLength 500
            $statusLines.Add("Skipped`t$logName`t$msg")
            Write-DiagError -Stage 'Export-EventLogs' -Target $logName -ErrorRecord $_
        }
    }

    $statusLines | Out-File -FilePath $statusPath -Encoding UTF8
}

function Export-SystemDiagnostics {
    param([string]$OutDir)
    Write-DiagStatus 'Collecting system diagnostics' Cyan

    $mimServiceHost = $null
    try {
        if (-not [string]::IsNullOrWhiteSpace($MimServiceUri)) {
            $mimServiceHost = ([System.Uri]$MimServiceUri).Host
        }
    } catch {}

    $servers = @($env:COMPUTERNAME,$SyncServer,$MimServiceServer,$mimServiceHost,$PcnSServer) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique

    # OS version for all known servers
    $osRows = foreach ($server in $servers) {
        try {
            Invoke-OnComputer -ComputerName $server -ScriptBlock {
                $os = Get-CimInstance Win32_OperatingSystem
                [pscustomobject]@{
                    ComputerName   = $env:COMPUTERNAME
                    Caption        = $os.Caption
                    Version        = $os.Version
                    BuildNumber    = $os.BuildNumber
                    OSArchitecture = $os.OSArchitecture
                    LastBootUpTime = $os.LastBootUpTime
                }
            }
        }
        catch {
            [pscustomobject]@{
                ComputerName   = $server
                Caption        = $null
                Version        = $null
                BuildNumber    = $null
                OSArchitecture = $null
                LastBootUpTime = $null
                Error          = (Limit-Text -Text $_.Exception.Message -MaxLength 500)
            }
        }
    }
    $osRows | Export-Csv -Path (Join-Path $OutDir 'OS_Version.csv') -NoTypeInformation -Encoding UTF8

    # MIM Sync file versions
    $syncFileRows = @()
    try {
        $syncFileRows = @(Invoke-OnComputer -ComputerName $SyncServer -ScriptBlock {
            $paths = @(
                'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Bin\miiserver.exe',
                'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\miisclient.exe',
                'C:\Program Files\Microsoft Azure AD Sync\Bin\miiserver.exe',
                'C:\Program Files\Microsoft Azure AD Sync\UIShell\miisclient.exe'
            )

            foreach ($path in $paths) {
                if (Test-Path -LiteralPath $path) {
                    $item = Get-Item -LiteralPath $path
                    [pscustomobject]@{
                        Component      = 'MIM Synchronization Service'
                        ComputerName   = $env:COMPUTERNAME
                        Path           = $path
                        ProductVersion = $item.VersionInfo.ProductVersion
                        FileVersion    = $item.VersionInfo.FileVersion
                    }
                }
            }
        })
        $syncFileRows | Export-Csv -Path (Join-Path $OutDir 'MIM_Sync_Build.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        Write-DiagError -Stage 'Export-SystemDiagnostics' -Target 'MIM Sync build version' -ErrorRecord $_
    }

    # MIM Service / Portal file versions and service status.
    # The MIM Service host is derived from -MimServiceUri unless -MimServiceServer is explicitly provided.
    $serviceTargets = @($MimServiceServer,$mimServiceHost) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique

    $serviceBuildRows = @()
    $serviceStatusRows = @()
    $portalSolutionRows = @()

    foreach ($server in $serviceTargets) {
        try {
            $result = Invoke-OnComputer -ComputerName $server -ScriptBlock {
                $serviceRows = @()
                $fileRows = @()
                $solutionRows = @()

                $svc = Get-CimInstance Win32_Service -Filter "Name='FIMService'" -ErrorAction SilentlyContinue

                if ($svc) {
                    $serviceRows += [pscustomobject]@{
                        ComputerName = $env:COMPUTERNAME
                        ServiceName  = $svc.Name
                        DisplayName  = $svc.DisplayName
                        State        = $svc.State
                        StartMode    = $svc.StartMode
                        StartName    = $svc.StartName
                        PathName     = $svc.PathName
                    }

                    $serviceExe = $null
                    if ($svc.PathName -match '^"([^"]+)"') {
                        $serviceExe = $matches[1]
                    }
                    elseif ($svc.PathName) {
                        $serviceExe = ($svc.PathName -split '\s+')[0]
                    }

                    if ($serviceExe -and (Test-Path -LiteralPath $serviceExe)) {
                        $item = Get-Item -LiteralPath $serviceExe
                        $fileRows += [pscustomobject]@{
                            Component      = 'MIM Service'
                            ComputerName   = $env:COMPUTERNAME
                            Path           = $serviceExe
                            ProductVersion = $item.VersionInfo.ProductVersion
                            FileVersion    = $item.VersionInfo.FileVersion
                        }
                    }
                }
                else {
                    $serviceRows += [pscustomobject]@{
                        ComputerName = $env:COMPUTERNAME
                        ServiceName  = 'FIMService'
                        DisplayName  = $null
                        State        = 'NotFound'
                        StartMode    = $null
                        StartName    = $null
                        PathName     = $null
                    }
                }

                $candidateFiles = @(
                    @{ Component = 'MIM Service'; Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Service\Microsoft.ResourceManagement.Service.exe' },
                    @{ Component = 'MIM Service'; Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Service\Microsoft.ResourceManagement.ServiceHost.exe' },
                    @{ Component = 'MIM Service'; Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Service\Microsoft.ResourceManagement.dll' },
                    @{ Component = 'MIM Portal';  Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Portal\bin\Microsoft.IdentityManagement.WebUI.Controls.dll' },
                    @{ Component = 'MIM Portal';  Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Portal\bin\Microsoft.ResourceManagement.WebServices.dll' },
                    @{ Component = 'MIM Portal';  Path = 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Portal\Microsoft.IdentityManagement.WebUI.Controls.dll' }
                )

                foreach ($candidate in $candidateFiles) {
                    $path = $candidate.Path
                    if (Test-Path -LiteralPath $path) {
                        $item = Get-Item -LiteralPath $path
                        $fileRows += [pscustomobject]@{
                            Component      = $candidate.Component
                            ComputerName   = $env:COMPUTERNAME
                            Path           = $path
                            ProductVersion = $item.VersionInfo.ProductVersion
                            FileVersion    = $item.VersionInfo.FileVersion
                        }
                    }
                }

                # Best-effort SharePoint solution list for MIM Portal.
                try {
                    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
                    if (Get-Command Get-SPSolution -ErrorAction SilentlyContinue) {
                        $solutionRows = @(Get-SPSolution -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -match 'fim|mim|identity' } |
                            Select-Object @{Name='ComputerName';Expression={$env:COMPUTERNAME}},
                                Name,
                                Deployed,
                                SolutionId,
                                DeployedServers,
                                LastOperationResult,
                                LastOperationTime)
                    }
                }
                catch {}

                [pscustomobject]@{
                    ServiceStatus = $serviceRows
                    FileVersions  = $fileRows | Sort-Object Component,Path -Unique
                    PortalSolutions = $solutionRows
                }
            }

            if ($result) {
                $serviceStatusRows += @($result.ServiceStatus)
                $serviceBuildRows  += @($result.FileVersions)
                $portalSolutionRows += @($result.PortalSolutions)
            }
        }
        catch {
            Write-DiagError -Stage 'Export-SystemDiagnostics' -Target "MIM Service / Portal build version on $server" -ErrorRecord $_
            $serviceStatusRows += [pscustomobject]@{
                ComputerName = $server
                ServiceName  = 'FIMService'
                DisplayName  = $null
                State        = 'CollectionFailed'
                StartMode    = $null
                StartName    = $null
                PathName     = $null
            }
        }
    }

    $serviceStatusRows | Export-Csv -Path (Join-Path $OutDir 'MIM_Service_Status.csv') -NoTypeInformation -Encoding UTF8
    $serviceBuildRows  | Export-Csv -Path (Join-Path $OutDir 'MIM_Service_Portal_Build.csv') -NoTypeInformation -Encoding UTF8
    $portalSolutionRows | Export-Csv -Path (Join-Path $OutDir 'MIM_Portal_SharePointSolutions.csv') -NoTypeInformation -Encoding UTF8

    # SharePoint version / build information for MIM Portal hosting server, when available.
    $sharePointBuildRows = @()
    foreach ($server in $serviceTargets) {
        try {
            $sharePointBuildRows += @(Invoke-OnComputer -ComputerName $server -ScriptBlock {
                $rows = @()
                try {
                    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
                    if (Get-Command Get-SPFarm -ErrorAction SilentlyContinue) {
                        $farm = Get-SPFarm -ErrorAction SilentlyContinue
                        if ($farm) {
                            $rows += [pscustomobject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Component      = 'SharePoint Farm'
                                BuildVersion   = $farm.BuildVersion.ToString()
                                ProductVersion = $null
                                FileVersion    = $null
                                Path           = $null
                                Source         = 'Get-SPFarm'
                            }
                        }
                    }
                } catch {}

                foreach ($path in @(
                    'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.dll',
                    'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll',
                    'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI\Microsoft.SharePoint.dll'
                )) {
                    if (Test-Path -LiteralPath $path) {
                        $item = Get-Item -LiteralPath $path
                        $rows += [pscustomobject]@{
                            ComputerName   = $env:COMPUTERNAME
                            Component      = 'Microsoft.SharePoint.dll'
                            BuildVersion   = $null
                            ProductVersion = $item.VersionInfo.ProductVersion
                            FileVersion    = $item.VersionInfo.FileVersion
                            Path           = $path
                            Source         = 'FileVersion'
                        }
                    }
                }

                $paths = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                )
                $products = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -match 'SharePoint' } |
                    Select-Object @{Name='ComputerName';Expression={$env:COMPUTERNAME}},
                        @{Name='Component';Expression={$_.DisplayName}},
                        @{Name='BuildVersion';Expression={$null}},
                        @{Name='ProductVersion';Expression={$_.DisplayVersion}},
                        @{Name='FileVersion';Expression={$null}},
                        @{Name='Path';Expression={$_.InstallLocation}},
                        @{Name='Source';Expression={'InstalledProduct'}}
                $rows += @($products)
                return $rows
            })
        }
        catch {
            Write-DiagError -Stage 'Export-SystemDiagnostics' -Target "SharePoint build version on $server" -ErrorRecord $_
        }
    }
    $sharePointBuildRows | Sort-Object ComputerName,Component,ProductVersion,FileVersion -Unique | Export-Csv -Path (Join-Path $OutDir 'SharePoint_Build.csv') -NoTypeInformation -Encoding UTF8

    # Installed products for all known MIM-related servers
    $installedProductRows = @()
    foreach ($server in $servers) {
        try {
            $installedProductRows += @(Invoke-OnComputer -ComputerName $server -ScriptBlock {
                $paths = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                )

                Get-ItemProperty $paths -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.DisplayName -match 'Microsoft Identity Manager|Forefront Identity Manager|Microsoft Forefront Identity Manager|MIM|FIM'
                    } |
                    Select-Object @{Name='ComputerName';Expression={$env:COMPUTERNAME}},
                        DisplayName,
                        DisplayVersion,
                        Publisher,
                        InstallDate,
                        InstallLocation,
                        UninstallString
            })
        }
        catch {
            Write-DiagError -Stage 'Export-SystemDiagnostics' -Target "Installed products on $server" -ErrorRecord $_
        }
    }

    $installedProductRows |
        Sort-Object ComputerName,DisplayName,DisplayVersion -Unique |
        Export-Csv -Path (Join-Path $OutDir 'MIM_InstalledProducts.csv') -NoTypeInformation -Encoding UTF8

    # Backward-compatible combined build file
    @($syncFileRows + $serviceBuildRows) |
        Sort-Object ComputerName,Component,Path -Unique |
        Export-Csv -Path (Join-Path $OutDir 'MIM_BuildVersion.csv') -NoTypeInformation -Encoding UTF8
}


function Export-NetworkDiagnostics {
    param([string]$OutDir)
    Write-DiagStatus 'Running network diagnostics' Cyan

    $targets = @()
    $nameTargets = @()

    try {
        $u = [System.Uri]$MimServiceUri
        $targets += [pscustomobject]@{Target=$u.Host;Port=$u.Port;Purpose='MIM Service'}
        $nameTargets += [pscustomobject]@{Role='MIMService';HostName=$u.Host;Source='MimServiceUri'}
    } catch {}

    if ($SyncServer) {
        $targets += [pscustomobject]@{Target=$SyncServer;Port=135;Purpose='MIM Sync RPC'}
        $targets += [pscustomobject]@{Target=$SyncServer;Port=5985;Purpose='PowerShell Remoting'}
        $nameTargets += [pscustomobject]@{Role='FIMSynchronizationService';HostName=$SyncServer;Source='SyncServer parameter'}
    }

    if ($PcnSServer) {
        $nameTargets += [pscustomobject]@{Role='PCNS';HostName=$PcnSServer;Source='PcnSServer parameter'}
    }

    $domainName = $GetObjDomainName
    if ([string]::IsNullOrWhiteSpace($domainName)) {
        try { $domainName = ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()).Name } catch {}
    }

    $dcNames = @()

    if ($domainName) {
        try {
            $dcText = nltest /dsgetdc:$domainName 2>&1
            $dcText | Out-File -FilePath (Join-Path $OutDir 'nltest_dsgetdc.txt') -Encoding UTF8
            $dcName = ($dcText | Where-Object { $_ -match 'DC:' } | Select-Object -First 1) -replace '.*DC:\s*\\*',''
            $dcName = $dcName.Trim()
            if ($dcName) { $dcNames += $dcName }
        } catch { Write-DiagError -Stage 'Export-NetworkDiagnostics' -Target 'nltest' -ErrorRecord $_ }

        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $dcs = @(Get-ADDomainController -Filter * -Server $domainName -ErrorAction Stop |
                Select-Object @{Name='Domain';Expression={$domainName}},
                    HostName,
                    Name,
                    IPv4Address,
                    Site,
                    IsGlobalCatalog,
                    OperatingSystem,
                    Enabled)

            $dcs | Export-Csv -Path (Join-Path $OutDir 'DomainControllers.csv') -NoTypeInformation -Encoding UTF8
            $dcNames += @($dcs | Select-Object -ExpandProperty HostName)
        }
        catch {
            Write-DiagError -Stage 'Export-NetworkDiagnostics' -Target 'Get-ADDomainController' -ErrorRecord $_
            [pscustomobject]@{
                Domain = $domainName
                Error  = (Limit-Text -Text $_.Exception.Message -MaxLength 500)
            } | Export-Csv -Path (Join-Path $OutDir 'DomainControllers.csv') -NoTypeInformation -Encoding UTF8
        }
    }

    foreach ($dcName in ($dcNames | Where-Object { $_ } | Sort-Object -Unique)) {
        $nameTargets += [pscustomobject]@{Role='DomainController';HostName=$dcName;Source='nltest/Get-ADDomainController'}
        foreach ($p in @(53,88,135,389,445,636,3268,3269)) {
            $targets += [pscustomobject]@{Target=$dcName;Port=$p;Purpose='AD DS / DC'}
        }
    }

    $maConn = Join-Path $ConfigDir 'ManagementAgents\MA_Connection_Summary.csv'
    if (Test-Path $maConn) {
        Import-Csv $maConn |
            Where-Object { $_.SqlServer } |
            Select-Object -ExpandProperty SqlServer -Unique |
            ForEach-Object {
                $targets += [pscustomobject]@{Target=$_;Port=1433;Purpose='SQL Server from MA config'}
                $nameTargets += [pscustomobject]@{Role='SQLServerFromMAConfig';HostName=$_;Source='MA_Connection_Summary.csv'}
            }
    }

    # Name resolution summary from the execution server perspective.
    $nameResolutionRows = foreach ($entry in ($nameTargets | Where-Object { $_.HostName } | Sort-Object Role,HostName -Unique)) {
        try {
            $addresses = @([System.Net.Dns]::GetHostAddresses($entry.HostName) |
                ForEach-Object { $_.IPAddressToString }) |
                Where-Object { $_ }

            $fqdn = $null
            try { $fqdn = ([System.Net.Dns]::GetHostEntry($entry.HostName)).HostName } catch {}

            [pscustomobject]@{
                Role      = $entry.Role
                HostName  = $entry.HostName
                FQDN      = $fqdn
                IPAddress = ($addresses -join '; ')
                Source    = $entry.Source
                Resolved  = $true
                Error     = $null
            }
        }
        catch {
            [pscustomobject]@{
                Role      = $entry.Role
                HostName  = $entry.HostName
                FQDN      = $null
                IPAddress = $null
                Source    = $entry.Source
                Resolved  = $false
                Error     = (Limit-Text -Text $_.Exception.Message -MaxLength 500)
            }
        }
    }

    $nameResolutionRows | Export-Csv -Path (Join-Path $OutDir 'Server_NameResolution.csv') -NoTypeInformation -Encoding UTF8

    $targets |
        Where-Object { $_.Target } |
        Sort-Object Target,Port -Unique |
        ForEach-Object {
            try {
                [pscustomobject]@{
                    Target           = $_.Target
                    Port             = $_.Port
                    Purpose          = $_.Purpose
                    TcpTestSucceeded = Test-TcpPortQuiet -ComputerName $_.Target -Port $_.Port
                }
            }
            catch {
                [pscustomobject]@{
                    Target           = $_.Target
                    Port             = $_.Port
                    Purpose          = $_.Purpose
                    TcpTestSucceeded = $false
                    Error            = (Limit-Text -Text $_.Exception.Message -MaxLength 500)
                }
            }
        } |
        Export-Csv -Path (Join-Path $OutDir 'NetworkTests.csv') -NoTypeInformation -Encoding UTF8
}


function Export-ObjectDiagnostics {
    param([string]$OutDir)
    Write-DiagStatus "Running OBJ mode diagnostics for DN: $GetObjADdn" Cyan
    $adDir = Join-Path $OutDir 'OBJECT_ADDS'; $csDir = Join-Path $OutDir 'OBJECT_CONNECTOR_SPACE'; $mvDir = Join-Path $OutDir 'OBJECT_METAVERSE'
    New-Item -Path $adDir -ItemType Directory -Force | Out-Null; New-Item -Path $csDir -ItemType Directory -Force | Out-Null; New-Item -Path $mvDir -ItemType Directory -Force | Out-Null
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-DiagStatus "Prompting for AD credential: $DomainAdminName" Cyan
        Write-DiagStatus "A credential dialog will be displayed. If it is hidden, check behind the PowerShell window or press Alt+Tab." DarkYellow
        $cred = Get-Credential -UserName $DomainAdminName -Message 'Enter credentials for AD object collection.'
        Write-DiagStatus "AD credential input completed. Collecting AD object from $GetObjDomainName." Cyan
        $obj = Get-ADObject -Server $GetObjDomainName -Identity $GetObjADdn -Credential $cred -Properties * -ErrorAction Stop
        Write-DiagStatus "AD object collected. Exporting AD object files." Cyan
        $obj | Format-List * | Out-File -FilePath (Join-Path $adDir 'ADDS_Object.txt') -Encoding UTF8
        $obj | Export-Clixml -Path (Join-Path $adDir 'ADDS_Object.clixml')
        $obj | Select-Object * | Export-Csv -Path (Join-Path $adDir 'ADDS_Object.csv') -NoTypeInformation -Encoding UTF8
    } catch { Write-DiagError -Stage 'Export-ObjectDiagnostics' -Target "ADDS $GetObjADdn" -ErrorRecord $_ }
    $remoteTemp = New-RemoteWorkPath -Name "Object"
    $scriptBlock = {
        param($RemoteOut,$DN)
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
        $ErrorActionPreference = 'SilentlyContinue'
        $WarningPreference = 'SilentlyContinue'
        $ProgressPreference = 'SilentlyContinue'

        New-Item -Path $RemoteOut -ItemType Directory -Force | Out-Null
        $csDir = Join-Path $RemoteOut 'ConnectorSpace'
        $mvDir = Join-Path $RemoteOut 'Metaverse'
        New-Item -Path $csDir -ItemType Directory -Force | Out-Null
        New-Item -Path $mvDir -ItemType Directory -Force | Out-Null

        function New-SafeRemoteFileName {
            param([string]$Name)
            $safe = $Name -replace '[\/:*?"<>|\s]+','_'
            $safe = $safe.Trim('_')
            if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'NoName' }
            return $safe
        }

        function Limit-RemoteText {
            param([string]$Text,[int]$MaxLength=1000)
            if ($null -eq $Text) { return $null }
            if ($Text.Length -le $MaxLength) { return $Text }
            return $Text.Substring(0,$MaxLength) + ' ...[truncated]'
        }

        function Export-AnyObject {
            param($Object,[string]$BasePath)
            try { $Object | Format-List * | Out-File -FilePath ($BasePath + '.txt') -Encoding UTF8 -Width 300 } catch {}
            try { $Object | Export-Clixml -Path ($BasePath + '.clixml') } catch {}
            try { $Object | ConvertTo-Json -Depth 12 | Out-File -FilePath ($BasePath + '.json') -Encoding UTF8 -Width 300 } catch {}
        }

        function Export-ObjectPropertySummary {
            param($Object,[string]$Path)
            try {
                $rows = foreach ($p in $Object.PSObject.Properties) {
                    [pscustomobject]@{
                        Name        = $p.Name
                        TypeName    = if ($null -ne $p.Value) { $p.Value.GetType().FullName } else { $null }
                        ValueSample = Limit-RemoteText -Text ([string]$p.Value) -MaxLength 800
                    }
                }
                $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            } catch {}
        }

        function Get-MethodResultNoArgs {
            param($Object,[string]$MethodName)
            try {
                $method = $Object.PSObject.Methods | Where-Object { $_.Name -eq $MethodName } | Select-Object -First 1
                if ($null -ne $method) { return $method.Invoke() }
            } catch {}
            return $null
        }

        function Get-CandidateMvIds {
            param($Object)
            $candidates = New-Object System.Collections.Generic.List[string]
            if ($null -eq $Object) { return @() }

            $knownPropNames = @(
                'MVObjectID','MvObjectID','MVObjectId','MVObjectGuid','MVGuid','MVObjectGuidString',
                'MetaverseObjectID','MetaverseObjectId','MetaverseGuid','MetaverseObjectGuid',
                'ConnectedMVObjectID','ConnectedMVObjectId','ConnectedMVObjectGuid'
            )

            foreach ($name in $knownPropNames) {
                try {
                    $v = $Object.PSObject.Properties[$name].Value
                    if ($v) { $candidates.Add([string]$v) }
                } catch {}
            }

            foreach ($p in $Object.PSObject.Properties) {
                try {
                    $n = [string]$p.Name
                    $v = $p.Value
                    if ($null -eq $v) { continue }

                    if ($n -match '(?i)mv|metaverse') {
                        $candidates.Add([string]$v)
                    }

                    $s = [string]$v
                    foreach ($m in [regex]::Matches($s,'(?i)(urn:uuid:)?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')) {
                        $candidates.Add($m.Value)
                    }

                    if ($n -match '(?i)mv|metaverse' -and ($v -isnot [string])) {
                        foreach ($child in $v.PSObject.Properties) {
                            $cv = $child.Value
                            if ($cv) {
                                foreach ($m in [regex]::Matches(([string]$cv),'(?i)(urn:uuid:)?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')) {
                                    $candidates.Add($m.Value)
                                }
                            }
                        }
                    }
                } catch {}
            }

            return @($candidates | ForEach-Object { ($_ -replace '^urn:uuid:','').Trim('{}') } | Where-Object { $_ } | Sort-Object -Unique)
        }

        function Invoke-CommandIfParametersExist {
            param(
                [string]$CommandName,
                [hashtable]$Parameters
            )
            $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
            if ($null -eq $cmd) { return $null }
            foreach ($key in @($Parameters.Keys)) {
                if (-not $cmd.Parameters.ContainsKey($key)) { return $null }
            }
            try { return & $CommandName @Parameters } catch { return $null }
        }

        function Try-GetMvObjectById {
            param([string]$CandidateId)
            if ([string]::IsNullOrWhiteSpace($CandidateId)) { return $null }
            $id = $CandidateId -replace '^urn:uuid:',''
            $id = $id.Trim('{}')

            $attempts = @(
                @{ ID = $id },
                @{ Id = $id },
                @{ ObjectID = $id },
                @{ ObjectId = $id },
                @{ Guid = $id },
                @{ Identifier = $id },
                @{ MVObjectID = $id },
                @{ MVObjectId = $id }
            )
            foreach ($params in $attempts) {
                $r = Invoke-CommandIfParametersExist -CommandName 'Get-MVObject' -Parameters $params
                if ($r) { return $r }
            }
            return $null
        }

        function Try-SearchMvObjectByAnchor {
            param([string]$Anchor)
            if ([string]::IsNullOrWhiteSpace($Anchor)) { return $null }
            $attrs = @('accountName','AccountName','sAMAccountName','uid','displayName')
            $commands = @('Get-MVObject','Search-MVObject')

            foreach ($cmdName in $commands) {
                $cmd = Get-Command $cmdName -ErrorAction SilentlyContinue
                if ($null -eq $cmd) { continue }

                foreach ($attr in $attrs) {
                    $attempts = @(
                        @{ Attribute = $attr; Value = $Anchor },
                        @{ AttributeName = $attr; AttributeValue = $Anchor },
                        @{ ObjectType = 'person'; Attribute = $attr; Value = $Anchor },
                        @{ ObjectType = 'person'; AttributeName = $attr; AttributeValue = $Anchor },
                        @{ Filter = "$attr='$Anchor'" },
                        @{ XPath = "/person[$attr='$Anchor']" }
                    )
                    foreach ($params in $attempts) {
                        $r = Invoke-CommandIfParametersExist -CommandName $cmdName -Parameters $params
                        if ($r) { return $r }
                    }
                }
            }
            return $null
        }

        $moduleStatus = @()
        try {
            Import-Module LithnetMIISAutomation -Force -ErrorAction Stop
            $moduleStatus += [pscustomobject]@{ Name='LithnetMIISAutomation'; Status='Loaded'; Error=$null }
        }
        catch {
            $moduleStatus += [pscustomobject]@{ Name='LithnetMIISAutomation'; Status='LoadFailed'; Error=Limit-RemoteText -Text $_.Exception.Message -MaxLength 1000 }
            $moduleStatus | Export-Csv -Path (Join-Path $RemoteOut 'ObjectDiagnostics_ModuleStatus.csv') -NoTypeInformation -Encoding UTF8
            return
        }
        $moduleStatus | Export-Csv -Path (Join-Path $RemoteOut 'ObjectDiagnostics_ModuleStatus.csv') -NoTypeInformation -Encoding UTF8

        Get-Command -Module LithnetMIISAutomation -ErrorAction SilentlyContinue |
            Select-Object Name,CommandType,ModuleName |
            Sort-Object Name |
            Export-Csv -Path (Join-Path $RemoteOut 'LithnetMIISAutomation_Commands.csv') -NoTypeInformation -Encoding UTF8

        $anchor = $null
        if ($DN -match '^CN=([^,]+)') { $anchor = $matches[1] }

        $found = @()
        $mvRows = @()
        foreach($ma in Get-ManagementAgent){
            $safe = New-SafeRemoteFileName $ma.Name
            $csObj = $null
            $errs = @()
            foreach($attempt in 1..4){
                try {
                    switch($attempt){
                        1 { $csObj = Get-CSObject -MA $ma.Name -DN $DN -ErrorAction Stop }
                        2 { $csObj = Get-CSObject -ManagementAgent $ma.Name -DN $DN -ErrorAction Stop }
                        3 { $csObj = Get-CSObject -ManagementAgentName $ma.Name -DN $DN -ErrorAction Stop }
                        4 { $csObj = Get-CSObject -ConnectorName $ma.Name -DN $DN -ErrorAction Stop }
                    }
                    if($csObj){break}
                } catch { $errs += (Limit-RemoteText -Text $_.Exception.Message -MaxLength 500) }
            }

            if($csObj){
                Export-AnyObject -Object $csObj -BasePath (Join-Path $csDir "$safe`_CSObject")
                Export-ObjectPropertySummary -Object $csObj -Path (Join-Path $csDir "$safe`_CSObject_PropertySummary.csv")
                ($csObj.PSObject.Methods | Select-Object Name,MemberType,OverloadDefinitions | Sort-Object Name) |
                    Export-Csv -Path (Join-Path $csDir "$safe`_CSObject_Methods.csv") -NoTypeInformation -Encoding UTF8

                $found += [pscustomobject]@{ MAName=$ma.Name; Status='Found'; SafeName=$safe }

                $directMv = $null
                foreach($pname in @('MVObject','MvObject','MetaverseObject','ConnectedMVObject')){
                    try { if($csObj.PSObject.Properties[$pname].Value){ $directMv = $csObj.PSObject.Properties[$pname].Value; break } } catch {}
                }
                if(-not $directMv){
                    foreach($mname in @('GetMVObject','GetMvObject','GetMetaverseObject','GetConnectedMVObject')){
                        $directMv = Get-MethodResultNoArgs -Object $csObj -MethodName $mname
                        if($directMv){ break }
                    }
                }

                if($directMv){
                    Export-AnyObject -Object $directMv -BasePath (Join-Path $mvDir "$safe`_MVObject_Direct")
                    Export-ObjectPropertySummary -Object $directMv -Path (Join-Path $mvDir "$safe`_MVObject_Direct_PropertySummary.csv")
                    $mvRows += [pscustomobject]@{ MAName=$ma.Name; Method='DirectFromCSObject'; Candidate=$null; Status='Found' }
                    continue
                }

                $candidateIds = @(Get-CandidateMvIds -Object $csObj)
                $candidateIds | ForEach-Object { [pscustomobject]@{ MAName=$ma.Name; CandidateMVId=$_ } } |
                    Export-Csv -Path (Join-Path $mvDir "$safe`_CandidateMVIds.csv") -NoTypeInformation -Encoding UTF8

                $resolved = $false
                foreach($candidate in $candidateIds){
                    $mvObj = Try-GetMvObjectById -CandidateId $candidate
                    if($mvObj){
                        Export-AnyObject -Object $mvObj -BasePath (Join-Path $mvDir "$safe`_MVObject_$candidate")
                        Export-ObjectPropertySummary -Object $mvObj -Path (Join-Path $mvDir "$safe`_MVObject_$candidate`_PropertySummary.csv")
                        $mvRows += [pscustomobject]@{ MAName=$ma.Name; Method='Get-MVObjectByCandidateId'; Candidate=$candidate; Status='Found' }
                        $resolved = $true
                        break
                    }
                }

                if(-not $resolved -and $anchor){
                    $mvObj = Try-SearchMvObjectByAnchor -Anchor $anchor
                    if($mvObj){
                        Export-AnyObject -Object $mvObj -BasePath (Join-Path $mvDir "$safe`_MVObject_SearchByAnchor_$anchor")
                        Export-ObjectPropertySummary -Object $mvObj -Path (Join-Path $mvDir "$safe`_MVObject_SearchByAnchor_$anchor`_PropertySummary.csv")
                        $mvRows += [pscustomobject]@{ MAName=$ma.Name; Method='SearchByAnchor'; Candidate=$anchor; Status='Found' }
                        $resolved = $true
                    }
                }

                if(-not $resolved){
                    $mvRows += [pscustomobject]@{ MAName=$ma.Name; Method='BestEffort'; Candidate=($candidateIds -join ';'); Status='NotResolved' }
                }
            }
            else {
                [pscustomobject]@{MAName=$ma.Name; Status='NotFoundOrCommandFailed'; Errors=($errs -join ' | ')} |
                    Export-Csv -Path (Join-Path $csDir "$safe`_SearchErrors.csv") -NoTypeInformation -Encoding UTF8
            }
        }

        $found | Export-Csv -Path (Join-Path $RemoteOut 'ConnectorSpace_SearchSummary.csv') -NoTypeInformation -Encoding UTF8
        $mvRows | Export-Csv -Path (Join-Path $RemoteOut 'Metaverse_SearchSummary.csv') -NoTypeInformation -Encoding UTF8
        $mvRows | Export-Csv -Path (Join-Path $mvDir 'Metaverse_SearchSummary.csv') -NoTypeInformation -Encoding UTF8
    }
    $objDiagJob = $null
    try {
        Write-DiagStatus "Connector Space / Metaverse object diagnostics started on $SyncServer. This may take several minutes." Cyan

        if (Test-IsLocalComputer -ComputerName $SyncServer) {
            $objDiagJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList @($remoteTemp, $GetObjADdn)
        }
        else {
            $objDiagJob = Start-Job -ScriptBlock {
                param($ComputerName, $ScriptBlockText, $RemoteOut, $DN)
                $sb = [scriptblock]::Create($ScriptBlockText)
                Invoke-Command -ComputerName $ComputerName -ScriptBlock $sb -ArgumentList @($RemoteOut, $DN) -ErrorAction Stop
            } -ArgumentList @($SyncServer, $scriptBlock.ToString(), $remoteTemp, $GetObjADdn)
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($objDiagJob.State -eq 'Running') {
            Start-Sleep -Seconds 30
            $elapsedText = $sw.Elapsed.ToString('hh\:mm\:ss')
            Write-DiagStatus "Connector Space / Metaverse object diagnostics is still running on $SyncServer. Elapsed: $elapsedText. Please wait." DarkYellow
        }

        Receive-Job -Job $objDiagJob -ErrorAction Stop | Out-Null
        Write-DiagStatus "Connector Space / Metaverse object diagnostics completed on $SyncServer. Copying files." Cyan

        Copy-RemoteFolderToLocal -ComputerName $SyncServer -RemotePath (Join-Path $remoteTemp 'ConnectorSpace') -LocalPath $csDir
        Copy-RemoteFolderToLocal -ComputerName $SyncServer -RemotePath (Join-Path $remoteTemp 'Metaverse') -LocalPath $mvDir
        Copy-RemoteFolderToLocal -ComputerName $SyncServer -RemotePath $remoteTemp -LocalPath (Join-Path $OutDir 'OBJECT_SUMMARY')
    }
    catch { Write-DiagError -Stage 'Export-ObjectDiagnostics' -Target "CS/MV $GetObjADdn" -ErrorRecord $_ }
    finally {
        if ($null -ne $objDiagJob) { Remove-Job -Job $objDiagJob -Force -ErrorAction SilentlyContinue }
        Remove-RemoteTempFolder -ComputerName $SyncServer -RemotePath $remoteTemp -Stage 'Export-ObjectDiagnostics-CleanupRemoteTemp'
    }
}


function Export-MetaverseExtensionConfig {
    param([string]$OutDir)

    if ($ObjectMode) { return }

    Write-DiagStatus "Collecting Metaverse extension configuration from $SyncServer" Cyan

    $localOut = Join-Path $OutDir 'MetaverseExtension'
    New-Item -Path $localOut -ItemType Directory -Force | Out-Null

    function Convert-DataTableToObjects {
        param([System.Data.DataTable]$DataTable)
        foreach ($row in $DataTable.Rows) {
            $obj = [ordered]@{}
            foreach ($col in $DataTable.Columns) {
                $value = $row[$col.ColumnName]
                if ($value -is [System.DBNull]) { $value = $null }
                $obj[$col.ColumnName] = $value
            }
            [pscustomobject]$obj
        }
    }

    function Invoke-SyncSqlQuery {
        param(
            [string]$ConnectionString,
            [string]$Sql,
            [int]$TimeoutSeconds = 120
        )

        $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $TimeoutSeconds
        $table = New-Object System.Data.DataTable
        try {
            $conn.Open()
            $reader = $cmd.ExecuteReader()
            $table.Load($reader)
            return ,$table
        }
        finally {
            if ($conn.State -ne 'Closed') { $conn.Close() }
        }
    }

    $statusRows = @()
    $configRows = @()
    $rawConfigRows = @()
    $registeredRows = @()
    $dllInfoRows = @()

    try {
        # Avoid Kerberos double-hop issues:
        # 1. Read Sync service registry from the Sync server by remoting.
        # 2. Query the Sync DB directly from the execution server.
        # 3. Query DLL file information from the Sync server by remoting.
        $regScript = {
            $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\FIMSynchronizationService\Parameters'
            $syncParams = Get-ItemProperty -Path $regPath -ErrorAction Stop
            [pscustomobject]@{
                Server               = [string]$syncParams.Server
                DBName               = [string]$syncParams.DBName
                Path                 = [string]$syncParams.Path
                SQLInstance          = [string]$syncParams.SQLInstance
                ServerAuthentication = [string]$syncParams.ServerAuthentication
                ServerLocation       = [string]$syncParams.ServerLocation
            }
        }

        $syncInfo = Invoke-OnComputer -ComputerName $SyncServer -ScriptBlock $regScript
        if ($syncInfo -is [array]) { $syncInfo = $syncInfo | Select-Object -First 1 }

        $sqlServer = [string]$syncInfo.Server
        $database  = [string]$syncInfo.DBName
        $syncRoot  = ([string]$syncInfo.Path).TrimEnd('\')
        $sqlInstance = [string]$syncInfo.SQLInstance
        $sqlTarget = $sqlServer
        if (-not [string]::IsNullOrWhiteSpace($sqlInstance)) {
            if ($sqlTarget -notmatch '\\') { $sqlTarget = "$sqlTarget\$sqlInstance" }
        }

        $statusRows += [pscustomobject]@{
            Item = 'Registry'
            Status = 'Success'
            Detail = "Server=$sqlServer; SQLTarget=$sqlTarget; DBName=$database; Path=$syncRoot; ServerAuthentication=$($syncInfo.ServerAuthentication); ServerLocation=$($syncInfo.ServerLocation)"
        }

        $connectionString = "Server=$sqlTarget;Database=$database;Integrated Security=True;TrustServerCertificate=True;"

        $serverConfigSql = @"
SELECT
    CONVERT(nvarchar(max), mv_extension_dll_xml) AS mv_extension_dll_xml
FROM dbo.mms_server_configuration;
"@

        $serverConfigTable = Invoke-SyncSqlQuery -ConnectionString $connectionString -Sql $serverConfigSql
        $serverConfigObjects = @(Convert-DataTableToObjects -DataTable $serverConfigTable)
        $serverConfigObjects | Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_RawServerConfiguration.csv') -NoTypeInformation -Encoding UTF8

        $statusRows += [pscustomobject]@{
            Item = 'mms_server_configuration.mv_extension_dll_xml'
            Status = 'Success'
            Detail = "Rows=$($serverConfigObjects.Count)"
        }

        $configuredDllNames = @()
        $configIndex = 0
        foreach ($row in $serverConfigObjects) {
            $xmlText = [string]$row.mv_extension_dll_xml
            $rawConfigRows += [pscustomobject]@{
                Source = 'dbo.mms_server_configuration.mv_extension_dll_xml'
                SqlServer = $sqlServer
                Database = $database
                SyncRoot = $syncRoot
                RawXml = $xmlText
            }

            if ([string]::IsNullOrWhiteSpace($xmlText)) {
                $configRows += [pscustomobject]@{
                    Source = 'dbo.mms_server_configuration.mv_extension_dll_xml'
                    ConfigIndex = $configIndex
                    Enabled = $false
                    ExtensionName = $null
                    ApplicationProtection = $null
                    DllPath = $null
                    DllExists = $false
                    SqlServer = $sqlServer
                    Database = $database
                    SyncRoot = $syncRoot
                    ParseStatus = 'Empty'
                    RawXml = $xmlText
                }
                $configIndex++
                continue
            }

            try {
                [xml]$xml = $xmlText
                $nodes = @()
                if ($xml.extension) { $nodes += $xml.extension }
                if ($xml.extensions -and $xml.extensions.extension) { $nodes += @($xml.extensions.extension) }
                if ($nodes.Count -eq 0 -and $xml.DocumentElement) { $nodes += $xml.DocumentElement }

                foreach ($node in $nodes) {
                    $dllName = $null
                    $appProtection = $null
                    try { $dllName = [string]$node.'assembly-name' } catch {}
                    try { $appProtection = [string]$node.'application-protection' } catch {}
                    if ([string]::IsNullOrWhiteSpace($dllName)) {
                        try { $dllName = [string]$node.assemblyName } catch {}
                    }

                    if (-not [string]::IsNullOrWhiteSpace($dllName)) { $configuredDllNames += $dllName }
                    $dllPath = if (-not [string]::IsNullOrWhiteSpace($dllName)) { Join-Path (Join-Path $syncRoot 'Extensions') $dllName } else { $null }

                    $configRows += [pscustomobject]@{
                        Source = 'dbo.mms_server_configuration.mv_extension_dll_xml'
                        ConfigIndex = $configIndex
                        Enabled = -not [string]::IsNullOrWhiteSpace($dllName)
                        ExtensionName = $dllName
                        ApplicationProtection = $appProtection
                        DllPath = $dllPath
                        DllExists = $false
                        SqlServer = $sqlServer
                        Database = $database
                        SyncRoot = $syncRoot
                        ParseStatus = 'Parsed'
                        RawXml = $xmlText
                    }
                    $configIndex++
                }
            }
            catch {
                $configRows += [pscustomobject]@{
                    Source = 'dbo.mms_server_configuration.mv_extension_dll_xml'
                    ConfigIndex = $configIndex
                    Enabled = $false
                    ExtensionName = $null
                    ApplicationProtection = $null
                    DllPath = $null
                    DllExists = $false
                    SqlServer = $sqlServer
                    Database = $database
                    SyncRoot = $syncRoot
                    ParseStatus = "ParseFailed: $($_.Exception.Message)"
                    RawXml = $xmlText
                }
                $configIndex++
            }
        }

        $extensionSql = @"
SELECT *
FROM dbo.mms_extensions
ORDER BY file_name;
"@
        try {
            $extensionTable = Invoke-SyncSqlQuery -ConnectionString $connectionString -Sql $extensionSql
            $extensionObjects = @(Convert-DataTableToObjects -DataTable $extensionTable)

            foreach ($ext in $extensionObjects) {
                $fileName = [string]$ext.file_name
                if (-not [string]::IsNullOrWhiteSpace($fileName)) { $configuredDllNames += $fileName }
                $dllPath = if (-not [string]::IsNullOrWhiteSpace($fileName)) { Join-Path (Join-Path $syncRoot 'Extensions') $fileName } else { $null }
                $base = [ordered]@{
                    SqlServer = $sqlServer
                    Database = $database
                    SyncRoot = $syncRoot
                    FileName = $fileName
                    DllPath = $dllPath
                    DllExists = $false
                }
                foreach ($prop in $ext.PSObject.Properties) {
                    if (-not $base.Contains($prop.Name)) { $base[$prop.Name] = $prop.Value }
                }
                $registeredRows += [pscustomobject]$base
            }

            $statusRows += [pscustomobject]@{
                Item = 'dbo.mms_extensions'
                Status = 'Success'
                Detail = "Rows=$($extensionObjects.Count)"
            }
        }
        catch {
            $statusRows += [pscustomobject]@{
                Item = 'dbo.mms_extensions'
                Status = 'Failed'
                Detail = $_.Exception.Message
            }
        }

        if ($configRows.Count -eq 0) {
            $configRows += [pscustomobject]@{
                Source = 'dbo.mms_server_configuration.mv_extension_dll_xml'
                ConfigIndex = 0
                Enabled = $false
                ExtensionName = $null
                ApplicationProtection = $null
                DllPath = $null
                DllExists = $false
                SqlServer = $sqlServer
                Database = $database
                SyncRoot = $syncRoot
                ParseStatus = 'NoRows'
                RawXml = $null
            }
        }

        $uniqueDllNames = @($configuredDllNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        if ($uniqueDllNames.Count -gt 0) {
            $dllScript = {
                param([string]$SyncRootRemote, [string[]]$DllNames)

                foreach ($dllName in $DllNames) {
                    $extRoot = Join-Path $SyncRootRemote 'Extensions'
                    $dllPath = Join-Path $extRoot $dllName
                    $exists = Test-Path -LiteralPath $dllPath
                    $item = if ($exists) { Get-Item -LiteralPath $dllPath -ErrorAction SilentlyContinue } else { $null }
                    $hash = $null
                    if ($item) {
                        try { $hash = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $hash = $null }
                    }

                    [pscustomobject]@{
                        DllName        = $dllName
                        DllPath        = $dllPath
                        DllExists      = [bool]$exists
                        Length         = if ($item) { $item.Length } else { $null }
                        LastWriteTime  = if ($item) { $item.LastWriteTime } else { $null }
                        ProductVersion = if ($item) { $item.VersionInfo.ProductVersion } else { $null }
                        FileVersion    = if ($item) { $item.VersionInfo.FileVersion } else { $null }
                        SHA256         = $hash
                    }
                }
            }

            try {
                $dllInfoRows = @(Invoke-OnComputer -ComputerName $SyncServer -ScriptBlock $dllScript -ArgumentList @($syncRoot, $uniqueDllNames))
                $statusRows += [pscustomobject]@{
                    Item = 'ExtensionDllFileInfo'
                    Status = 'Success'
                    Detail = "Rows=$($dllInfoRows.Count)"
                }
            }
            catch {
                $statusRows += [pscustomobject]@{
                    Item = 'ExtensionDllFileInfo'
                    Status = 'Failed'
                    Detail = $_.Exception.Message
                }
            }
        }
        else {
            $statusRows += [pscustomobject]@{
                Item = 'ExtensionDllFileInfo'
                Status = 'Skipped'
                Detail = 'No configured or registered extension DLL names found.'
            }
        }

        # Reflect DllExists from remote file information into config/registered rows.
        $dllInfoLookup = @{}
        foreach ($d in $dllInfoRows) {
            if ($d -and $d.DllName) { $dllInfoLookup[[string]$d.DllName] = $d }
        }

        foreach ($row in $configRows) {
            if ($row.ExtensionName -and $dllInfoLookup.ContainsKey([string]$row.ExtensionName)) {
                $row.DllExists = [bool]$dllInfoLookup[[string]$row.ExtensionName].DllExists
            }
        }
        foreach ($row in $registeredRows) {
            if ($row.FileName -and $dllInfoLookup.ContainsKey([string]$row.FileName)) {
                $row.DllExists = [bool]$dllInfoLookup[[string]$row.FileName].DllExists
            }
        }

        $configRows | Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_Config.csv') -NoTypeInformation -Encoding UTF8
        $rawConfigRows | Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_RawConfig.csv') -NoTypeInformation -Encoding UTF8
        $registeredRows | Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_RegisteredDlls.csv') -NoTypeInformation -Encoding UTF8
        $dllInfoRows |
            Where-Object { $_ -ne $null -and -not [string]::IsNullOrWhiteSpace($_.DllName) } |
            Sort-Object DllName,DllPath -Unique |
            Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_DllInfo.csv') -NoTypeInformation -Encoding UTF8
    }
    catch {
        $statusRows += [pscustomobject]@{
            Item = 'MetaverseExtensionCollection'
            Status = 'Failed'
            Detail = $_.Exception.Message
        }
        Write-DiagError -Stage 'Export-MetaverseExtensionConfig' -Target $SyncServer -ErrorRecord $_
    }
    finally {
        if ($statusRows.Count -eq 0) {
            $statusRows += [pscustomobject]@{
                Item = 'MetaverseExtensionCollection'
                Status = 'Unknown'
                Detail = 'No status rows were generated.'
            }
        }
        $statusRows | Export-Csv -Path (Join-Path $localOut 'Metaverse_Extension_CollectionStatus.csv') -NoTypeInformation -Encoding UTF8
        $statusRows | Format-List | Out-File -FilePath (Join-Path $localOut 'Metaverse_Extension_CollectionStatus.txt') -Encoding UTF8
    }
}

function Convert-CsvFileToHtmlSection {
    param(
        [string]$Title,
        [string]$Path,
        [int]$MaxRows = 200
    )
    $titleEncoded = [System.Net.WebUtility]::HtmlEncode($Title)
    $pathEncoded = [System.Net.WebUtility]::HtmlEncode($Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return "<section><h2>$titleEncoded</h2><p class='missing'>Not found: $pathEncoded</p></section>"
    }
    try {
        $rows = @(Import-Csv -LiteralPath $Path)
        $count = $rows.Count
        if ($count -eq 0) {
            return "<section><h2>$titleEncoded</h2><p>File exists but contains no rows.</p><p class='path'>$pathEncoded</p></section>"
        }
        $fragment = $rows | Select-Object -First $MaxRows | ConvertTo-Html -Fragment
        $note = if ($count -gt $MaxRows) { "<p class='note'>Showing first $MaxRows of $count rows.</p>" } else { "<p class='note'>Rows: $count</p>" }
        return "<section><h2>$titleEncoded</h2><p class='path'>$pathEncoded</p>$note$fragment</section>"
    }
    catch {
        $err = [System.Net.WebUtility]::HtmlEncode((Limit-Text -Text $_.Exception.Message -MaxLength 500))
        return "<section><h2>$titleEncoded</h2><p class='error'>Failed to render CSV: $err</p><p class='path'>$pathEncoded</p></section>"
    }
}

function Convert-TextFileToHtmlSection {
    param(
        [string]$Title,
        [string]$Path,
        [int]$MaxLength = 6000
    )
    $titleEncoded = [System.Net.WebUtility]::HtmlEncode($Title)
    $pathEncoded = [System.Net.WebUtility]::HtmlEncode($Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return "<section><h2>$titleEncoded</h2><p class='missing'>Not found: $pathEncoded</p></section>"
    }
    try {
        $text = Get-Content -LiteralPath $Path -Raw
        if ($text.Length -gt $MaxLength) { $text = $text.Substring(0,$MaxLength) + "`r`n... [truncated]" }
        $encoded = [System.Net.WebUtility]::HtmlEncode($text)
        return "<section><h2>$titleEncoded</h2><p class='path'>$pathEncoded</p><pre>$encoded</pre></section>"
    }
    catch {
        $err = [System.Net.WebUtility]::HtmlEncode((Limit-Text -Text $_.Exception.Message -MaxLength 500))
        return "<section><h2>$titleEncoded</h2><p class='error'>Failed to render text: $err</p><p class='path'>$pathEncoded</p></section>"
    }
}

function New-MimDiagnosticHtmlReport {
    param([string]$ReportPath)
    if ($ObjectMode) { return }
    Write-DiagStatus 'Creating HTML summary report' Cyan

    $sections = @()
    $sections += Convert-CsvFileToHtmlSection -Title 'Run Summary' -Path (Join-Path $Root 'RunSummary.csv') -MaxRows 50
    $sections += Convert-CsvFileToHtmlSection -Title 'Management Agents' -Path (Join-Path $ConfigDir 'ManagementAgent_List_From_SyncServer.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'MA Connection Summary' -Path (Join-Path $ConfigDir 'ManagementAgents\MA_Connection_Summary.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'AD MA OU Scope' -Path (Join-Path $ConfigDir 'ManagementAgents\ADMA_OU_Scope.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'Synchronization Rules' -Path (Join-Path $ConfigDir 'SynchronizationRules\All_SynchronizationRules.csv') -MaxRows 300
    $sections += Convert-CsvFileToHtmlSection -Title 'Synchronization Rule Attribute Flows' -Path (Join-Path $ConfigDir 'SynchronizationRules\All_SynchronizationRule_AttributeFlows.csv') -MaxRows 500
    $sections += Convert-CsvFileToHtmlSection -Title 'SynchronizationRule / Workflow / MPR / Set Map' -Path (Join-Path $ConfigDir 'All_SyncRule_Workflow_MPR_Set_Map.csv') -MaxRows 500
    $sections += Convert-CsvFileToHtmlSection -Title 'Workflow Definitions' -Path (Join-Path $ConfigDir 'WorkflowDefinitions_Summary.csv') -MaxRows 300
    $sections += Convert-CsvFileToHtmlSection -Title 'Management Policy Rules' -Path (Join-Path $ConfigDir 'ManagementPolicyRules_Summary.csv') -MaxRows 300
    $sections += Convert-CsvFileToHtmlSection -Title 'Sets' -Path (Join-Path $ConfigDir 'Sets_Summary.csv') -MaxRows 300
    $sections += Convert-CsvFileToHtmlSection -Title 'Metaverse Extension Configuration' -Path (Join-Path $ConfigDir 'MetaverseExtension\Metaverse_Extension_Config.csv') -MaxRows 100
    $sections += Convert-CsvFileToHtmlSection -Title 'Metaverse Extension DLL Info' -Path (Join-Path $ConfigDir 'MetaverseExtension\Metaverse_Extension_DllInfo.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'Registered Extension DLLs' -Path (Join-Path $ConfigDir 'MetaverseExtension\Metaverse_Extension_RegisteredDlls.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'Metaverse Extension Collection Status' -Path (Join-Path $ConfigDir 'MetaverseExtension\Metaverse_Extension_CollectionStatus.csv') -MaxRows 50
    $sections += Convert-CsvFileToHtmlSection -Title 'PCNS Collection Status' -Path (Join-Path $ConfigDir 'PCNS\PCNS_CollectionStatus.csv') -MaxRows 50
    $sections += Convert-CsvFileToHtmlSection -Title 'PCNS Related Services' -Path (Join-Path $ConfigDir 'PCNS\PCNS_RelatedServices.csv') -MaxRows 100
    $sections += Convert-CsvFileToHtmlSection -Title 'PCNS Target Network Tests' -Path (Join-Path $ConfigDir 'PCNS\PCNS_TargetNetworkTests.csv') -MaxRows 200
    $sections += Convert-TextFileToHtmlSection -Title 'PCNS Configuration Text' -Path (Join-Path $ConfigDir 'PCNS\pcnscfg_list.txt') -MaxLength 8000
    $sections += Convert-CsvFileToHtmlSection -Title 'MIM Build Version' -Path (Join-Path $DiagDir 'MIM_BuildVersion.csv') -MaxRows 100
    $sections += Convert-CsvFileToHtmlSection -Title 'MIM Service / Portal Build' -Path (Join-Path $DiagDir 'MIM_Service_Portal_Build.csv') -MaxRows 100
    $sections += Convert-CsvFileToHtmlSection -Title 'SharePoint Build' -Path (Join-Path $DiagDir 'SharePoint_Build.csv') -MaxRows 100
    $sections += Convert-CsvFileToHtmlSection -Title 'Domain Controllers' -Path (Join-Path $DiagDir 'DomainControllers.csv') -MaxRows 200
    $sections += Convert-CsvFileToHtmlSection -Title 'Network Tests' -Path (Join-Path $DiagDir 'NetworkTests.csv') -MaxRows 300

    $generated = [System.Net.WebUtility]::HtmlEncode((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    $rootEncoded = [System.Net.WebUtility]::HtmlEncode($Root)
    $style = @'
<style>
body { font-family: Segoe UI, Meiryo, Arial, sans-serif; margin: 24px; color: #242424; }
h1 { border-bottom: 3px solid #555; padding-bottom: 8px; }
h2 { margin-top: 28px; border-left: 6px solid #777; padding-left: 10px; }
table { border-collapse: collapse; width: 100%; font-size: 12px; margin-top: 8px; }
th, td { border: 1px solid #ccc; padding: 4px 6px; vertical-align: top; }
th { background: #f2f2f2; position: sticky; top: 0; }
pre { background: #f7f7f7; border: 1px solid #ddd; padding: 10px; white-space: pre-wrap; overflow-wrap: anywhere; }
.path { color: #666; font-size: 12px; }
.note { color: #555; }
.missing { color: #a66; }
.error { color: #b00020; font-weight: 600; }
section { margin-bottom: 24px; }
</style>
'@
    $html = @(
        '<!DOCTYPE html>',
        '<html>',
        '<head>',
        '<meta charset="utf-8" />',
        '<title>MIM Diagnostic Summary</title>',
        $style,
        '</head>',
        '<body>',
        '<h1>MIM Diagnostic Summary</h1>',
        "<p><strong>Generated:</strong> $generated</p>",
        "<p><strong>Output root:</strong> $rootEncoded</p>",
        ($sections -join "`r`n"),
        '</body>',
        '</html>'
    )
    $html -join "`r`n" | Out-File -FilePath $ReportPath -Encoding UTF8
}

try {
    Write-DiagStatus 'Get-MIMDiagData.ps1 started' Green
    Write-DiagStatus "Output root: $Root" Green
    $modeText = if ($ObjectMode) { 'OBJ' } elseif ($FIMServiceOnly) { 'FIMServiceOnly' } elseif ($FIMSyncOnly) { 'FIMSyncOnly' } else { 'ALL' }
    Write-DiagStatus "Mode: $modeText" Green
    Write-DiagStatus "SyncServer: $SyncServer" Green
    if ([string]::IsNullOrWhiteSpace($MimServiceUri)) { Write-DiagStatus "MimServiceUri: <not specified / skipped>" Green } else { Write-DiagStatus "MimServiceUri: $MimServiceUri" Green }
    if ($SkipPCNS -or [string]::IsNullOrWhiteSpace($PcnSServer)) { Write-DiagStatus "PcnSServer: <not specified / skipped>" Green } else { Write-DiagStatus "PcnSServer: $PcnSServer" Green }
    if ($ObjectMode) {
        Export-ObjectDiagnostics -OutDir $SyncDataDir
    }
    else {
        if (-not $FIMServiceOnly) {
            Export-SyncDataAll -OutDir $SyncDataDir
            Export-ManagementAgentConfigs -OutDir $ConfigDir
            Export-MetaverseExtensionConfig -OutDir $ConfigDir
        }
        else {
            Write-DiagStatus 'Skipping FIMSynchronizationService-side collection because -FIMServiceOnly was specified.' DarkYellow
        }

        if (-not $FIMSyncOnly) {
            try { Export-MimServiceConfig -OutDir $ConfigDir } catch { Write-DiagError -Stage 'Export-MimServiceConfig' -Target $MimServiceUri -ErrorRecord $_ }
            Export-PCNSConfig -OutDir $ConfigDir
        }
        else {
            Write-DiagStatus 'Skipping FIMService-side and PCNS collection because -FIMSyncOnly was specified.' DarkYellow
        }

        Export-EventLogs -OutDir $EventLogDir
        Export-SystemDiagnostics -OutDir $DiagDir
        Export-NetworkDiagnostics -OutDir $DiagDir
    }
    try {
        $localWorkRoot = Join-Path $Root '_REMOTE_WORK'
        if (Test-Path -LiteralPath $localWorkRoot) {
            Remove-Item -LiteralPath $localWorkRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    [pscustomobject]@{StartTime=$ScriptStartTime;EndTime=Get-Date;Mode=$modeText;OutputRoot=$Root;SyncServer=$SyncServer;MimServiceUri=$MimServiceUri;PcnSServer=$PcnSServer;SkipPCNS=$SkipPCNS.IsPresent;ErrorLog=$Global:DiagErrorCsv} | Export-Csv -Path (Join-Path $Root 'RunSummary.csv') -NoTypeInformation -Encoding UTF8
    if (-not $ObjectMode) {
        try { New-MimDiagnosticHtmlReport -ReportPath (Join-Path $Root 'MIM_Diagnostic_Report.html') } catch { Write-DiagError -Stage 'New-MimDiagnosticHtmlReport' -Target $Root -ErrorRecord $_ }
    }
    Write-DiagStatus 'Get-MIMDiagData.ps1 completed' Green
    try { Stop-Transcript | Out-Null; $script:TranscriptStopped = $true } catch {}
    try {
        $zip = "$Root.zip"
        if (Test-Path $zip) { Remove-Item $zip -Force }
        Compress-Archive -Path $Root -DestinationPath $zip -Force -ErrorAction Stop
        Write-DiagStatus "ZIP output: $zip" Green
        Remove-Item -LiteralPath $Root -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch { Write-DiagError -Stage 'Compress-Archive' -Target $Root -ErrorRecord $_ }
}
catch { Write-DiagError -Stage 'Main' -Target 'Script' -ErrorRecord $_; Write-DiagStatus 'Get-MIMDiagData.ps1 stopped after recording an error. See Get-MIMDiagData_Errors.csv.' DarkYellow }
finally { if (-not $script:TranscriptStopped) { try { Stop-Transcript | Out-Null } catch {} } }
