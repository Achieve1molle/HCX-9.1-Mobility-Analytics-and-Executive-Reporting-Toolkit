<#
.SYNOPSIS
  HCX 9.1 Mobility Analytics and Executive Reporting Toolkit, Rev 1.0
.DESCRIPTION
  PowerShell 7 WPF utility for authenticating to an HCX 9.1 Manager, discovering
  mobility-group reporting endpoints, collecting time-filtered migration history,
  normalizing group and VM records, and exporting executive HTML, raw CSV, summary
  CSV, JSON, and Excel when the ImportExcel module is available.

  All generated files are placed beneath one per-launch run folder. Debug logging,
  sanitized REST request/response artifacts, endpoint discovery evidence, raw API
  captures, reports, and exports remain together for repeatability and auditability.

  HCX 9.x operations use REST and do not use VMware.VimAutomation.Hcx.
.NOTES
  Requires Windows, PowerShell 7+, and STA mode. No credentials or tokens are written
  to disk. HCX endpoint shapes vary by build and deployment; discovery is evidence-
  driven and records sanitized response shapes for troubleshooting.
#>
[CmdletBinding()]
param([switch]$NoRelaunch)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0

if ($PSVersionTable.PSVersion.Major -lt 7 -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if (-not $NoRelaunch) {
        $pwsh = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source
        if (-not $pwsh) { $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source }
        if (-not $pwsh) { throw 'PowerShell 7 or later is required.' }
        & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File $PSCommandPath -NoRelaunch
        exit $LASTEXITCODE
    }
}
if (-not $IsWindows) { throw 'This WPF utility requires Windows.' }
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml,System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:AppName = 'HCX91-Mobility-Analytics'
$script:AppVersion = '1.0.38'
$script:RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:LaunchBase = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:OutputBase = $script:LaunchBase
$script:RunDir = $null
$script:Dirs = @{}
$script:LogFile = $null
$script:TranscriptFile = $null
$script:TranscriptStarted = $false
$script:DebugLoggingEnabled = $true
$script:DebugSequence = 0
$script:LastHtml = $null
$script:LastExcel = $null
$script:RawGroups = @()
$script:RawWorkloads = @()
$script:Normalized = @()
$script:Hcx = [ordered]@{ BaseUri=''; Session=$null; Connected=$false; User=''; Headers=@{}; EndpointProfile=''; SelectedEndpoint=''; EndpointEvidence=@() ; VcGuid='231a7ede-4651-4d9b-87d2-c2d1603ba12a' }

# Initialize WPF control references before any logging or diagnostics execute.
# Set-StrictMode otherwise throws when startup logging evaluates an unset variable.
$script:Window       = $null
$script:txtHcx       = $null
$script:txtUser      = $null
$script:txtPassword  = $null
$script:btnConnect   = $null
$script:btnDisconnect= $null
$script:dpStart      = $null
$script:dpEnd        = $null
$script:txtOutputPath= $null
$script:txtWaveName  = $null
$script:txtWaveEstimatedVms = $null
$script:btnBrowse    = $null
$script:chkDebug     = $null
$script:lblEndpoint  = $null
$script:btnCollect   = $null
$script:gridData     = $null
$script:txtLog       = $null
$script:pbProgress   = $null
$script:lblStatus    = $null
$script:btnOpenHtml  = $null
$script:btnOpenRun   = $null
$script:btnClose     = $null
function Initialize-RunFolder {
    param([Parameter(Mandatory)][string]$BasePath)
    $candidate = [Environment]::ExpandEnvironmentVariables($BasePath.Trim())
    if (-not $candidate) { throw 'Select or enter an output base path.' }
    if (-not (Test-Path -LiteralPath $candidate)) { New-Item -ItemType Directory -Path $candidate -Force | Out-Null }
    $candidate = (Resolve-Path -LiteralPath $candidate).Path
    $test = Join-Path $candidate ('.hcx-write-test-' + [guid]::NewGuid().ToString('N'))
    try { [IO.File]::WriteAllText($test,'test'); Remove-Item -LiteralPath $test -Force }
    catch { throw "Output path is not writable: $candidate" }

    $script:OutputBase = $candidate
    $script:RunDir = Join-Path $script:OutputBase ("$($script:AppName)-Run-$($script:RunStamp)")
    $script:Dirs = [ordered]@{
        Root=$script:RunDir
        Logs=(Join-Path $script:RunDir 'Logs')
        Debug=(Join-Path $script:RunDir 'Debug-Artifacts')
        Raw=(Join-Path $script:RunDir 'Raw-API')
        Reports=(Join-Path $script:RunDir 'Reports')
        Exports=(Join-Path $script:RunDir 'Exports')
        Config=(Join-Path $script:RunDir 'Configuration')
    }
    foreach ($path in $script:Dirs.Values) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    $script:LogFile = Join-Path $script:Dirs.Logs ("HCX91-Mobility-Analytics-$($script:RunStamp).log")
    $script:TranscriptFile = Join-Path $script:Dirs.Logs ("HCX91-PowerShell-Transcript-$($script:RunStamp).log")
}
Initialize-RunFolder -BasePath $script:LaunchBase

function DoEvents { try { [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background) } catch {} }
function Protect-HcxDiagnosticText {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    $safe = $Text
    $safe = $safe -replace '(?i)("?(?:password|passwd|pwd|token|access_token|refresh_token|authorization|x-hm-authorization|cookie|set-cookie|xsrf-token|csrf-token)"?\s*[:=]\s*")([^"]+)(")','$1********$3'
    $safe = $safe -replace '(?i)((?:authorization|x-hm-authorization|cookie|set-cookie|xsrf-token|csrf-token)\s*[:=]\s*)([^;\r\n]+)','$1********'
    $safe = $safe -replace '(?i)(Basic|Bearer)\s+[A-Za-z0-9+/=._-]+','$1 ********'
    try { if ($script:txtPassword -and $script:txtPassword.Password) { $safe = $safe -replace [regex]::Escape($script:txtPassword.Password),'********' } } catch {}
    return $safe
}
function ConvertTo-HcxSanitizedObject {
    param($Object)
    if ($null -eq $Object) { return $null }
    try { return (Protect-HcxDiagnosticText ($Object | ConvertTo-Json -Depth 80 -Compress)) | ConvertFrom-Json -AsHashtable -Depth 80 }
    catch { return Protect-HcxDiagnosticText ([string]$Object) }
}
function Write-HcxDebug {
    param([string]$Message,[string]$Category='GENERAL')
    if (-not $script:DebugLoggingEnabled) { return }
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') [DEBUG] [$Category] $(Protect-HcxDiagnosticText $Message)"
    Add-Content -LiteralPath $script:LogFile -Value $line
    if ($null -ne $script:txtLog) { $script:txtLog.AppendText($line + [Environment]::NewLine); $script:txtLog.ScrollToEnd(); DoEvents }
}
function Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','PASS')][string]$Level='INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') [$Level] $(Protect-HcxDiagnosticText $Message)"
    Add-Content -LiteralPath $script:LogFile -Value $line
    if ($null -ne $script:txtLog) { $script:txtLog.AppendText($line + [Environment]::NewLine); $script:txtLog.ScrollToEnd(); DoEvents }
}
function Save-HcxDebugArtifact {
    param([string]$Category,[string]$Operation,$Data,[hashtable]$Metadata=@{})
    if (-not $script:DebugLoggingEnabled) { return $null }
    $script:DebugSequence++
    $safeCategory = $Category -replace '[^A-Za-z0-9_-]','_'
    $safeOperation = $Operation -replace '[^A-Za-z0-9_-]','_'
    $path = Join-Path $script:Dirs.Debug ('{0:d4}-{1}-{2}-{3}.json' -f $script:DebugSequence,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'),$safeCategory,$safeOperation)
    [ordered]@{ CapturedAt=(Get-Date).ToString('o'); Category=$Category; Operation=$Operation; Metadata=(ConvertTo-HcxSanitizedObject $Metadata); Data=(ConvertTo-HcxSanitizedObject $Data) } |
        ConvertTo-Json -Depth 90 | Set-Content -LiteralPath $path -Encoding utf8BOM
    Write-HcxDebug "Artifact saved: $path" 'ARTIFACT'
    return $path
}
function Start-HcxDiagnosticTranscript {
    if ($script:TranscriptStarted) { return }
    try { Start-Transcript -LiteralPath $script:TranscriptFile -IncludeInvocationHeader -Force | Out-Null; $script:TranscriptStarted=$true }
    catch { Write-HcxDebug "Transcript could not start: $($_.Exception.Message)" 'TRANSCRIPT' }
}
function Stop-HcxDiagnosticTranscript { if ($script:TranscriptStarted) { try { Stop-Transcript | Out-Null } catch {}; $script:TranscriptStarted=$false } }
function Wait-Tcp {
    param([string]$HostName,[int]$Port=443,[int]$Seconds=15)
    $end=(Get-Date).AddSeconds($Seconds)
    while((Get-Date) -lt $end){try{$c=[Net.Sockets.TcpClient]::new();$a=$c.BeginConnect($HostName,$Port,$null,$null);if($a.AsyncWaitHandle.WaitOne(1000)){$c.EndConnect($a);$c.Close();return $true};$c.Close()}catch{}}
    return $false
}
function ConvertFrom-HcxJsonIfNeeded {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string]) {
        $text=$InputObject.Trim()
        if (($text.StartsWith('{') -and $text.EndsWith('}')) -or ($text.StartsWith('[') -and $text.EndsWith(']'))) {
            try { return $text | ConvertFrom-Json -AsHashtable -Depth 100 }
            catch { return $InputObject }
        }
    }
    return $InputObject
}
function Get-HcxPropertyValue {
    param($Object,[string]$Name)
    $Object=ConvertFrom-HcxJsonIfNeeded $Object
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return ConvertFrom-HcxJsonIfNeeded $Object[$Name] }
        $key=@($Object.Keys|Where-Object{[string]$_ -ieq $Name}|Select-Object -First 1)
        if ($key) { return ConvertFrom-HcxJsonIfNeeded $Object[$key[0]] }
        return $null
    }
    $p=$Object.PSObject.Properties[$Name]; if($p){return ConvertFrom-HcxJsonIfNeeded $p.Value}; return $null
}
function Get-HcxFirstValue {
    param($Object,[string[]]$Names)
    foreach ($name in $Names){$value=Get-HcxPropertyValue $Object $name;if($null -ne $value-and-not[string]::IsNullOrWhiteSpace([string]$value)){return $value}}
    return $null
}
function Get-HcxResponseItems {
    param($Response)
    $Response=ConvertFrom-HcxJsonIfNeeded $Response
    if($null -eq $Response){return @()}
    foreach($name in 'items','elements','results','content','list','groups','migrations','workloads','data'){
        $value=Get-HcxPropertyValue $Response $name
        if($null -ne $value){
            if($name -eq 'data'){
                foreach($inner in 'items','elements','results','content','groups','migrations','workloads'){$x=Get-HcxPropertyValue $value $inner;if($null -ne $x){return @($x)}}
            }
            return @($value)
        }
    }
    if($Response-is [System.Collections.IEnumerable] -and $Response -isnot [string] -and $Response -isnot [System.Collections.IDictionary]){return @($Response)}
    return @($Response)
}
function Test-HcxTransientTransportError {
    param($ErrorRecord)
    $m=[Collections.Generic.List[string]]::new();$e=$ErrorRecord.Exception;while($e){$m.Add([string]$e.Message);$e=$e.InnerException}
    return (($m-join' | ') -match '(?i)response ended prematurely|connection.*closed|forcibly closed|unexpected end|HTTP/2.*error|transport connection|request was aborted|error occurred while sending')
}
function Invoke-HcxRest {
    param([string]$Method,[string]$Path,$Body=$null,[hashtable]$Headers=@{},[switch]$AllowFailure,[int]$MaxAttempts=3)
    if(-not $script:Hcx.BaseUri){throw 'HCX base URI is not initialized.'}
    $uri=if($Path -match '^https?://'){$Path}else{$script:Hcx.BaseUri.TrimEnd('/')+'/'+$Path.TrimStart('/')}
    $all=@{Accept='application/json'};foreach ($k in $script:Hcx.Headers.Keys){$all[$k]=$script:Hcx.Headers[$k]};foreach ($k in $Headers.Keys){$all[$k]=$Headers[$k]}
    $requestId=[guid]::NewGuid().ToString('N');$requestArtifact=Save-HcxDebugArtifact 'REST-REQUEST' $Method ([ordered]@{RequestId=$requestId;Method=$Method;Uri=$uri;Headers=$all;Body=$Body})
    $last=$null
    for($attempt=1;$attempt -le $MaxAttempts;$attempt++){
        try{
            $p=@{Method=$Method;Uri=$uri;Headers=$all;SkipCertificateCheck=$true;ErrorAction='Stop';TimeoutSec=120}
            if($script:Hcx.Session){$p.WebSession=$script:Hcx.Session};if($null -ne $Body){$p.Body=$Body|ConvertTo-Json -Depth 80 -Compress;$p.ContentType='application/json'}
            if((Get-Command Invoke-RestMethod).Parameters.ContainsKey('DisableKeepAlive')){$p.DisableKeepAlive=$true};if((Get-Command Invoke-RestMethod).Parameters.ContainsKey('HttpVersion')){$p.HttpVersion='1.1'}
            $response=Invoke-RestMethod @p;$response=ConvertFrom-HcxJsonIfNeeded $response
            Save-HcxDebugArtifact 'REST-RESPONSE' $Method $response @{RequestId=$requestId;Uri=$uri;Attempt=$attempt;RequestArtifact=$requestArtifact}|Out-Null
            return $response
        }catch{
            $last=$_;$status='';try{$status=[int]$_.Exception.Response.StatusCode}catch{};$detail=$_.Exception.Message;try{if($_.ErrorDetails.Message){$detail=[string]$_.ErrorDetails.Message}}catch{}
            Save-HcxDebugArtifact 'REST-ATTEMPT-FAILURE' $Method ([ordered]@{RequestId=$requestId;Attempt=$attempt;Uri=$uri;Status=$status;Detail=$detail})|Out-Null
            if($AllowFailure){Write-HcxDebug "Optional endpoint unavailable: $Path HTTP=$status $detail" 'REST-OPTIONAL';return $null}
            if((Test-HcxTransientTransportError $_) -and $attempt -lt $MaxAttempts){Start-Sleep -Milliseconds (300*$attempt);continue}
            break
        }
    }
    if($AllowFailure){return $null};throw "HCX REST $Method $Path failed: $($last.Exception.Message)"
}
function Connect-Hcx91Rest {
    param([string]$Fqdn,[string]$User,[string]$Password)
    $hostName=($Fqdn-replace'^https?://','').TrimEnd('/')
    if(-not $hostName-or-not $User-or-not $Password){throw 'HCX Manager FQDN, username, and password are required.'}
    if(-not (Wait-Tcp $hostName 443 15)){throw"TCP 443 is not reachable on $hostName."}
    $script:Hcx.BaseUri='https://'+$hostName;$script:Hcx.Session=[Microsoft.PowerShell.Commands.WebRequestSession]::new();$script:Hcx.Headers=@{}
    $uri=$script:Hcx.BaseUri+'/hybridity/api/sessions';$json=@{username=$User;password=$Password}|ConvertTo-Json -Compress;$pair=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$User`:$Password"));$failures=@()
    foreach ($profile in @(@{Name='HCX JSON session';Headers=@{Accept='application/json'}},@{Name='HCX JSON session with Basic bootstrap';Headers=@{Accept='application/json';Authorization="Basic $pair"}})){
        try{$r=Invoke-WebRequest -Method POST -Uri $uri -Headers $profile.Headers -Body $json -ContentType 'application/json' -WebSession $script:Hcx.Session -SkipCertificateCheck -TimeoutSec 90 -ErrorAction Stop;$token=[string]$r.Headers['x-hm-authorization'];if(-not $token){$token=[string]$r.Headers['X-HM-Authorization']};if(-not $token){throw 'x-hm-authorization response header was missing.'};$script:Hcx.Headers=@{'x-hm-authorization'=$token};$script:Hcx.EndpointProfile=$profile.Name;break}catch{$failures+="$($profile.Name): $($_.Exception.Message)"}
    }
    if (-not $script:Hcx.Headers.ContainsKey('x-hm-authorization')) { throw ('HCX authentication failed. ' + ($failures -join ' | ')) }
    $script:Hcx.User=$User;$script:Hcx.Connected=$true;Log "Authenticated to $hostName using HCX x-hm-authorization." PASS
}
function Disconnect-HcxRest { $script:Hcx.Connected=$false;$script:Hcx.Session=$null;$script:Hcx.Headers=@{};$script:Hcx.SelectedEndpoint='';$script:Hcx.EndpointEvidence=@();Log 'Disconnected from HCX.' INFO }

function ConvertTo-DateTimeSafe {
    param($Value)
    if($null -eq $Value-or[string]::IsNullOrWhiteSpace([string]$Value)){return $null}
    if($Value-is [datetime]){return $Value}
    $number=0L;if([long]::TryParse([string]$Value,[ref]$number)){
        try{if($number-gt100000000000){return[DateTimeOffset]::FromUnixTimeMilliseconds($number).LocalDateTime};if($number-gt1000000000){return[DateTimeOffset]::FromUnixTimeSeconds($number).LocalDateTime}}catch{}
    }
    $dt=[datetime]::MinValue;if([datetime]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)){return $dt}
    if([datetime]::TryParse([string]$Value,[ref]$dt)){return $dt};return $null
}
function ConvertTo-BytesSafe {
    param($Value,[string]$Name='')
    if($null -eq $Value){return 0.0};$d=0.0;if(-not[double]::TryParse(([string]$Value-replace',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){return 0.0}
    if($Name -match '(?i)kb'){return $d*1KB};if($Name -match '(?i)mb'){return $d*1MB};if($Name -match '(?i)gb'){return $d*1GB};if($Name -match '(?i)tb'){return $d*1TB};return $d
}
function Get-HcxMetricValue {
    param($Object,[string[]]$Names)
    foreach ($name in $Names){$v=Get-HcxPropertyValue $Object $name;if($null -ne $v-and-not[string]::IsNullOrWhiteSpace([string]$v)){return $v}}
    foreach($container in 'info','summary','countSummary','resourceSummary','entity','migration','details','statistics','stats','progress','transfer','transferStats','migrationStats','config','networkParams'){
        $node=Get-HcxPropertyValue $Object $container;if($node){foreach ($name in $Names){$v=Get-HcxPropertyValue $node $name;if($null -ne $v-and-not[string]::IsNullOrWhiteSpace([string]$v)){return $v}}}
    }
    return $null
}
function Test-HcxMobilityShape {
    param($Response)
    $items=@(Get-HcxResponseItems $Response);if($items.Count-eq0){return 0};$score=0
    foreach ($item in @($items|Select-Object -First 3)){$names=if($item-is [System.Collections.IDictionary]){@($item.Keys)}else{@($item.PSObject.Properties.Name)};$joined=$names-join' ';if($joined -match '(?i)migration|workload|group|entity|vm'){ $score+=5 };if($joined -match '(?i)status|state'){ $score+=2 };if($joined -match '(?i)start|end|time|transfer|progress'){ $score+=2 }}
    return ($score + $items.Count)
}
function Find-HcxMobilityEndpoint {
    if ([string]::IsNullOrWhiteSpace([string]$script:Hcx.VcGuid)) {
        throw 'The HCX vCenter GUID is not configured.'
    }

    $encodedVcGuid = [uri]::EscapeDataString([string]$script:Hcx.VcGuid)
    $path = "/hybridity/api/v2/mobility-groups/query?vcGuid=$encodedVcGuid"
    $body = [ordered]@{
        pageParameters = [ordered]@{
            pageNumber = 1
            pageSize = 500
            sortBy = @(
                [ordered]@{
                    field = 'mobilityGroupCreationTime'
                    direction = 'DESC'
                }
            )
        }
        filter = [ordered]@{
            includePermissions = $true
        }
    }

    Write-HcxDebug "Using browser-confirmed HCX 9.1 v2 mobility query for vcGuid=$($script:Hcx.VcGuid)." 'ENDPOINT-DISCOVERY'
    $response = Invoke-HcxRest -Method 'POST' -Path $path -Body $body
    $items = @(Get-HcxResponseItems -Response $response)
    $score = Test-HcxMobilityShape -Response $response

    $script:Hcx.SelectedEndpoint = "POST $path"
    $script:Hcx.EndpointEvidence = @(
        [pscustomobject]@{
            Method = 'POST'
            Path = $path
            ItemCount = $items.Count
            Score = $score
            Result = 'Accepted'
        }
    )

    Save-HcxDebugArtifact -Category 'DISCOVERY' -Operation 'MOBILITY-ENDPOINTS' -Data $script:Hcx.EndpointEvidence | Out-Null
    Log "Selected reporting endpoint: $($script:Hcx.SelectedEndpoint); items=$($items.Count); score=$score." PASS

    return [pscustomobject]@{
        Method = 'POST'
        Path = $path
        Body = $body
        Score = $score
        ItemCount = $items.Count
        Response = $response
    }
}

function Get-HcxIntentForVm {
 param([string]$GroupId,[string]$MigrationId,[string]$EntityId)
 try {
  $r=Invoke-HcxRest -Method 'POST' -Path '/hybridity/api/mobility/groups/intents' -Body ([ordered]@{filters=[ordered]@{groups=@([ordered]@{migrationGroupId=$GroupId})}})
  foreach($g in @(Get-HcxResponseItems $r)){foreach($i in @((Get-HcxPropertyValue $g 'migrations'))){
   $mid=[string](Get-HcxPropertyValue $i 'migrationId');$e=Get-HcxPropertyValue $i 'entity';$eid=[string](Get-HcxPropertyValue $e 'entityId')
   if(($MigrationId-and$mid-eq$MigrationId)-or($EntityId-and$eid-eq$EntityId)){return $i}
  }}
 }catch{Write-HcxDebug "Intent enrichment failed for groupId=$GroupId. $($_.Exception.Message)" 'GROUP-INTENT'}
 return $null
}
function Get-HcxWorkloadsForGroup {
    param($Group)

    $groupId = [string](Get-HcxMetricValue -Object $Group -Names @('id','groupId','mobilityGroupId','uuid'))
    if ([string]::IsNullOrWhiteSpace($groupId)) {
        Write-HcxDebug 'Mobility group did not contain an ID; VM-detail query skipped.' 'GROUP-DETAIL'
        return @()
    }

    $path = "/hybridity/api/v2/mobility-groups/$groupId/migrations/query?vcGuid=$($script:Hcx.VcGuid)"
    $body = [ordered]@{
        pageParameters = [ordered]@{
            pageNumber = 1
            pageSize = 500
            sortBy = @([ordered]@{ field = 'entityName'; direction = 'ASC' })
        }
        filter = [ordered]@{ archiveScope = 'EXCLUDE_ARCHIVE' }
    }

    try {
        Write-HcxDebug "Querying browser-confirmed HCX 9.1 v2 VM-detail endpoint for groupId=$groupId." 'GROUP-DETAIL'
        $response = Invoke-HcxRest -Method 'POST' -Path $path -Body $body
        $items = @(Get-HcxResponseItems -Response $response)
        Write-HcxDebug "VM-detail query completed for groupId=$groupId; items=$($items.Count)." 'GROUP-DETAIL'
        return $items
    }
    catch {
        Write-HcxDebug "VM-detail query failed for groupId=$groupId. $($_.Exception.Message)" 'GROUP-DETAIL'
        return @()
    }
}
function Normalize-HcxMigrationRecord {
    param($Group, $Workload)


    $info = Get-HcxPropertyValue -Object $Workload -Name 'info'
    $entity = Get-HcxPropertyValue -Object $info -Name 'entity'
    $summary = Get-HcxPropertyValue -Object $Workload -Name 'summary'
    $resourceSummary = Get-HcxPropertyValue -Object $summary -Name 'resourceSummary'
    $progressSummary = Get-HcxPropertyValue -Object $summary -Name 'progressSummary'
    $transferSummary = Get-HcxPropertyValue -Object $progressSummary -Name 'transfer'
    $totalSummary = Get-HcxPropertyValue -Object $progressSummary -Name 'total'
    $errorSummary = Get-HcxPropertyValue -Object $progressSummary -Name 'error'
    $warningSummary = Get-HcxPropertyValue -Object $progressSummary -Name 'warning'
    $groupIdForIntent=[string](Get-HcxMetricValue $Group @('id','groupId','mobilityGroupId','uuid'))
    $migrationIdForIntent=[string](Get-HcxPropertyValue $info 'id')
    $entityIdForIntent=[string](Get-HcxPropertyValue $entity 'entityId')
    $intent=Get-HcxIntentForVm -GroupId $groupIdForIntent -MigrationId $migrationIdForIntent -EntityId $entityIdForIntent
    $intentEntity=Get-HcxPropertyValue $intent 'entity';$intentSummary=Get-HcxPropertyValue $intentEntity 'summary'
    $networkParams=Get-HcxPropertyValue $intent 'networkParams';$networkMappings=@((Get-HcxPropertyValue $networkParams 'networkMappings'))

    $vm = [string](Get-HcxPropertyValue -Object $entity -Name 'entityName')
    if (-not $vm) { $vm = [string](Get-HcxMetricValue -Object $Workload -Names @('entityName','vmName','name','displayName','workloadName')) }
    $groupName = [string](Get-HcxMetricValue -Object $Group -Names @(
        'name','displayName','groupName'
    ))

    $status = [string](Get-HcxPropertyValue -Object $progressSummary -Name 'status')
    if (-not $status) { $status = [string](Get-HcxMetricValue -Object $Workload -Names @('status','state','migrationState','overallStatus')) }
    if (-not $status) {
        $status = [string](Get-HcxMetricValue -Object $Group -Names @(
            'status','state','migrationState'
        ))
    }

    $startValue = Get-HcxPropertyValue -Object $progressSummary -Name 'startTimestamp'
    if (-not $startValue) { $startValue = Get-HcxMetricValue -Object $Workload -Names @(
        'startTime','startedAt','startDate','migrationStartTime',
        'executionStartTime','createdTime'
    ) }
    $start = ConvertTo-DateTimeSafe -Value $startValue
    if (-not $start) {
        $startValue = Get-HcxMetricValue -Object $Group -Names @(
            'startTime','startedAt','startDate','migrationStartTime',
            'executionStartTime','migrationStartTimestamp','createdTimestamp','createdTime'
        )
        $start = ConvertTo-DateTimeSafe -Value $startValue
    }

    $endValue = Get-HcxPropertyValue -Object $progressSummary -Name 'lastUpdatedTimestamp'
    if (-not $endValue) { $endValue = Get-HcxMetricValue -Object $Workload -Names @(
        'endTime','completedAt','completionTime','endDate',
        'migrationEndTime','lastUpdatedTime','lastUpdatedTimestamp','updatedTime'
    ) }
    $end = ConvertTo-DateTimeSafe -Value $endValue
    if (-not $end) {
        $endValue = Get-HcxMetricValue -Object $Group -Names @('endTime','completedAt','completionTime','lastUpdatedTimestamp','lastUpdatedTime','updatedTime')
        $end = ConvertTo-DateTimeSafe -Value $endValue
    }

    $durationSec = 0.0
    if ($start -and $end) {
        $durationSec = ($end - $start).TotalSeconds
    }
    else {
        $rawDuration = Get-HcxMetricValue -Object $Workload -Names @(
            'durationSeconds','elapsedSeconds','duration','elapsedTime'
        )
        [void][double]::TryParse([string]$rawDuration, [ref]$durationSec)
    }

    $storageRaw = Get-HcxMetricValue -Object $Workload -Names @(
        'bytesTransferred','transferredBytes','diskSize','storageBytes',
        'dataTransferred','totalBytes','disk'
    )
    $storageName = 'bytes'
    if ($null -eq $storageRaw) {
        $storageRaw = Get-HcxMetricValue -Object $Workload -Names @(
            'diskSizeGB','storageGB'
        )
        $storageName = 'gb'
    }
    $storage = ConvertTo-BytesSafe -Value $storageRaw -Name $storageName

    $memoryRaw = Get-HcxMetricValue -Object $Workload -Names @(
        'mem','memorySize','memoryBytes','memorySizeMB','memoryMB','memoryGB'
    )
    $memoryName = 'bytes'
    $memoryMbValue = Get-HcxMetricValue -Object $Workload -Names @('memorySizeMB','memoryMB')
    $memoryGbValue = Get-HcxMetricValue -Object $Workload -Names @('memoryGB')
    if ($null -ne $memoryMbValue) {
        $memoryName = 'mb'
    }
    elseif ($null -ne $memoryGbValue) {
        $memoryName = 'gb'
    }
    $memory = ConvertTo-BytesSafe -Value $memoryRaw -Name $memoryName
    $exactMemory = Get-HcxPropertyValue -Object $resourceSummary -Name 'mem'
    if ($null -ne $exactMemory) { $memory = [double]$exactMemory }

    $cpu = 0
    $cpuValue = Get-HcxMetricValue -Object $Workload -Names @(
        'cpu','numCpu','cpuCount','vCpu','vCPUs'
    )
    [void][int]::TryParse([string]$cpuValue, [ref]$cpu)
    $exactCpu = Get-HcxPropertyValue -Object $resourceSummary -Name 'cpu'
    if ($null -ne $exactCpu) { $cpu = [int]$exactCpu }
    $exactDisk = Get-HcxPropertyValue -Object $resourceSummary -Name 'disk'
    if ($null -ne $exactDisk) { $storage = [double]$exactDisk }

    $errorMessage = [string](Get-HcxMetricValue -Object $Workload -Names @(
        'errorMessage','error','failureReason','message','lastError','reason'
    ))
    $network = [string](Get-HcxMetricValue -Object $Workload -Names @(
        'destNetworkName','destinationNetworkName','networkName',
        'destinationNetwork','networkMappings'
    ))
    $guest = [string](Get-HcxMetricValue -Object $Workload -Names @(
        'guestFullName','guestOS','guestId','osName','operatingSystem'
    ))
    if(-not$guest){$guest=[string](Get-HcxPropertyValue $intentSummary 'guestFullName')}
    $guestId=[string](Get-HcxPropertyValue $intentSummary 'guestId');$guestHostName=[string](Get-HcxPropertyValue $intentSummary 'guestHostName')
    $destNetworks=@($networkMappings|ForEach-Object{[string](Get-HcxPropertyValue $_ 'destNetworkName')}|Where-Object{$_}|Select-Object -Unique)-join'; '
    if($destNetworks){$network=$destNetworks}
    $destNetworkIds=@($networkMappings|ForEach-Object{[string](Get-HcxPropertyValue $_ 'destNetworkId')}|Where-Object{$_}|Select-Object -Unique)-join'; '
    $destNetworkTypes=@($networkMappings|ForEach-Object{[string](Get-HcxPropertyValue $_ 'destNetworkType')}|Where-Object{$_}|Select-Object -Unique)-join'; '
    $mappingDetails=@($networkMappings|ForEach-Object{$src=[string](Get-HcxPropertyValue $_ 'srcNetworkName');$dst=[string](Get-HcxPropertyValue $_ 'destNetworkName');if($src-and$dst){"$src to $dst"}}|Where-Object{$_})-join'; '
    $placement=@((Get-HcxPropertyValue $intent 'placement')|ForEach-Object{[string](Get-HcxPropertyValue $_ 'name')}|Where-Object{$_})-join'; '
    $migrationType = [string](Get-HcxMetricValue -Object $Workload -Names @(
        'migrationType','type','switchoverType','transferType'
    ))
    $attempt = [string](Get-HcxMetricValue -Object $Workload -Names @(
        'attempt','attemptNumber','retryCount','sequence'
    ))
    $groupId = [string](Get-HcxMetricValue -Object $Group -Names @(
        'id','groupId','mobilityGroupId','uuid'
    ))

    $progressMessage = [string](Get-HcxPropertyValue -Object $progressSummary -Name 'message')
    $errorMessages = @((Get-HcxPropertyValue -Object $errorSummary -Name 'messages')) -join '; '
    if (-not $errorMessage) { $errorMessage = if ($errorMessages) { $errorMessages } elseif ($status -match 'ERROR|FAIL|CANCEL') { $progressMessage } else { '' } }
    $warningMessages = @((Get-HcxPropertyValue -Object $warningSummary -Name 'messages')) -join '; '
    $actualBytesTransferred = [double](Get-HcxPropertyValue -Object $totalSummary -Name 'bytesTransferred')
    $provisionedBytes = [double](Get-HcxPropertyValue -Object $totalSummary -Name 'totalBytes')
    if ($provisionedBytes -le 0) { $provisionedBytes = $storage }
    $throughputBasis = if ($actualBytesTransferred -gt 0) { $actualBytesTransferred } else { $provisionedBytes }
    $throughputIsEstimated = ($actualBytesTransferred -le 0 -and $throughputBasis -gt 0)
    $averageMBps = if ($durationSec -gt 0 -and $throughputBasis -gt 0) { [math]::Round(($throughputBasis / 1MB) / $durationSec, 2) } else { 0 }
    $guestDisplay = if ($guest) { $guest } else { 'Not returned by migration API' }
    $networkDisplay = if ($network) { $network } else { '' }

    [pscustomobject][ordered]@{
        GroupName         = $groupName
        GroupId           = $groupId
        MigrationId       = [string](Get-HcxPropertyValue -Object $info -Name 'id')
        VMEntityId         = [string](Get-HcxPropertyValue -Object $entity -Name 'entityId')
        VMName            = $vm
        Status            = $status
        MigrationType     = $migrationType
        StartTime         = $start
        EndTime           = $end
        DurationMinutes   = [math]::Round(($durationSec / 60), 2)
        StorageBytes      = [math]::Round($storage, 0)
        StorageGB         = [math]::Round(($storage / 1GB), 2)
        ProvisionedGB     = [math]::Round(($provisionedBytes / 1GB), 2)
        ActualBytesTransferred = [math]::Round($actualBytesTransferred, 0)
        AverageThroughputMBps = $averageMBps
        ThroughputBasis   = if ($throughputIsEstimated) { 'Estimated from provisioned bytes / elapsed time' } elseif ($actualBytesTransferred -gt 0) { 'HCX bytesTransferred / elapsed time' } else { 'Not available' }
        MemoryBytes       = [math]::Round($memory, 0)
        MemoryGB          = [math]::Round(($memory / 1GB), 2)
        vCPU              = $cpu
        ComputeScore      = [math]::Round(($cpu * [math]::Max(($durationSec / 60), 1)), 2)
        DestinationNetwork= $networkDisplay
        DestinationNetworkId=$destNetworkIds
        DestinationNetworkType=$destNetworkTypes
        NetworkMappingDetails=$mappingDetails
        DestinationPlacement=$placement
        GuestOS           = $guestDisplay
        GuestId=$guestId
        GuestHostName=$guestHostName
        ProgressMessage   = $progressMessage
        WarningMessages   = $warningMessages
        ErrorMessage      = $errorMessage
        Attempt           = $attempt
        Raw               = $Workload
    }
}


function ConvertTo-HcxGroupSummary {
    param($Group)

    $info = Get-HcxPropertyValue -Object $Group -Name 'info'
    $summary = Get-HcxPropertyValue -Object $Group -Name 'summary'
    $count = Get-HcxPropertyValue -Object $summary -Name 'countSummary'
    $resource = Get-HcxPropertyValue -Object $summary -Name 'resourceSummary'
    $source = Get-HcxPropertyValue -Object $info -Name 'source'
    $destination = Get-HcxPropertyValue -Object $info -Name 'destination'

    $created = ConvertTo-DateTimeSafe (Get-HcxPropertyValue -Object $info -Name 'createdTimestamp')
    $started = ConvertTo-DateTimeSafe (Get-HcxPropertyValue -Object $info -Name 'migrationStartTimestamp')
    $updated = ConvertTo-DateTimeSafe (Get-HcxPropertyValue -Object $info -Name 'lastUpdatedTimestamp')
    $effective = if ($started) { $started } elseif ($updated) { $updated } else { $created }
    $elapsedMinutes = $null
    if ($started -and $updated -and $updated -ge $started) {
        $elapsedMinutes = [math]::Round(($updated - $started).TotalMinutes, 2)
    }

    $memoryBytes = [double](Get-HcxPropertyValue -Object $resource -Name 'totalMem')
    $diskBytes = [double](Get-HcxPropertyValue -Object $resource -Name 'totalDisk')

    [pscustomobject][ordered]@{
        GroupName = [string](Get-HcxPropertyValue -Object $info -Name 'name')
        GroupId = [string](Get-HcxPropertyValue -Object $info -Name 'id')
        State = [string](Get-HcxPropertyValue -Object $summary -Name 'state')
        ConfigurationStatus = [string](Get-HcxPropertyValue -Object $summary -Name 'configurationStatus')
        EffectiveTimestamp = $effective
        CreatedTimestamp = $created
        MigrationStartTimestamp = $started
        LastUpdatedTimestamp = $updated
        GroupElapsedMinutes = $elapsedMinutes
        GroupElapsedDisplay = if ($null -ne $elapsedMinutes) { '{0:N2} min' -f $elapsedMinutes } else { 'Not returned' }
        SourceSite = [string](Get-HcxPropertyValue -Object $source -Name 'name')
        SourceVCenter = [string](Get-HcxPropertyValue -Object $source -Name 'infraManagerName')
        DestinationSite = [string](Get-HcxPropertyValue -Object $destination -Name 'name')
        DestinationVCenter = [string](Get-HcxPropertyValue -Object $destination -Name 'infraManagerName')
        TotalVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVms')
        DraftVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsDraft')
        QueuedVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsQueued')
        WarningVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsWarning')
        CancelledVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsCancelled')
        ErrorVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsError')
        InTransferVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsInTransfer')
        InSwitchoverVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsInSwitchover')
        WaitingSwitchoverVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsWaitingSwitchover')
        CompletedVMs = [int](Get-HcxPropertyValue -Object $count -Name 'totalVmsCompleted')
        TotalVcpu = [int](Get-HcxPropertyValue -Object $resource -Name 'totalCpus')
        TotalMemoryBytes = $memoryBytes
        TotalMemoryGB = [math]::Round($memoryBytes / 1GB, 2)
        TotalMemoryDisplay = if ($memoryBytes -gt 0) { '{0:N2} GB' -f ($memoryBytes / 1GB) } else { 'Not returned' }
        TotalDiskBytes = $diskBytes
        TotalDiskGB = [math]::Round($diskBytes / 1GB, 2)
        TotalDiskDisplay = if ($diskBytes -gt 0) { '{0:N2} GB' -f ($diskBytes / 1GB) } else { 'Not returned' }
        Username = [string](Get-HcxPropertyValue -Object $info -Name 'username')
        ServiceMeshId = [string](Get-HcxPropertyValue -Object $info -Name 'servicemeshId')
    }
}

function Test-MigratedStatus { param([string]$Status) return $Status -match '(?i)complete|completed|migrated|success|succeeded|done' }
function Test-ErrorStatus { param([string]$Status,[string]$Error) return ($Status -match '(?i)error|failed|failure|canceled|cancelled')-or-not[string]::IsNullOrWhiteSpace($Error) }
function Get-HcxMobilityData {
    param([datetime]$StartDate,[datetime]$EndDate)
    if(-not $script:Hcx.Connected){throw 'Connect to HCX first.'};$selected=Find-HcxMobilityEndpoint;$groups=@(Get-HcxResponseItems $selected.Response);$records=[Collections.Generic.List[object]]::new();$groupIndex=0
    foreach ($group in $groups){$groupIndex++;$workloads=@(Get-HcxWorkloadsForGroup $group);if($workloads.Count-eq0){$workloads=@($group)};foreach ($w in $workloads){$n=Normalize-HcxMigrationRecord $group $w;$effective=if($n.StartTime){$n.StartTime}elseif($n.EndTime){$n.EndTime}else{$null};if($effective -and $effective -ge $StartDate -and $effective -le $EndDate){$records.Add($n)}};if($script:pbProgress){$script:pbProgress.Value=[math]::Round(($groupIndex/[math]::Max($groups.Count,1))*100,0)};DoEvents}
    $script:RawGroups = $groups
    $script:GroupSummaries = @(
        foreach ($group in $groups) {
            $groupSummary = ConvertTo-HcxGroupSummary -Group $group
            if ($groupSummary.EffectiveTimestamp -and $groupSummary.EffectiveTimestamp -ge $StartDate -and $groupSummary.EffectiveTimestamp -le $EndDate) {
                $groupSummary
            }
        }
    )
    $script:Normalized = @($records)
    $rawPath=Join-Path $script:Dirs.Raw ("HCX91-Mobility-Raw-$($script:RunStamp).json");[ordered]@{CapturedAt=(Get-Date).ToString('o');SourceEndpoint=$script:Hcx.SelectedEndpoint;StartDate=$StartDate;EndDate=$EndDate;Groups=(ConvertTo-HcxSanitizedObject $groups);Normalized=(ConvertTo-HcxSanitizedObject $script:Normalized)}|ConvertTo-Json -Depth 90|Set-Content -LiteralPath $rawPath -Encoding utf8BOM
    Log "Time-filtered collection completed. GroupsExamined=$($groups.Count); GroupSummariesIncluded=$($script:GroupSummaries.Count); VMRecordsIncluded=$($script:Normalized.Count); Start=$StartDate; End=$EndDate." PASS
    return $script:Normalized
}
function Format-Bytes { param([double]$Bytes) if($Bytes-ge1TB){'{0:N2} TB'-f($Bytes/1TB)}elseif($Bytes-ge1GB){'{0:N2} GB'-f($Bytes/1GB)}elseif($Bytes-ge1MB){'{0:N2} MB'-f($Bytes/1MB)}else{'{0:N0} bytes'-f$Bytes} }
function ConvertTo-HtmlEncoded { param([AllowNull()]$Value) [Net.WebUtility]::HtmlEncode([string]$Value) }
function New-SvgBarChart {
    param(
        [string]$Title,
        $Items,
        [string]$LabelProperty,
        [string]$ValueProperty,
        [string]$ValueSuffix = ''
    )

    $data = @($Items | Where-Object { $null -ne $_ })
    $encodedTitle = ConvertTo-HtmlEncoded -Value $Title

    if ($data.Count -eq 0) {
        return "<section class='card chartcard'><h2>$encodedTitle</h2><div class='chartscroll'><p>No data available.</p></div></section>"
    }

    $maximumResult = $data | Measure-Object -Property $ValueProperty -Maximum
    $maximum = [double]$maximumResult.Maximum
    if ($maximum -le 0) { $maximum = 1.0 }

    $rows = foreach ($item in $data) {
        $labelValue = Get-HcxPropertyValue -Object $item -Name $LabelProperty
        $metricValue = Get-HcxPropertyValue -Object $item -Name $ValueProperty
        $encodedLabel = ConvertTo-HtmlEncoded -Value $labelValue
        $numericValue = 0.0
        [void][double]::TryParse(
            [string]$metricValue,
            [Globalization.NumberStyles]::Any,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$numericValue
        )
        $width = [math]::Round(($numericValue / $maximum) * 100, 1)
        $displayValue = '{0:N2}{1}' -f $numericValue, $ValueSuffix
        "<div class='barrow'><div class='barlabel'>$encodedLabel</div><div class='bartrack'><div class='barfill' style='width:$width%'></div></div><div class='barvalue'>$displayValue</div></div>"
    }

    return "<section class='card chartcard'><h2>$encodedTitle</h2><div class='chartscroll'>$($rows -join [Environment]::NewLine)</div></section>"
}
function ConvertTo-HtmlTable {
    param(
        $Rows,
        [string[]]$Columns,
        [string]$Empty = 'No data available.'
    )

    $data = @($Rows | Where-Object { $null -ne $_ })
    if ($data.Count -eq 0) {
        $encodedEmpty = ConvertTo-HtmlEncoded -Value $Empty
        return "<p>$encodedEmpty</p>"
    }

    $head = @(
        foreach ($column in $Columns) {
            '<th>{0}</th>' -f (ConvertTo-HtmlEncoded -Value $column)
        }
    ) -join ''

    $body = @(
        foreach ($row in $data) {
            $cells = foreach ($column in $Columns) {
                $value = Get-HcxPropertyValue -Object $row -Name $column
                if ($value -is [datetime]) {
                    $value = $value.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
                }
                '<td>{0}</td>' -f (ConvertTo-HtmlEncoded -Value $value)
            }
            '<tr>{0}</tr>' -f ($cells -join '')
        }
    ) -join [Environment]::NewLine

    return "<div class='tablewrap'><table><thead><tr>$head</tr></thead><tbody>$body</tbody></table></div>"
}
function New-HcxWaveCompletionChart {
    param(
        [string]$WaveName,
        [int]$CompletedVms,
        [int]$EstimatedVms
    )
    if ([string]::IsNullOrWhiteSpace($WaveName) -or $EstimatedVms -le 0) { return '' }
    $rawPercent = [math]::Round(($CompletedVms / [double]$EstimatedVms) * 100, 1)
    $chartPercent = [math]::Min([math]::Max($rawPercent, 0), 100)
    $remaining = [math]::Max($EstimatedVms - $CompletedVms, 0)
    $circumference = 314.159
    $dash = [math]::Round(($chartPercent / 100) * $circumference, 2)
    $gap = [math]::Round($circumference - $dash, 2)
    $encodedWave = ConvertTo-HtmlEncoded -Value $WaveName
    $statusText = if ($CompletedVms -gt $EstimatedVms) { 'Completed count exceeds the current wave estimate.' } elseif ($CompletedVms -eq $EstimatedVms) { 'Wave estimate reached.' } else { "$remaining VM(s) remain against the current estimate." }
    return @"
<section class='card wave-progress'><h2>$encodedWave Completion</h2>
<div style='display:grid;grid-template-columns:260px 1fr;gap:28px;align-items:center'>
<div style='text-align:center'><svg width='230' height='230' viewBox='0 0 140 140' role='img' aria-label='$encodedWave completion $rawPercent percent'>
<circle cx='70' cy='70' r='50' fill='none' stroke='#20353f' stroke-width='18'/>
<circle cx='70' cy='70' r='50' fill='none' stroke='#65d48a' stroke-width='18' stroke-linecap='round' transform='rotate(-90 70 70)' stroke-dasharray='$dash $gap'/>
<text x='70' y='66' text-anchor='middle' fill='#e6e6e6' font-size='22' font-weight='700'>$rawPercent%</text>
<text x='70' y='86' text-anchor='middle' fill='#9fb7c3' font-size='9'>COMPLETE</text></svg></div>
<div><div class='metrics' style='margin-bottom:10px'><div class='metric'><div class='value'>$CompletedVms</div><div class='label'>Unique VMs Completed</div></div><div class='metric'><div class='value'>$EstimatedVms</div><div class='label'>Estimated Wave VMs</div></div><div class='metric'><div class='value'>$remaining</div><div class='label'>Estimated Remaining</div></div></div><p class='note'>$statusText Completion is calculated from unique VM names with a successful migration status in the selected reporting period.</p></div>
</div></section>
"@
}

function Get-HcxTransportAnalytics {
    try {
        $response = Invoke-HcxRest -Method 'GET' -Path "/hybridity/api/interconnect/underlay/serviceMeshHealth?vcGuid=$($script:Hcx.VcGuid)"
        $rows = [Collections.Generic.List[object]]::new()
        foreach ($mesh in @(Get-HcxResponseItems -Response $response)) {
            $serviceMeshId = [string](Get-HcxPropertyValue -Object $mesh -Name 'serviceMeshId')
            $pmtuByUplink = @{}
            if (-not [string]::IsNullOrWhiteSpace($serviceMeshId)) {
                try {
                    $pmtuResponse = Invoke-HcxRest -Method 'GET' -Path "/hybridity/api/interconnect/underlay/pmtu/serviceMesh/${serviceMeshId}?vcGuid=$($script:Hcx.VcGuid)"
                    foreach ($pmtuItem in @(Get-HcxResponseItems -Response $pmtuResponse)) {
                        $pmtuName = [string](Get-HcxPropertyValue -Object $pmtuItem -Name 'uplinkName')
                        if ($pmtuName) { $pmtuByUplink[$pmtuName] = $pmtuItem }
                    }
                }
                catch { Write-HcxDebug "Path MTU collection unavailable for service mesh $serviceMeshId. $($_.Exception.Message)" 'TRANSPORT-MTU' }
            }
            foreach ($uplink in @((Get-HcxPropertyValue -Object $mesh -Name 'uplinkMetrics'))) {
                $measured = Get-HcxPropertyValue -Object $uplink -Name 'measuredData'
                $bandwidth = Get-HcxPropertyValue -Object $measured -Name 'maxAvailableBandwidth'
                $value = Get-HcxPropertyValue -Object $bandwidth -Name 'value'
                $latency = Get-HcxPropertyValue -Object $measured -Name 'latency'
                $loss = Get-HcxPropertyValue -Object $measured -Name 'lossRate'
                $uplinkName = [string](Get-HcxPropertyValue -Object $uplink -Name 'uplinkName')
                $remoteUplinkName = [string](Get-HcxPropertyValue -Object $uplink -Name 'remoteUplinkName')
                $pmtu = if ($pmtuByUplink.ContainsKey($uplinkName)) { $pmtuByUplink[$uplinkName] } elseif ($pmtuByUplink.ContainsKey($remoteUplinkName)) { $pmtuByUplink[$remoteUplinkName] } else { $null }
                $pmtuMetrics = @((Get-HcxPropertyValue -Object $pmtu -Name 'pmtuMetrics'))
                $serviceHealth = @((Get-HcxPropertyValue -Object $bandwidth -Name 'services') | ForEach-Object {
                    '{0}: {1}' -f (Get-HcxPropertyValue -Object $_ -Name 'serviceType'), (Get-HcxPropertyValue -Object $_ -Name 'status')
                }) -join '; '
                $rows.Add([pscustomobject][ordered]@{
                    UplinkName = $uplinkName
                    RemoteUplinkName = $remoteUplinkName
                    UploadMbps = [math]::Round(([double](Get-HcxPropertyValue -Object $value -Name 'uploadKbps')) / 1000, 2)
                    DownloadMbps = [math]::Round(([double](Get-HcxPropertyValue -Object $value -Name 'downloadKbps')) / 1000, 2)
                    LatencyMs = [double](Get-HcxPropertyValue -Object $latency -Name 'value')
                    LossPercent = [double](Get-HcxPropertyValue -Object $loss -Name 'value')
                    DiscoveredMTU = Get-HcxPropertyValue -Object $pmtu -Name 'discoveredMtu'
                    CurrentMTU = Get-HcxPropertyValue -Object $pmtu -Name 'currentMtu'
                    ConfiguredMTU = Get-HcxPropertyValue -Object $pmtu -Name 'configuredMtu'
                    RxPathMTU = @($pmtuMetrics | ForEach-Object { Get-HcxPropertyValue -Object $_ -Name 'rxPmtu' } | Where-Object { $null -ne $_ }) -join '; '
                    TxPathMTU = @($pmtuMetrics | ForEach-Object { Get-HcxPropertyValue -Object $_ -Name 'txPmtu' } | Where-Object { $null -ne $_ }) -join '; '
                    ServiceHealth = $serviceHealth
                })
            }
        }
        return @($rows)
    }
    catch {
        Write-HcxDebug "Transport analytics unavailable. $($_.Exception.Message)" 'TRANSPORT'
        return @()
    }
}
function New-HcxReportModel {
    param($Records, [datetime]$StartDate, [datetime]$EndDate)

    $allRecords = @($Records | Where-Object { $null -ne $_ })
    $migrated = @($allRecords | Where-Object {
        $statusValue = Get-HcxPropertyValue -Object $_ -Name 'Status'
        Test-MigratedStatus -Status ([string]$statusValue)
    })
    $errors = @($allRecords | Where-Object {
        $statusValue = Get-HcxPropertyValue -Object $_ -Name 'Status'
        $errorValue = Get-HcxPropertyValue -Object $_ -Name 'ErrorMessage'
        Test-ErrorStatus -Status ([string]$statusValue) -Error ([string]$errorValue)
    })

    $laterMoved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($errorRecord in $errors) {
        $errorVm = [string](Get-HcxPropertyValue -Object $errorRecord -Name 'VMName')
        $errorEnd = Get-HcxPropertyValue -Object $errorRecord -Name 'EndTime'
        if ([string]::IsNullOrWhiteSpace($errorVm)) { continue }
        $laterSuccess = @($migrated | Where-Object {
            $candidateVm = [string](Get-HcxPropertyValue -Object $_ -Name 'VMName')
            $candidateEnd = Get-HcxPropertyValue -Object $_ -Name 'EndTime'
            ($candidateVm -ieq $errorVm) -and (($null -eq $errorEnd) -or ($candidateEnd -gt $errorEnd))
        })
        if ($laterSuccess.Count -gt 0) { [void]$laterMoved.Add($errorVm) }
    }

    $daily = @($migrated | Where-Object { $null -ne $_.EndTime } |
        Group-Object { $_.EndTime.ToString('yyyy-MM-dd') } |
        ForEach-Object { [pscustomobject]@{ Date = $_.Name; Migrated = $_.Count } } |
        Sort-Object Date)
    $dailyOs = @($migrated | Where-Object { $null -ne $_.EndTime } |
        Group-Object { $_.EndTime.ToString('yyyy-MM-dd') + '|' + $_.GuestOS } |
        ForEach-Object {
            $parts = $_.Name -split '\|', 2
            [pscustomobject]@{ Date = $parts[0]; GuestOS = $parts[1]; Migrated = $_.Count }
        } | Sort-Object Date, GuestOS)
    $maxRecord = @($migrated | Sort-Object DurationMinutes -Descending | Select-Object -First 1)

    $storageSum = if ($migrated.Count -gt 0) { [double](($migrated | Measure-Object StorageBytes -Sum).Sum) } else { 0.0 }
    $memorySum = if ($migrated.Count -gt 0) { [double](($migrated | Measure-Object MemoryBytes -Sum).Sum) } else { 0.0 }
    $cpuSum = if ($migrated.Count -gt 0) { [int](($migrated | Measure-Object vCPU -Sum).Sum) } else { 0 }
    $averageMinutes = if ($migrated.Count -gt 0) { [math]::Round([double](($migrated | Measure-Object DurationMinutes -Average).Average), 2) } else { 0.0 }

    [pscustomobject]@{
        StartDate = $StartDate
        EndDate = $EndDate
        All = $allRecords
        Migrated = $migrated
        Errors = $errors
        ErrorLaterMoved = $laterMoved
        TotalMigrated = $migrated.Count
        TotalRecords = $allRecords.Count
        TotalStorageBytes = $storageSum
        TotalMemoryBytes = $memorySum
        TotalVcpu = $cpuSum
        AverageMinutes = $averageMinutes
        MaxMinutes = if ($maxRecord.Count -gt 0) { $maxRecord[0].DurationMinutes } else { 0 }
        MaxVm = if ($maxRecord.Count -gt 0) { $maxRecord[0].VMName } else { '' }
        Daily = $daily
        DailyOS = $dailyOs
        TopTime = @($migrated | Sort-Object DurationMinutes -Descending | Select-Object -First 5)
        TopStorage = @($migrated | Sort-Object StorageBytes -Descending | Select-Object -First 5)
        TopCompute = @($migrated | Sort-Object ComputeScore -Descending | Select-Object -First 5)
    }
}
function Export-HcxReports {
    param($Records,[datetime]$StartDate,[datetime]$EndDate,[string]$WaveName,[int]$WaveEstimatedVms)
    $model = New-HcxReportModel -Records $Records -StartDate $StartDate -EndDate $EndDate
    $waveNameClean = if ([string]::IsNullOrWhiteSpace($WaveName)) { 'Unspecified Wave' } else { $WaveName.Trim() }
    $uniqueCompletedVms = @($model.Migrated | Where-Object { -not [string]::IsNullOrWhiteSpace($_.VMName) } | Select-Object -ExpandProperty VMName -Unique)
    $waveCompletedVms = $uniqueCompletedVms.Count
    $wavePercentComplete = if ($WaveEstimatedVms -gt 0) { [math]::Round(($waveCompletedVms / [double]$WaveEstimatedVms) * 100, 1) } else { $null }
    $waveRemainingVms = if ($WaveEstimatedVms -gt 0) { [math]::Max($WaveEstimatedVms - $waveCompletedVms, 0) } else { $null }
    $waveCompletionChart = New-HcxWaveCompletionChart -WaveName $waveNameClean -CompletedVms $waveCompletedVms -EstimatedVms $WaveEstimatedVms
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss', [Globalization.CultureInfo]::InvariantCulture)
    $base = "HCX91-Mobility-Executive-$stamp"
    $rawCsv = Join-Path $script:Dirs.Exports "$base-Raw.csv"
    $summaryCsv = Join-Path $script:Dirs.Exports "$base-Summary.csv"
    $dailyCsv = Join-Path $script:Dirs.Exports "$base-Daily.csv"
    $dailyGuestOsCsv = Join-Path $script:Dirs.Exports "$base-Daily-GuestOS.csv"
    $groupSummaryCsv = Join-Path $script:Dirs.Exports "$base-Mobility-Groups.csv"


    $groupColumns = @('GroupName','GroupId','State','ConfigurationStatus','EffectiveTimestamp','CreatedTimestamp','MigrationStartTimestamp','LastUpdatedTimestamp','GroupElapsedMinutes','GroupElapsedDisplay','SourceSite','SourceVCenter','DestinationSite','DestinationVCenter','TotalVMs','DraftVMs','QueuedVMs','WarningVMs','CancelledVMs','ErrorVMs','InTransferVMs','InSwitchoverVMs','WaitingSwitchoverVMs','CompletedVMs','TotalVcpu','TotalMemoryGB','TotalMemoryDisplay','TotalDiskGB','TotalDiskDisplay','Username','ServiceMeshId')
    if ($script:GroupSummaries.Count -gt 0) {
        $script:GroupSummaries | Select-Object $groupColumns | Export-Csv -LiteralPath $groupSummaryCsv -NoTypeInformation -Encoding utf8BOM
    }
    else {
        ('"' + ($groupColumns -join '","') + '"') | Set-Content -LiteralPath $groupSummaryCsv -Encoding utf8BOM
    }

    $rawColumns = @('GroupName','GroupId','VMName','Status','MigrationType','StartTime','EndTime','DurationMinutes','StorageBytes','StorageGB','MemoryBytes','MemoryGB','vCPU','ComputeScore','DestinationNetwork','ErrorMessage','Attempt')
    if ($model.All.Count -gt 0) {
        $model.All | Select-Object $rawColumns | Export-Csv -LiteralPath $rawCsv -NoTypeInformation -Encoding utf8BOM
    }
    else {
        ('"' + ($rawColumns -join '","') + '"') | Set-Content -LiteralPath $rawCsv -Encoding utf8BOM
    }

    [pscustomobject]@{
        WaveName = $waveNameClean
        WaveEstimatedVMs = $WaveEstimatedVms
        WaveCompletedUniqueVMs = $waveCompletedVms
        WaveRemainingEstimatedVMs = $waveRemainingVms
        WavePercentComplete = $wavePercentComplete
        StartDate = $StartDate
        EndDate = $EndDate
        Records = $model.TotalRecords
        Migrated = $model.TotalMigrated
        Errors = $model.Errors.Count
        ErrorsLaterMoved = $model.ErrorLaterMoved.Count
        TotalStorageGB = [math]::Round($model.TotalStorageBytes / 1GB, 2)
        TotalMemoryGB = [math]::Round($model.TotalMemoryBytes / 1GB, 2)
        TotalVcpu = $model.TotalVcpu
        AverageMinutes = $model.AverageMinutes
        MaximumMinutes = $model.MaxMinutes
        MaximumDurationVM = $model.MaxVm
    } | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding utf8BOM

    if ($model.Daily.Count -gt 0) {
        $model.Daily | Export-Csv -LiteralPath $dailyCsv -NoTypeInformation -Encoding utf8BOM
    }
    else {
        '"Date","Migrated"' | Set-Content -LiteralPath $dailyCsv -Encoding utf8BOM
    }

    if ($model.DailyOS.Count -gt 0) {
        $model.DailyOS | Export-Csv -LiteralPath $dailyGuestOsCsv -NoTypeInformation -Encoding utf8BOM
    }
    else {
        '"Date","GuestOS","Migrated"' | Set-Content -LiteralPath $dailyGuestOsCsv -Encoding utf8BOM
    }

    $script:LastExcel=$null
    $script:LastExcel = $null
    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            Import-Module -Name ImportExcel -ErrorAction Stop
            $xlsx = Join-Path -Path $script:Dirs.Exports -ChildPath "$base.xlsx"

            $model.All |
                Select-Object * -ExcludeProperty Raw |
                Export-Excel -Path $xlsx -WorksheetName 'All Migrations' -AutoSize -FreezeTopRow -TableName 'AllMigrations'

            $model.Migrated |
                Select-Object * -ExcludeProperty Raw |
                Export-Excel -Path $xlsx -WorksheetName 'Migrated' -AutoSize -FreezeTopRow -TableName 'Migrated'

            $model.Errors |
                Select-Object * -ExcludeProperty Raw |
                Export-Excel -Path $xlsx -WorksheetName 'Errors' -AutoSize -FreezeTopRow -TableName 'Errors'

            $model.Daily |
                Export-Excel -Path $xlsx -WorksheetName 'Daily' -AutoSize -TableName 'Daily'

            $model.DailyOS |
                Export-Excel -Path $xlsx -WorksheetName 'Daily Guest OS' -AutoSize -TableName 'DailyGuestOS'

            $script:GroupSummaries |
                Select-Object $groupColumns |
                Export-Excel -Path $xlsx -WorksheetName 'Mobility Groups' -AutoSize -FreezeTopRow -TableName 'MobilityGroups'

            $script:LastExcel = $xlsx
            Log -Message "Excel workbook created: $xlsx" -Level PASS
        }
        catch {
            Log -Message ("Excel export failed; CSV exports remain available. " + $_.Exception.Message) -Level WARN
        }
    }
    else {
        Log -Message 'ImportExcel is not installed. CSV files were created; Excel export was skipped.' -Level WARN
    }

    $errorsView=foreach ($e in $model.Errors){[pscustomobject]@{VMName=$e.VMName;GroupName=$e.GroupName;Status=$e.Status;ErrorMessage=$e.ErrorMessage;LaterMigrated=$model.ErrorLaterMoved.Contains($e.VMName);EndTime=$e.EndTime}}
    $topTimeChart = New-SvgBarChart `
        -Title 'Top Five VMs by Migration Time' `
        -Items $model.TopTime `
        -LabelProperty 'VMName' `
        -ValueProperty 'DurationMinutes' `
        -ValueSuffix ' min'

    $topStorageData = @(
        $model.TopStorage | ForEach-Object {
            [pscustomobject]@{
                VMName = $_.VMName
                StorageGB = $_.StorageGB
            }
        }
    )

    $topStorageChart = New-SvgBarChart `
        -Title 'Top Five VMs by Storage Transferred' `
        -Items $topStorageData `
        -LabelProperty 'VMName' `
        -ValueProperty 'StorageGB' `
        -ValueSuffix ' GB'

    $topComputeChart = New-SvgBarChart `
        -Title 'Top Five VMs by Compute Transfer Score' `
        -Items $model.TopCompute `
        -LabelProperty 'VMName' `
        -ValueProperty 'ComputeScore' `
        -ValueSuffix ''

    $dailyChart = New-SvgBarChart `
        -Title 'Daily Migrated VM Count' `
        -Items $model.Daily `
        -LabelProperty 'Date' `
        -ValueProperty 'Migrated' `
        -ValueSuffix ''

    $migratedVcpuTotal = [int](($model.Migrated | Measure-Object -Property vCPU -Sum).Sum)
    $migratedMemoryBytesTotal = [double](($model.Migrated | Measure-Object -Property MemoryBytes -Sum).Sum)
    $migratedMemoryTbTotal = [math]::Round($migratedMemoryBytesTotal / 1TB, 2)
    $migratedDiskBytesTotal = [double](($model.Migrated | Measure-Object -Property StorageBytes -Sum).Sum)
    $migratedDiskTbTotal = [math]::Round($migratedDiskBytesTotal / 1TB, 2)
    $groupTable = ConvertTo-HtmlTable -Rows $script:GroupSummaries -Columns @('GroupName','State','MigrationStartTimestamp','LastUpdatedTimestamp','GroupElapsedDisplay','SourceVCenter','DestinationVCenter','TotalVMs','CompletedVMs','ErrorVMs','CancelledVMs','TotalVcpu','TotalMemoryDisplay','TotalDiskDisplay')
    $groupCount = @($script:GroupSummaries).Count
    $groupVmTotal = [int](($script:GroupSummaries | Measure-Object -Property TotalVMs -Sum).Sum)
    $groupCompletedTotal = [int](($script:GroupSummaries | Measure-Object -Property CompletedVMs -Sum).Sum)
    $groupErrorTotal = [int](($script:GroupSummaries | Measure-Object -Property ErrorVMs -Sum).Sum)
    $groupCancelledTotal = [int](($script:GroupSummaries | Measure-Object -Property CancelledVMs -Sum).Sum)
    $groupVcpuTotal = [int](($script:GroupSummaries | Measure-Object -Property TotalVcpu -Sum).Sum)
    $groupMemoryBytesKnown = [double](($script:GroupSummaries | Measure-Object -Property TotalMemoryBytes -Sum).Sum)
    $groupDiskBytesKnown = [double](($script:GroupSummaries | Measure-Object -Property TotalDiskBytes -Sum).Sum)
    $durationRows = @($script:GroupSummaries | Where-Object { $null -ne $_.GroupElapsedMinutes })
    $groupAverageElapsed = if ($durationRows.Count -gt 0) { [math]::Round([double](($durationRows | Measure-Object -Property GroupElapsedMinutes -Average).Average), 2) } else { $null }
    $groupMaximumElapsed = if ($durationRows.Count -gt 0) { [math]::Round([double](($durationRows | Measure-Object -Property GroupElapsedMinutes -Maximum).Maximum), 2) } else { $null }
    $groupAverageElapsedDisplay = if ($null -ne $groupAverageElapsed) { '{0:N2} min' -f $groupAverageElapsed } else { 'Not returned' }
    $groupMaximumElapsedDisplay = if ($null -ne $groupMaximumElapsed) { '{0:N2} min' -f $groupMaximumElapsed } else { 'Not returned' }

    $migratedVmRows = foreach ($migratedRecord in @($model.Migrated | Sort-Object EndTime,VMName)) {
        $effectiveNetwork = [string]$migratedRecord.DestinationNetwork
        if ([string]::IsNullOrWhiteSpace($effectiveNetwork)) {
            $networkMatch = @($model.All | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.DestinationNetwork) -and
                ((-not [string]::IsNullOrWhiteSpace($migratedRecord.VMEntityId) -and $_.VMEntityId -eq $migratedRecord.VMEntityId) -or $_.VMName -eq $migratedRecord.VMName)
            } | Sort-Object EndTime -Descending | Select-Object -First 1)
            if ($networkMatch.Count -gt 0) { $effectiveNetwork = [string]$networkMatch[0].DestinationNetwork }
        }
        [pscustomobject][ordered]@{
            VMName=$migratedRecord.VMName; VMEntityId=$migratedRecord.VMEntityId; GroupName=$migratedRecord.GroupName
            MigrationType=$migratedRecord.MigrationType; StartTime=$migratedRecord.StartTime; EndTime=$migratedRecord.EndTime
            DurationMinutes=$migratedRecord.DurationMinutes; ProvisionedGB=$migratedRecord.ProvisionedGB; MemoryGB=$migratedRecord.MemoryGB
            vCPU=$migratedRecord.vCPU; AverageThroughputMBps=$migratedRecord.AverageThroughputMBps
            ThroughputBasis=$migratedRecord.ThroughputBasis; DestinationNetwork=$effectiveNetwork
        }
    }
    $migratedVmColumns = @('VMName','VMEntityId','GroupName','MigrationType','StartTime','EndTime','DurationMinutes','ProvisionedGB','MemoryGB','vCPU','AverageThroughputMBps','ThroughputBasis')
    if (@($migratedVmRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DestinationNetwork) }).Count -gt 0) { $migratedVmColumns += 'DestinationNetwork' }
    $migratedVmTable = ConvertTo-HtmlTable -Rows $migratedVmRows -Columns $migratedVmColumns
    $allVmCount = @($model.All).Count
    $completedVmCount = @($model.Migrated).Count
    $failedVmCount = @($model.Errors).Count
    $maxVm = @($model.Migrated | Sort-Object DurationMinutes -Descending | Select-Object -First 1)
    $maxVmText = if ($maxVm.Count) { "$($maxVm[0].VMName) at $($maxVm[0].DurationMinutes) minutes" } else { 'Not available' }

    $transportRows = @(Get-HcxTransportAnalytics)
    $transportSection = ''
    if ($transportRows.Count -gt 0) {
        $transportTable = ConvertTo-HtmlTable -Rows $transportRows -Columns @('UplinkName','RemoteUplinkName','UploadMbps','DownloadMbps','LatencyMs','LossPercent','DiscoveredMTU','CurrentMTU','ConfiguredMTU','RxPathMTU','TxPathMTU','ServiceHealth')
        $transportChart = New-SvgBarChart -Title 'Available Upload Bandwidth by HCX Uplink' -Items $transportRows -LabelProperty 'UplinkName' -ValueProperty 'UploadMbps' -ValueSuffix ' Mbps'
        $transportPeakUpload = [math]::Round([double](($transportRows | Measure-Object -Property UploadMbps -Maximum).Maximum), 2)
        $transportPeakDownload = [math]::Round([double](($transportRows | Measure-Object -Property DownloadMbps -Maximum).Maximum), 2)
        $transportMaxLatency = [math]::Round([double](($transportRows | Measure-Object -Property LatencyMs -Maximum).Maximum), 2)
        $transportMaxLoss = [math]::Round([double](($transportRows | Measure-Object -Property LossPercent -Maximum).Maximum), 3)
        $transportMtuValues = @($transportRows | ForEach-Object { $_.CurrentMTU } | Where-Object { $null -ne $_ -and [string]$_ -ne '' } | Select-Object -Unique)
        $transportMtuSummary = if ($transportMtuValues.Count -gt 0) { ' | Current Path MTU: <strong>' + ($transportMtuValues -join ', ') + '</strong>' } else { '' }
        $transportSummary = "<p>Peak available upload: <strong>$transportPeakUpload Mbps</strong> | Peak available download: <strong>$transportPeakDownload Mbps</strong> | Maximum observed latency: <strong>$transportMaxLatency ms</strong> | Maximum observed packet loss: <strong>$transportMaxLoss%</strong>$transportMtuSummary</p>"
        $transportServiceItems = @($transportRows | ForEach-Object { $_.ServiceHealth -split ';' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
        $transportServiceList = if ($transportServiceItems.Count -gt 0) { '<ul>' + (@($transportServiceItems | ForEach-Object { '<li>' + (ConvertTo-HtmlEncoded $_) + '</li>' }) -join '') + '</ul>' } else { '<p>No service health status was returned.</p>' }
        $transportMetricCards = @($transportRows | ForEach-Object {
            $uplink = ConvertTo-HtmlEncoded $_.UplinkName
            "<div class='transport-metric'><div class='transport-uplink'>$uplink</div><div class='transport-value'>$($_.UploadMbps) Mbps</div><div class='transport-label'>Available Upload</div><div class='transport-value'>$($_.DownloadMbps) Mbps</div><div class='transport-label'>Available Download</div><div class='transport-inline'><span>Latency <strong>$($_.LatencyMs) ms</strong></span><span>Loss <strong>$($_.LossPercent)%</strong></span><span>Path MTU <strong>$($_.CurrentMTU)</strong></span></div></div>"
        }) -join ''
        $transportSection = "<section class='card transport-card'><h2>Transport Analytics</h2>$transportSummary<div class='transport-layout'><div class='transport-metrics'>$transportMetricCards</div><div class='transport-health'><h3>Service Health</h3>$transportServiceList</div></div></section>"
    }
    $networkRows = @($model.All | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DestinationNetwork) })
    $networkSection = ''
    if ($networkRows.Count -gt 0) {
        $networkTable = ConvertTo-HtmlTable -Rows $networkRows -Columns @('VMName','GroupName','DestinationNetwork','DestinationNetworkId','DestinationNetworkType','NetworkMappingDetails','DestinationPlacement')
        $networkSection = "<section class='card'><h2>VM Network Mappings and Placement</h2>$networkTable</section>"
    }
    $htmlPath=Join-Path $script:Dirs.Reports "$base.html";$generated=Get-Date
    $html=@"
<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>$waveNameClean - HCX 9.1 Mobility Executive Summary</title><style>
:root{--bg:#071015;--panel:#0d1b22;--panel2:#132832;--line:#42606e;--text:#e6e6e6;--muted:#9fb7c3;--aqua:#76c7d8;--green:#65d48a;--red:#ff756c;--gold:#f1c75b}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:14px 'Segoe UI',Arial,sans-serif}.hero{padding:30px 38px;background:linear-gradient(135deg,#0b2530,#173e4b)}h1{margin:0 0 8px;font-size:30px}h2{color:var(--aqua);font-size:20px;margin-top:0}.subtitle{color:#c8d9e0}.wrap{padding:22px 32px}.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:18px}.metric,.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:16px}.metric .value{font-size:25px;font-weight:700;color:var(--green)}.metric .label{color:var(--muted);margin-top:5px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px}.card{margin-bottom:14px}.chartcard{min-height:255px;max-height:360px;overflow-y:auto;overflow-x:hidden;scrollbar-gutter:stable}.chartscroll{width:100%;min-width:0}.barrow{display:grid;grid-template-columns:minmax(80px,120px) minmax(80px,240px) 76px;gap:8px;align-items:center;margin:9px 0}.barlabel{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bartrack{height:18px;background:#20353f;border-radius:3px;overflow:hidden}.barfill{height:100%;background:linear-gradient(90deg,#2786a5,#76c7d8)}.barvalue{text-align:right;color:var(--muted)}.tablewrap{height:300px;max-height:300px;overflow-x:auto;overflow-y:scroll;max-width:100%;scrollbar-gutter:stable both-edges;border:1px solid #294550;border-radius:5px;overscroll-behavior:contain}table{width:max-content;min-width:100%;border-collapse:collapse;font-size:12px;table-layout:auto}th,td{min-width:max-content;white-space:nowrap;overflow-wrap:normal}th{position:sticky;top:0;z-index:2;background:#2b3740;color:#fff;text-align:left}th,td{border:1px solid #3c5561;padding:7px;vertical-align:top}tr:nth-child(even){background:#0a171d}.note{color:var(--muted)}.error{color:var(--red)}.transport-card>p{text-align:center}.transport-layout{display:grid;grid-template-columns:minmax(0,2fr) minmax(240px,1fr);gap:20px;align-items:stretch;max-width:1200px;margin:16px auto 0}.transport-metrics{display:flex;justify-content:center;align-items:center}.transport-metric{width:100%;max-width:720px;text-align:center;padding:18px;border:1px solid #294550;border-radius:8px;background:var(--panel2)}.transport-uplink{color:var(--aqua);font-weight:700;margin-bottom:12px}.transport-value{font-size:24px;font-weight:700;color:var(--green);margin-top:8px}.transport-label{color:var(--muted);margin-bottom:6px}.transport-inline{display:flex;justify-content:center;gap:24px;flex-wrap:wrap;margin-top:16px}.transport-health{padding:16px 20px;border:1px solid #294550;border-radius:8px;background:var(--panel2)}.transport-health h3{margin:0 0 10px;color:var(--aqua)}.transport-health ul{margin:0;padding-left:20px}.transport-health li{margin:7px 0}@media(max-width:800px){.transport-layout{grid-template-columns:1fr}}footer{padding:20px 32px;color:var(--muted);border-top:1px solid var(--line)}</style></head><body>
<header class='hero'><h1>$waveNameClean - HCX 9.1 Mobility Executive Summary</h1><div class='subtitle'>Wave estimate: $WaveEstimatedVms VM(s) | Reporting period: $($StartDate.ToString('yyyy-MM-dd')) through $($EndDate.ToString('yyyy-MM-dd HH:mm:ss')) | HCX Manager: $(ConvertTo-HtmlEncoded $script:Hcx.BaseUri)</div></header><main class='wrap'>
<section class='metrics'><div class='metric'><div class='value'>$groupCount</div><div class='label'>Mobility Groups</div></div><div class='metric'><div class='value'>$groupVmTotal</div><div class='label'>Total Scheduled VMs</div></div><div class='metric'><div class='value'>$($model.TotalMigrated)</div><div class='label'>Successfully Migrated VMs</div></div><div class='metric'><div class='value'>$groupErrorTotal</div><div class='label'>Error VMs</div></div><div class='metric'><div class='value'>$groupCancelledTotal</div><div class='label'>Cancelled VMs</div></div><div class='metric'><div class='value'>$migratedVcpuTotal</div><div class='label'>Migrated vCPU</div></div><div class='metric'><div class='value'>$migratedMemoryTbTotal TB</div><div class='label'>Migrated Memory</div></div><div class='metric'><div class='value'>$migratedDiskTbTotal TB</div><div class='label'>Migrated Disk</div></div><div class='metric'><div class='value'>$($model.AverageMinutes) min</div><div class='label'>Average VM Migration Time</div></div><div class='metric'><div class='value'>$($model.MaxMinutes) min</div><div class='label'>Maximum VM Migration Time</div></div></section>
$waveCompletionChart
<section class='card'><h2>Executive Summary</h2><p>For <strong>$waveNameClean</strong>, the current estimate is <strong>$WaveEstimatedVms VM(s)</strong>, with <strong>$waveCompletedVms unique VM(s) completed</strong> and an estimated completion of <strong>$(if ($null -ne $wavePercentComplete) { "$wavePercentComplete%" } else { 'Not calculated' })</strong>. During the selected reporting period, HCX returned <strong>$allVmCount VM migration records</strong> across <strong>$groupCount mobility groups</strong>. Of those records, <strong>$completedVmCount completed successfully</strong>, <strong>$failedVmCount were in an error or cancelled state</strong>, and the group summaries represented <strong>$groupVmTotal total VMs</strong>.</p><p>The largest known VM migration duration was <strong>$maxVmText</strong>. All successful migration records in the selected reporting period represented <strong>$migratedMemoryTbTotal TB of memory</strong>, <strong>$migratedDiskTbTotal TB of provisioned disk</strong>, and <strong>$migratedVcpuTotal vCPU</strong>. The maximum individual VM migration time was <strong>$($model.MaxMinutes) minutes</strong>, with an average individual VM migration time of <strong>$($model.AverageMinutes) minutes</strong>.</p><p class='note'>Estimated transfer rates use provisioned storage divided by elapsed migration time when measured transfer bytes are unavailable.</p></section>
<div class='grid'>$dailyChart$topTimeChart$topStorageChart$topComputeChart</div>
<section class='card'><h2>Compute Transfer Score</h2><p>The Compute Transfer Score is a comparative workload-effort index calculated for each successfully migrated VM as <strong>vCPU multiplied by migration duration in minutes</strong>. The calculation uses a minimum duration of one minute: <strong>Compute Transfer Score = vCPU x max(Migration Minutes, 1)</strong>. For example, a VM with 4 vCPU that migrated in 30 minutes receives a score of 120.</p><p class='note'>A higher score indicates that a migration combined more virtual CPU capacity with a longer migration duration. The score is useful for comparing relative migration effort and identifying workloads that consumed more compute-time during migration. The score is not a network-throughput measurement, CPU-utilization percentage, or HCX health rating.</p></section>
$transportSection
<section class='card'><h2>Mobility Group Summary</h2>$groupTable</section>
<section class='card'><h2>VMs Migrated Successfully by Name</h2>$migratedVmTable</section>
$networkSection
<section class='card'><h2>Migration Errors and Recovery Status</h2>$(ConvertTo-HtmlTable $errorsView @('VMName','GroupName','Status','ErrorMessage','LaterMigrated','EndTime'))</section>
<section class='card'><h2>All Time-Filtered VM Migration Records</h2>$(ConvertTo-HtmlTable ($model.All|Select-Object * -ExcludeProperty Raw) @('VMName','VMEntityId','GroupName','Status','MigrationType','StartTime','EndTime','DurationMinutes','ProvisionedGB','MemoryGB','vCPU','AverageThroughputMBps','ThroughputBasis','DestinationNetwork','ProgressMessage','WarningMessages','ErrorMessage'))</section>
</main><footer>Generated by Rev $script:AppVersion at $($generated.ToString('yyyy-MM-dd HH:mm:ss')) | Run folder: $(ConvertTo-HtmlEncoded $script:RunDir)</footer></body></html>
"@
    $html = $html.Replace('\<','<').Replace('\>','>')
    [IO.File]::WriteAllText($htmlPath, $html, [Text.UTF8Encoding]::new($false))
    $script:LastHtml = $htmlPath

    $requiredArtifacts = @($htmlPath, $groupSummaryCsv, $rawCsv, $summaryCsv, $dailyCsv, $dailyGuestOsCsv)
    $artifactStatus = foreach ($artifactPath in $requiredArtifacts) {
        $exists = Test-Path -LiteralPath $artifactPath -PathType Leaf
        $length = if ($exists) { (Get-Item -LiteralPath $artifactPath).Length } else { 0 }
        [pscustomobject]@{ Path = $artifactPath; Exists = $exists; Length = $length }
    }
    $invalidArtifacts = @($artifactStatus | Where-Object { -not $_.Exists -or $_.Length -le 0 })
    if ($invalidArtifacts.Count -gt 0) {
        throw ('Required report artifacts were not created: ' + (($invalidArtifacts.Path) -join ' | '))
    }

    $manifestPath = Join-Path $script:Dirs.Reports "$base-Artifact-Manifest.csv"
    $manifestRows = foreach ($artifact in $artifactStatus) {
        $item = Get-Item -LiteralPath $artifact.Path
        [pscustomobject]@{
            Name = $item.Name
            FullName = $item.FullName
            Length = $item.Length
            SHA256 = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash
            LastWriteTime = $item.LastWriteTime
        }
    }
    $manifestRows | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding utf8BOM
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf) -or (Get-Item -LiteralPath $manifestPath).Length -le 0) {
        throw "Artifact manifest was not created: $manifestPath"
    }

    Log -Message ("Executive HTML report created: $htmlPath; bytes=$((Get-Item -LiteralPath $htmlPath).Length)") -Level PASS
    [pscustomobject]@{
        WaveName = $waveNameClean
        WaveEstimatedVMs = $WaveEstimatedVms
        WaveCompletedUniqueVMs = $waveCompletedVms
        WavePercentComplete = $wavePercentComplete
        Html = $htmlPath
        GroupSummaryCsv = $groupSummaryCsv
        RawCsv = $rawCsv
        SummaryCsv = $summaryCsv
        DailyCsv = $dailyCsv
        DailyGuestOsCsv = $dailyGuestOsCsv
        Excel = $script:LastExcel
        Manifest = $manifestPath
        Model = $model
    }
}

$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="HCX 9.1 Mobility Analytics and Executive Reporting" Height="860" Width="1380" MinHeight="720" MinWidth="1100" WindowStartupLocation="CenterScreen" Background="#071015" Foreground="#E6E6E6" FontFamily="Segoe UI">
<Window.Resources><Style TargetType="Button"><Setter Property="Background" Value="#2B3740"/><Setter Property="Foreground" Value="#F1F4F6"/><Setter Property="BorderBrush" Value="#5F7482"/><Setter Property="Padding" Value="10,5"/><Setter Property="Margin" Value="4"/><Setter Property="MinHeight" Value="30"/></Style><Style TargetType="TextBlock"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="Margin" Value="4"/></Style><Style TargetType="TextBox"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="BorderBrush" Value="#607D8B"/><Setter Property="Padding" Value="5"/></Style><Style TargetType="PasswordBox"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="BorderBrush" Value="#607D8B"/><Setter Property="Padding" Value="5"/></Style><Style TargetType="GroupBox"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="Background" Value="#0D1B22"/><Setter Property="BorderBrush" Value="#2B3740"/><Setter Property="Margin" Value="5"/><Setter Property="Padding" Value="8"/></Style><Style TargetType="DataGrid"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="RowBackground" Value="#071015"/><Setter Property="AlternatingRowBackground" Value="#0D1B22"/><Setter Property="BorderBrush" Value="#607D8B"/></Style><Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#2B3740"/><Setter Property="Foreground" Value="#F1F4F6"/><Setter Property="FontWeight" Value="SemiBold"/></Style></Window.Resources>
<Grid Margin="10"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="170"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<GroupBox Header="HCX 9.1 Connection"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="115"/><ColumnDefinition Width="2*"/><ColumnDefinition Width="95"/><ColumnDefinition Width="1.4*"/><ColumnDefinition Width="90"/><ColumnDefinition Width="1.4*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="HCX Manager" VerticalAlignment="Center"/><TextBox x:Name="txtHcx" Grid.Column="1"/><TextBlock Grid.Column="2" Text="Username" VerticalAlignment="Center"/><TextBox x:Name="txtUser" Grid.Column="3"/><TextBlock Grid.Column="4" Text="Password" VerticalAlignment="Center"/><PasswordBox x:Name="txtPassword" Grid.Column="5"/><StackPanel Grid.Column="6" Orientation="Horizontal"><Button x:Name="btnConnect" Content="Connect" Width="95"/><Button x:Name="btnDisconnect" Content="Disconnect" Width="95"/></StackPanel></Grid></GroupBox>
<GroupBox Grid.Row="1" Header="Reporting Scope and Output"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="100"/><ColumnDefinition Width="145"/><ColumnDefinition Width="85"/><ColumnDefinition Width="145"/><ColumnDefinition Width="95"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="34"/><RowDefinition Height="34"/></Grid.RowDefinitions><TextBlock Text="Start Date" VerticalAlignment="Center"/><DatePicker x:Name="dpStart" Grid.Column="1"/><TextBlock Grid.Column="2" Text="End Date" VerticalAlignment="Center"/><DatePicker x:Name="dpEnd" Grid.Column="3"/><TextBlock Grid.Column="4" Text="Output Base" VerticalAlignment="Center"/><TextBox x:Name="txtOutputPath" Grid.Column="5"/><Button x:Name="btnBrowse" Grid.Column="6" Content="Select Folder"/><CheckBox x:Name="chkDebug" Grid.Row="1" Grid.Column="1" Content="Debug enabled" IsChecked="True" Foreground="#E6E6E6" VerticalAlignment="Center"/><TextBlock Grid.Row="1" Grid.Column="2" Text="Endpoint" VerticalAlignment="Center"/><TextBlock x:Name="lblEndpoint" Grid.Row="1" Grid.Column="3" Grid.ColumnSpan="3" Text="Not discovered" Foreground="#76C7D8" VerticalAlignment="Center"/><Button x:Name="btnCollect" Grid.Row="1" Grid.Column="6" Content="Collect and Generate Reports" MinWidth="205" IsEnabled="False"/><TextBlock Grid.Row="2" Grid.Column="0" Text="Wave Name" VerticalAlignment="Center"/><TextBox x:Name="txtWaveName" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" ToolTip="Example: Wave 0"/><TextBlock Grid.Row="2" Grid.Column="3" Text="Estimated Wave VMs" VerticalAlignment="Center"/><TextBox x:Name="txtWaveEstimatedVms" Grid.Row="2" Grid.Column="4" Width="90" HorizontalAlignment="Left" ToolTip="Estimated total number of VMs planned for this wave"/><TextBlock Grid.Row="2" Grid.Column="5" Text="Used to calculate wave completion percentage" Foreground="#9FB7C3" VerticalAlignment="Center"/></Grid></GroupBox>
<DataGrid x:Name="gridData" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"><DataGrid.Columns><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" Width="160"/><DataGridTextColumn Header="Group" Binding="{Binding GroupName}" Width="180"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="105"/><DataGridTextColumn Header="Type" Binding="{Binding MigrationType}" Width="120"/><DataGridTextColumn Header="Start" Binding="{Binding StartTime}" Width="145"/><DataGridTextColumn Header="End" Binding="{Binding EndTime}" Width="145"/><DataGridTextColumn Header="Minutes" Binding="{Binding DurationMinutes}" Width="75"/><DataGridTextColumn Header="Storage GB" Binding="{Binding StorageGB}" Width="85"/><DataGridTextColumn Header="Memory GB" Binding="{Binding MemoryGB}" Width="85"/><DataGridTextColumn Header="vCPU" Binding="{Binding vCPU}" Width="55"/><DataGridTextColumn Header="Destination Network" Binding="{Binding DestinationNetwork}" Width="180"/><DataGridTextColumn Header="Error" Binding="{Binding ErrorMessage}" Width="300"/></DataGrid.Columns></DataGrid>
<GroupBox Grid.Row="3" Header="Log"><TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Background="#071015" Foreground="#E6E6E6"/></GroupBox>
<Grid Grid.Row="4"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Orientation="Horizontal"><ProgressBar x:Name="pbProgress" Width="360" Height="18" Minimum="0" Maximum="100" Margin="6"/><TextBlock x:Name="lblStatus" Text="Ready" Foreground="#76C7D8" VerticalAlignment="Center"/></StackPanel><StackPanel Grid.Column="1" Orientation="Horizontal"><Button x:Name="btnOpenHtml" Content="Open HTML" IsEnabled="False"/><Button x:Name="btnOpenRun" Content="Open Run Folder"/><Button x:Name="btnClose" Content="Close"/></StackPanel></Grid>
</Grid></Window>
'@
$script:Window=[Windows.Markup.XamlReader]::Parse($xaml)
foreach ($n in @('txtHcx','txtUser','txtPassword','btnConnect','btnDisconnect','dpStart','dpEnd','txtOutputPath','txtWaveName','txtWaveEstimatedVms','btnBrowse','chkDebug','lblEndpoint','btnCollect','gridData','txtLog','pbProgress','lblStatus','btnOpenHtml','btnOpenRun','btnClose')){Set-Variable -Scope Script -Name $n -Value $script:Window.FindName($n)}
$script:dpStart.SelectedDate=(Get-Date).Date.AddDays(-14);$script:dpEnd.SelectedDate=(Get-Date).Date;$script:txtOutputPath.Text=$script:OutputBase;$script:txtWaveName.Text='Wave 0';$script:txtWaveEstimatedVms.Text='0'
if([string]::IsNullOrWhiteSpace($script:txtUser.Text)){$script:txtUser.Text='administrator@vsphere.local'}
Start-HcxDiagnosticTranscript
Write-HcxDebug "Debug logging enabled. RunFolder=$script:RunDir; Log=$script:LogFile; Transcript=$script:TranscriptFile" 'STARTUP'
Log "HCX 9.1 Mobility Analytics started. All output is centralized under $script:RunDir" PASS
$script:btnConnect.Add_Click({try{$script:Window.Cursor='Wait';$script:btnConnect.IsEnabled=$false;Connect-Hcx91Rest $script:txtHcx.Text.Trim() $script:txtUser.Text.Trim() $script:txtPassword.Password;$script:btnCollect.IsEnabled=$true;$script:lblStatus.Text='Connected';$script:lblStatus.Foreground='LightGreen'}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'HCX connection failed','OK','Error')|Out-Null}finally{$script:txtPassword.Clear();$script:btnConnect.IsEnabled=$true;$script:Window.Cursor=$null}})
$script:btnDisconnect.Add_Click({Disconnect-HcxRest;$script:btnCollect.IsEnabled=$false;$script:lblEndpoint.Text='Not discovered';$script:lblStatus.Text='Disconnected'})
$script:chkDebug.Add_Checked({$script:DebugLoggingEnabled=$true;Log 'Debug logging enabled.' INFO});$script:chkDebug.Add_Unchecked({Log 'Debug logging disabled.' INFO;$script:DebugLoggingEnabled=$false})
$script:btnBrowse.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;$d.Description='Select base folder for the per-launch HCX reporting folder';$d.SelectedPath=$script:txtOutputPath.Text;if($d.ShowDialog()-eq[System.Windows.Forms.DialogResult]::OK){try{$old=$script:LogFile;Stop-HcxDiagnosticTranscript;Initialize-RunFolder $d.SelectedPath;$script:txtOutputPath.Text=$script:OutputBase;Start-HcxDiagnosticTranscript;if(Test-Path -LiteralPath $old){Copy-Item -LiteralPath $old -Destination (Join-Path$script:Dirs.Logs('PrePathChange-'+[IO.Path]::GetFileName($old))) -Force };Log"Output path applied. Active run folder: $script:RunDir" PASS}catch{[Windows.MessageBox]::Show($_.Exception.Message,'Output path failed','OK','Error')|Out-Null}}})
$script:btnCollect.Add_Click({try{$script:Window.Cursor='Wait';$script:btnCollect.IsEnabled=$false;$script:pbProgress.Value=0;$script:lblStatus.Text='Collecting HCX mobility data...';$start=[datetime]$script:dpStart.SelectedDate;$end=([datetime]$script:dpEnd.SelectedDate).Date.AddDays(1).AddTicks(-1);if($end -lt $start){throw 'End date must be on or after start date.'};$waveName=$script:txtWaveName.Text.Trim();if([string]::IsNullOrWhiteSpace($waveName)){throw 'Wave Name is required.'};$waveEstimatedVms=0;if(-not [int]::TryParse($script:txtWaveEstimatedVms.Text.Trim(),[ref]$waveEstimatedVms)-or$waveEstimatedVms-lt 1){throw 'Estimated Wave VMs must be a whole number greater than zero.'};$records=Get-HcxMobilityData $start $end;$script:gridData.ItemsSource=$records;$script:lblEndpoint.Text=$script:Hcx.SelectedEndpoint;$script:lblStatus.Text='Generating reports...';$result=Export-HcxReports $records $start $end $waveName $waveEstimatedVms;$script:pbProgress.Value=100;$script:btnOpenHtml.IsEnabled=$true;$script:lblStatus.Text="$($result.WaveName): $($result.WaveCompletedUniqueVMs)/$($result.WaveEstimatedVMs) complete ($($result.WavePercentComplete)%); $($result.Model.Errors.Count) error record(s)";$script:lblStatus.Foreground='LightGreen';if (-not (Test-Path -LiteralPath $result.Html -PathType Leaf)) { throw "Verified HTML report is missing: $($result.Html)" };[Windows.MessageBox]::Show("Reports created successfully.`n`nHTML: $($result.Html)`nRaw CSV: $($result.RawCsv)`nManifest: $($result.Manifest)`nRun folder: $script:RunDir",'HCX reporting complete')|Out-Null}catch{Log $_.Exception.Message ERROR;Save-HcxDebugArtifact 'EXCEPTION' 'COLLECT-REPORT'([ordered]@{Message=$_.Exception.Message;Stack=$_.ScriptStackTrace})|Out-Null;$script:lblStatus.Text='Failed';$script:lblStatus.Foreground='Tomato';[Windows.MessageBox]::Show($_.Exception.Message,'Collection failed','OK','Error')|Out-Null}finally{$script:btnCollect.IsEnabled=$script:Hcx.Connected;$script:Window.Cursor=$null}})
$script:btnOpenHtml.Add_Click({if($script:LastHtml -and (Test-Path -LiteralPath $script:LastHtml)){Invoke-Item $script:LastHtml}});$script:btnOpenRun.Add_Click({Invoke-Item $script:RunDir});$script:btnClose.Add_Click({Disconnect-HcxRest;Stop-HcxDiagnosticTranscript;$script:Window.Close()});$script:Window.Add_Closing({Stop-HcxDiagnosticTranscript})
$null=$script:Window.ShowDialog();Stop-HcxDiagnosticTranscript















