#Requires -Version 7.4
# SharePoint Storage Tree - scanner. Read-only: it never changes anything in SharePoint.
# Emits report schema v1.1 (see viewer/storagetree.html header for the contract).

$script:GeneratorVersion = '0.1.0'
$script:Activity = 'SharePoint Storage Tree'

function ConvertTo-SPSTInt64 {
    # SMTotalSize comes back as a FieldLookupValue (bytes in .LookupId); the stream sizes come back as strings.
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return $null }
    if ($Value -is [Microsoft.SharePoint.Client.FieldLookupValue]) { return [int64]$Value.LookupId }
    [int64]$Value
}

function ConvertTo-SPSTIso {
    param($Date)
    if ($null -eq $Date) { return $null }
    ([datetime]$Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Get-SPSTMajorVersion {
    # "16.0" -> 16; "0.3" (draft only) -> 1. Approximate where minor versions exist.
    param([string]$Label)
    if ($Label -match '^(\d+)\.(\d+)$') { return [math]::Max(1, [int]$Matches[1]) }
    1
}

function Get-SPSTWebScan {
    # Scans the web of the CURRENT PnP connection. $State = @{ calls; warnings } shared across webs.
    param(
        [hashtable]$State,
        [switch]$IsTop,            # top web: read site collection usage (covers subsites too)
        [string]$ParentUrl,
        [int]$VersionDetailMinVersions,
        [double]$VersionDetailMinHistoryMB,
        [int]$VersionBatchSize,
        [string]$Label = 'site'
    )
    $ctx = Get-PnPContext
    $warnings = $State.warnings
    Write-Progress -Activity $script:Activity -Status "$($Label): reading site"
    $web = Get-PnPWeb -Includes Title, Url, LastItemUserModifiedDate, CurrentUser; $State.calls++
    $partialReasons = @()
    # Site usage is only readable by site collection admins, so it doubles as the "full scan?" check.
    # (CurrentUser.IsSiteAdmin can't be used: with a read-only app registration it is always false.)
    $storageUsed = $null
    try { $usage = [int64](Get-PnPSite -Includes Usage).Usage.Storage; $State.calls++; if ($IsTop) { $storageUsed = $usage } }
    catch {
        $partialReasons += "scanned as $($web.CurrentUser.Email), who is not a site collection admin (site total not readable; libraries with their own permissions may be missing)"
        $warnings.Add("$($Label): site usage not readable ($($_.Exception.Message)). Scan as a site collection admin for a complete report.")
    }

    # ---- libraries (all document libraries, including hidden/system ones) ----------------------
    # Root folder path comes from DefaultViewUrl: asking for RootFolder on every list needs more than AllSites.Read.
    $lists = @(Get-PnPList -Includes DefaultViewUrl, BaseTemplate, Hidden, ItemCount | Where-Object { $_.BaseType -eq 'DocumentLibrary' }); $State.calls++
    $libRoot = {
        param($l)
        $u = [string]$l.DefaultViewUrl
        if ($u -match '^(.*?)/Forms/[^/]+$') { return $Matches[1] }
        if ($u -match '^(.*)/[^/]+\.aspx$') { return $Matches[1] }
        $u.TrimEnd('/')
    }
    $holdLib = $lists | Where-Object { (& $libRoot $_) -match '/PreservationHoldLibrary$' -or $_.Title -eq 'Preservation Hold Library' }

    $fields = 'FileLeafRef', 'FileRef', 'FileDirRef', 'FSObjType', 'File_x0020_Size', 'SMTotalSize', 'SMTotalFileStreamSize', '_UIVersionString', 'Modified', 'Created', 'Editor'
    $files = [System.Collections.Generic.List[object]]::new()
    $refs = [System.Collections.Generic.List[object]]::new()   # parallel to $files: server-relative url + history estimate
    $libsOut = [System.Collections.Generic.List[object]]::new()
    $noSM = 0; $overflow = 0; $i = 0
    foreach ($l in $lists) {
        $i++
        Write-Progress -Activity $script:Activity -Status "$($Label): library $i of $($lists.Count) ($($l.Title))" -PercentComplete ($i / $lists.Count * 60)
        $root = & $libRoot $l
        $items = @()
        if ($l.ItemCount -gt 0) {
            try { $items = @(Get-PnPListItem -List $l -PageSize 5000 -Fields $fields); $State.calls += [math]::Max(1, [math]::Ceiling($l.ItemCount / 5000)) }
            catch { $warnings.Add("$($Label): library '$($l.Title)': $($_.Exception.Message)"); $partialReasons += "library '$($l.Title)' not readable"; continue }
        }
        # safety net: if the view URL didn't give the real root, use the shortest folder the items live in
        $dirs = @($items | ForEach-Object { [string]$_.FieldValues.FileDirRef } | Where-Object { $_ })
        if ($dirs -and ($dirs | Where-Object { -not $_.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) })) {
            $root = ($dirs | Sort-Object Length | Select-Object -First 1).TrimEnd('/')
        }
        $n = 0
        foreach ($it in $items) {
            $fv = $it.FieldValues
            if ([int]$fv.FSObjType -ne 0) { continue }
            $size = ConvertTo-SPSTInt64 $fv.SMTotalFileStreamSize
            if ($null -eq $size) { $size = ConvertTo-SPSTInt64 $fv.File_x0020_Size }
            if ($null -eq $size) { $size = 0 }
            $total = ConvertTo-SPSTInt64 $fv.SMTotalSize
            if ($null -eq $total) { $noSM++; $total = $size }                     # storage metrics not computed yet
            elseif ($total -lt $size) { $overflow++; $total = $size }            # Int32 LookupId overflow (> 2 GB)
            $dir = [string]$fv.FileDirRef
            $folder = if ($dir.Length -gt $root.Length) { $dir.Substring($root.Length).Trim('/') } else { '' }
            $label = [string]$fv._UIVersionString
            $editor = $fv.Editor
            $files.Add([ordered]@{
                library      = $l.Title
                folder       = $folder
                name         = [string]$fv.FileLeafRef
                size         = $size
                totalSize    = $total
                versionCount = Get-SPSTMajorVersion $label
                versionLabel = $label
                modified     = ConvertTo-SPSTIso $fv.Modified
                modifiedBy   = if ($editor -is [Microsoft.SharePoint.Client.FieldUserValue]) { $editor.LookupValue } else { [string]$editor }
                created      = ConvertTo-SPSTIso $fv.Created
                url          = [string]$fv.FileRef
            })
            $refs.Add(@{ ref = [string]$fv.FileRef; hist = $total - $size; old = (Get-SPSTMajorVersion $label) - 1 })
            $n++
        }
        $libsOut.Add([ordered]@{ title = $l.Title; url = $root; itemCount = $n; hidden = [bool]$l.Hidden; template = [int]$l.BaseTemplate })
    }

    # ---- per-version detail for files with real history (batched: one request per $VersionBatchSize files)
    $want = @(for ($k = 0; $k -lt $files.Count; $k++) {
        $r = $refs[$k]
        if ($r.old -ge 1 -and ($r.old -ge $VersionDetailMinVersions -or $r.hist -ge $VersionDetailMinHistoryMB * 1MB)) { $k }
    })
    for ($b = 0; $b -lt $want.Count; $b += $VersionBatchSize) {
        $chunk = $want[$b..([math]::Min($b + $VersionBatchSize, $want.Count) - 1)]
        Write-Progress -Activity $script:Activity -Status "$($Label): version detail, $([math]::Min($b + $VersionBatchSize, $want.Count)) of $($want.Count) files" -PercentComplete (60 + $b / [math]::Max(1, $want.Count) * 40)
        $pending = foreach ($k in $chunk) {
            $f = $ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($refs[$k].ref))
            $v = $f.Versions; $ctx.Load($v)
            @{ k = $k; v = $v }
        }
        try { Invoke-PnPQuery -RetryCount 5; $State.calls++ }
        catch {
            # one bad file fails the whole batch: fall back to one request per file for this batch
            $warnings.Add("$($Label): version batch failed ($($_.Exception.Message)); retrying file by file")
            $pending = foreach ($k in $chunk) {
                try {
                    $f = $ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($refs[$k].ref))
                    $v = $f.Versions; $ctx.Load($v); Invoke-PnPQuery -RetryCount 5; $State.calls++
                    @{ k = $k; v = $v }
                } catch { $warnings.Add("Versions of $($refs[$k].ref): $($_.Exception.Message)") }
            }
        }
        foreach ($p in $pending) {
            $list = @($p.v | ForEach-Object {
                $sz = if ($_.PSObject.Properties['Length'] -and $_.Length) { [int64]$_.Length } else { [int64]$_.Size }
                [ordered]@{ label = $_.VersionLabel; size = $sz; created = ConvertTo-SPSTIso $_.Created }
            })
            $files[$p.k].versions = $list
            $files[$p.k].versionCount = $list.Count + 1     # exact, includes minor versions
        }
    }

    if ($noSM) { $warnings.Add("$($Label): $noSM files had no storage metrics yet (new or recently changed); their version history is counted as 0.") }
    if ($overflow) { $warnings.Add("$($Label): $overflow files reported a total size smaller than their current size (likely > 2 GB); counted as current size only.") }
    $siteOut = [ordered]@{ url = $web.Url; title = $web.Title }
    if ($ParentUrl) { $siteOut.parentUrl = $ParentUrl }
    if ($IsTop) { $siteOut.storageUsed = $storageUsed }
    $siteOut.lastModified = ConvertTo-SPSTIso $web.LastItemUserModifiedDate
    $siteOut.deep = $true
    if ($partialReasons) { $siteOut.partial = $true; $siteOut.partialReason = ($partialReasons -join '; ') }
    if ($holdLib) { $siteOut.preservationHold = $true }
    $siteOut.libraries = $libsOut
    $siteOut.files = $files
    @{ site = $siteOut; versionDetail = $want.Count; partial = [bool]$partialReasons }
}

function Invoke-SPSTSite {
    # Connect to one site collection and scan it (+ subsites). Returns @{ results = [scan...]; top = scan }.
    param([string]$Url, [string]$ClientId, [switch]$IncludeSubsites, [switch]$Connect, [hashtable]$Opt, [string]$Label = 'site')
    if ($Connect) { Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId }
    $results = [System.Collections.Generic.List[object]]::new()
    $top = Get-SPSTWebScan @Opt -IsTop -Label $Label
    $results.Add($top)
    if ($IncludeSubsites) {
        # Walk the subsite tree with GetSubwebsForCurrentUser (works with AllSites.Read; Web.Webs does not).
        $childUrls = {
            $c = Get-PnPContext; $s = $c.Web.GetSubwebsForCurrentUser($null); $c.Load($s); $c.ExecuteQuery(); $Opt.State.calls++
            @($s | ForEach-Object { $_.ServerRelativeUrl })
        }
        $origin = ([uri]$Url).GetLeftPart([UriPartial]::Authority)
        $queue = [System.Collections.Generic.Queue[string]]::new()
        try { & $childUrls | ForEach-Object { $queue.Enqueue($origin + $_) } }
        catch { $Opt.State.warnings.Add("Subsites of $Url not listed: $($_.Exception.Message)") }
        $n = 0
        while ($queue.Count) {
            $u = $queue.Dequeue().TrimEnd('/'); $n++
            try {
                if (-not $ClientId) { throw 'subsites need -ClientId (or a connection made with -ClientId) to sign in to each subsite' }
                Connect-PnPOnline -Url $u -Interactive -ClientId $ClientId
                $results.Add((Get-SPSTWebScan @Opt -ParentUrl $u.Substring(0, $u.LastIndexOf('/')) -Label "$Label, subsite $n"))
                & $childUrls | ForEach-Object { $queue.Enqueue($origin + $_) }
            } catch {
                $Opt.State.warnings.Add("Subsite $u not scanned: $($_.Exception.Message)")
                $top.site.partial = $true
                $top.site.partialReason = (@($top.site.partialReason, "subsite $u not readable") | Where-Object { $_ }) -join '; '
            }
        }
        if ($n) { Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId }
    }
    @{ results = $results; top = $top }
}

function Get-SPStorageTree {
    <#
    .SYNOPSIS
    Scans SharePoint storage and writes a report (JSON) for the SharePoint Storage Tree viewer:
    one site, a site with its subsites, or the whole tenant.

    .DESCRIPTION
    Read-only. For each site scanned in detail it reads every document library (including hidden and system
    ones, so totals match the site's storage figure), every file's current and total size, and per-version
    sizes for files with a lot of history. Tenant mode lists every site with its storage total (one call) and
    scans the sites you choose in detail.

    Sign-in: pass -ClientId to sign in interactively (your tenant's own Entra ID app registration, see
    Register-PnPEntraIDAppForInteractiveLogin), or, in -Site mode, connect first with Connect-PnPOnline.
    Site mode works best as a site collection admin. Tenant mode needs a SharePoint Administrator; detailed
    scans still need site collection admin rights on each chosen site (an admin role alone can't read content).
    App permission (delegated SharePoint): AllSites.Read is enough for -Site and -IncludeSubsites;
    -Tenant needs AllSites.FullControl (the admin site list refuses Read and Manage).

    .EXAMPLE
    Get-SPStorageTree -Site https://contoso.sharepoint.com/sites/ProjectHub -ClientId <app-id>

    .EXAMPLE
    Get-SPStorageTree -Site https://contoso.sharepoint.com/sites/ProjectHub -ClientId <app-id> -IncludeSubsites

    .EXAMPLE
    Get-SPStorageTree -Tenant contoso -ClientId <app-id> -Top 5
    Every site's total, plus the 5 largest sites scanned in detail.

    .EXAMPLE
    Get-SPStorageTree -Tenant contoso -ClientId <app-id> -DeepScan https://contoso.sharepoint.com/sites/Finance
    #>
    [CmdletBinding(DefaultParameterSetName = 'Site')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Site', Position = 0)][string]$Site,
        # tenant name ("contoso") or admin center URL
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][string]$Tenant,
        # tenant mode: sites to scan in detail
        [Parameter(ParameterSetName = 'Tenant')][string[]]$DeepScan = @(),
        # tenant mode: also scan the N largest sites in detail
        [Parameter(ParameterSetName = 'Tenant')][ValidateRange(0, 1000)][int]$Top = 0,
        [string]$ClientId,
        [string]$Out = 'report.json',
        # also scan every subsite (of the site, or of each site scanned in detail)
        [switch]$IncludeSubsites,
        # fetch exact per-version sizes for files with at least this many old versions...
        [int]$VersionDetailMinVersions = 10,
        # ...or at least this much version history (MB, estimated from SMTotalSize)
        [double]$VersionDetailMinHistoryMB = 50,
        # files per version-detail request
        [ValidateRange(1, 200)][int]$VersionBatchSize = 50,
        [switch]$Pretty,
        # open the report in the viewer when the scan finishes (writes <Out>.html next to the JSON)
        [switch]$Open
    )
    $ErrorActionPreference = 'Stop'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $state = @{ calls = 0; warnings = [System.Collections.Generic.List[string]]::new() }
    $opt = @{ State = $state; VersionDetailMinVersions = $VersionDetailMinVersions; VersionDetailMinHistoryMB = $VersionDetailMinHistoryMB; VersionBatchSize = $VersionBatchSize }
    $sitesOut = [System.Collections.Generic.List[object]]::new()
    $scans = [System.Collections.Generic.List[object]]::new()     # every web scanned in detail
    $tenantName = $null; $failed = @()

    if ($PSCmdlet.ParameterSetName -eq 'Site') {
        $Site = $Site.TrimEnd('/')
        if (-not $ClientId) {
            $conn = try { Get-PnPConnection } catch { $null }
            if (-not $conn) { throw "Not connected. Pass -ClientId, or run Connect-PnPOnline -Url $Site first." }
            if ($conn.Url.TrimEnd('/') -ne $Site) { throw "Connected to $($conn.Url), not $Site. Connect to the site you want to scan, or pass -ClientId." }
            $ClientId = $conn.ClientId
        }
        $r = Invoke-SPSTSite -Url $Site -ClientId $ClientId -IncludeSubsites:$IncludeSubsites -Connect:([bool]$PSBoundParameters.ClientId) -Opt $opt
        $r.results | ForEach-Object { $scans.Add($_); $sitesOut.Add($_.site) }
        $scope = if ($IncludeSubsites) { 'subsites' } else { 'site' }
    }
    else {
        if (-not $ClientId) { throw 'Tenant mode needs -ClientId (it signs in to the admin center and to each site scanned in detail).' }
        $adminUrl = if ($Tenant -match '^https://') { $Tenant.TrimEnd('/') } else { "https://$Tenant-admin.sharepoint.com" }
        $tenantName = ([uri]$adminUrl).Host -replace '-admin\.sharepoint\.com$', ''
        Write-Progress -Activity $script:Activity -Status 'Reading the tenant site list'
        Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $ClientId
        $all = @(Get-PnPTenantSite | Where-Object { $_.Template -notmatch '^RedirectSite' }); $state.calls++
        $all = @($all | Sort-Object StorageUsageCurrent -Descending)
        $byUrl = @{}; foreach ($s in $all) { $byUrl[$s.Url.TrimEnd('/').ToLower()] = $s }

        $deep = [System.Collections.Generic.List[string]]::new()
        foreach ($u in $DeepScan) {
            $k = $u.TrimEnd('/').ToLower()
            if ($byUrl.ContainsKey($k)) { if (-not $deep.Contains($k)) { $deep.Add($k) } }
            else { $state.warnings.Add("-DeepScan $u is not in the tenant site list (check the URL); skipped.") }
        }
        foreach ($s in ($all | Select-Object -First $Top)) { $k = $s.Url.TrimEnd('/').ToLower(); if (-not $deep.Contains($k)) { $deep.Add($k) } }

        $deepDone = @{}
        $n = 0
        foreach ($k in $deep) {
            $n++; $s = $byUrl[$k]
            try {
                $r = Invoke-SPSTSite -Url $s.Url -ClientId $ClientId -IncludeSubsites:$IncludeSubsites -Connect -Opt $opt -Label "site $n of $($deep.Count) ($($s.Title))"
                $r.top.site.template = $s.Template
                $r.results | ForEach-Object { $scans.Add($_) }
                $deepDone[$k] = $r.results
            } catch {
                $failed += $s.Url
                $state.warnings.Add("Detailed scan of $($s.Url) failed, kept as summary only: $($_.Exception.Message) (detailed scans need site collection admin rights on the site)")
            }
        }
        foreach ($s in $all) {
            $k = $s.Url.TrimEnd('/').ToLower()
            if ($deepDone.ContainsKey($k)) { $deepDone[$k] | ForEach-Object { $sitesOut.Add($_.site) }; continue }
            $o = [ordered]@{ url = $s.Url; title = $s.Title; template = $s.Template
                             storageUsed = [int64]$s.StorageUsageCurrent * 1MB
                             lastModified = ConvertTo-SPSTIso $s.LastContentModifiedDate; deep = $false }
            if ($failed -contains $s.Url) { $o.note = 'detailed scan failed (no access)' }
            $sitesOut.Add($o)
        }
        $scope = 'tenant'
    }

    # ---- write ---------------------------------------------------------------------------------
    $report = [ordered]@{
        schemaVersion = '1.1'
        generator     = [ordered]@{ name = 'SPStorageTree'; version = $script:GeneratorVersion }
        scannedAt     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        scope         = $scope
    }
    if ($tenantName) { $report.tenant = $tenantName }
    $report.sites = $sitesOut
    $report.scan = [ordered]@{ seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1); requests = $state.calls
                               versionDetailFiles = [int]($scans | ForEach-Object { $_.versionDetail } | Measure-Object -Sum).Sum
                               warnings = $state.warnings }
    $json = $report | ConvertTo-Json -Depth 8 -Compress:(-not $Pretty)
    $outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
    [IO.File]::WriteAllText($outPath, $json, [Text.UTF8Encoding]::new($false))
    Write-Progress -Activity $script:Activity -Completed
    foreach ($w in $state.warnings) { Write-Warning $w }

    $allFiles = @($scans | ForEach-Object { $_.site.files })
    $sumTotal = [int64]($allFiles | ForEach-Object { $_.totalSize } | Measure-Object -Sum).Sum
    $tops = @($scans | Where-Object { $_.site.Contains('storageUsed') })
    $used = [int64]($tops | ForEach-Object { $_.site.storageUsed } | Measure-Object -Sum).Sum
    $summary = [ordered]@{
        Scope         = $scope
        Report        = $outPath
        Sites         = $sitesOut.Count
        DetailedWebs  = $scans.Count
        Files         = $allFiles.Count
        FilesBytes    = $sumTotal
        Coverage      = if ($used) { '{0:P1}' -f ($sumTotal / $used) } else { 'n/a' }
        VersionDetail = $report.scan.versionDetailFiles
        Partial       = [bool]($scans | Where-Object { $_.site.partial })
        Requests      = $state.calls
        Seconds       = [math]::Round($sw.Elapsed.TotalSeconds, 1)
    }
    if ($scope -eq 'tenant') { $summary.TenantBytes = [int64](($sitesOut | ForEach-Object { $_.storageUsed }) | Measure-Object -Sum).Sum; $summary.DeepScanFailed = $failed.Count }
    if ($Open) { $summary.Viewer = (Show-SPStorageTree -Path $outPath).FullName }
    [pscustomobject]$summary
}

function Show-SPStorageTree {
    <#
    .SYNOPSIS
    Opens one or more SPStorageTree reports in the viewer.

    .DESCRIPTION
    Writes a copy of the offline viewer with the report data embedded (one self-contained .html file, no
    network, nothing uploaded) and opens it in your default browser. Several reports are combined into one
    view. The .html file contains your tenant's file names and sizes: treat it like the report itself.

    .EXAMPLE
    Show-SPStorageTree .\report.json

    .EXAMPLE
    Show-SPStorageTree .\finance.json, .\hr.json -OutFile storage.html -NoOpen
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Report', 'FullName')][string[]]$Path,
        # default: next to the first report, same name with .html
        [string]$OutFile,
        [switch]$NoOpen
    )
    begin { $all = [System.Collections.Generic.List[string]]::new() }
    process { foreach ($p in $Path) { $all.Add((Resolve-Path -LiteralPath $p).ProviderPath) } }
    end {
        $viewer = Join-Path $PSScriptRoot 'viewer\storagetree.html'
        if (-not (Test-Path $viewer)) { throw "Viewer not found at $viewer. Reinstall the module." }
        $json = @($all | ForEach-Object {
            $t = [IO.File]::ReadAllText($_).TrimStart([char]0xFEFF)
            $null = $t | ConvertFrom-Json -ErrorAction Stop   # validate
            $t
        })
        $data = if ($json.Count -eq 1) { $json[0] } else { '[' + ($json -join ',') + ']' }
        $data = $data -replace '</', '<\/'                    # can't close the script tag early
        $html = [IO.File]::ReadAllText($viewer)
        if ($html -notmatch '<!--SPST:DATA-->') { throw 'This viewer version has no data slot.' }
        $html = $html.Replace('<!--SPST:DATA-->', "<script>window.SPST_REPORT = $data;</script>")
        if (-not $OutFile) { $OutFile = [IO.Path]::ChangeExtension($all[0], '.html') }
        $outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
        [IO.File]::WriteAllText($outPath, $html, [Text.UTF8Encoding]::new($false))
        if (-not $NoOpen) { Invoke-Item -LiteralPath $outPath }
        Get-Item -LiteralPath $outPath
    }
}

Export-ModuleMember -Function Get-SPStorageTree, Show-SPStorageTree
