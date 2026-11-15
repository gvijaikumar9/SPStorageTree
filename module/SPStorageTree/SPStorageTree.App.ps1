# SharePoint Storage Tree - Windows app (Start-SPStorageTree) and the batch engine behind it.
# Dot-sourced by SPStorageTree.psm1. The window only collects choices; the work is done by
# Invoke-SPSTBatch / Get-SPSTSiteList in a background runspace so the window never freezes.

function Get-SPSTSettingsPath {
    $base = if ($env:APPDATA) { $env:APPDATA } else { Join-Path $HOME '.config' }
    Join-Path (Join-Path $base 'SPStorageTree') 'settings.json'
}

function Get-SPSTSiteList {
    # Every site in the tenant with its storage total (needs SharePoint admin + AllSites.FullControl).
    param([Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][string]$ClientId)
    $ErrorActionPreference = 'Stop'
    $adminUrl = if ($Tenant -match '^https://') { $Tenant.TrimEnd('/') } else { "https://$Tenant-admin.sharepoint.com" }
    Write-Progress -Activity $script:Activity -Status 'Signing in and reading the site list'
    Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $ClientId
    $sites = @(Get-PnPTenantSite | Where-Object { $_.Template -notmatch '^RedirectSite' })
    Write-Progress -Activity $script:Activity -Completed
    @($sites | Sort-Object StorageUsageCurrent -Descending | ForEach-Object {
        [pscustomobject]@{
            Title      = if ($_.Title) { $_.Title } else { $_.Url }
            SizeMB     = [double]$_.StorageUsageCurrent
            Url        = $_.Url.TrimEnd('/')
            # the tenant API returns UTC (often with Kind Unspecified); show it in the user's local time
            LastChange = if ($_.LastContentModifiedDate) { [datetime]::SpecifyKind([datetime]$_.LastContentModifiedDate, 'Utc').ToLocalTime() } else { $null }
            Template   = $_.Template
        }
    })
}

function Get-SPSTBatchResult {
    # The batch result from a background job's output (or $null). PowerShell 7 runs Where-Object once even
    # when the input is $null, so null items must be guarded before touching .PSObject.
    param($Output)
    @($Output | Where-Object { $null -ne $_ -and $_.PSObject.Properties['Html'] }) | Select-Object -Last 1
}

function Invoke-SPSTBatch {
    # Runs one scan job for the app: several sites (one report each, combined in the viewer) or the tenant.
    param(
        [Parameter(Mandatory)][ValidateSet('Sites', 'Tenant')][string]$Mode,
        [string]$Tenant,
        [Parameter(Mandatory)][string]$ClientId,
        [string[]]$Sites = @(),
        [int]$Top = 0,
        [switch]$IncludeSubsites,
        [switch]$ExcludeUserNames,
        [Parameter(Mandatory)][string]$OutDir,
        [switch]$NoOpen,
        [hashtable]$Sync = @{}
    )
    $ErrorActionPreference = 'Stop'
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reports = [System.Collections.Generic.List[string]]::new()
    $failed = [System.Collections.Generic.List[string]]::new()
    $common = @{ ClientId = $ClientId; IncludeSubsites = [bool]$IncludeSubsites; ExcludeUserNames = [bool]$ExcludeUserNames }

    if ($Mode -eq 'Tenant') {
        if (-not $Tenant) { throw 'Tenant mode needs the tenant name.' }
        $name = ($Tenant -replace '^https://', '' -replace '-admin\.sharepoint\.com/?$', '')
        $out = Join-Path $OutDir "tenant-$name-$ts.json"
        $Sync.overall = "Whole tenant ($name)"
        $null = Get-SPStorageTree -Tenant $Tenant -DeepScan $Sites -Top $Top -Out $out @common
        $reports.Add($out)
    } else {
        $list = @($Sites | Where-Object { $_ } | ForEach-Object { $_.Trim().TrimEnd('/') } | Select-Object -Unique)
        if (-not $list) { throw 'Pick at least one site.' }
        for ($i = 0; $i -lt $list.Count; $i++) {
            $u = $list[$i]
            $Sync.overall = "Site $($i + 1) of $($list.Count): $u"
            $slug = (($u -split '/')[-1] -replace '[^A-Za-z0-9_-]', '-')
            if (-not $slug -or $u -match '\.sharepoint\.com$') { $slug = 'root' }
            $out = Join-Path $OutDir "site-$slug-$ts.json"
            try { $null = Get-SPStorageTree -Site $u -Out $out @common; $reports.Add($out) }
            catch {
                $why = $_.Exception.Message
                if ($why -match 'NotFound|\(404\)') { $why = 'site not found, check the URL' }
                elseif ($why -match 'unauthori|denied|\(401\)|\(403\)') { $why = 'no access (you need to be site collection admin there)' }
                $failed.Add($u); Write-Warning "$u not scanned: $why"
            }
        }
    }
    if (-not $reports.Count) { throw 'Nothing was scanned. The reasons are in the message log at the bottom of the window.' }
    $Sync.overall = 'Opening the viewer'
    $html = Show-SPStorageTree -Path $reports -OutFile (Join-Path $OutDir "storage-$ts.html") -NoOpen:$NoOpen
    [pscustomobject]@{ Html = $html.FullName; Reports = @($reports); Failed = @($failed); OutDir = $OutDir }
}

function Start-SPStorageTree {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Opens a window; changes nothing.')]
    <#
    .SYNOPSIS
    Opens the SharePoint Storage Tree window: sign in, pick sites (or the whole tenant), scan, and the viewer opens.

    .DESCRIPTION
    Windows only. The same read-only scanner as Get-SPStorageTree, with a window instead of parameters.
    SharePoint admins (app with AllSites.FullControl) get a list of every site with its size to tick from.
    Everyone else pastes the site URLs they want to scan. Reports and the viewer page are saved in the
    output folder (default: Documents\SPStorageTree).

    .EXAMPLE
    Start-SPStorageTree
    #>
    [CmdletBinding()]
    param()
    if (-not $IsWindows) { throw 'The window is Windows only. On macOS/Linux use Get-SPStorageTree and Show-SPStorageTree.' }
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    try { [void][System.Windows.Forms.Application]::SetHighDpiMode([System.Windows.Forms.HighDpiMode]::SystemAware) } catch { Write-Verbose 'High-DPI mode already set.' }
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $S = @{ busy = $false; ps = $null; handle = $null; task = $null; warnIdx = 0; errIdx = 0; signedIn = $false
            sync = [hashtable]::Synchronized(@{ overall = '' }); module = $MyInvocation.MyCommand.Module.ModuleBase }
    $settingsPath = Get-SPSTSettingsPath
    $cfg = @{ tenant = ''; clientId = ''; outDir = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'SPStorageTree') }
    try { (Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { if ($_.Value) { $cfg[$_.Name] = [string]$_.Value } } } catch { Write-Verbose 'No saved settings yet.' }

    # ---- layout ------------------------------------------------------------------------------
    $font = [System.Drawing.Font]::new('Segoe UI', 9)
    $f = [System.Windows.Forms.Form]@{ Text = 'SharePoint Storage Tree'; Size = [System.Drawing.Size]::new(980, 720); MinimumSize = [System.Drawing.Size]::new(820, 600)
                                       StartPosition = 'CenterScreen'; Font = $font }
    function Add-Ctl($type, $props) { $c = New-Object "System.Windows.Forms.$type"; foreach ($k in $props.Keys) { $c.$k = $props[$k] }; $f.Controls.Add($c); $c }
    $AL = [System.Windows.Forms.AnchorStyles]
    $TL = $AL::Top -bor $AL::Left; $TLR = $TL -bor $AL::Right; $ALL = $TLR -bor $AL::Bottom; $BL = $AL::Bottom -bor $AL::Left; $BLR = $BL -bor $AL::Right; $BR = $AL::Bottom -bor $AL::Right

    Add-Ctl Label @{ Text = 'Tenant name'; Location = '16,18'; AutoSize = $true } | Out-Null
    $tTenant = Add-Ctl TextBox @{ Location = '16,38'; Width = 170; Text = $cfg.tenant; PlaceholderText = 'contoso' }
    Add-Ctl Label @{ Text = 'App (client) ID'; Location = '200,18'; AutoSize = $true } | Out-Null
    $tClient = Add-Ctl TextBox @{ Location = '200,38'; Width = 290; Text = $cfg.clientId; PlaceholderText = '00000000-0000-0000-0000-000000000000' }
    $bSign = Add-Ctl Button @{ Text = 'Sign in and list sites'; Location = '504,36'; Width = 170; Height = 27 }
    [void](Add-Ctl Label @{ Text = 'SharePoint admins: sign in to pick from every site. Others: paste site URLs below.'; Location = '688,30'; Size = '270,40'; Anchor = $TLR; ForeColor = [System.Drawing.Color]::DimGray })

    $rSites = Add-Ctl RadioButton @{ Text = 'Scan the ticked sites in detail'; Location = '16,78'; AutoSize = $true; Checked = $true }
    $rTenant = Add-Ctl RadioButton @{ Text = 'Whole tenant: every site''s total, plus the ticked sites and the largest'; Location = '16,100'; AutoSize = $true; Enabled = $false }
    $nTop = Add-Ctl NumericUpDown @{ Location = '0,98'; Width = 55; Minimum = 0; Maximum = 1000; Value = 3; Enabled = $false }
    $lTop = Add-Ctl Label @{ Text = 'in detail'; Location = '0,101'; AutoSize = $true }; $nTop.Left = $rTenant.Right + 2; $lTop.Left = $nTop.Right + 6

    $tFilter = Add-Ctl TextBox @{ Location = '16,130'; Width = 260; PlaceholderText = 'Filter the list (type part of a name)' }
    $bAll = Add-Ctl Button @{ Text = 'Tick all'; Location = '288,128'; Width = 80 }
    $bNone = Add-Ctl Button @{ Text = 'Tick none'; Location = '374,128'; Width = 80 }
    $lCount = Add-Ctl Label @{ Text = ''; Location = '466,132'; AutoSize = $true; ForeColor = [System.Drawing.Color]::DimGray }

    $table = [System.Data.DataTable]::new()
    [void]$table.Columns.Add('Scan', [bool]); [void]$table.Columns.Add('Site', [string]); [void]$table.Columns.Add('Size (MB)', [double])
    [void]$table.Columns.Add('URL', [string]); [void]$table.Columns.Add('Last change', [datetime])
    $grid = Add-Ctl DataGridView @{ Location = '16,160'; Size = '932,250'; Anchor = $ALL; AllowUserToAddRows = $false; AllowUserToDeleteRows = $false
                                    RowHeadersVisible = $false; SelectionMode = 'FullRowSelect'; AutoSizeColumnsMode = 'Fill'; BackgroundColor = [System.Drawing.SystemColors]::Window }
    $grid.DataSource = $table.DefaultView

    [void](Add-Ctl Label @{ Text = 'Or paste site URLs, one per line'; Location = '16,418'; AutoSize = $true; Anchor = $BL })
    $tUrls = Add-Ctl TextBox @{ Location = '16,436'; Size = '932,50'; Multiline = $true; AcceptsReturn = $true; ScrollBars = 'Vertical'; Anchor = $BLR; PlaceholderText = 'https://contoso.sharepoint.com/sites/ProjectHub' }
    $cSub = Add-Ctl CheckBox @{ Text = 'Include subsites'; Location = '16,492'; AutoSize = $true; Anchor = $BL }
    $cNames = Add-Ctl CheckBox @{ Text = 'Leave out people''s names'; Location = '170,492'; AutoSize = $true; Anchor = $BL }

    [void](Add-Ctl Label @{ Text = 'Save reports in'; Location = '16,524'; AutoSize = $true; Anchor = $BL })
    $tOut = Add-Ctl TextBox @{ Location = '110,521'; Width = 622; Text = $cfg.outDir; Anchor = $BLR }
    $bBrowse = Add-Ctl Button @{ Text = 'Browse...'; Location = '744,519'; Width = 100; Anchor = $BR }
    $bOpenDir = Add-Ctl Button @{ Text = 'Open folder'; Location = '852,519'; Width = 96; Anchor = $BR }

    $bar = Add-Ctl ProgressBar @{ Location = '16,560'; Size = '716,18'; Anchor = $BLR }
    $lStatus = Add-Ctl Label @{ Text = 'Ready.'; Location = '16,582'; Size = '716,20'; Anchor = $BLR; AutoEllipsis = $true }
    $bScan = Add-Ctl Button @{ Text = 'Scan'; Location = '744,556'; Size = '100,46'; Anchor = $BR; Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold) }
    $bCancel = Add-Ctl Button @{ Text = 'Cancel'; Location = '852,556'; Size = '96,46'; Anchor = $BR; Enabled = $false }
    $tLog = Add-Ctl TextBox @{ Location = '16,610'; Size = '932,60'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; Anchor = $BLR; BackColor = [System.Drawing.SystemColors]::Control }
    $f.AcceptButton = $bScan

    # ---- helpers -----------------------------------------------------------------------------
    $log = { param($m) $tLog.AppendText("[$(Get-Date -Format HH:mm:ss)] $m`r`n") }
    $msg = { param($m, $icon = 'Warning') [void][System.Windows.Forms.MessageBox]::Show($f, $m, 'SharePoint Storage Tree', 'OK', $icon) }
    $guidOk = { $tClient.Text.Trim() -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$' }
    $tenantOk = { $tTenant.Text.Trim() -match '^[A-Za-z0-9-]+$' }
    $saveCfg = {
        try {
            $d = Split-Path $settingsPath; if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force $d | Out-Null }
            @{ tenant = $tTenant.Text.Trim(); clientId = $tClient.Text.Trim(); outDir = $tOut.Text.Trim() } | ConvertTo-Json | Set-Content -LiteralPath $settingsPath -Encoding utf8
        } catch { Write-Verbose "Settings not saved: $($_.Exception.Message)" }
    }
    $toFront = {
        # sign-in happens in the browser, which takes the foreground; bring the window back afterwards
        if ($f.WindowState -eq 'Minimized') { $f.WindowState = 'Normal' }
        $f.TopMost = $true; $f.Activate(); $f.BringToFront(); $f.TopMost = $false
    }
    $updateCount = {
        $n = @($table.Rows | Where-Object { $_['Scan'] }).Count
        $shown = $table.DefaultView.Count
        $lCount.Text = if (-not $table.Rows.Count) { '' }
                       elseif ($shown -lt $table.Rows.Count) { "$n of $($table.Rows.Count) sites ticked  |  $shown shown by the filter" }
                       else { "$n of $($table.Rows.Count) sites ticked" }
    }
    $setBusy = {
        param([bool]$b)
        $S.busy = $b
        foreach ($c in $bSign, $bScan, $tTenant, $tClient, $rSites, $grid, $tUrls, $cSub, $cNames, $tOut, $bBrowse, $bAll, $bNone, $tFilter) { $c.Enabled = -not $b }
        $rTenant.Enabled = (-not $b) -and $S.signedIn; $nTop.Enabled = (-not $b) -and $S.signedIn -and $rTenant.Checked
        $bCancel.Enabled = $b
        $bar.Style = if ($b) { 'Marquee' } else { 'Blocks' }; if (-not $b) { $bar.Value = 0 }
    }
    $runBg = {
        # run a script in the background runspace; $onDone gets the output (or $null + error text)
        param([string]$script, [hashtable]$args2, [scriptblock]$onDone, [string]$what)
        if (-not $S.ps) {
            $iss = [initialsessionstate]::CreateDefault2(); $iss.ExecutionPolicy = 'Bypass'
            $rs = [runspacefactory]::CreateRunspace($iss); $rs.Open()
            $rs.SessionStateProxy.SetVariable('SPSTSync', $S.sync)
            $S.ps = [powershell]::Create(); $S.ps.Runspace = $rs
            [void]$S.ps.AddScript("Import-Module -Name '$((Join-Path $S.module 'SPStorageTree.psd1') -replace "'", "''")' -Force").Invoke()
            if ($S.ps.HadErrors) { throw "Couldn't load the module in the background: $($S.ps.Streams.Error[0])" }
        }
        $S.ps.Commands.Clear(); $S.ps.Streams.ClearStreams(); $S.warnIdx = 0; $S.errIdx = 0; $S.sync.overall = ''
        [void]$S.ps.AddScript($script).AddArgument($args2)
        $S.task = @{ onDone = $onDone; what = $what; started = Get-Date }
        $S.handle = $S.ps.BeginInvoke()
        & $setBusy $true
        & $log "$what..."
    }

    # ---- timer: progress, warnings, completion -------------------------------------------------
    $timer = [System.Windows.Forms.Timer]@{ Interval = 300 }
    $timer.Add_Tick({
        if (-not $S.handle) { return }
        $ps = $S.ps
        $p = $ps.Streams.Progress | Select-Object -Last 1
        $parts = @($S.sync.overall, $(if ($p -and $p.RecordType -ne 'Completed') { $p.StatusDescription })) | Where-Object { $_ }
        $el = [int]((Get-Date) - $S.task.started).TotalSeconds
        $lStatus.Text = "$(if ($parts) { $parts -join '  |  ' } else { $S.task.what })   ($el s)"
        if ($p -and $p.PercentComplete -ge 0 -and $p.RecordType -ne 'Completed') { $bar.Style = 'Blocks'; $bar.Value = [math]::Min(100, $p.PercentComplete) } else { $bar.Style = 'Marquee' }
        while ($S.warnIdx -lt $ps.Streams.Warning.Count) { & $log "Warning: $($ps.Streams.Warning[$S.warnIdx].Message)"; $S.warnIdx++ }
        if ($S.handle.IsCompleted) {
            $out = $null; $err = $null
            try { $out = $ps.EndInvoke($S.handle) } catch { $err = $_.Exception.InnerException.Message; if (-not $err) { $err = $_.Exception.Message } }
            while ($S.warnIdx -lt $ps.Streams.Warning.Count) { & $log "Warning: $($ps.Streams.Warning[$S.warnIdx].Message)"; $S.warnIdx++ }
            if (-not $err -and $ps.Streams.Error.Count) { $err = ($ps.Streams.Error | ForEach-Object { $_.Exception.Message } | Select-Object -Unique) -join "`n" }
            $S.handle = $null
            & $setBusy $false
            # never let a UI bug take the window down: log it and tell the user instead
            try { & $S.task.onDone $out $err }
            catch {
                & $log "Internal error: $($_.Exception.Message)"
                $lStatus.Text = 'Something went wrong. See the message log below.'
                & $msg "Something went wrong in SharePoint Storage Tree:`n`n$($_.Exception.Message)`n`nPlease report it at https://github.com/gvijaikumar9/SPStorageTree/issues" 'Error'
            }
        }
    })
    $timer.Start()

    # ---- events --------------------------------------------------------------------------------
    $bSign.Add_Click({
        if (-not (& $tenantOk)) { & $msg 'Enter the tenant name, the part before .sharepoint.com (for example contoso).'; return }
        if (-not (& $guidOk)) { & $msg 'Enter the Application (client) ID of your PnP PowerShell app registration. It is a GUID. See the README for the one command that creates it.'; return }
        & $saveCfg
        & $runBg 'param($a) & (Get-Module SPStorageTree) { param($a) Get-SPSTSiteList @a } $a' @{ Tenant = $tTenant.Text.Trim(); ClientId = $tClient.Text.Trim() } {
            param($out, $err)
            & $toFront
            if ($err) {
                $S.signedIn = $false; $rTenant.Enabled = $false; $rSites.Checked = $true
                & $log "Couldn't list the tenant's sites: $err"
                $lStatus.Text = 'No site list. Paste site URLs below instead.'
                & $msg "Couldn't list the sites in the tenant.`n`nThis needs a SharePoint Administrator and an app registration with AllSites.FullControl.`n`nYou can still scan sites you own: paste their URLs in the box below and click Scan.`n`nDetails: $err"
                return
            }
            $table.Rows.Clear()
            foreach ($site in $out) { [void]$table.Rows.Add($false, $site.Title, $site.SizeMB, $site.Url, $(if ($site.LastChange) { [datetime]$site.LastChange } else { [DBNull]::Value })) }
            $S.signedIn = $true; $rTenant.Enabled = $true
            $grid.Columns['Size (MB)'].DefaultCellStyle.Format = 'N0'; $grid.Columns['Size (MB)'].DefaultCellStyle.Alignment = 'MiddleRight'
            $grid.Columns['Scan'].FillWeight = 12; $grid.Columns['Size (MB)'].FillWeight = 22; $grid.Columns['Last change'].FillWeight = 30
            foreach ($c in 'Site', 'Size (MB)', 'URL', 'Last change') { $grid.Columns[$c].ReadOnly = $true }
            $total = ($out | Measure-Object SizeMB -Sum).Sum
            $lStatus.Text = "Signed in. $($out.Count) sites, $('{0:N0}' -f $total) MB in total. Tick the sites to scan."
            & $log "Listed $($out.Count) sites."
            & $updateCount
        } 'Signing in: finish the sign-in in your browser (if it asks which account, pick your admin account)'
    })
    $tFilter.Add_TextChanged({
        $q = $tFilter.Text.Trim() -replace "'", "''" -replace '([\[\]%*])', '[$1]'
        $table.DefaultView.RowFilter = if ($q) { "Site LIKE '%$q%' OR URL LIKE '%$q%'" } else { '' }
        & $updateCount
    })
    $bAll.Add_Click({ foreach ($r in $table.DefaultView) { $r['Scan'] = $true }; & $updateCount })
    $bNone.Add_Click({ foreach ($r in $table.Rows) { $r['Scan'] = $false }; & $updateCount })
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { [void]$grid.CommitEdit('CurrentCellChange') } })
    $grid.Add_CellValueChanged({ & $updateCount })
    $rTenant.Add_CheckedChanged({ $nTop.Enabled = $rTenant.Checked -and $rTenant.Enabled })
    $bBrowse.Add_Click({
        $d = [System.Windows.Forms.FolderBrowserDialog]@{ Description = 'Folder for reports'; SelectedPath = $tOut.Text }
        if ($d.ShowDialog($f) -eq 'OK') { $tOut.Text = $d.SelectedPath }
    })
    $bOpenDir.Add_Click({ $p = $tOut.Text.Trim(); if ($p -and (Test-Path -LiteralPath $p)) { Invoke-Item -LiteralPath $p } else { & $msg 'That folder does not exist yet. It is created on the first scan.' 'Information' } })
    $bCancel.Add_Click({
        if ($S.ps -and $S.handle) { & $log 'Cancelling...'; $lStatus.Text = 'Cancelling...'; [void]$S.ps.BeginStop($null, $null) }
    })
    $bScan.Add_Click({
        if (-not (& $guidOk)) { & $msg 'Enter the Application (client) ID of your PnP PowerShell app registration (a GUID).'; return }
        $ticked = @($table.Rows | Where-Object { $_['Scan'] } | ForEach-Object { [string]$_['URL'] })
        $pasted = @($tUrls.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        # a site URL typed into the filter box (easy mistake): offer to include it, in either mode
        $fq = $tFilter.Text.Trim().TrimEnd('/')
        if ($fq -match '^https://[A-Za-z0-9-]+\.sharepoint\.com(/.*)?$' -and $fq -notin $pasted -and $fq -notin $ticked) {
            $how = if ($rTenant.Checked) { 'Add it to the sites scanned in detail?' } else { 'Scan this site?' }
            $ans = [System.Windows.Forms.MessageBox]::Show($f, "That URL is in the filter box, which only filters the list.`n`n$how`n`n$fq", 'SharePoint Storage Tree', 'YesNoCancel', 'Question')
            if ($ans -eq 'Cancel') { return }
            if ($ans -eq 'Yes') { $tUrls.Text = (@($pasted) + $fq) -join "`r`n"; $pasted = @($pasted) + $fq }
            $tFilter.Text = ''
        }
        $bad = @($pasted | Where-Object { $_ -notmatch '^https://[A-Za-z0-9-]+\.sharepoint\.com(/.*)?$' })
        if ($bad) { & $msg "These don't look like SharePoint Online site URLs:`n`n$($bad -join "`n")"; return }
        $out = $tOut.Text.Trim(); if (-not $out) { & $msg 'Choose a folder for the reports.'; return }
        $mode = if ($rTenant.Checked) { 'Tenant' } else { 'Sites' }
        $sites = @($ticked + $pasted | Select-Object -Unique)
        if ($mode -eq 'Sites' -and -not $sites) { & $msg 'Tick at least one site in the list, or paste a site URL.' 'Information'; return }
        if ($mode -eq 'Tenant' -and -not (& $tenantOk)) { & $msg 'Enter the tenant name.'; return }
        & $saveCfg
        $a = @{ Mode = $mode; ClientId = $tClient.Text.Trim(); Sites = $sites; OutDir = $out; IncludeSubsites = $cSub.Checked; ExcludeUserNames = $cNames.Checked }
        if ($mode -eq 'Tenant') { $a.Tenant = $tTenant.Text.Trim(); $a.Top = [int]$nTop.Value }
        $what = if ($mode -eq 'Tenant') { "Scanning the whole tenant ($($sites.Count) ticked + largest $([int]$nTop.Value) in detail)" } else { "Scanning $($sites.Count) site(s)" }
        $what += '. If a browser window opens, finish the sign-in there'
        & $runBg 'param($a) $a.Sync = $SPSTSync; & (Get-Module SPStorageTree) { param($a) Invoke-SPSTBatch @a } $a' $a {
            param($out, $err)
            $r = Get-SPSTBatchResult $out
            if ($err -or -not $r -or $r.Failed.Count) { & $toFront }   # only jump in front of the viewer when there's something to read
            if ($err -and -not $r) {
                $lStatus.Text = 'Scan stopped.'; & $log "Error: $err"
                if ($err -notmatch 'stopped') { & $msg "The scan didn't finish.`n`n$err" }
                return
            }
            $lStatus.Text = "Done. Opened the viewer. Reports are in $($r.OutDir)"
            & $log "Done: $($r.Reports.Count) report(s), viewer $($r.Html)"
            if ($r.Failed.Count) { & $msg "Done, but these sites couldn't be scanned (usually: you are not site collection admin there):`n`n$($r.Failed -join "`n")" 'Information' }
        } $what
    })
    $f.Add_FormClosing({
        param($src, $e)
        $null = $src
        if ($S.busy) {
            $ans = [System.Windows.Forms.MessageBox]::Show($f, 'A scan is still running. Stop it and close?', 'SharePoint Storage Tree', 'YesNo', 'Question')
            if ($ans -ne 'Yes') { $e.Cancel = $true; return }
            try { $S.ps.Stop() } catch { Write-Verbose 'Scan already stopped.' }
        }
        $timer.Stop()
        if ($S.ps) { try { $S.ps.Runspace.Dispose(); $S.ps.Dispose() } catch { Write-Verbose 'Background runspace already closed.' } }
    })

    if ($env:SPST_APP_SCREENSHOT) {
        # test hook: render the window to a PNG without user interaction, then close
        $f.Add_Shown({
            foreach ($row in @(@($true, 'Intranet Demo', 140, 'https://contoso.sharepoint.com/sites/intranet-demo', [datetime]'2026-11-03'),
                               @($false, 'List Dashboard Demo', 11, 'https://contoso.sharepoint.com/sites/listdash', [datetime]'2026-11-11'),
                               @($true, 'PowerShell Demo', 2, 'https://contoso.sharepoint.com/sites/PSDemo', [datetime]'2026-11-15'))) { [void]$table.Rows.Add($row) }
            $S.signedIn = $true; $rTenant.Enabled = $true; $tTenant.Text = 'contoso'; $tClient.Text = '00000000-0000-0000-0000-000000000000'
            $grid.Columns['Size (MB)'].DefaultCellStyle.Format = 'N0'; $grid.Columns['Scan'].FillWeight = 12; $grid.Columns['Size (MB)'].FillWeight = 22; $grid.Columns['Last change'].FillWeight = 30
            & $updateCount; $lStatus.Text = 'Signed in. 3 sites, 153 MB in total. Tick the sites to scan.'
            $f.Refresh(); [System.Windows.Forms.Application]::DoEvents()
            $bmp = [System.Drawing.Bitmap]::new($f.Width, $f.Height); $f.DrawToBitmap($bmp, [System.Drawing.Rectangle]::new(0, 0, $f.Width, $f.Height))
            $bmp.Save($env:SPST_APP_SCREENSHOT); $bmp.Dispose(); $f.Close()
        })
    }
    [void]$f.ShowDialog()
    $f.Dispose()
}
