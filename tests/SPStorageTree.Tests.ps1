#Requires -Modules Pester
# Offline tests: no SharePoint connection needed. Run .\build.ps1 first (copies the viewer into the module).
BeforeAll {
    $script:ModuleDir = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'module') 'SPStorageTree'
    Import-Module (Join-Path $ModuleDir 'SPStorageTree.psd1') -Force
    $script:Guid = '00000000-0000-0000-0000-000000000001'
    function New-Report([string]$Path, [string]$FileName = 'Plan.docx') {
        @{ schemaVersion = '1.1'; scope = 'site'; scannedAt = '2026-11-15T00:00:00Z'
           sites = @(@{ url = 'https://contoso.sharepoint.com/sites/x'; title = 'X'; deep = $true
                        libraries = @(@{ title = 'Documents'; url = '/sites/x/Shared Documents' })
                        files = @(@{ library = 'Documents'; folder = ''; name = $FileName; size = 10; totalSize = 30; versionCount = 3 }) }) } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding utf8
        $Path
    }
}

Describe 'Module' {
    It 'has a valid manifest' {
        { Test-ModuleManifest (Join-Path $ModuleDir 'SPStorageTree.psd1') } | Should -Not -Throw
    }
    It 'exports exactly the three public commands' {
        (Get-Command -Module SPStorageTree).Name | Sort-Object | Should -Be @('Get-SPStorageTree', 'Show-SPStorageTree', 'Start-SPStorageTree')
    }
    It 'keeps the app engine private' {
        Get-Command Invoke-SPSTBatch, Get-SPSTSiteList -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
    It 'ships the viewer with a data slot' {
        $v = Join-Path (Join-Path $ModuleDir 'viewer') 'storagetree.html'
        $v | Should -Exist
        Get-Content -LiteralPath $v -Raw | Should -Match '<!--SPST:DATA-->'
    }
}

Describe 'Helpers' {
    It 'reads SMTotalSize from a FieldLookupValue' {
        InModuleScope SPStorageTree {
            $v = [Microsoft.SharePoint.Client.FieldLookupValue]::new(); $v.LookupId = 9284
            ConvertTo-SPSTInt64 $v | Should -Be 9284
        }
    }
    It 'converts sizes and treats blanks as missing' -TestCases @(
        @{ In = '123'; Out = 123 }, @{ In = 0; Out = 0 }, @{ In = $null; Out = $null }, @{ In = ''; Out = $null }
    ) {
        param($In, $Out)
        InModuleScope SPStorageTree -Parameters @{ In = $In; Out = $Out } { param($In, $Out) ConvertTo-SPSTInt64 $In | Should -Be $Out }
    }
    It 'reads the major version from the label' -TestCases @(
        @{ L = '16.0'; N = 16 }, @{ L = '1.0'; N = 1 }, @{ L = '0.3'; N = 1 }, @{ L = ''; N = 1 }, @{ L = 'abc'; N = 1 }
    ) {
        param($L, $N)
        InModuleScope SPStorageTree -Parameters @{ L = $L; N = $N } { param($L, $N) Get-SPSTMajorVersion $L | Should -Be $N }
    }
    It 'formats dates as UTC ISO and never throws on bad input' {
        InModuleScope SPStorageTree {
            ConvertTo-SPSTIso ([datetime]::new(2026, 11, 15, 10, 30, 0, [DateTimeKind]::Utc)) | Should -Be '2026-11-15T10:30:00Z'
            ConvertTo-SPSTIso 'not a date' | Should -BeNullOrEmpty
            ConvertTo-SPSTIso $null | Should -BeNullOrEmpty
        }
    }
    It 'finds the library root from the default view URL' -TestCases @(
        @{ U = '/sites/x/Shared Documents/Forms/AllItems.aspx'; R = '/sites/x/Shared Documents' }
        @{ U = '/sites/x/SitePages/Forms/ByAuthor.aspx'; R = '/sites/x/SitePages' }
        @{ U = '/sites/x/_catalogs/masterpage/Forms/AllItems.aspx'; R = '/sites/x/_catalogs/masterpage' }
        @{ U = '/sites/x/ClientSideAssets/AllItems.aspx'; R = '/sites/x/ClientSideAssets' }
        @{ U = '/sites/x/_catalogs/appfiles'; R = '/sites/x/_catalogs/appfiles' }
        @{ U = '/sites/x/Forms Library/Forms/All Forms.aspx'; R = '/sites/x/Forms Library' }
    ) {
        param($U, $R)
        InModuleScope SPStorageTree -Parameters @{ U = $U; R = $R } { param($U, $R) Get-SPSTLibraryRoot $U | Should -Be $R }
    }
    It 'refuses an output folder that does not exist' {
        InModuleScope SPStorageTree -Parameters @{ P = (Join-Path $TestDrive 'nope/report.json') } {
            param($P) { Resolve-SPSTOutPath $P } | Should -Throw '*doesn''t exist*'
        }
    }
    It 'writes files atomically (no temp file left behind)' {
        InModuleScope SPStorageTree -Parameters @{ P = (Join-Path $TestDrive 'atomic.json') } {
            param($P)
            Write-SPSTFile -Path $P -Text '{"a":1}'
            Get-Content -LiteralPath $P -Raw | Should -Be '{"a":1}'
            @(Get-ChildItem -LiteralPath (Split-Path $P) -Filter '*.tmp').Count | Should -Be 0
        }
    }
}

Describe 'Code safety' {
    It 'no function reassigns one of its own parameters under a different case' {
        # PowerShell variable names ignore case: '$label = ...' silently overwrites a '$Label' parameter.
        $files = Get-ChildItem $ModuleDir -Filter *.ps*1 | Where-Object Extension -in '.psm1', '.ps1'
        $hits = foreach ($file in $files) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            foreach ($fn in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
                $params = @($fn.Body.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
                $assigned = $fn.Body.FindAll({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $false) |
                    ForEach-Object { $_.Left } | Where-Object { $_ -is [System.Management.Automation.Language.VariableExpressionAst] } |
                    ForEach-Object { $_.VariablePath.UserPath }
                foreach ($a in $assigned) {
                    $p = $params | Where-Object { $_ -ieq $a -and $_ -cne $a }
                    if ($p) { "$($fn.Name): `$$a overwrites parameter `$$p" }
                }
            }
        }
        $hits | Should -BeNullOrEmpty
    }
}

Describe 'Get-SPStorageTree input validation (no network)' {
    It 'rejects <Name>' -TestCases @(
        @{ Name = 'a non-SharePoint site URL'; Params = @{ Site = 'https://example.com/sites/x'; ClientId = '00000000-0000-0000-0000-000000000001' } }
        @{ Name = 'an http site URL'; Params = @{ Site = 'http://contoso.sharepoint.com/sites/x'; ClientId = '00000000-0000-0000-0000-000000000001' } }
        @{ Name = 'a bad tenant name'; Params = @{ Tenant = 'contoso!'; ClientId = '00000000-0000-0000-0000-000000000001' } }
        @{ Name = 'a ClientId that is not a GUID'; Params = @{ Site = 'https://contoso.sharepoint.com/sites/x'; ClientId = 'my-app' } }
        @{ Name = 'a bad -DeepScan URL'; Params = @{ Tenant = 'contoso'; DeepScan = 'contoso/sites/x'; ClientId = '00000000-0000-0000-0000-000000000001' } }
        @{ Name = 'a zero version threshold'; Params = @{ Site = 'https://contoso.sharepoint.com/sites/x'; ClientId = '00000000-0000-0000-0000-000000000001'; VersionDetailMinVersions = 0 } }
    ) {
        param($Name, $Params)
        { Get-SPStorageTree @Params } | Should -Throw
    }
    It 'fails on a missing output folder before signing in' {
        { Get-SPStorageTree -Site https://contoso.sharepoint.com/sites/x -ClientId $Guid -Out (Join-Path $TestDrive 'missing/r.json') } |
            Should -Throw '*Output folder*'
    }
}

Describe 'Show-SPStorageTree' {
    It 'embeds the report and writes the html next to it' {
        $r = New-Report (Join-Path $TestDrive 'report.json')
        $h = Show-SPStorageTree $r -NoOpen
        $h.FullName | Should -Be (Join-Path $TestDrive 'report.html')
        $html = Get-Content -LiteralPath $h.FullName -Raw
        $html | Should -Match 'window\.SPST_REPORT = \{'
        $html | Should -Not -Match '<!--SPST:DATA-->'
    }
    It 'cannot be broken out of by a hostile file name' {
        $evil = '</script><script>alert(1)</script><!--x.docx'
        $r = New-Report (Join-Path $TestDrive 'evil.json') $evil
        $html = Get-Content -LiteralPath (Show-SPStorageTree $r -NoOpen).FullName -Raw
        $data = $html.Substring($html.IndexOf('window.SPST_REPORT'))
        $data = $data.Substring(0, $data.IndexOf('</script>'))
        $data | Should -Not -Match '<'                        # nothing in the data can open or close a tag
        $lt = '\' + 'u' + '003c'                              # the JSON escape for '<', built from pieces
        $data | Should -Match ([regex]::Escape("$lt/script>$($lt)script>alert(1)"))
        # and the escaped JSON still parses back to the original name
        $json = $data.Substring($data.IndexOf('{')).TrimEnd().TrimEnd(';')
        ($json | ConvertFrom-Json).sites[0].files[0].name | Should -Be $evil
    }
    It 'combines several reports into one array' {
        $a = New-Report (Join-Path $TestDrive 'a.json'); $b = New-Report (Join-Path $TestDrive 'b.json')
        $html = Get-Content -LiteralPath (Show-SPStorageTree $a, $b -OutFile (Join-Path $TestDrive 'ab.html') -NoOpen).FullName -Raw
        $html | Should -Match 'window\.SPST_REPORT = \[\{'
    }
    It 'rejects invalid JSON and non-report JSON' {
        $bad = Join-Path $TestDrive 'bad.json'; Set-Content -LiteralPath $bad '{ not json'
        { Show-SPStorageTree $bad -NoOpen } | Should -Throw '*not valid JSON*'
        $other = Join-Path $TestDrive 'other.json'; Set-Content -LiteralPath $other '{"hello":1}'
        { Show-SPStorageTree $other -NoOpen } | Should -Throw '*not an SPStorageTree report*'
    }
}
