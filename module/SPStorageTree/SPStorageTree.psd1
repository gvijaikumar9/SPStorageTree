@{
    RootModule        = 'SPStorageTree.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '5b6f0d0e-4b7a-4c55-9a1e-3f0c2d7e8a41'
    Author            = 'G Vijai Kumar'
    Copyright         = '(c) G Vijai Kumar. MIT License.'
    Description       = 'TreeSize for SharePoint: scans a SharePoint site and writes a storage report (where the storage goes, version-history bloat, stale/ROT content, Copilot readiness) for an offline HTML treemap viewer. Read-only.'
    PowerShellVersion = '7.4'
    RequiredModules   = @(@{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' })
    FunctionsToExport = @('Get-SPStorageTree', 'Show-SPStorageTree')
    FileList          = @('SPStorageTree.psd1', 'SPStorageTree.psm1', 'viewer\storagetree.html', 'LICENSE')
    CmdletsToExport   = @()
    AliasesToExport   = @()
    VariablesToExport = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('SharePoint', 'SharePointOnline', 'Storage', 'PnP', 'Copilot', 'Governance', 'Treemap', 'VersionHistory', 'M365')
            LicenseUri = 'https://github.com/gvijaikumar9/SPStorageTree/blob/main/LICENSE'
            ProjectUri = 'https://github.com/gvijaikumar9/SPStorageTree'
        }
    }
}
