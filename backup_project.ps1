[CmdletBinding()]
param(
    [ValidateSet('Incremental', 'Archive', 'Pull')]
    [string]$Mode,

    [switch]$ConfigureOnly,

    [switch]$PreviewOnly,

    [switch]$AutoConfirm
)

$ErrorActionPreference = 'Stop'

function Write-Section {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 66)
    Write-Host $Text
    Write-Host ('=' * 66)
}

function Get-ScriptRoot {
    $PSScriptRoot
}

function Test-IsCloudBackedProjectRoot {
    param([string]$ProjectRoot)

    $cloudParents = @(
        (Join-Path $HOME 'iCloudDrive'),
        (Join-Path $HOME 'OneDrive')
    )

    $projectFull = [IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\')
    foreach ($parent in $cloudParents) {
        if (-not (Test-Path -LiteralPath $parent)) { continue }
        $parentFull = [IO.Path]::GetFullPath($parent).TrimEnd('\')
        if ($projectFull.StartsWith($parentFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    $false
}

function Get-GlobalConfigPaths {
    $baseDir = Join-Path $env:APPDATA 'ResearchProjectBackup'
    [pscustomobject]@{
        BaseDir    = $baseDir
        RootsFile  = Join-Path $baseDir 'roots.json'
    }
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Get-TemplateConfig {
    $configPath = Join-Path (Get-ScriptRoot) 'backup_project.config.json'
    if (-not (Test-Path -LiteralPath $configPath)) {
        throw "Template config file not found: $configPath"
    }

    Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
}

function Get-PathCandidates {
    param(
        [string[]]$ExactCandidates,
        [string[]]$FallbackParents,
        [string[]]$NameHints
    )

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($candidate in $ExactCandidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ((Test-Path -LiteralPath $candidate) -and $seen.Add($candidate)) {
            $score = 200
            if ((Split-Path -Leaf $candidate) -match 'workspace') { $score += 50 }
            if ($candidate -match 'onedrive|icloud') { $score += 20 }
            $results.Add([pscustomobject]@{ Path = $candidate; Score = $score; Reason = 'Exact match' })
        }
    }

    foreach ($parent in $FallbackParents) {
        if (-not (Test-Path -LiteralPath $parent)) { continue }
        Get-ChildItem -LiteralPath $parent -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $name = $_.Name.ToLowerInvariant()
                $NameHints | Where-Object { $name -like "*$_*" }
            } |
            ForEach-Object {
                if ($seen.Add($_.FullName)) {
                    $score = 100
                    if ($_.Name -match 'workspace') { $score += 40 }
                    if ($_.FullName -match 'onedrive|icloud') { $score += 20 }
                    $results.Add([pscustomobject]@{ Path = $_.FullName; Score = $score; Reason = 'Name hint match' })
                }
            }
    }

    $results | Sort-Object -Property Score, Path -Descending -Unique
}

function Get-LikelyRoots {
    param([string]$ProjectRoot)

    $userHome = $HOME
    $documents = [Environment]::GetFolderPath('MyDocuments')
    $oneDrive = Join-Path $userHome 'OneDrive'
    $iCloud = Join-Path $userHome 'iCloudDrive'

    $localExact = @(
        (Join-Path $userHome 'Workspace'),
        (Join-Path $documents 'Workspace'),
        (Split-Path -Parent $ProjectRoot)
    )

    $cloudExact = @(
        (Join-Path $oneDrive 'Workspace'),
        (Join-Path $iCloud 'Workspace'),
        (Join-Path $oneDrive '文档\Workspace'),
        (Join-Path $iCloud 'Documents\Workspace')
    )

    $localFallback = @($userHome, $documents)
    $cloudFallback = @($oneDrive, $iCloud, (Join-Path $oneDrive '文档'), (Join-Path $iCloud 'Documents'))

    [pscustomobject]@{
        Local = Get-PathCandidates -ExactCandidates $localExact -FallbackParents $localFallback -NameHints @('workspace', 'research', 'project')
        Cloud = Get-PathCandidates -ExactCandidates $cloudExact -FallbackParents $cloudFallback -NameHints @('workspace', 'backup', 'archive', 'research', 'project')
    }
}

function Add-NestedCloudCandidates {
    param(
        [object[]]$CloudCandidates,
        [string]$LocalRoot
    )

    $results = New-Object System.Collections.Generic.List[object]
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($candidate in $CloudCandidates) {
        if ($seen.Add($candidate.Path)) {
            $results.Add($candidate)
        }
    }

    $localLeaf = Split-Path -Leaf $LocalRoot
    if ([string]::IsNullOrWhiteSpace($localLeaf)) {
        return $results | Sort-Object -Property Score, Path -Descending -Unique
    }

    foreach ($candidate in $CloudCandidates) {
        $nested = Join-Path $candidate.Path $localLeaf
        if ((Test-Path -LiteralPath $nested) -and $seen.Add($nested)) {
            $results.Add([pscustomobject]@{
                Path   = $nested
                Score  = ([int]$candidate.Score + 25)
                Reason = "Matches local root name '$localLeaf'"
            })
        }
    }

    $results | Sort-Object -Property Score, Path -Descending -Unique
}

function Show-Candidates {
    param(
        [string]$Label,
        [object[]]$Candidates
    )

    Write-Host "$Label candidates:"
    if (-not $Candidates -or $Candidates.Count -eq 0) {
        Write-Host '  (none found automatically)'
        return
    }

    for ($i = 0; $i -lt $Candidates.Count; $i++) {
        $item = $Candidates[$i]
        Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $item.Path, $item.Reason)
    }
}

function Select-Root {
    param(
        [string]$Label,
        [object[]]$Candidates
    )

    Show-Candidates -Label $Label -Candidates $Candidates

    if ($Candidates.Count -eq 1) {
        Write-Host "Automatically selected $Label root: $($Candidates[0].Path)"
        return $Candidates[0].Path
    }

    while ($true) {
        Write-Host ''
        $choice = Read-Host "Choose $Label root number, or press Enter to type a path"
        if ([string]::IsNullOrWhiteSpace($choice)) {
            $manual = Read-Host "Enter full $Label root path"
            if (Test-Path -LiteralPath $manual) {
                return (Resolve-Path -LiteralPath $manual).Path
            }
            Write-Host "Path not found: $manual" -ForegroundColor Yellow
            continue
        }

        $index = 0
        if ([int]::TryParse($choice, [ref]$index) -and $index -ge 1 -and $index -le $Candidates.Count) {
            return $Candidates[$index - 1].Path
        }

        Write-Host 'Invalid selection.' -ForegroundColor Yellow
    }
}

function Save-RootSelection {
    param(
        [string]$LocalRoot,
        [string]$CloudRoot
    )

    $paths = Get-GlobalConfigPaths
    Ensure-Directory -Path $paths.BaseDir
    @{
        localWorkspaceRoot = $LocalRoot
        cloudWorkspaceRoot = $CloudRoot
        savedAt            = (Get-Date).ToString('s')
        savedFromComputer  = $env:COMPUTERNAME
    } | ConvertTo-Json | Set-Content -LiteralPath $paths.RootsFile -Encoding UTF8
}

function Load-RootSelection {
    $paths = Get-GlobalConfigPaths
    if (-not (Test-Path -LiteralPath $paths.RootsFile)) {
        return $null
    }

    Get-Content -LiteralPath $paths.RootsFile -Raw | ConvertFrom-Json
}

function Test-ProjectUnderRoot {
    param(
        [string]$ProjectRoot,
        [string]$WorkspaceRoot
    )

    $projectFull = [IO.Path]::GetFullPath($ProjectRoot)
    $rootFull = [IO.Path]::GetFullPath($WorkspaceRoot)
    if (-not $rootFull.EndsWith('\')) {
        $rootFull += '\'
    }
    $projectFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RelativeProjectPath {
    param(
        [string]$ProjectRoot,
        [string]$WorkspaceRoot
    )

    $projectUri = [Uri](([IO.Path]::GetFullPath($ProjectRoot)).TrimEnd('\') + '\')
    $rootUri = [Uri](([IO.Path]::GetFullPath($WorkspaceRoot)).TrimEnd('\') + '\')
    $relative = $rootUri.MakeRelativeUri($projectUri).ToString()
    [Uri]::UnescapeDataString($relative).TrimEnd('/') -replace '/', '\'
}

function Resolve-Roots {
    param(
        [string]$ProjectRoot,
        [switch]$ForceConfigure
    )

    $saved = Load-RootSelection
    if (-not $ForceConfigure -and $saved) {
        $localOk = Test-Path -LiteralPath $saved.localWorkspaceRoot
        $cloudOk = Test-Path -LiteralPath $saved.cloudWorkspaceRoot
        $projectFits = $localOk -and (Test-ProjectUnderRoot -ProjectRoot $ProjectRoot -WorkspaceRoot $saved.localWorkspaceRoot)

        if ($localOk -and $cloudOk -and $projectFits) {
            return [pscustomobject]@{
                LocalRoot = (Resolve-Path -LiteralPath $saved.localWorkspaceRoot).Path
                CloudRoot = (Resolve-Path -LiteralPath $saved.cloudWorkspaceRoot).Path
                Source    = 'Saved configuration'
            }
        }
    }

    $candidates = Get-LikelyRoots -ProjectRoot $ProjectRoot
    Write-Section 'First-run root configuration'
    Write-Host "Current project root: $ProjectRoot"
    Write-Host ''

    $localRoot = Select-Root -Label 'local workspace' -Candidates $candidates.Local
    $cloudCandidates = Add-NestedCloudCandidates -CloudCandidates $candidates.Cloud -LocalRoot $localRoot
    $cloudRoot = Select-Root -Label 'cloud workspace' -Candidates $cloudCandidates

    if ([IO.Path]::GetFullPath($localRoot).TrimEnd('\') -ieq [IO.Path]::GetFullPath($cloudRoot).TrimEnd('\')) {
        throw "Local workspace root and cloud workspace root cannot be the same path.`nRoot: $localRoot"
    }

    if (-not (Test-ProjectUnderRoot -ProjectRoot $ProjectRoot -WorkspaceRoot $localRoot)) {
        throw "The current project is not inside the selected local workspace root.`nProject: $ProjectRoot`nLocal root: $localRoot"
    }

    $previewRelative = Get-RelativeProjectPath -ProjectRoot $ProjectRoot -WorkspaceRoot $localRoot
    $previewDestination = Join-Path $cloudRoot $previewRelative
    Write-Host ''
    Write-Host 'Destination preview before saving:' -ForegroundColor Cyan
    Write-Host "  $previewDestination"
    $confirm = Read-Host 'Save these roots? [Y/n]'
    if ($confirm -match '^(n|no)$') {
        return Resolve-Roots -ProjectRoot $ProjectRoot -ForceConfigure
    }

    Save-RootSelection -LocalRoot $localRoot -CloudRoot $cloudRoot

    [pscustomobject]@{
        LocalRoot = $localRoot
        CloudRoot = $cloudRoot
        Source    = 'Newly saved configuration'
    }
}

function Get-MatchingFiles {
    param(
        [string]$ProjectRoot,
        [object]$Config
    )

    $includeSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($ext in $Config.includedExtensions) {
        if ($ext) { [void]$includeSet.Add($ext) }
    }

    $excludeDirSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($dir in $Config.excludedDirectories) {
        if ($dir) { [void]$excludeDirSet.Add($dir) }
    }

    $excludeNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $Config.excludedFileNames) {
        if ($name) { [void]$excludeNameSet.Add($name) }
    }

    $includeNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $Config.includedFileNames) {
        if ($name) { [void]$includeNameSet.Add($name) }
    }

    # Enumerate candidate files once, then filter with readable rules from the JSON config.
    $files = Get-ChildItem -LiteralPath $ProjectRoot -Recurse -Force -File -ErrorAction SilentlyContinue |
        Where-Object {
            $relative = $_.FullName.Substring($ProjectRoot.Length).TrimStart('\')
            $parentRelative = Split-Path -Path $relative -Parent
            $segments = if ($parentRelative) { $parentRelative -split '\\' } else { @() }

            foreach ($segment in $segments) {
                if ($excludeDirSet.Contains($segment)) { return $false }
            }

            foreach ($pattern in $Config.excludedFilePatterns) {
                if ($_.Name -like $pattern) { return $false }
            }

            if ($excludeNameSet.Contains($_.Name)) {
                return $false
            }

            $includeSet.Contains($_.Extension) -or $includeNameSet.Contains($_.Name)
        }

    return $files
}

function Group-FilesByDirectory {
    param([System.IO.FileInfo[]]$Files)

    $groups = @{}
    foreach ($file in $Files) {
        $dir = $file.DirectoryName
        if (-not $groups.ContainsKey($dir)) {
            $groups[$dir] = New-Object System.Collections.Generic.List[string]
        }
        $groups[$dir].Add($file.Name)
    }
    $groups
}

function Get-FilesNeedingCopy {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$SourceRoot,
        [string]$DestinationRoot
    )

    $needed = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationRoot $relativePath
        if (Test-FileNeedsCopy -SourceFile $file -DestinationPath $destinationPath) {
            $needed.Add($file)
        }
    }

    $needed
}

function Get-PlannedCopyRecords {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$SourceRoot,
        [string]$DestinationRoot
    )

    $records = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationRoot $relativePath
        $destinationState = Get-FileState -Path $destinationPath

        $status = if (-not $destinationState.Exists) { 'New File' } else { 'Updated' }
        $records.Add([pscustomobject]@{
            Status       = $status
            RelativePath = $relativePath
        })
    }

    $records
}

function Get-FileState {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Exists            = $false
            Length            = $null
            LastWriteTimeUtc  = $null
        }
    }

    $item = Get-Item -LiteralPath $Path -Force
    [pscustomobject]@{
        Exists            = $true
        Length            = $item.Length
        LastWriteTimeUtc  = $item.LastWriteTimeUtc
    }
}

function Get-IncludedDirectoryFiles {
    param(
        [string]$SourceDir,
        [string]$ProjectRoot,
        [string[]]$ExcludedDirectories
    )

    $excludeDirSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($dir in $ExcludedDirectories) {
        if ($dir) { [void]$excludeDirSet.Add($dir) }
    }

    $records = New-Object System.Collections.Generic.List[object]
    $sourceFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -Force -File -ErrorAction SilentlyContinue
    foreach ($file in $sourceFiles) {
        $relativePath = $file.FullName.Substring($ProjectRoot.Length).TrimStart('\')
        $parentRelative = Split-Path -Path $relativePath -Parent
        $segments = if ($parentRelative) { $parentRelative -split '\\' } else { @() }

        $skip = $false
        foreach ($segment in $segments) {
            if ($excludeDirSet.Contains($segment)) {
                $skip = $true
                break
            }
        }
        if ($skip) { continue }

        $records.Add($file)
    }

    $records
}

function New-DestinationSnapshot {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$SourceRoot,
        [string]$DestinationRoot
    )

    $snapshot = @{}
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationRoot $relativePath
        $snapshot[$relativePath] = Get-FileState -Path $destinationPath
    }
    $snapshot
}

function Test-FileNeedsCopy {
    param(
        [System.IO.FileInfo]$SourceFile,
        [string]$DestinationPath
    )

    $destinationState = Get-FileState -Path $DestinationPath
    if (-not $destinationState.Exists) {
        return $true
    }

    $deltaSeconds = [math]::Abs(($SourceFile.LastWriteTimeUtc - $destinationState.LastWriteTimeUtc).TotalSeconds)
    $sourceIsOlder = $SourceFile.LastWriteTimeUtc -lt $destinationState.LastWriteTimeUtc.AddSeconds(-2)

    if ($sourceIsOlder) {
        return $false
    }

    if ($SourceFile.Length -ne $destinationState.Length) {
        return $true
    }

    if ($deltaSeconds -le 2.0) {
        return $false
    }

    try {
        $sourceHash = (Get-FileHash -LiteralPath $SourceFile.FullName -Algorithm MD5).Hash
        $destinationHash = (Get-FileHash -LiteralPath $DestinationPath -Algorithm MD5).Hash
        return ($sourceHash -ne $destinationHash)
    }
    catch {
        # If the destination cannot be hashed reliably, fall back to copying.
        return $true
    }
}

function Get-ActualCopiedFileRecords {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$SourceRoot,
        [string]$DestinationRoot,
        [hashtable]$BeforeSnapshot
    )

    $records = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Files) {
        $relativePath = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationRoot $relativePath
        $beforeState = $BeforeSnapshot[$relativePath]
        $afterState = Get-FileState -Path $destinationPath

        if (-not $afterState.Exists) { continue }

        $status = $null
        if (-not $beforeState.Exists) {
            $status = 'New File'
        }
        elseif ($beforeState.Length -ne $afterState.Length -or $beforeState.LastWriteTimeUtc -ne $afterState.LastWriteTimeUtc) {
            $status = 'Updated'
        }

        if ($status) {
            $records.Add([pscustomobject]@{
                Status       = $status
                RelativePath = $relativePath
            })
        }
    }

    $records
}

function Invoke-RobocopyBatch {
    param(
        [string]$SourceDir,
        [string]$TargetDir,
        [string[]]$Files,
        [string]$ProjectRoot
    )

    Ensure-Directory -Path $TargetDir

    $argList = @(
        $SourceDir,
        $TargetDir
    ) + $Files + @(
        '/R:1',
        '/W:1',
        '/XO',
        '/FFT',
        '/COPY:DAT',
        '/DCOPY:DAT',
        '/FP',
        '/NP',
        '/NJH',
        '/NJS'
    )

    & robocopy @argList | Out-Null
    $code = $LASTEXITCODE
    if ($code -ge 8) {
        throw "robocopy failed for $SourceDir -> $TargetDir with exit code $code"
    }
}

function Invoke-FileCopyWithRetry {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$MaxAttempts = 3
    )

    $destinationDir = Split-Path -Parent $DestinationPath
    Ensure-Directory -Path $destinationDir

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if (Test-Path -LiteralPath $DestinationPath) {
                $destinationItem = Get-Item -LiteralPath $DestinationPath -Force
                if (($destinationItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                    Remove-Item -LiteralPath $DestinationPath -Force
                }
                elseif ($destinationItem.IsReadOnly) {
                    $destinationItem.IsReadOnly = $false
                }
            }

            Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force

            try {
                $sourceItem = Get-Item -LiteralPath $SourcePath -Force
                $copiedItem = Get-Item -LiteralPath $DestinationPath -Force
                $copiedItem.LastWriteTimeUtc = $sourceItem.LastWriteTimeUtc
            }
            catch {
                # Some cloud-backed files refuse timestamp updates even after content is copied.
            }
            return
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw "Failed to copy $SourcePath to $DestinationPath after $MaxAttempts attempts. $($_.Exception.Message)"
            }
            Start-Sleep -Seconds 1
        }
    }
}

function Copy-IncludedDirectories {
    param(
        [string]$ProjectRoot,
        [string]$DestinationRoot,
        [object]$Config
    )

    $copiedDirCount = 0
    foreach ($dirName in $Config.includedDirectories) {
        # Some directories are always copied whole, even if their files do not match the extension list.
        $sourceDir = Join-Path $ProjectRoot $dirName
        if (-not (Test-Path -LiteralPath $sourceDir)) { continue }

        foreach ($file in (Get-IncludedDirectoryFiles -SourceDir $sourceDir -ProjectRoot $ProjectRoot -ExcludedDirectories $Config.excludedDirectories)) {
            $relativePath = $file.FullName.Substring($ProjectRoot.Length).TrimStart('\')
            $destinationPath = Join-Path $DestinationRoot $relativePath
            if (Test-FileNeedsCopy -SourceFile $file -DestinationPath $destinationPath) {
                Invoke-FileCopyWithRetry -SourcePath $file.FullName -DestinationPath $destinationPath
            }
        }
        $copiedDirCount++
    }

    $copiedDirCount
}

function Write-CopiedFilesReport {
    param([object[]]$CopiedFiles)

    Write-Section 'Files copied this run'
    if (-not $CopiedFiles -or $CopiedFiles.Count -eq 0) {
        Write-Host 'No files needed copying.'
        return
    }

    foreach ($record in $CopiedFiles) {
        Write-Host ("[{0}] {1}" -f $record.Status, $record.RelativePath)
    }
}

function Write-PlannedFilesReport {
    param([object[]]$PlannedFiles)

    Write-Section 'Files queued for copy'
    if (-not $PlannedFiles -or $PlannedFiles.Count -eq 0) {
        Write-Host 'No files need copying.'
        return
    }

    foreach ($record in $PlannedFiles) {
        Write-Host ("[{0}] {1}" -f $record.Status, $record.RelativePath)
    }
}

function Start-Backup {
    param(
        [string]$ProjectRoot,
        [ValidateSet('Incremental', 'Archive', 'Pull')]
        [string]$Mode,
        [switch]$PreviewOnly,
        [switch]$AutoConfirm
    )

    $config = Get-TemplateConfig
    $roots = Resolve-Roots -ProjectRoot $ProjectRoot

    if (Test-ProjectUnderRoot -ProjectRoot $ProjectRoot -WorkspaceRoot $roots.CloudRoot) {
        throw "This script is running from inside the cloud backup copy.`nRun the backup from the local workspace project instead.`nCurrent project: $ProjectRoot"
    }

    $relativeProjectPath = Get-RelativeProjectPath -ProjectRoot $ProjectRoot -WorkspaceRoot $roots.LocalRoot

    if ([string]::IsNullOrWhiteSpace($relativeProjectPath)) {
        throw 'Could not determine relative project path under the local workspace root.'
    }

    $cloudProjectBase =
        if ([string]::IsNullOrWhiteSpace($config.incrementalFolderName)) {
            Join-Path $roots.CloudRoot $relativeProjectPath
        }
        else {
            Join-Path $roots.CloudRoot (Join-Path $config.incrementalFolderName $relativeProjectPath)
        }

    if ($Mode -eq 'Archive') {
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $sourceBase = $ProjectRoot
        $destinationBase = Join-Path $roots.CloudRoot (Join-Path $config.snapshotFolderName (Join-Path $stamp $relativeProjectPath))
    }
    elseif ($Mode -eq 'Pull') {
        $sourceBase = $cloudProjectBase
        $destinationBase = $ProjectRoot
        if (-not (Test-Path -LiteralPath $sourceBase)) {
            throw "Cloud backup folder not found for this project.`nExpected path: $sourceBase"
        }
    }
    else {
        $sourceBase = $ProjectRoot
        $destinationBase = $cloudProjectBase
    }

    if ([IO.Path]::GetFullPath($sourceBase).TrimEnd('\') -ieq [IO.Path]::GetFullPath($destinationBase).TrimEnd('\')) {
        throw "Source and destination resolve to the same project folder.`nSource: $sourceBase`nDestination: $destinationBase"
    }

    Write-Section 'Backup plan'
    Write-Host "Configuration source : $($roots.Source)"
    Write-Host "Local workspace root : $($roots.LocalRoot)"
    Write-Host "Cloud workspace root : $($roots.CloudRoot)"
    Write-Host "Current project root : $ProjectRoot"
    Write-Host "Relative project path: $relativeProjectPath"
    Write-Host "Backup mode          : $Mode"
    Write-Host "Source path          : $sourceBase"
    Write-Host "Destination          : $destinationBase"

    Ensure-Directory -Path $destinationBase

    $files = @(Get-MatchingFiles -ProjectRoot $sourceBase -Config $config)
    $pendingMap = @{}
    foreach ($file in (Get-FilesNeedingCopy -Files $files -SourceRoot $sourceBase -DestinationRoot $destinationBase)) {
        $relativePath = $file.FullName.Substring($sourceBase.Length).TrimStart('\')
        $pendingMap[$relativePath] = $file
    }

    foreach ($dirName in $config.includedDirectories) {
        $sourceDir = Join-Path $sourceBase $dirName
        if (-not (Test-Path -LiteralPath $sourceDir)) { continue }
        foreach ($file in (Get-IncludedDirectoryFiles -SourceDir $sourceDir -ProjectRoot $sourceBase -ExcludedDirectories $config.excludedDirectories)) {
            $relativePath = $file.FullName.Substring($sourceBase.Length).TrimStart('\')
            $destinationPath = Join-Path $destinationBase $relativePath
            if (Test-FileNeedsCopy -SourceFile $file -DestinationPath $destinationPath) {
                $pendingMap[$relativePath] = $file
            }
        }
    }

    $pendingFiles = @($pendingMap.Values)
    $plannedFiles = @(Get-PlannedCopyRecords -Files $pendingFiles -SourceRoot $sourceBase -DestinationRoot $destinationBase)
    $fileGroups = Group-FilesByDirectory -Files $pendingFiles

    Write-Section 'Scanning tracked project files'
    Write-PlannedFilesReport -PlannedFiles $plannedFiles

    if ($PreviewOnly) {
        return [pscustomobject]@{
            FileCount         = $files.Count
            PendingFileCount  = $pendingFiles.Count
            DirectoryBatches  = 0
            IncludedDirCount  = 0
            CopiedFileCount   = 0
            Destination       = $destinationBase
            RelativeProject   = $relativeProjectPath
            LocalRoot         = $roots.LocalRoot
            CloudRoot         = $roots.CloudRoot
            Mode              = $Mode
            Status            = 'PREVIEW'
        }
    }

    if ($pendingFiles.Count -gt 0 -and -not $AutoConfirm) {
        Write-Host ''
        $confirm = Read-Host 'Proceed with copy? [Y/n]'
        if ($confirm -match '^(n|no)$') {
            return [pscustomobject]@{
                FileCount         = $files.Count
                PendingFileCount  = $pendingFiles.Count
                DirectoryBatches  = 0
                IncludedDirCount  = 0
                CopiedFileCount   = 0
                Destination       = $destinationBase
                RelativeProject   = $relativeProjectPath
                LocalRoot         = $roots.LocalRoot
                CloudRoot         = $roots.CloudRoot
                Mode              = $Mode
                Status            = 'CANCELLED'
            }
        }
    }

    $beforeSnapshot = New-DestinationSnapshot -Files $pendingFiles -SourceRoot $sourceBase -DestinationRoot $destinationBase

    $groupCount = 0
    foreach ($sourceDir in ($fileGroups.Keys | Sort-Object)) {
        $relativeDir = $sourceDir.Substring($sourceBase.Length).TrimStart('\')
        $targetDir = if ($relativeDir) { Join-Path $destinationBase $relativeDir } else { $destinationBase }
        Invoke-RobocopyBatch -SourceDir $sourceDir -TargetDir $targetDir -Files ($fileGroups[$sourceDir].ToArray()) -ProjectRoot $sourceBase
        $groupCount++
    }

    $copiedFiles = @(Get-ActualCopiedFileRecords -Files $pendingFiles -SourceRoot $sourceBase -DestinationRoot $destinationBase -BeforeSnapshot $beforeSnapshot)

    Write-CopiedFilesReport -CopiedFiles $copiedFiles

    [pscustomobject]@{
        FileCount         = $files.Count
        PendingFileCount  = $pendingFiles.Count
        DirectoryBatches  = $groupCount
        IncludedDirCount  = 0
        CopiedFileCount   = $copiedFiles.Count
        Destination       = $destinationBase
        RelativeProject   = $relativeProjectPath
        LocalRoot         = $roots.LocalRoot
        CloudRoot         = $roots.CloudRoot
        Mode              = $Mode
        Status            = 'SUCCESS'
    }
}

try {
    $projectRoot = Get-ScriptRoot

    if (Test-IsCloudBackedProjectRoot -ProjectRoot $projectRoot) {
        throw "This script is being run from inside a cloud-backed backup copy.`nRun it from the local workspace project folder instead.`nCurrent project: $projectRoot"
    }

    if ($ConfigureOnly) {
        $result = Resolve-Roots -ProjectRoot $projectRoot -ForceConfigure
        Write-Section 'Configuration saved'
        Write-Host "Local workspace root : $($result.LocalRoot)"
        Write-Host "Cloud workspace root : $($result.CloudRoot)"
        exit 0
    }

    if (-not $Mode) {
        throw 'No backup mode was provided.'
    }

    $summary = Start-Backup -ProjectRoot $projectRoot -Mode $Mode -PreviewOnly:$PreviewOnly -AutoConfirm:$AutoConfirm

    Write-Section 'Backup summary'
    Write-Host "Status              : $($summary.Status)"
    Write-Host "Mode                : $($summary.Mode)"
    Write-Host "Local root          : $($summary.LocalRoot)"
    Write-Host "Cloud root          : $($summary.CloudRoot)"
    Write-Host "Relative project    : $($summary.RelativeProject)"
    Write-Host "Destination         : $($summary.Destination)"
    Write-Host "Included files      : $($summary.FileCount)"
    Write-Host "Pending files       : $($summary.PendingFileCount)"
    Write-Host "Copied this run     : $($summary.CopiedFileCount)"
    Write-Host "Directory batches   : $($summary.DirectoryBatches)"
    Write-Host "Included directories: $($summary.IncludedDirCount)"
    exit 0
}
catch {
    Write-Section 'Backup summary'
    Write-Host 'Status: FAILED' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
