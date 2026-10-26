param(
    [switch]$SelfTest,
    [string]$StartFolder
)

$ErrorActionPreference = 'Stop'

$script:AppName = '批量重命名助手'
$script:AppVersion = '1.0.0'
$script:ScriptPath = $PSCommandPath
$script:ToolRoot = $PSScriptRoot
$script:TempPrefix = '.~batchrename-'

$localAppData = $env:LOCALAPPDATA
if ([string]::IsNullOrWhiteSpace($localAppData)) {
    $localAppData = [System.IO.Path]::GetTempPath()
}
$script:AppDataDir = Join-Path $localAppData 'BatchRenameHelper'
$script:JournalPath = Join-Path $script:AppDataDir 'last-operation.json'

$script:ToolFilePaths = @(
    $script:ScriptPath
    (Join-Path $script:ToolRoot 'BatchRename.ps1')
    (Join-Path $script:ToolRoot '一键批量重命名.bat')
    (Join-Path $script:ToolRoot '使用说明.md')
    (Join-Path $script:ToolRoot '界面预览.png')
)

function Write-AppLog {
    param(
        [string]$Context,
        [System.Exception]$Exception
    )

    try {
        if (-not [System.IO.Directory]::Exists($script:AppDataDir)) {
            [void][System.IO.Directory]::CreateDirectory($script:AppDataDir)
        }
        $logPath = Join-Path $script:AppDataDir 'error.log'
        $text = '[{0}] {1}{2}{3}{2}{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Context, [Environment]::NewLine, $Exception.ToString()
        [System.IO.File]::AppendAllText($logPath, $text, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        # 日志写入失败时不影响主流程。
    }
}

function Get-NaturalSortKey {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    $parts = [regex]::Split($Text, '(\d+)')
    $builder = New-Object System.Text.StringBuilder
    foreach ($part in $parts) {
        if ([string]::IsNullOrEmpty($part)) {
            continue
        }
        if ($part -match '^\d+$') {
            if ($part.Length -lt 30) {
                [void]$builder.Append($part.PadLeft(30, '0'))
            }
            else {
                [void]$builder.Append($part)
            }
        }
        else {
            [void]$builder.Append($part.ToLowerInvariant())
        }
    }
    return $builder.ToString()
}

function Sort-FileList {
    param(
        [object[]]$Files,
        [ValidateSet('Name', 'Modified', 'Created')]
        [string]$SortBy = 'Name',
        [switch]$Descending
    )

    if ($null -eq $Files -or $Files.Count -eq 0) {
        return @()
    }

    $desc = $Descending.IsPresent

    if ($SortBy -eq 'Modified') {
        return @($Files | Sort-Object `
            @{ Expression = { $_.LastWriteTime }; Descending = $desc }, `
            @{ Expression = { Get-NaturalSortKey -Text $_.Name }; Descending = $desc })
    }

    if ($SortBy -eq 'Created') {
        return @($Files | Sort-Object `
            @{ Expression = { $_.CreationTime }; Descending = $desc }, `
            @{ Expression = { Get-NaturalSortKey -Text $_.Name }; Descending = $desc })
    }

    return @($Files | Sort-Object `
        @{ Expression = { Get-NaturalSortKey -Text $_.Name }; Descending = $desc }, `
        @{ Expression = { $_.Name }; Descending = $desc })
}

function Format-NewFileName {
    param(
        [string]$OriginalName,
        [string]$BaseName,
        [string]$Separator,
        [int]$Number,
        [int]$Digits
    )

    $extension = [System.IO.Path]::GetExtension($OriginalName)
    $numberText = $Number.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    if ($Digits -gt 0) {
        $numberText = $numberText.PadLeft($Digits, '0')
    }
    return ($BaseName + $Separator + $numberText + $extension)
}

function Test-NewBaseName {
    param([string]$BaseName)

    $cleanName = ''
    if ($null -ne $BaseName) {
        $cleanName = $BaseName.Trim()
    }

    if ($cleanName.Length -eq 0) {
        return [pscustomobject]@{
            Ok        = $false
            CleanName = ''
            Message   = '请输入新名称。'
        }
    }

    if ($cleanName -match '[<>:"/\\|?*\x00-\x1F]') {
        return [pscustomobject]@{
            Ok        = $false
            CleanName = $cleanName
            Message   = '新名称不能包含下列字符：< > : " / \ | ? *'
        }
    }

    if ($cleanName.Length -gt 180) {
        return [pscustomobject]@{
            Ok        = $false
            CleanName = $cleanName
            Message   = '新名称太长，请控制在 180 个字符以内。'
        }
    }

    return [pscustomobject]@{
        Ok        = $true
        CleanName = $cleanName
        Message   = ''
    }
}

function New-RenamePlan {
    param(
        [object[]]$Entries,
        [string]$Folder
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $operations = New-Object System.Collections.Generic.List[object]
    $targetMap = @{}
    $sourceMap = @{}
    $noChangeCount = 0
    $folderFull = $null

    try {
        if ([string]::IsNullOrWhiteSpace($Folder)) {
            throw '请选择文件夹。'
        }
        if (-not [System.IO.Directory]::Exists($Folder)) {
            throw "文件夹不存在：$Folder"
        }
        $folderFull = [System.IO.Path]::GetFullPath($Folder)
    }
    catch {
        return [pscustomobject]@{
            Ok            = $false
            Errors        = @($_.Exception.Message)
            Operations    = @()
            NoChangeCount = 0
        }
    }

    foreach ($entry in @($Entries)) {
        if ($null -eq $entry -or $null -eq $entry.File) {
            $errors.Add('有一个待处理文件的信息不完整，已跳过。')
            continue
        }

        $file = $entry.File
        $sourcePath = [string]$file.FullName
        $newName = [string]$entry.NewName
        $sourceDir = $null
        try {
            $sourceDir = [System.IO.Path]::GetDirectoryName($sourcePath)
        }
        catch {
            $errors.Add("文件路径无效：$sourcePath")
            continue
        }

        if (-not [string]::Equals($sourceDir, $folderFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $errors.Add("文件不在当前文件夹中：$($file.Name)")
            continue
        }

        if (-not [System.IO.File]::Exists($sourcePath)) {
            $errors.Add("文件不存在或已被移动：$($file.Name)")
            continue
        }

        if ([string]::IsNullOrWhiteSpace($newName) -or
            $newName -match '[<>:"/\\|?*\x00-\x1F]' -or
            $newName -eq '.' -or
            $newName -eq '..') {
            $errors.Add("生成的文件名无效：$newName")
            continue
        }

        if ($newName.Length -gt 255) {
            $errors.Add("文件名过长：$newName")
            continue
        }

        if (($folderFull.Length + 1 + $newName.Length) -gt 255) {
            $errors.Add("完整路径过长，请缩短新名称或选择层级较浅的文件夹：$newName")
            continue
        }

        $sourceKey = $sourcePath.ToLowerInvariant()
        if ($sourceMap.ContainsKey($sourceKey)) {
            $errors.Add("同一个文件被重复加入：$($file.Name)")
            continue
        }
        $sourceMap[$sourceKey] = $true

        $targetPath = [System.IO.Path]::Combine($folderFull, $newName)
        $targetKey = $targetPath.ToLowerInvariant()
        if ($targetMap.ContainsKey($targetKey)) {
            $errors.Add("有两个文件会得到相同的名称：$newName")
        }
        else {
            $targetMap[$targetKey] = $sourcePath
        }

        if ($newName -cne $file.Name) {
            $operations.Add([pscustomobject]@{
                OriginalPath = $sourcePath
                NewPath      = $targetPath
                OriginalName = $file.Name
                NewName      = $newName
            })
        }
        else {
            $noChangeCount++
        }
    }

    $movingSources = @{}
    foreach ($operation in $operations) {
        $movingSources[$operation.OriginalPath.ToLowerInvariant()] = $true
    }

    foreach ($operation in $operations) {
        if ([System.IO.Directory]::Exists($operation.NewPath)) {
            $errors.Add("目标名称已被一个文件夹占用：$($operation.NewName)")
            continue
        }

        if ([System.IO.File]::Exists($operation.NewPath)) {
            $targetKey = $operation.NewPath.ToLowerInvariant()
            if (-not $movingSources.ContainsKey($targetKey)) {
                $errors.Add("目标文件已存在：$($operation.NewName)")
            }
        }
    }

    return [pscustomobject]@{
        Ok            = ($errors.Count -eq 0)
        Errors        = $errors.ToArray()
        Operations    = $operations.ToArray()
        NoChangeCount = $noChangeCount
    }
}

function Invoke-RenameTransaction {
    param([object[]]$Operations)

    if ($null -eq $Operations -or $Operations.Count -eq 0) {
        return
    }

    $states = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($operation in $Operations) {
            $directory = [System.IO.Path]::GetDirectoryName($operation.OriginalPath)
            $tempPath = [System.IO.Path]::Combine(
                $directory,
                ($script:TempPrefix + [guid]::NewGuid().ToString('N') + '.tmp')
            )
            [System.IO.File]::Move($operation.OriginalPath, $tempPath)
            [void]$states.Add([pscustomobject]@{
                Operation = $operation
                TempPath  = $tempPath
                Phase     = 'Staged'
            })
        }

        foreach ($state in $states) {
            [System.IO.File]::Move($state.TempPath, $state.Operation.NewPath)
            $state.Phase = 'Moved'
        }
    }
    catch {
        $failureMessage = $_.Exception.Message
        $rollbackErrors = New-Object System.Collections.Generic.List[string]

        foreach ($state in $states) {
            if ($state.Phase -ne 'Moved') {
                continue
            }

            try {
                $directory = [System.IO.Path]::GetDirectoryName($state.Operation.OriginalPath)
                $rollbackTemp = [System.IO.Path]::Combine(
                    $directory,
                    ($script:TempPrefix + [guid]::NewGuid().ToString('N') + '.tmp')
                )
                [System.IO.File]::Move($state.Operation.NewPath, $rollbackTemp)
                $state.TempPath = $rollbackTemp
                $state.Phase = 'RollbackStaged'
            }
            catch {
                $rollbackErrors.Add("无法暂存新文件 $($state.Operation.NewName)：$($_.Exception.Message)")
            }
        }

        foreach ($state in $states) {
            if ($state.Phase -ne 'Staged' -and $state.Phase -ne 'RollbackStaged') {
                continue
            }

            try {
                if (-not [System.IO.File]::Exists($state.TempPath)) {
                    throw '临时文件不存在'
                }
                if ([System.IO.File]::Exists($state.Operation.OriginalPath)) {
                    throw '原文件名已被占用'
                }
                [System.IO.File]::Move($state.TempPath, $state.Operation.OriginalPath)
            }
            catch {
                $rollbackErrors.Add("无法恢复 $($state.Operation.OriginalName)：$($_.Exception.Message)")
            }
        }

        $message = "重命名失败：$failureMessage"
        if ($rollbackErrors.Count -eq 0) {
            $message += [Environment]::NewLine + [Environment]::NewLine + '已自动恢复原来的文件名。'
        }
        else {
            $message += [Environment]::NewLine + [Environment]::NewLine + '自动恢复未完全成功，请到文件夹中检查：'
            foreach ($item in ($rollbackErrors | Select-Object -First 6)) {
                $message += [Environment]::NewLine + '· ' + $item
            }
            $message += [Environment]::NewLine + [Environment]::NewLine + "以 $($script:TempPrefix) 开头的文件是本次操作的临时文件，请勿直接删除。"
        }

        throw $message
    }
}

function Save-RenameJournal {
    param(
        [string]$JournalPath,
        [string]$Folder,
        [object[]]$Operations
    )

    $parent = [System.IO.Path]::GetDirectoryName($JournalPath)
    if (-not [System.IO.Directory]::Exists($parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($operation in $Operations) {
        [void]$pairs.Add([pscustomobject]@{
            original = $operation.OriginalPath
            renamed  = $operation.NewPath
        })
    }

    $data = [pscustomobject]@{
        version   = 1
        createdAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        folder    = $Folder
        pairs     = $pairs.ToArray()
    }

    $json = $data | ConvertTo-Json -Depth 6
    $tempPath = $JournalPath + '.tmp'
    [System.IO.File]::WriteAllText($tempPath, $json, (New-Object System.Text.UTF8Encoding($true)))

    if ([System.IO.File]::Exists($JournalPath)) {
        [System.IO.File]::Delete($JournalPath)
    }
    [System.IO.File]::Move($tempPath, $JournalPath)
}

function Read-RenameJournal {
    param([string]$JournalPath)

    if (-not [System.IO.File]::Exists($JournalPath)) {
        return [pscustomobject]@{
            Present = $false
            Valid   = $false
            Data    = $null
            Pairs   = @()
            Folder  = $null
            Error   = $null
        }
    }

    try {
        $json = [System.IO.File]::ReadAllText($JournalPath)
        if ([string]::IsNullOrWhiteSpace($json)) {
            throw '撤销记录为空。'
        }

        $data = $json | ConvertFrom-Json
        $pairs = @($data.pairs)
        if ($pairs.Count -eq 0) {
            throw '撤销记录中没有文件。'
        }

        foreach ($pair in $pairs) {
            if ([string]::IsNullOrWhiteSpace([string]$pair.original) -or
                [string]::IsNullOrWhiteSpace([string]$pair.renamed)) {
                throw '撤销记录格式不完整。'
            }
        }

        $folder = [string]$data.folder
        if ([string]::IsNullOrWhiteSpace($folder)) {
            $folder = [System.IO.Path]::GetDirectoryName([string]$pairs[0].original)
        }

        return [pscustomobject]@{
            Present = $true
            Valid   = $true
            Data    = $data
            Pairs   = $pairs
            Folder  = $folder
            Error   = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Present = $true
            Valid   = $false
            Data    = $null
            Pairs   = @()
            Folder  = $null
            Error   = $_.Exception.Message
        }
    }
}

function Remove-RenameJournal {
    param([string]$JournalPath)

    if ([System.IO.File]::Exists($JournalPath)) {
        [System.IO.File]::Delete($JournalPath)
    }
}

function New-UndoPlan {
    param([string]$JournalPath)

    $journal = Read-RenameJournal -JournalPath $JournalPath
    if (-not $journal.Present) {
        return [pscustomobject]@{
            Ok         = $false
            Errors     = @('没有可撤销的记录。')
            Operations = @()
            Journal    = $journal
            Folder     = $null
            PairCount  = 0
        }
    }

    if (-not $journal.Valid) {
        return [pscustomobject]@{
            Ok         = $false
            Errors     = @("撤销记录无法读取：$($journal.Error)")
            Operations = @()
            Journal    = $journal
            Folder     = $null
            PairCount  = 0
        }
    }

    $errors = New-Object System.Collections.Generic.List[string]
    $entries = New-Object System.Collections.Generic.List[object]

    foreach ($pair in $journal.Pairs) {
        $renamed = [string]$pair.renamed
        $original = [string]$pair.original

        if (-not [System.IO.File]::Exists($renamed)) {
            $errors.Add("文件已不在预期位置：$renamed")
            continue
        }

        try {
            $renamedDir = [System.IO.Path]::GetDirectoryName($renamed)
            $originalDir = [System.IO.Path]::GetDirectoryName($original)
        }
        catch {
            $errors.Add("撤销记录中的路径无效：$renamed")
            continue
        }

        if (-not [string]::Equals($renamedDir, $journal.Folder, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($originalDir, $journal.Folder, [System.StringComparison]::OrdinalIgnoreCase)) {
            $errors.Add("撤销记录中的文件不在同一文件夹：$renamed")
            continue
        }

        $originalName = [System.IO.Path]::GetFileName($original)
        [void]$entries.Add([pscustomobject]@{
            File    = (New-Object System.IO.FileInfo -ArgumentList $renamed)
            NewName = $originalName
        })
    }

    if ($errors.Count -gt 0) {
        return [pscustomobject]@{
            Ok         = $false
            Errors     = $errors.ToArray()
            Operations = @()
            Journal    = $journal
            Folder     = $journal.Folder
            PairCount  = $journal.Pairs.Count
        }
    }

    $plan = New-RenamePlan -Entries $entries.ToArray() -Folder $journal.Folder
    return [pscustomobject]@{
        Ok         = $plan.Ok
        Errors     = @($plan.Errors)
        Operations = @($plan.Operations)
        Journal    = $journal
        Folder     = $journal.Folder
        PairCount  = $journal.Pairs.Count
    }
}

function Get-VisibleFileList {
    param([string]$Folder)

    $items = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop)
    $excluded = @{}
    foreach ($path in $script:ToolFilePaths) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }
        try {
            $fullPath = [System.IO.Path]::GetFullPath($path)
            $excluded[$fullPath.ToLowerInvariant()] = $true
        }
        catch {
        }
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($file in $items) {
        if ($file.Name.StartsWith($script:TempPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        if ($excluded.ContainsKey($file.FullName.ToLowerInvariant())) {
            continue
        }
        [void]$result.Add($file)
    }
    return $result.ToArray()
}

function Assert-SelfTest {
    param(
        [bool]$Condition,
        [string]$Name
    )

    if (-not $Condition) {
        throw "ASSERT FAILED: $Name"
    }
}

function Invoke-SelfTest {
    $testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('BatchRenameSelfTest-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($testRoot)
    $success = $false

    try {
        'one' | Set-Content -LiteralPath (Join-Path $testRoot 'photo1.txt') -Encoding ASCII
        'two' | Set-Content -LiteralPath (Join-Path $testRoot 'photo2.txt') -Encoding ASCII
        'ten' | Set-Content -LiteralPath (Join-Path $testRoot 'photo10.txt') -Encoding ASCII

        $photoFiles = @(Get-ChildItem -LiteralPath $testRoot -File | Where-Object { $_.Name -like 'photo*.txt' })
        $sorted = @(Sort-FileList -Files $photoFiles -SortBy 'Name')
        $sortedNames = ($sorted | ForEach-Object { $_.Name }) -join '|'
        Assert-SelfTest ($sortedNames -eq 'photo1.txt|photo2.txt|photo10.txt') 'natural filename sort'

        $descSorted = @(Sort-FileList -Files $photoFiles -SortBy 'Name' -Descending)
        $descNames = ($descSorted | ForEach-Object { $_.Name }) -join '|'
        Assert-SelfTest ($descNames -eq 'photo10.txt|photo2.txt|photo1.txt') 'descending filename sort'

        $entries = New-Object System.Collections.Generic.List[object]
        $index = 0
        foreach ($file in $sorted) {
            $index++
            $newName = Format-NewFileName -OriginalName $file.Name -BaseName 'Demo' -Separator ' ' -Number $index -Digits 1
            [void]$entries.Add([pscustomobject]@{
                File    = $file
                NewName = $newName
            })
        }

        $plan = New-RenamePlan -Entries $entries.ToArray() -Folder $testRoot
        Assert-SelfTest ($plan.Ok -eq $true) ('valid plan: ' + (($plan.Errors) -join '; '))
        Assert-SelfTest ($plan.Operations.Count -eq 3) 'plan operation count'
        Assert-SelfTest ($plan.Operations[0].NewName -eq 'Demo 1.txt') 'generated new name'

        Invoke-RenameTransaction -Operations $plan.Operations
        Assert-SelfTest (Test-Path -LiteralPath (Join-Path $testRoot 'Demo 1.txt')) 'renamed file exists'
        Assert-SelfTest (Test-Path -LiteralPath (Join-Path $testRoot 'Demo 3.txt') -PathType Leaf) 'third target exists'
        Assert-SelfTest (-not (Test-Path -LiteralPath (Join-Path $testRoot 'photo1.txt'))) 'old file moved'

        $journalPath = Join-Path $testRoot 'journal.json'
        Save-RenameJournal -JournalPath $journalPath -Folder $testRoot -Operations $plan.Operations
        $journal = Read-RenameJournal -JournalPath $journalPath
        Assert-SelfTest ($journal.Valid -eq $true) 'journal readable'
        Assert-SelfTest ($journal.Pairs.Count -eq 3) 'journal pair count'

        $undoPlan = New-UndoPlan -JournalPath $journalPath
        Assert-SelfTest ($undoPlan.Ok -eq $true) ('valid undo plan: ' + (($undoPlan.Errors) -join '; '))
        Invoke-RenameTransaction -Operations $undoPlan.Operations
        Assert-SelfTest (Test-Path -LiteralPath (Join-Path $testRoot 'photo1.txt')) 'undo restored original'
        Assert-SelfTest (-not (Test-Path -LiteralPath (Join-Path $testRoot 'Demo 1.txt'))) 'undo removed target'
        Remove-RenameJournal -JournalPath $journalPath
        Assert-SelfTest (-not (Test-Path -LiteralPath $journalPath)) 'journal removed'

        $swapDir = Join-Path $testRoot 'swap'
        [void][System.IO.Directory]::CreateDirectory($swapDir)
        'content-a' | Set-Content -LiteralPath (Join-Path $swapDir 'A.txt') -Encoding ASCII
        'content-b' | Set-Content -LiteralPath (Join-Path $swapDir 'B.txt') -Encoding ASCII
        $swapOps = @(
            [pscustomobject]@{
                OriginalPath = (Join-Path $swapDir 'A.txt')
                NewPath      = (Join-Path $swapDir 'B.txt')
                OriginalName = 'A.txt'
                NewName      = 'B.txt'
            }
            [pscustomobject]@{
                OriginalPath = (Join-Path $swapDir 'B.txt')
                NewPath      = (Join-Path $swapDir 'A.txt')
                OriginalName = 'B.txt'
                NewName      = 'A.txt'
            }
        )
        Invoke-RenameTransaction -Operations $swapOps
        Assert-SelfTest ((Get-Content -LiteralPath (Join-Path $swapDir 'A.txt') -Raw).Trim() -eq 'content-b') 'swap contents A'
        Assert-SelfTest ((Get-Content -LiteralPath (Join-Path $swapDir 'B.txt') -Raw).Trim() -eq 'content-a') 'swap contents B'

        $collisionName = 'Demo 1.txt'
        'external' | Set-Content -LiteralPath (Join-Path $testRoot $collisionName) -Encoding ASCII
        $sourceFile = Get-Item -LiteralPath (Join-Path $testRoot 'photo1.txt')
        $collisionPlan = New-RenamePlan -Entries @(
            [pscustomobject]@{
                File    = $sourceFile
                NewName = $collisionName
            }
        ) -Folder $testRoot
        Assert-SelfTest ($collisionPlan.Ok -eq $false) 'external collision blocked'

        $invalidBase = Test-NewBaseName -BaseName 'bad/name'
        Assert-SelfTest ($invalidBase.Ok -eq $false) 'invalid base name blocked'
        $emptyBase = Test-NewBaseName -BaseName '   '
        Assert-SelfTest ($emptyBase.Ok -eq $false) 'empty base name blocked'

        $paddedName = Format-NewFileName -OriginalName 'photo.JPG' -BaseName 'Demo' -Separator '_' -Number 7 -Digits 3
        Assert-SelfTest ($paddedName -eq 'Demo_007.JPG') 'zero padded name and extension case'
        $noSeparatorName = Format-NewFileName -OriginalName 'readme' -BaseName 'Note' -Separator '' -Number 42 -Digits 1
        Assert-SelfTest ($noSeparatorName -eq 'Note42') 'empty separator and no extension'

        $success = $true
        Write-Host 'SELFTEST PASSED'
    }
    catch {
        Write-Host ('SELFTEST FAILED: ' + $_.Exception.Message)
        Write-Host $_.ScriptStackTrace
    }
    finally {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    return $success
}

if ($SelfTest) {
    $testSucceeded = Invoke-SelfTest
    if ($testSucceeded) {
        exit 0
    }
    exit 1
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

    $bgColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $whiteColor = [System.Drawing.Color]::White
    $headerColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $accentColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $accentHoverColor = [System.Drawing.Color]::FromArgb(29, 78, 216)
    $borderColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $textColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $mutedColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $dangerColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $successColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
    $disabledBackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $disabledTextColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $gridHeaderColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $gridAltColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $gridSelectionColor = [System.Drawing.Color]::FromArgb(239, 246, 255)
    $gridLineColor = [System.Drawing.Color]::FromArgb(226, 232, 240)

    $script:SuppressPreview = $false
    $script:CurrentFolder = ''
    $script:AllFiles = @()

    function New-SecondaryButton {
        param(
            [string]$Text,
            [int]$ButtonWidth = 96
        )

        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Width = $ButtonWidth
        $button.Height = 32
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $borderColor
        $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
        $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
        $button.BackColor = $whiteColor
        $button.ForeColor = $textColor
        $button.UseVisualStyleBackColor = $false
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        return $button
    }

    function New-PrimaryButton {
        param([string]$Text)

        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 0
        $button.FlatAppearance.MouseOverBackColor = $accentHoverColor
        $button.FlatAppearance.MouseDownBackColor = $accentHoverColor
        $button.BackColor = $accentColor
        $button.ForeColor = $whiteColor
        $button.UseVisualStyleBackColor = $false
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $button.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
        return $button
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:AppName)  v$($script:AppVersion)"
    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $initialWidth = [Math]::Min(1000, [Math]::Max(880, $workingArea.Width - 40))
    $initialHeight = [Math]::Min(700, [Math]::Max(540, $workingArea.Height - 40))
    $form.Size = New-Object System.Drawing.Size($initialWidth, $initialHeight)
    $form.MinimumSize = New-Object System.Drawing.Size(880, 540)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor = $bgColor
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.KeyPreview = $true
    $form.Icon = [System.Drawing.SystemIcons]::Application

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = [System.Windows.Forms.DockStyle]::Fill
    $root.ColumnCount = 1
    $root.RowCount = 2
    $root.Margin = New-Object System.Windows.Forms.Padding(0)
    $root.Padding = New-Object System.Windows.Forms.Padding(0)
    $root.BackColor = $bgColor
    [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 74)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $form.Controls.Add($root)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Fill
    $header.Margin = New-Object System.Windows.Forms.Padding(0)
    $header.BackColor = $headerColor

    $accentBar = New-Object System.Windows.Forms.Panel
    $accentBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $accentBar.Height = 4
    $accentBar.BackColor = $accentColor
    $header.Controls.Add($accentBar)

    $headerTitle = New-Object System.Windows.Forms.Label
    $headerTitle.AutoSize = $true
    $headerTitle.Text = $script:AppName
    $headerTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 16, [System.Drawing.FontStyle]::Bold)
    $headerTitle.ForeColor = $whiteColor
    $headerTitle.BackColor = [System.Drawing.Color]::Transparent
    $headerTitle.Location = New-Object System.Drawing.Point(24, 12)
    $header.Controls.Add($headerTitle)

    $headerSubtitle = New-Object System.Windows.Forms.Label
    $headerSubtitle.AutoSize = $true
    $headerSubtitle.Text = '按「名称 + 序号」批量命名  ·  自动保留扩展名  ·  防冲突  ·  可撤销'
    $headerSubtitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $headerSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $headerSubtitle.BackColor = [System.Drawing.Color]::Transparent
    $headerSubtitle.Location = New-Object System.Drawing.Point(26, 45)
    $header.Controls.Add($headerSubtitle)

    [void]$root.Controls.Add($header, 0, 0)

    $content = New-Object System.Windows.Forms.TableLayoutPanel
    $content.Dock = [System.Windows.Forms.DockStyle]::Fill
    $content.ColumnCount = 1
    $content.RowCount = 5
    $content.Margin = New-Object System.Windows.Forms.Padding(0)
    $content.Padding = New-Object System.Windows.Forms.Padding(18, 14, 18, 12)
    $content.BackColor = $bgColor
    [void]$content.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$content.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
    [void]$content.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
    [void]$content.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
    [void]$content.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$content.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))
    [void]$root.Controls.Add($content, 0, 1)

    $folderRow = New-Object System.Windows.Forms.TableLayoutPanel
    $folderRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $folderRow.ColumnCount = 5
    $folderRow.RowCount = 1
    $folderRow.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
    [void]$folderRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
    [void]$folderRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$folderRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 112)))
    [void]$folderRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 104)))
    [void]$folderRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 66)))
    [void]$folderRow.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = '文件夹'
    $folderLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $folderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $folderLabel.ForeColor = $textColor
    $folderLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)
    $folderRow.Controls.Add($folderLabel, 0, 0)

    $folderBox = New-Object System.Windows.Forms.TextBox
    $folderBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $folderBox.ReadOnly = $true
    $folderBox.BackColor = $whiteColor
    $folderBox.ForeColor = $textColor
    $folderBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $folderBox.Margin = New-Object System.Windows.Forms.Padding(0, 7, 8, 7)
    $folderRow.Controls.Add($folderBox, 1, 0)

    $btnBrowse = New-SecondaryButton -Text '选择文件夹…' -ButtonWidth 112
    $btnBrowse.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnBrowse.Margin = New-Object System.Windows.Forms.Padding(0, 7, 6, 7)
    $folderRow.Controls.Add($btnBrowse, 2, 0)

    $btnSelectFiles = New-SecondaryButton -Text '选择文件…' -ButtonWidth 104
    $btnSelectFiles.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnSelectFiles.Margin = New-Object System.Windows.Forms.Padding(0, 7, 6, 7)
    $folderRow.Controls.Add($btnSelectFiles, 3, 0)

    $btnRefresh = New-SecondaryButton -Text '刷新' -ButtonWidth 66
    $btnRefresh.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnRefresh.Margin = New-Object System.Windows.Forms.Padding(0, 7, 0, 7)
    $folderRow.Controls.Add($btnRefresh, 4, 0)

    [void]$content.Controls.Add($folderRow, 0, 0)

    $optionsGroup = New-Object System.Windows.Forms.GroupBox
    $optionsGroup.Text = '命名规则'
    $optionsGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
    $optionsGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
    $optionsGroup.BackColor = $whiteColor
    $optionsGroup.ForeColor = $textColor
    $optionsGroup.Size = New-Object System.Drawing.Size(960, 124)
    [void]$content.Controls.Add($optionsGroup, 0, 1)

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = '新名称'
    $nameLabel.AutoSize = $true
    $nameLabel.BackColor = [System.Drawing.Color]::Transparent
    $nameLabel.ForeColor = $textColor
    $nameLabel.Location = New-Object System.Drawing.Point(14, 32)
    $optionsGroup.Controls.Add($nameLabel)

    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Text = '某某'
    $nameBox.Location = New-Object System.Drawing.Point(62, 27)
    $nameBox.Size = New-Object System.Drawing.Size(240, 28)
    $nameBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $nameBox.BackColor = $whiteColor
    $nameBox.ForeColor = $textColor
    $optionsGroup.Controls.Add($nameBox)

    $startLabel = New-Object System.Windows.Forms.Label
    $startLabel.Text = '起始序号'
    $startLabel.AutoSize = $true
    $startLabel.BackColor = [System.Drawing.Color]::Transparent
    $startLabel.ForeColor = $textColor
    $startLabel.Location = New-Object System.Drawing.Point(322, 32)
    $optionsGroup.Controls.Add($startLabel)

    $startNumberInput = New-Object System.Windows.Forms.NumericUpDown
    $startNumberInput.Minimum = 0
    $startNumberInput.Maximum = 999999
    $startNumberInput.Value = 1
    $startNumberInput.Location = New-Object System.Drawing.Point(388, 27)
    $startNumberInput.Size = New-Object System.Drawing.Size(64, 28)
    $startNumberInput.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right
    $optionsGroup.Controls.Add($startNumberInput)

    $digitsLabel = New-Object System.Windows.Forms.Label
    $digitsLabel.Text = '序号位数'
    $digitsLabel.AutoSize = $true
    $digitsLabel.BackColor = [System.Drawing.Color]::Transparent
    $digitsLabel.ForeColor = $textColor
    $digitsLabel.Location = New-Object System.Drawing.Point(468, 32)
    $optionsGroup.Controls.Add($digitsLabel)

    $digitsInput = New-Object System.Windows.Forms.NumericUpDown
    $digitsInput.Minimum = 1
    $digitsInput.Maximum = 6
    $digitsInput.Value = 1
    $digitsInput.Location = New-Object System.Drawing.Point(530, 27)
    $digitsInput.Size = New-Object System.Drawing.Size(52, 28)
    $digitsInput.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right
    $optionsGroup.Controls.Add($digitsInput)

    $separatorLabel = New-Object System.Windows.Forms.Label
    $separatorLabel.Text = '分隔符'
    $separatorLabel.AutoSize = $true
    $separatorLabel.BackColor = [System.Drawing.Color]::Transparent
    $separatorLabel.ForeColor = $textColor
    $separatorLabel.Location = New-Object System.Drawing.Point(598, 32)
    $optionsGroup.Controls.Add($separatorLabel)

    $separatorCombo = New-Object System.Windows.Forms.ComboBox
    $separatorCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $separatorCombo.Location = New-Object System.Drawing.Point(650, 27)
    $separatorCombo.Size = New-Object System.Drawing.Size(100, 28)
    $separatorCombo.Items.Add('空格') | Out-Null
    $separatorCombo.Items.Add('下划线 (_)') | Out-Null
    $separatorCombo.Items.Add('短横线 (-)') | Out-Null
    $separatorCombo.Items.Add('无') | Out-Null
    $separatorCombo.SelectedIndex = 0
    $optionsGroup.Controls.Add($separatorCombo)

    $sortLabel = New-Object System.Windows.Forms.Label
    $sortLabel.Text = '排序'
    $sortLabel.AutoSize = $true
    $sortLabel.BackColor = [System.Drawing.Color]::Transparent
    $sortLabel.ForeColor = $textColor
    $sortLabel.Location = New-Object System.Drawing.Point(14, 74)
    $optionsGroup.Controls.Add($sortLabel)

    $sortCombo = New-Object System.Windows.Forms.ComboBox
    $sortCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $sortCombo.Location = New-Object System.Drawing.Point(54, 70)
    $sortCombo.Size = New-Object System.Drawing.Size(150, 28)
    $sortCombo.Items.Add('文件名（自然顺序）') | Out-Null
    $sortCombo.Items.Add('修改时间') | Out-Null
    $sortCombo.Items.Add('创建时间') | Out-Null
    $sortCombo.SelectedIndex = 0
    $optionsGroup.Controls.Add($sortCombo)

    $directionLabel = New-Object System.Windows.Forms.Label
    $directionLabel.Text = '顺序'
    $directionLabel.AutoSize = $true
    $directionLabel.BackColor = [System.Drawing.Color]::Transparent
    $directionLabel.ForeColor = $textColor
    $directionLabel.Location = New-Object System.Drawing.Point(218, 74)
    $optionsGroup.Controls.Add($directionLabel)

    $directionCombo = New-Object System.Windows.Forms.ComboBox
    $directionCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $directionCombo.Location = New-Object System.Drawing.Point(254, 70)
    $directionCombo.Size = New-Object System.Drawing.Size(82, 28)
    $directionCombo.Items.Add('升序') | Out-Null
    $directionCombo.Items.Add('降序') | Out-Null
    $directionCombo.SelectedIndex = 0
    $optionsGroup.Controls.Add($directionCombo)

    $formatHint = New-Object System.Windows.Forms.Label
    $formatHint.AutoSize = $false
    $formatHint.AutoEllipsis = $true
    $formatHint.Text = '示例：某某 1.txt'
    $formatHint.Location = New-Object System.Drawing.Point(354, 74)
    $formatHint.Size = New-Object System.Drawing.Size(300, 22)
    $formatHint.ForeColor = $mutedColor
    $formatHint.BackColor = [System.Drawing.Color]::Transparent
    $optionsGroup.Controls.Add($formatHint)

    $btnSelectAll = New-SecondaryButton -Text '全选' -ButtonWidth 72
    $btnSelectAll.Size = New-Object System.Drawing.Size(72, 30)
    $btnSelectAll.Location = New-Object System.Drawing.Point(794, 68)
    $btnSelectAll.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $optionsGroup.Controls.Add($btnSelectAll)

    $btnSelectNone = New-SecondaryButton -Text '全不选' -ButtonWidth 72
    $btnSelectNone.Size = New-Object System.Drawing.Size(72, 30)
    $btnSelectNone.Location = New-Object System.Drawing.Point(874, 68)
    $btnSelectNone.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $optionsGroup.Controls.Add($btnSelectNone)

    $optionsGroup.add_Resize({
        $allX = $optionsGroup.ClientSize.Width - 166
        $btnSelectAll.Location = New-Object System.Drawing.Point($allX, 68)
        $btnSelectNone.Location = New-Object System.Drawing.Point(($allX + 80), 68)
    })

    $previewHeader = New-Object System.Windows.Forms.TableLayoutPanel
    $previewHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
    $previewHeader.ColumnCount = 2
    $previewHeader.RowCount = 1
    $previewHeader.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)
    [void]$previewHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$previewHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$previewHeader.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $previewLabel = New-Object System.Windows.Forms.Label
    $previewLabel.Text = '文件预览'
    $previewLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $previewLabel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $previewLabel.ForeColor = $textColor
    $previewLabel.Margin = New-Object System.Windows.Forms.Padding(2, 0, 0, 0)
    $previewHeader.Controls.Add($previewLabel, 0, 0)

    $dropHint = New-Object System.Windows.Forms.Label
    $dropHint.Text = '可将文件或文件夹直接拖入本窗口'
    $dropHint.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dropHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $dropHint.ForeColor = $mutedColor
    $dropHint.Margin = New-Object System.Windows.Forms.Padding(0)
    $previewHeader.Controls.Add($dropHint, 1, 0)

    [void]$content.Controls.Add($previewHeader, 0, 2)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    $grid.BackgroundColor = $whiteColor
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.GridColor = $gridLineColor
    $grid.RowHeadersVisible = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.AllowUserToOrderColumns = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $grid.ColumnHeadersHeight = 38
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.RowTemplate.Height = 32
    $grid.EnableHeadersVisualStyles = $false
    $grid.ShowCellErrors = $false
    $grid.ShowEditingIcon = $false
    $grid.DefaultCellStyle.BackColor = $whiteColor
    $grid.DefaultCellStyle.ForeColor = $textColor
    $grid.DefaultCellStyle.SelectionBackColor = $gridSelectionColor
    $grid.DefaultCellStyle.SelectionForeColor = $textColor
    $grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $gridAltColor
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $gridHeaderColor
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $gridHeaderColor
    $grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)

    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.Name = 'Selected'
    $checkColumn.HeaderText = '选择'
    $checkColumn.Width = 54
    $checkColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $checkColumn.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    $checkColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    $checkColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $checkColumn.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    [void]$grid.Columns.Add($checkColumn)

    $oldColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $oldColumn.Name = 'OriginalName'
    $oldColumn.HeaderText = '原文件名'
    $oldColumn.ReadOnly = $true
    $oldColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $oldColumn.FillWeight = 48
    $oldColumn.MinimumWidth = 150
    $oldColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    [void]$grid.Columns.Add($oldColumn)

    $newColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $newColumn.Name = 'NewName'
    $newColumn.HeaderText = '新文件名'
    $newColumn.ReadOnly = $true
    $newColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $newColumn.FillWeight = 52
    $newColumn.MinimumWidth = 160
    $newColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    [void]$grid.Columns.Add($newColumn)

    $timeColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $timeColumn.Name = 'Modified'
    $timeColumn.HeaderText = '修改时间'
    $timeColumn.ReadOnly = $true
    $timeColumn.Width = 138
    $timeColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $timeColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    $timeColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $timeColumn.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    [void]$grid.Columns.Add($timeColumn)

    [void]$content.Controls.Add($grid, 0, 3)

    $footer = New-Object System.Windows.Forms.TableLayoutPanel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footer.ColumnCount = 3
    $footer.RowCount = 1
    $footer.Margin = New-Object System.Windows.Forms.Padding(0)
    [void]$footer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$footer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 148)))
    [void]$footer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))
    [void]$footer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $statusLabel.AutoEllipsis = $true
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $statusLabel.Text = '请选择文件夹，或把文件 / 文件夹拖入窗口。'
    $statusLabel.ForeColor = $mutedColor
    $statusLabel.Margin = New-Object System.Windows.Forms.Padding(2, 0, 8, 0)
    $footer.Controls.Add($statusLabel, 0, 0)

    $btnUndo = New-SecondaryButton -Text '撤销上次' -ButtonWidth 138
    $btnUndo.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnUndo.Margin = New-Object System.Windows.Forms.Padding(0, 8, 10, 8)
    $btnUndo.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)
    $btnUndo.Enabled = $false
    $footer.Controls.Add($btnUndo, 1, 0)

    $btnRename = New-PrimaryButton -Text '开始重命名'
    $btnRename.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnRename.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $footer.Controls.Add($btnRename, 2, 0)

    [void]$content.Controls.Add($footer, 0, 4)

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 12000
    $toolTip.InitialDelay = 350
    $toolTip.ReshowDelay = 100
    $toolTip.ShowAlways = $true
    $toolTip.SetToolTip($folderBox, '也可以直接把文件或文件夹拖入窗口。')
    $toolTip.SetToolTip($startNumberInput, '第一个勾选文件的序号。')
    $toolTip.SetToolTip($digitsInput, '填 1 时生成 1、2、3；填 2 时生成 01、02、03。')
    $toolTip.SetToolTip($sortCombo, '编号会严格按照下方列表的顺序生成。')
    $toolTip.SetToolTip($btnSelectAll, '勾选列表中的全部文件。编号只分配给已勾选的文件。')
    $toolTip.SetToolTip($btnSelectNone, '取消勾选全部文件，不会改动任何文件。')
    $toolTip.SetToolTip($grid, '取消勾选的文件不会被重命名；序号会重新连续编号。')
    $toolTip.SetToolTip($btnUndo, '暂时没有可撤销的记录。')
    $toolTip.SetToolTip($btnRename, '所有文件都先检查冲突，确认后才会执行。')

    function Get-SelectedSeparator {
        $selected = [string]$separatorCombo.SelectedItem
        switch ($selected) {
            '下划线 (_)' { return '_' }
            '短横线 (-)' { return '-' }
            '无' { return '' }
            default { return ' ' }
        }
    }

    function Get-SortSettings {
        $sortBy = 'Name'
        switch ([string]$sortCombo.SelectedItem) {
            '修改时间' { $sortBy = 'Modified' }
            '创建时间' { $sortBy = 'Created' }
        }
        return [pscustomobject]@{
            SortBy     = $sortBy
            Descending = ([string]$directionCombo.SelectedItem -eq '降序')
        }
    }

    function Show-AppMessage {
        param(
            [string]$Message,
            [string]$Title = '提示',
            [switch]$IsError
        )

        $icon = [System.Windows.Forms.MessageBoxIcon]::Information
        if ($IsError) {
            $icon = [System.Windows.Forms.MessageBoxIcon]::Error
        }
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $icon
        )
    }

    function Set-StatusMessage {
        param(
            [string]$Message,
            [bool]$IsError = $false
        )

        $statusLabel.Text = $Message
        if ($IsError) {
            $statusLabel.ForeColor = $dangerColor
        }
        else {
            $statusLabel.ForeColor = $mutedColor
        }
        $toolTip.SetToolTip($statusLabel, $Message)
    }

    function Set-RenameButtonEnabled {
        param([bool]$Enabled)

        $btnRename.Enabled = $Enabled
        if ($Enabled) {
            $btnRename.BackColor = $accentColor
            $btnRename.ForeColor = $whiteColor
        }
        else {
            $btnRename.BackColor = $disabledBackColor
            $btnRename.ForeColor = $disabledTextColor
        }
    }

    function Get-CheckedRows {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) {
                continue
            }
            if ([bool]$row.Cells[0].Value) {
                [void]$rows.Add($row)
            }
        }
        return $rows.ToArray()
    }

    function Get-CheckedPathSet {
        $set = @{}
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow -or -not [bool]$row.Cells[0].Value) {
                continue
            }
            $file = $row.Tag
            if ($null -ne $file) {
                $set[$file.FullName.ToLowerInvariant()] = $true
            }
        }
        return $set
    }

    function Update-Preview {
        if ($script:SuppressPreview) {
            return
        }

        $baseCheck = Test-NewBaseName -BaseName $nameBox.Text
        $separator = Get-SelectedSeparator
        $startNumber = [int]$startNumberInput.Value
        $digits = [int]$digitsInput.Value
        $checkedCount = 0
        $index = 0
        $firstName = $null
        $lastName = $null
        $exampleExtension = ''

        $script:SuppressPreview = $true
        try {
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) {
                    continue
                }

                $file = $row.Tag
                if ($null -eq $file) {
                    continue
                }

                $isChecked = [bool]$row.Cells[0].Value
                if (-not $isChecked) {
                    $row.Cells[2].Value = '—'
                    continue
                }

                $checkedCount++
                if ($exampleExtension.Length -eq 0) {
                    $exampleExtension = [System.IO.Path]::GetExtension($file.Name)
                }

                if ($baseCheck.Ok) {
                    $newName = Format-NewFileName `
                        -OriginalName $file.Name `
                        -BaseName $baseCheck.CleanName `
                        -Separator $separator `
                        -Number ($startNumber + $index) `
                        -Digits $digits
                    $row.Cells[2].Value = $newName
                    if ($null -eq $firstName) {
                        $firstName = $newName
                    }
                    $lastName = $newName
                }
                else {
                    $row.Cells[2].Value = ''
                }
                $index++
            }
        }
        finally {
            $script:SuppressPreview = $false
        }

        $sampleBase = $baseCheck.CleanName
        if (-not $baseCheck.Ok) {
            $sampleBase = '某某'
        }
        if ($exampleExtension.Length -eq 0) {
            $exampleExtension = '.txt'
        }
        $sampleNumber = ([int]$startNumber).ToString().PadLeft($digits, '0')
        $formatHint.Text = '示例：' + $sampleBase + $separator + $sampleNumber + $exampleExtension

        if ($grid.Rows.Count -eq 0) {
            if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
                Set-StatusMessage '请选择文件夹，或把文件 / 文件夹拖入窗口。'
            }
            else {
                Set-StatusMessage '这个文件夹中没有可处理的文件（本工具自身的文件会被自动排除）。' $true
            }
            $btnRename.Text = '开始重命名'
            Set-RenameButtonEnabled $false
            return
        }

        if (-not $baseCheck.Ok) {
            Set-StatusMessage $baseCheck.Message $true
            $btnRename.Text = '开始重命名'
            Set-RenameButtonEnabled $false
            return
        }

        if ($checkedCount -eq 0) {
            Set-StatusMessage '请至少勾选一个文件。'
            $btnRename.Text = '开始重命名'
            Set-RenameButtonEnabled $false
            return
        }

        $btnRename.Text = "开始重命名 ($checkedCount)"
        Set-RenameButtonEnabled $true

        if ($checkedCount -eq 1) {
            Set-StatusMessage "已勾选 1 个文件 · 新名称：「$firstName」"
        }
        else {
            Set-StatusMessage "已勾选 $checkedCount 个文件 · 新名称从 「$firstName」 到 「$lastName」"
        }
    }

    function Get-EntriesFromGrid {
        param(
            [string]$BaseName,
            [string]$Separator,
            [int]$StartNumber,
            [int]$Digits
        )

        $entries = New-Object System.Collections.Generic.List[object]
        $index = 0
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow -or -not [bool]$row.Cells[0].Value) {
                continue
            }
            $file = $row.Tag
            if ($null -eq $file) {
                continue
            }
            $newName = Format-NewFileName `
                -OriginalName $file.Name `
                -BaseName $BaseName `
                -Separator $Separator `
                -Number ($StartNumber + $index) `
                -Digits $Digits
            [void]$entries.Add([pscustomobject]@{
                File    = $file
                NewName = $newName
            })
            $index++
        }
        return $entries.ToArray()
    }

    function Show-FileGrid {
        param(
            [object[]]$Files,
            [hashtable]$CheckedPaths
        )

        $script:SuppressPreview = $true
        $grid.SuspendLayout()
        try {
            $grid.Rows.Clear()
            foreach ($file in @($Files)) {
                $rowIndex = $grid.Rows.Add()
                $row = $grid.Rows[$rowIndex]
                $isChecked = $true
                if ($null -ne $CheckedPaths) {
                    $isChecked = $CheckedPaths.ContainsKey($file.FullName.ToLowerInvariant())
                }
                $row.Cells[0].Value = $isChecked
                $row.Cells[1].Value = $file.Name
                $row.Cells[2].Value = '—'
                $row.Cells[3].Value = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                $row.Tag = $file
            }
            $grid.ClearSelection()
        }
        finally {
            $grid.ResumeLayout()
            $script:SuppressPreview = $false
        }

        Update-Preview
    }

    function Refresh-GridFromMemory {
        param([hashtable]$CheckedPaths)

        if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
            Update-Preview
            return
        }

        $settings = Get-SortSettings
        $sortedFiles = @(Sort-FileList `
            -Files $script:AllFiles `
            -SortBy $settings.SortBy `
            -Descending:$settings.Descending)
        Show-FileGrid -Files $sortedFiles -CheckedPaths $CheckedPaths
    }

    function Load-Folder {
        param(
            [string]$Folder,
            [hashtable]$CheckedPaths,
            [string[]]$OnlyFiles
        )

        try {
            if ([string]::IsNullOrWhiteSpace($Folder)) {
                throw '请选择文件夹。'
            }
            if (-not [System.IO.Directory]::Exists($Folder)) {
                throw "文件夹不存在：$Folder"
            }

            $fullFolder = [System.IO.Path]::GetFullPath($Folder)
            $files = @(Get-VisibleFileList -Folder $fullFolder)
            $checkSet = @{}

            if ($null -ne $OnlyFiles -and $OnlyFiles.Count -gt 0) {
                $wanted = @{}
                foreach ($path in $OnlyFiles) {
                    try {
                        $fullPath = [System.IO.Path]::GetFullPath($path)
                        $wanted[$fullPath.ToLowerInvariant()] = $true
                    }
                    catch {
                    }
                }
                foreach ($file in $files) {
                    if ($wanted.ContainsKey($file.FullName.ToLowerInvariant())) {
                        $checkSet[$file.FullName.ToLowerInvariant()] = $true
                    }
                }
                if ($checkSet.Count -eq 0) {
                    throw '选中的文件不在这个文件夹中，或已被本工具自动排除。'
                }
            }
            elseif ($null -ne $CheckedPaths) {
                foreach ($file in $files) {
                    if ($CheckedPaths.ContainsKey($file.FullName.ToLowerInvariant())) {
                        $checkSet[$file.FullName.ToLowerInvariant()] = $true
                    }
                }
            }
            else {
                foreach ($file in $files) {
                    $checkSet[$file.FullName.ToLowerInvariant()] = $true
                }
            }

            $script:CurrentFolder = $fullFolder
            $script:AllFiles = $files
            $folderBox.Text = $fullFolder

            $settings = Get-SortSettings
            $sortedFiles = @(Sort-FileList `
                -Files $files `
                -SortBy $settings.SortBy `
                -Descending:$settings.Descending)
            Show-FileGrid -Files $sortedFiles -CheckedPaths $checkSet

            if ($files.Count -eq 0) {
                Set-StatusMessage '这个文件夹中没有可处理的文件（本工具自身的文件会被自动排除）。' $true
            }

            $nameBox.Focus()
            $nameBox.SelectAll()
            return $true
        }
        catch {
            $message = $_.Exception.Message
            Show-AppMessage "无法打开文件夹：`r`n$message" '打开失败' -IsError
            Set-StatusMessage $message $true
            return $false
        }
    }

    function Update-UndoButton {
        $journal = Read-RenameJournal -JournalPath $script:JournalPath
        if (-not $journal.Present) {
            $btnUndo.Text = '撤销上次'
            $btnUndo.Enabled = $false
            $toolTip.SetToolTip($btnUndo, '暂时没有可撤销的记录。')
            return
        }

        if ($journal.Valid) {
            $count = $journal.Pairs.Count
            $btnUndo.Text = "撤销上次 ($count)"
            $btnUndo.Enabled = $true
            $toolTip.SetToolTip($btnUndo, "撤销最近一次重命名（$($journal.Data.createdAt)）。")
        }
        else {
            $btnUndo.Text = '撤销上次'
            $btnUndo.Enabled = $true
            $toolTip.SetToolTip($btnUndo, "撤销记录无法读取：$($journal.Error)")
        }
    }

    function Set-AllCheckboxes {
        param([bool]$IsChecked)

        $grid.EndEdit()
        $script:SuppressPreview = $true
        try {
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) {
                    continue
                }
                $row.Cells[0].Value = $IsChecked
            }
        }
        finally {
            $script:SuppressPreview = $false
        }
        Update-Preview
    }

    Update-Preview
    Update-UndoButton

    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    $grid.add_CellValueChanged({
        param($sender, $e)
        if ($script:SuppressPreview) {
            return
        }
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) {
            Update-Preview
        }
    })

    $grid.add_CellContentClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) {
            [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    $nameBox.add_TextChanged({
        Update-Preview
    })

    $startNumberInput.add_ValueChanged({
        Update-Preview
    })

    $digitsInput.add_ValueChanged({
        Update-Preview
    })

    $separatorCombo.add_SelectedIndexChanged({
        Update-Preview
    })

    $sortChangedHandler = {
        try {
            if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
                return
            }
            $checked = Get-CheckedPathSet
            Refresh-GridFromMemory -CheckedPaths $checked
        }
        catch {
            Write-AppLog -Context '调整排序失败' -Exception $_.Exception
            Show-AppMessage "调整排序失败：`r`n$($_.Exception.Message)" '操作失败' -IsError
        }
    }
    $sortCombo.add_SelectedIndexChanged($sortChangedHandler)
    $directionCombo.add_SelectedIndexChanged($sortChangedHandler)

    $btnSelectAll.add_Click({
        Set-AllCheckboxes $true
    })

    $btnSelectNone.add_Click({
        Set-AllCheckboxes $false
    })

    $btnBrowse.add_Click({
        $dialog = $null
        try {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = '选择需要批量重命名的文件夹'
            $dialog.ShowNewFolderButton = $false
            if (-not [string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
                $dialog.SelectedPath = $script:CurrentFolder
            }
            if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                [void](Load-Folder -Folder $dialog.SelectedPath)
            }
        }
        catch {
            Write-AppLog -Context '选择文件夹失败' -Exception $_.Exception
            Show-AppMessage "选择文件夹失败：`r`n$($_.Exception.Message)" '打开失败' -IsError
        }
        finally {
            if ($null -ne $dialog) {
                $dialog.Dispose()
            }
        }
    })

    $btnSelectFiles.add_Click({
        $dialog = $null
        try {
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = '选择要批量重命名的文件'
            $dialog.Multiselect = $true
            $dialog.CheckFileExists = $true
            $dialog.Filter = '所有文件 (*.*)|*.*'
            $dialog.RestoreDirectory = $true
            if (-not [string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
                $dialog.InitialDirectory = $script:CurrentFolder
            }

            if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $paths = @($dialog.FileNames)
                if ($paths.Count -eq 0) {
                    return
                }

                $parentMap = @{}
                foreach ($path in $paths) {
                    $parent = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($path))
                    $parentMap[$parent.ToLowerInvariant()] = $parent
                }
                if ($parentMap.Count -ne 1) {
                    Show-AppMessage '请选择同一个文件夹中的文件。' '无法处理' -IsError
                    return
                }

                $folder = @($parentMap.Values)[0]
                [void](Load-Folder -Folder $folder -OnlyFiles $paths)
            }
        }
        catch {
            Write-AppLog -Context '选择文件失败' -Exception $_.Exception
            Show-AppMessage "选择文件失败：`r`n$($_.Exception.Message)" '打开失败' -IsError
        }
        finally {
            if ($null -ne $dialog) {
                $dialog.Dispose()
            }
        }
    })

    $btnRefresh.add_Click({
        if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
            Show-AppMessage '请先选择一个文件夹。' '提示'
            return
        }
        $checked = Get-CheckedPathSet
        [void](Load-Folder -Folder $script:CurrentFolder -CheckedPaths $checked)
    })

    $btnRename.add_Click({
        try {
            $checkedRows = @(Get-CheckedRows)
            if ($checkedRows.Count -eq 0) {
                Show-AppMessage '请先勾选需要重命名的文件。' '提示'
                return
            }
            if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
                Show-AppMessage '请先选择一个文件夹。' '提示'
                return
            }

            $baseCheck = Test-NewBaseName -BaseName $nameBox.Text
            if (-not $baseCheck.Ok) {
                Show-AppMessage $baseCheck.Message '新名称无效' -IsError
                $nameBox.Focus()
                return
            }

            $separator = Get-SelectedSeparator
            $startNumber = [int]$startNumberInput.Value
            $digits = [int]$digitsInput.Value
            $entries = @(Get-EntriesFromGrid `
                -BaseName $baseCheck.CleanName `
                -Separator $separator `
                -StartNumber $startNumber `
                -Digits $digits)

            if ($entries.Count -eq 0) {
                Show-AppMessage '请先勾选需要重命名的文件。' '提示'
                return
            }

            $plan = New-RenamePlan -Entries $entries -Folder $script:CurrentFolder
            if (-not $plan.Ok) {
                $allErrors = @($plan.Errors)
                $details = ($allErrors | Select-Object -First 8) -join [Environment]::NewLine
                if ($allErrors.Count -gt 8) {
                    $details += [Environment]::NewLine + "……另有 $($allErrors.Count - 8) 条错误未显示。"
                }
                Show-AppMessage ("发现名称冲突，本次没有修改任何文件：`r`n`r`n$details") '无法重命名' -IsError
                Set-StatusMessage '未执行重命名，请先解决名称冲突。' $true
                return
            }

            if ($plan.Operations.Count -eq 0) {
                Show-AppMessage '勾选的文件已经符合当前命名规则，无需重命名。' '无需处理'
                return
            }

            $count = $plan.Operations.Count
            $firstOperation = $plan.Operations[0]
            $lastOperation = $plan.Operations[$count - 1]
            $previewText = "第一个：$($firstOperation.OriginalName)  →  $($firstOperation.NewName)"
            if ($count -gt 1) {
                $previewText += "`r`n最后一个：$($lastOperation.OriginalName)  →  $($lastOperation.NewName)"
            }

            $question = "即将重命名 $count 个文件，扩展名将保持不变。`r`n`r`n$previewText`r`n`r`n原文件名会被记录，可在完成后撤销。是否继续？"
            $answer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                $question,
                '确认重命名',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }

            $successText = $null
            $errorText = $null
            $journalProblem = $null

            $btnRename.Enabled = $false
            $btnUndo.Enabled = $false
            $form.UseWaitCursor = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Invoke-RenameTransaction -Operations $plan.Operations

                try {
                    Save-RenameJournal `
                        -JournalPath $script:JournalPath `
                        -Folder $script:CurrentFolder `
                        -Operations $plan.Operations
                }
                catch {
                    $journalProblem = $_.Exception.Message
                    Write-AppLog -Context '保存撤销记录失败' -Exception $_.Exception
                    try {
                        Remove-RenameJournal -JournalPath $script:JournalPath
                    }
                    catch {
                    }
                }

                [void](Load-Folder -Folder $script:CurrentFolder)
                $successText = "已完成 $count 个文件的重命名。"
                if ($null -eq $journalProblem) {
                    $successText += "`r`n`r`n如需恢复，请点击「撤销上次」。"
                }
                else {
                    $successText += "`r`n`r`n注意：撤销记录保存失败，本次无法使用「撤销上次」。`r`n$journalProblem"
                }
            }
            catch {
                $errorText = $_.Exception.Message
                Write-AppLog -Context '执行重命名失败' -Exception $_.Exception
                try {
                    [void](Load-Folder -Folder $script:CurrentFolder)
                }
                catch {
                }
            }
            finally {
                $form.UseWaitCursor = $false
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
                Update-Preview
                Update-UndoButton
            }

            if ($null -ne $errorText) {
                Show-AppMessage $errorText '重命名未完成' -IsError
            }
            elseif ($null -ne $successText) {
                Show-AppMessage $successText '重命名完成'
            }
        }
        catch {
            Write-AppLog -Context '重命名操作异常' -Exception $_.Exception
            Show-AppMessage "操作失败：`r`n$($_.Exception.Message)" '操作失败' -IsError
            Update-Preview
            Update-UndoButton
        }
    })

    $btnUndo.add_Click({
        try {
            $journal = Read-RenameJournal -JournalPath $script:JournalPath
            if (-not $journal.Present) {
                Show-AppMessage '目前没有可撤销的记录。' '提示'
                Update-UndoButton
                return
            }
            if (-not $journal.Valid) {
                Show-AppMessage "撤销记录无法读取：`r`n$($journal.Error)" '无法撤销' -IsError
                return
            }

            $undoPlan = New-UndoPlan -JournalPath $script:JournalPath
            if (-not $undoPlan.Ok) {
                $allErrors = @($undoPlan.Errors)
                $details = ($allErrors | Select-Object -First 8) -join [Environment]::NewLine
                if ($allErrors.Count -gt 8) {
                    $details += [Environment]::NewLine + "……另有 $($allErrors.Count - 8) 条错误未显示。"
                }
                Show-AppMessage ("无法撤销：`r`n`r`n$details`r`n`r`n撤销记录仍保留，请检查文件是否被移动或重命名。") '无法撤销' -IsError
                return
            }

            if ($undoPlan.Operations.Count -eq 0) {
                Show-AppMessage '这些文件已经恢复原来的名称，无需撤销。' '无需处理'
                return
            }

            $count = $undoPlan.Operations.Count
            $question = "将撤销上一次对 $count 个文件的重命名，并恢复原来的文件名。`r`n`r`n文件夹：$($undoPlan.Folder)`r`n`r`n是否继续？"
            $answer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                $question,
                '确认撤销',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }

            $successText = $null
            $errorText = $null
            $journalProblem = $null

            $btnRename.Enabled = $false
            $btnUndo.Enabled = $false
            $form.UseWaitCursor = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Invoke-RenameTransaction -Operations $undoPlan.Operations

                try {
                    Remove-RenameJournal -JournalPath $script:JournalPath
                }
                catch {
                    $journalProblem = $_.Exception.Message
                    Write-AppLog -Context '删除撤销记录失败' -Exception $_.Exception
                }

                [void](Load-Folder -Folder $undoPlan.Folder)
                $successText = "已恢复 $count 个文件的原文件名。"
                if ($null -ne $journalProblem) {
                    $successText += "`r`n`r`n注意：撤销记录未能清除。`r`n$journalProblem"
                }
            }
            catch {
                $errorText = $_.Exception.Message
                Write-AppLog -Context '撤销失败' -Exception $_.Exception
                try {
                    if (-not [string]::IsNullOrWhiteSpace($undoPlan.Folder)) {
                        [void](Load-Folder -Folder $undoPlan.Folder)
                    }
                }
                catch {
                }
            }
            finally {
                $form.UseWaitCursor = $false
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
                Update-Preview
                Update-UndoButton
            }

            if ($null -ne $errorText) {
                Show-AppMessage $errorText '撤销未完成' -IsError
            }
            elseif ($null -ne $successText) {
                Show-AppMessage $successText '撤销完成'
            }
        }
        catch {
            Write-AppLog -Context '撤销操作异常' -Exception $_.Exception
            Show-AppMessage "撤销失败：`r`n$($_.Exception.Message)" '撤销失败' -IsError
            Update-UndoButton
        }
    })

    $dragEnterHandler = {
        param($sender, $eventArgs)
        if ($eventArgs.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $eventArgs.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
        else {
            $eventArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    }

    $dragDropHandler = {
        param($sender, $eventArgs)
        try {
            $rawData = $eventArgs.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($null -eq $rawData) {
                return
            }
            $paths = @([string[]]$rawData)
            $validPaths = @($paths | Where-Object { Test-Path -LiteralPath $_ })
            if ($validPaths.Count -eq 0) {
                Show-AppMessage '没有识别到可用的文件或文件夹。' '无法处理' -IsError
                return
            }

            $folderPaths = @($validPaths | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
            if ($folderPaths.Count -gt 1 -or ($folderPaths.Count -eq 1 -and $validPaths.Count -gt 1)) {
                Show-AppMessage '一次只能拖入一个文件夹；也可以拖入同一文件夹中的多个文件。' '无法处理' -IsError
                return
            }

            if ($folderPaths.Count -eq 1) {
                [void](Load-Folder -Folder $folderPaths[0])
                return
            }

            $parentMap = @{}
            foreach ($path in $validPaths) {
                $parent = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($path))
                $parentMap[$parent.ToLowerInvariant()] = $parent
            }
            if ($parentMap.Count -ne 1) {
                Show-AppMessage '请拖入同一个文件夹中的文件，或直接拖入一个文件夹。' '无法处理' -IsError
                return
            }

            $folder = @($parentMap.Values)[0]
            [void](Load-Folder -Folder $folder -OnlyFiles $validPaths)
        }
        catch {
            Write-AppLog -Context '拖放处理失败' -Exception $_.Exception
            Show-AppMessage "拖入文件失败：`r`n$($_.Exception.Message)" '无法处理' -IsError
        }
    }

    function Enable-FileDropRecursively {
        param([System.Windows.Forms.Control]$Control)

        try {
            $Control.AllowDrop = $true
            $Control.add_DragEnter($dragEnterHandler)
            $Control.add_DragDrop($dragDropHandler)
        }
        catch {
        }

        foreach ($child in $Control.Controls) {
            Enable-FileDropRecursively -Control $child
        }
    }

    Enable-FileDropRecursively -Control $form

    if (-not [string]::IsNullOrWhiteSpace($StartFolder) -and
        [System.IO.Directory]::Exists($StartFolder)) {
        try {
            [void](Load-Folder -Folder $StartFolder)
        }
        catch {
            Write-AppLog -Context '加载启动文件夹失败' -Exception $_.Exception
        }
    }

    $form.add_Shown({
        Update-Preview
        Update-UndoButton
        if ([string]::IsNullOrWhiteSpace($script:CurrentFolder)) {
            $btnBrowse.Focus()
        }
        else {
            $nameBox.Focus()
            $nameBox.SelectAll()
        }
    })

    $form.add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $btnRefresh.PerformClick()
        }
    })

    [System.Windows.Forms.Application]::Run($form)

    $toolTip.Dispose()
    $form.Dispose()
}
catch {
    Write-AppLog -Context '界面启动失败' -Exception $_.Exception
    try {
        [void][System.Windows.Forms.MessageBox]::Show(
            "程序启动失败：`r`n$($_.Exception.Message)`r`n`r`n日志位置：$($script:AppDataDir)\error.log",
            $script:AppName,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    catch {
    }
    exit 1
}
