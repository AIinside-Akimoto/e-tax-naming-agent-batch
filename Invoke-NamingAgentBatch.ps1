#requires -Version 5.1

<#
.SYNOPSIS
電子帳簿保存法向けの命名支援 API を呼び出し、抽出結果に基づいてファイルをリネームしてコピーします。

.DESCRIPTION
InputFolder 配下の .pdf / .txt / .csv ファイルを再帰的に走査し、各ファイルを命名支援 API に送信します。
API が返す documents[0].new_filename を使用して OutputFolder 配下へコピーします。

OutputFolder には InputFolder と同じサブフォルダー構造を再現します。
rename=false または new_filename が空の場合は、CopyNonRenamedFiles の指定に応じて
元の名前のままコピーするか、コピーせずにスキップします。

OrganizeSourceFilesAfterCopy を指定した場合は、以下を実施します。
- コピーできたファイル: 元の InputFolder 側ファイルを削除
- コピーできなかったファイル: InputFolder の親フォルダー配下の
  「対象外」（または ExcludedFolderName で指定した名前）へ
  サブフォルダー構造を維持して移動

API 呼び出しは Parallelism で指定した並列数で実行します。
既定値は 3 です。

ログには API の生レスポンス JSON と、抽出結果・コピー結果・移動結果・エラー情報を記録します。

.PARAMETER InputFolder
処理対象の入力フォルダーです。
このフォルダー配下のサブフォルダーも再帰的に処理します。

.PARAMETER OutputFolder
処理結果の出力先フォルダーです。
InputFolder と同じフォルダー構造でファイルを保存します。

.PARAMETER ApiBaseUrl
命名支援 API のベース URL です。
実際の呼び出し先は <ApiBaseUrl>/assist-naming です。

.PARAMETER ApiKey
API 呼び出しに使用する x-api-key ヘッダーの値です。

.PARAMETER Timeout
API 呼び出しのタイムアウト秒です。
既定値は 600 秒です。

.PARAMETER Parallelism
API 呼び出しの並列実行数です。
既定値は 3 です。

.PARAMETER CopyNonRenamedFiles
API が rename=false を返した場合でも、元のファイル名のまま OutputFolder にコピーします。
省略時はコピーしません。

.PARAMETER OrganizeSourceFilesAfterCopy
コピーできたファイルは元フォルダーから削除し、コピーできなかったファイルは
対象外フォルダーへ移動します。

.PARAMETER ExcludedFolderName
OrganizeSourceFilesAfterCopy が有効な場合に、コピーできなかったファイルの移動先として作成する
フォルダー名です。既定値は「対象外」です。

.PARAMETER LogFilePath
ログファイルの出力先パスです。
省略時は OutputFolder 配下に rename_log_yyyyMMdd_HHmmss.log を作成します。

.PARAMETER PassThru
処理結果のサマリー オブジェクトをパイプラインへ出力します。

.EXAMPLE
.\Invoke-NamingAgentBatch.ps1 `
    -InputFolder 'C:\work\input' `
    -OutputFolder 'C:\work\output' `
    -ApiBaseUrl 'https://example.contoso.com' `
    -ApiKey 'YOUR_API_KEY'

.EXAMPLE
.\Invoke-NamingAgentBatch.ps1 `
    -InputFolder 'C:\work\input' `
    -OutputFolder 'C:\work\output' `
    -ApiBaseUrl 'https://example.contoso.com' `
    -ApiKey 'YOUR_API_KEY' `
    -Parallelism 5 `
    -CopyNonRenamedFiles `
    -OrganizeSourceFilesAfterCopy `
    -Timeout 900 `
    -Verbose `
    -PassThru

.NOTES
対象拡張子:
- .pdf
- .txt
- .csv

API 前提:
- 認証は x-api-key ヘッダー
- エンドポイントは /assist-naming
- multipart/form-data の file パラメーターで単一ファイルを送信

WhatIf/Confirm について:
- SupportsShouldProcess は宣言していますが、並列実行との整合性を優先し、
  実装上は WhatIf を真偽値としてワーカーへ渡して疑似対応しています。
- 対話的な Confirm は並列実行と相性が悪いため、本スクリプトでは WhatIf を主な事前確認手段として想定しています。
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputFolder,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ApiBaseUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ApiKey,

    [Parameter()]
    [ValidateRange(1, 86400)]
    [int]$Timeout = 600,

    [Parameter()]
    [ValidateRange(1, 256)]
    [int]$Parallelism = 3,

    [Parameter()]
    [switch]$CopyNonRenamedFiles,

    [Parameter()]
    [switch]$OrganizeSourceFilesAfterCopy,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExcludedFolderName = '対象外',

    [Parameter()]
    [string]$LogFilePath,

    [Parameter()]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 許可対象の拡張子一覧
$script:AllowedExtensions = @('.pdf', '.txt', '.csv')

# プロセス内で共有するログ書き込み用ミューテックス名
$script:LogMutexName = 'Global\InvokeNamingAgentBatchLogMutex'

#------------------------------------------------------------
# ログ出力先を初期化する
#------------------------------------------------------------
function Initialize-LogFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputFolderPath,

        [Parameter()]
        [AllowEmptyString()]
        [string]$RequestedLogFilePath
    )

    if ([string]::IsNullOrWhiteSpace($RequestedLogFilePath)) {
        $resolvedLogPath = Join-Path -Path $OutputFolderPath -ChildPath ("rename_log_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    else {
        $resolvedLogPath = $RequestedLogFilePath
    }

    $logDirectory = Split-Path -Parent $resolvedLogPath
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory)) {
        $null = New-Item -ItemType Directory -Path $logDirectory -Force
    }

    if (-not (Test-Path -LiteralPath $resolvedLogPath)) {
        $null = New-Item -ItemType File -Path $resolvedLogPath -Force
    }

    return [System.IO.Path]::GetFullPath($resolvedLogPath)
}

#------------------------------------------------------------
# ミューテックスを使ってログファイルへ排他的に書き込む
#------------------------------------------------------------
function Write-LogTextSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$MutexName
    )

    $createdNew = $false
    $mutex = [System.Threading.Mutex]::new($false, $MutexName, [ref]$createdNew)

    try {
        $null = $mutex.WaitOne()
        Add-Content -LiteralPath $Path -Value $Text -Encoding UTF8
    }
    finally {
        try {
            $mutex.ReleaseMutex() | Out-Null
        }
        catch {
        }

        $mutex.Dispose()
    }
}

#------------------------------------------------------------
# ログファイルと Verbose 出力へメッセージを書き込む
#------------------------------------------------------------
function Write-LogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message

    Write-LogTextSafely -Path $script:LogFilePath -Text $line -MutexName $script:LogMutexName

    switch ($Level) {
        'ERROR' { Write-Warning -Message ('ERROR: ' + $Message) }
        'WARN'  { Write-Warning -Message $Message }
        'DEBUG' { Write-Debug -Message $Message }
        default { Write-Verbose -Message $Message }
    }
}

#------------------------------------------------------------
# ファイル名として使用できない文字を置換する
#------------------------------------------------------------
function ConvertTo-SafeFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $safeName = $Name -replace '[\\/:*?"<>|]', '_'
    $safeName = $safeName -replace '\s+', ' '
    $safeName = $safeName -replace '-{2,}', '-'
    $safeName = $safeName.Trim(' ', '.')

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return '_'
    }

    return $safeName
}

#------------------------------------------------------------
# 指定ディレクトリが存在しない場合は作成する
#------------------------------------------------------------
function New-DirectoryIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -Force
    }
}

#------------------------------------------------------------
# 入力フォルダーからの相対パスを取得する
#------------------------------------------------------------
function Get-RelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    if (-not $baseFullPath.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $baseFullPath += [System.IO.Path]::DirectorySeparatorChar
    }

    $childFullPath = [System.IO.Path]::GetFullPath($ChildPath)

    $baseUri = [System.Uri]::new($baseFullPath)
    $childUri = [System.Uri]::new($childFullPath)
    $relativeUri = $baseUri.MakeRelativeUri($childUri)

    return ([System.Uri]::UnescapeDataString($relativeUri.ToString()) -replace '/', [System.IO.Path]::DirectorySeparatorChar)
}

#------------------------------------------------------------
# 出力先に同名ファイルがある場合は連番付きのパスを返す
#------------------------------------------------------------
function Get-UniqueDestinationPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationDirectory,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $candidatePath = Join-Path -Path $DestinationDirectory -ChildPath $FileName
    if (-not (Test-Path -LiteralPath $candidatePath)) {
        return $candidatePath
    }

    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)

    for ($index = 1; $index -le 9999; $index++) {
        $candidatePath = Join-Path -Path $DestinationDirectory -ChildPath ('{0}_{1:D4}{2}' -f $nameWithoutExtension, $index, $extension)
        if (-not (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }
    }

    throw "一意な出力ファイル名を確保できませんでした。FileName=[$FileName]"
}

#------------------------------------------------------------
# ファイルをコピーし、元ファイルのタイムスタンプを復元する
#------------------------------------------------------------
function Copy-FilePreservingTimestamp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $destinationDirectory = Split-Path -Parent $DestinationPath
    New-DirectoryIfNeeded -Path $destinationDirectory

    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force

    $sourceItem = Get-Item -LiteralPath $SourcePath
    $destinationItem = Get-Item -LiteralPath $DestinationPath

    $destinationItem.CreationTime = $sourceItem.CreationTime
    $destinationItem.LastWriteTime = $sourceItem.LastWriteTime
    $destinationItem.LastAccessTime = $sourceItem.LastAccessTime
}

#------------------------------------------------------------
# 元ファイルを削除する
#------------------------------------------------------------
function Remove-SourceFileIfExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
}

#------------------------------------------------------------
# ファイルを移動し、移動先ディレクトリを自動作成する
#------------------------------------------------------------
function Move-FileEnsuringDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $destinationDirectory = Split-Path -Parent $DestinationPath
    New-DirectoryIfNeeded -Path $destinationDirectory

    Move-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
}

#------------------------------------------------------------
# 入力フォルダー配下の空サブフォルダーを削除する
#------------------------------------------------------------
function Remove-EmptySubdirectories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter()]
        [bool]$WhatIfEnabled = $false
    )

    $removedCount = 0

    $directories = @(
        Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -ErrorAction Stop |
            Sort-Object -Property @{ Expression = { $_.FullName.Length }; Descending = $true }, @{ Expression = { $_.FullName }; Descending = $true }
    )

    foreach ($directory in $directories) {
        $childItem = Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop | Select-Object -First 1
        if ($null -eq $childItem) {
            if (-not $WhatIfEnabled) {
                Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
            }

            $removedCount++
            Write-LogEntry -Level INFO -Message ('EMPTY_SOURCE_DIRECTORY_REMOVED path={0} whatif={1}' -f $directory.FullName, $WhatIfEnabled)
        }
    }

    return $removedCount
}

#------------------------------------------------------------
# API 呼び出し結果を扱いやすいオブジェクトへ整形する
#------------------------------------------------------------
function ConvertFrom-NamingApiResponse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawResponse
    )

    $jsonObject = $RawResponse | ConvertFrom-Json

    if (-not $jsonObject.documents -or $jsonObject.documents.Count -lt 1) {
        throw 'API レスポンスに documents[0] が存在しません。'
    }

    $document = $jsonObject.documents[0]

    return [pscustomobject]@{
        RawResponse      = $RawResponse
        UploadFileName   = ''
        Summary          = $jsonObject.summary
        Document         = $document
        Rename           = [bool]$document.rename
        NewFileName      = [string]$document.new_filename
        Notes            = [string]$document.notes
        OriginalFileName = [string]$document.original_filename
    }
}

#------------------------------------------------------------
# API アップロード時に使用する英数字のみのファイル名を生成する
#
# 備考:
# - API には元の日本語ファイル名を渡さず、ASCII 安全な名前で送信する
# - ローカル側の処理・ログ・コピー先判定は元ファイル名を利用する
#------------------------------------------------------------
function New-AsciiUploadFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OriginalPath
    )

    $extension = [System.IO.Path]::GetExtension($OriginalPath)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = '.bin'
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $token = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)

    return ('upload_{0}_{1}{2}' -f $timestamp, $token, $extension.ToLowerInvariant())
}

#------------------------------------------------------------
# 命名支援 API を呼び出して JSON 結果を取得する
#------------------------------------------------------------
function Invoke-NamingApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [uri]$ApiUrl,

        [Parameter(Mandatory = $true)]
        [string]$ApiKeyValue,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds
    )

    Add-Type -AssemblyName System.Net.Http

    $httpClient = [System.Net.Http.HttpClient]::new()
    try {
        $httpClient.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
        $httpClient.DefaultRequestHeaders.Accept.Clear()
        $null = $httpClient.DefaultRequestHeaders.Accept.Add(
            [System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json')
        )
        $null = $httpClient.DefaultRequestHeaders.Add('x-api-key', $ApiKeyValue)

        $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
        $memoryStream = [System.IO.MemoryStream]::new($fileBytes)
        try {
            $streamContent = [System.Net.Http.StreamContent]::new($memoryStream)
            $streamContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/octet-stream')

            $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
            try {
                $uploadFileName = New-AsciiUploadFileName -OriginalPath $FilePath
                $multipartContent.Add($streamContent, 'file', $uploadFileName)

                $response = $httpClient.PostAsync($ApiUrl, $multipartContent).GetAwaiter().GetResult()
                $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                if (-not $response.IsSuccessStatusCode) {
                    throw ('HTTP {0} {1} body={2}' -f [int]$response.StatusCode, $response.ReasonPhrase, $rawResponse)
                }

                if ([string]::IsNullOrWhiteSpace($rawResponse)) {
                    throw 'API レスポンス本文が空です。'
                }

                $parsedResponse = ConvertFrom-NamingApiResponse -RawResponse $rawResponse
                $parsedResponse | Add-Member -NotePropertyName 'UploadFileName' -NotePropertyValue $uploadFileName -Force
                return $parsedResponse
            }
            finally {
                $multipartContent.Dispose()
            }
        }
        finally {
            $memoryStream.Dispose()
        }
    }
    finally {
        $httpClient.Dispose()
    }
}

#------------------------------------------------------------
# 指定ファイルが処理対象拡張子かを判定する
#------------------------------------------------------------
function Test-SupportedExtension {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    return $script:AllowedExtensions -contains $Extension.ToLowerInvariant()
}

#------------------------------------------------------------
# 入力パラメーターを検証し、正規化済みの値を返す
#------------------------------------------------------------
function Resolve-ExecutionContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFolderPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputFolderPath,

        [Parameter(Mandatory = $true)]
        [string]$ApiBaseUrlValue
    )

    $resolvedInputFolder = [System.IO.Path]::GetFullPath($InputFolderPath)
    $resolvedOutputFolder = [System.IO.Path]::GetFullPath($OutputFolderPath)

    if (-not (Test-Path -LiteralPath $resolvedInputFolder)) {
        throw "InputFolder が見つかりません。Path=[$resolvedInputFolder]"
    }

    if (-not (Test-Path -LiteralPath $resolvedInputFolder -PathType Container)) {
        throw "InputFolder はフォルダーではありません。Path=[$resolvedInputFolder]"
    }

    New-DirectoryIfNeeded -Path $resolvedOutputFolder

    $apiUrl = [uri]($ApiBaseUrlValue.TrimEnd('/') + '/assist-naming')

    return [pscustomobject]@{
        InputFolder  = $resolvedInputFolder
        OutputFolder = $resolvedOutputFolder
        ApiUrl       = $apiUrl
    }
}

#------------------------------------------------------------
# 並列実行用の RunspacePool を使ってタスクを処理する
#------------------------------------------------------------
function Invoke-ParallelTaskCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [Parameter(Mandatory = $true)]
        [int]$ThrottleLimit,

        [Parameter(Mandatory = $true)]
        [scriptblock]$WorkerScript,

        [Parameter(Mandatory = $true)]
        [hashtable]$SharedArguments
    )

	if ($Items.Count -eq 0) {
		return @()
	}

    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $initialSessionState, $Host)
    $runspacePool.Open()

    $runningJobs = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($item in $Items) {
            $powerShell = [powershell]::Create()
            $powerShell.RunspacePool = $runspacePool

            $null = $powerShell.AddScript($WorkerScript)
            $null = $powerShell.AddArgument($item)
            $null = $powerShell.AddArgument($SharedArguments)

            $asyncHandle = $powerShell.BeginInvoke()

            $runningJobs.Add([pscustomobject]@{
                PowerShell = $powerShell
                Handle     = $asyncHandle
                SourceItem = $item
            })
        }

        $results = New-Object System.Collections.Generic.List[object]
        $completedCount = 0
        $totalCount = $runningJobs.Count

        while ($runningJobs.Count -gt 0) {
            for ($i = $runningJobs.Count - 1; $i -ge 0; $i--) {
                $job = $runningJobs[$i]

                if ($job.Handle.IsCompleted) {
                    try {
                        $outputItems = $job.PowerShell.EndInvoke($job.Handle)
                        foreach ($outputItem in $outputItems) {
                            $results.Add($outputItem)
                        }
                    }
                    finally {
                        $job.PowerShell.Dispose()
                        $runningJobs.RemoveAt($i)
                        $completedCount++

                        $percentComplete = if ($totalCount -eq 0) {
                            0
                        }
                        else {
                            [int](($completedCount / $totalCount) * 100)
                        }

                        Write-Progress `
                            -Activity '命名支援 API によるファイル処理' `
                            -Status ('完了 {0}/{1}' -f $completedCount, $totalCount) `
                            -PercentComplete $percentComplete
                    }
                }
            }

            if ($runningJobs.Count -gt 0) {
                Start-Sleep -Milliseconds 100
            }
        }

        Write-Progress -Activity '命名支援 API によるファイル処理' -Completed

        return $results
    }
    finally {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}

$resolvedContext = Resolve-ExecutionContext -InputFolderPath $InputFolder -OutputFolderPath $OutputFolder -ApiBaseUrlValue $ApiBaseUrl
$script:LogFilePath = Initialize-LogFile -OutputFolderPath $resolvedContext.OutputFolder -RequestedLogFilePath $LogFilePath

# 対象外ファイルの移動先ルートを決定する
$excludedRootFolder = Join-Path -Path ([System.IO.Path]::GetDirectoryName($resolvedContext.InputFolder)) -ChildPath $ExcludedFolderName

Write-LogEntry -Level INFO -Message ('InputFolder={0}' -f $resolvedContext.InputFolder)
Write-LogEntry -Level INFO -Message ('OutputFolder={0}' -f $resolvedContext.OutputFolder)
Write-LogEntry -Level INFO -Message ('ApiUrl={0}' -f $resolvedContext.ApiUrl.AbsoluteUri)
Write-LogEntry -Level INFO -Message ('Timeout={0}' -f $Timeout)
Write-LogEntry -Level INFO -Message ('Parallelism={0}' -f $Parallelism)
Write-LogEntry -Level INFO -Message ('CopyNonRenamedFiles={0}' -f $CopyNonRenamedFiles.IsPresent)
Write-LogEntry -Level INFO -Message ('OrganizeSourceFilesAfterCopy={0}' -f $OrganizeSourceFilesAfterCopy.IsPresent)
Write-LogEntry -Level INFO -Message ('ExcludedRootFolder={0}' -f $excludedRootFolder)
Write-LogEntry -Level INFO -Message ('LogFilePath={0}' -f $script:LogFilePath)

# 処理対象ファイルを再帰的に収集する
$targetFiles = @(
    Get-ChildItem -LiteralPath $resolvedContext.InputFolder -File -Recurse |
        Where-Object { Test-SupportedExtension -Extension $_.Extension } |
        Sort-Object -Property FullName
)

Write-LogEntry -Level INFO -Message ('TargetFiles={0}' -f $targetFiles.Count)

if ($targetFiles.Count -eq 0) {
    Write-LogEntry -Level WARN -Message '処理対象ファイルが見つかりませんでした。'

    $summaryObject = [pscustomobject]@{
        Total               = 0
        Renamed             = 0
        CopiedWithoutRename = 0
        SkippedNonRenamed   = 0
        MovedToExcluded     = 0
        DeletedSource              = 0
        RemovedEmptySourceDirectories = 0
        Errors                     = 0
    }

    $summaryJson = $summaryObject | ConvertTo-Json -Depth 10 -Compress
    Write-LogEntry -Level INFO -Message ('SUMMARY json={0}' -f $summaryJson)

    if ($PassThru.IsPresent) {
        $summaryObject
    }

    return
}

# 処理サマリーを保持する
$stats = [ordered]@{
    Total               = 0
    Renamed             = 0
    CopiedWithoutRename = 0
    SkippedNonRenamed   = 0
    MovedToExcluded            = 0
    DeletedSource              = 0
    RemovedEmptySourceDirectories = 0
    Errors                     = 0
}

#------------------------------------------------------------
# 並列ワーカー本体
#
# 備考:
# - 親スコープの関数は Runspace から直接参照できないため、
#   必要最小限のヘルパーをワーカー内にも定義する
# - WhatIf は真偽値として渡し、コピー・削除・移動を抑止する
#------------------------------------------------------------
$parallelWorker = {
    param(
        [Parameter(Mandatory = $true)]
        $File,

        [Parameter(Mandatory = $true)]
        [hashtable]$Shared
    )

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    #------------------------------------------------------------
    # ワーカー内でログファイルへ安全に書き込む
    #------------------------------------------------------------
    function Write-WorkerLogEntry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
            [string]$Level,

            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $true)]
            [string]$LogPath,

            [Parameter(Mandatory = $true)]
            [string]$MutexName
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $line = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message

        $createdNew = $false
        $mutex = [System.Threading.Mutex]::new($false, $MutexName, [ref]$createdNew)

        try {
            $null = $mutex.WaitOne()
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }
        finally {
            try {
                $mutex.ReleaseMutex() | Out-Null
            }
            catch {
            }

            $mutex.Dispose()
        }
    }

    #------------------------------------------------------------
    # ファイル名として使用できない文字を置換する
    #------------------------------------------------------------
    function ConvertTo-SafeFileNameLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        $safeName = $Name -replace '[\\/:*?"<>|]', '_'
        $safeName = $safeName -replace '\s+', ' '
        $safeName = $safeName -replace '-{2,}', '-'
        $safeName = $safeName.Trim(' ', '.')

        if ([string]::IsNullOrWhiteSpace($safeName)) {
            return '_'
        }

        return $safeName
    }

    #------------------------------------------------------------
    # 指定ディレクトリが存在しない場合は作成する
    #------------------------------------------------------------
    function New-DirectoryIfNeededLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            $null = New-Item -ItemType Directory -Path $Path -Force
        }
    }

    #------------------------------------------------------------
    # 入力フォルダーからの相対パスを取得する
    #------------------------------------------------------------
    function Get-RelativePathLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$BasePath,

            [Parameter(Mandatory = $true)]
            [string]$ChildPath
        )

        $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
        if (-not $baseFullPath.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
            $baseFullPath += [System.IO.Path]::DirectorySeparatorChar
        }

        $childFullPath = [System.IO.Path]::GetFullPath($ChildPath)

        $baseUri = [System.Uri]::new($baseFullPath)
        $childUri = [System.Uri]::new($childFullPath)
        $relativeUri = $baseUri.MakeRelativeUri($childUri)

        return ([System.Uri]::UnescapeDataString($relativeUri.ToString()) -replace '/', [System.IO.Path]::DirectorySeparatorChar)
    }

    #------------------------------------------------------------
    # 出力先に同名ファイルがある場合は連番付きのパスを返す
    #------------------------------------------------------------
    function Get-UniqueDestinationPathLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$DestinationDirectory,

            [Parameter(Mandatory = $true)]
            [string]$FileName
        )

        $candidatePath = Join-Path -Path $DestinationDirectory -ChildPath $FileName
        if (-not (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }

        $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $extension = [System.IO.Path]::GetExtension($FileName)

        for ($index = 1; $index -le 9999; $index++) {
            $candidatePath = Join-Path -Path $DestinationDirectory -ChildPath ('{0}_{1:D4}{2}' -f $nameWithoutExtension, $index, $extension)
            if (-not (Test-Path -LiteralPath $candidatePath)) {
                return $candidatePath
            }
        }

        throw "一意な出力ファイル名を確保できませんでした。FileName=[$FileName]"
    }

    #------------------------------------------------------------
    # ファイルをコピーし、元ファイルのタイムスタンプを復元する
    #------------------------------------------------------------
    function Copy-FilePreservingTimestampLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourcePath,

            [Parameter(Mandatory = $true)]
            [string]$DestinationPath
        )

        $destinationDirectory = Split-Path -Parent $DestinationPath
        New-DirectoryIfNeededLocal -Path $destinationDirectory

        Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force

        $sourceItem = Get-Item -LiteralPath $SourcePath
        $destinationItem = Get-Item -LiteralPath $DestinationPath

        $destinationItem.CreationTime = $sourceItem.CreationTime
        $destinationItem.LastWriteTime = $sourceItem.LastWriteTime
        $destinationItem.LastAccessTime = $sourceItem.LastAccessTime
    }

    #------------------------------------------------------------
    # 元ファイルを削除する
    #------------------------------------------------------------
    function Remove-SourceFileIfExistsLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }
    }

    #------------------------------------------------------------
    # ファイルを移動し、移動先ディレクトリを自動作成する
    #------------------------------------------------------------
    function Move-FileEnsuringDirectoryLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourcePath,

            [Parameter(Mandatory = $true)]
            [string]$DestinationPath
        )

        $destinationDirectory = Split-Path -Parent $DestinationPath
        New-DirectoryIfNeededLocal -Path $destinationDirectory

        Move-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    }

    #------------------------------------------------------------
    # API 呼び出し結果を扱いやすいオブジェクトへ整形する
    #------------------------------------------------------------
    function ConvertFrom-NamingApiResponseLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$RawResponse
        )

        $jsonObject = $RawResponse | ConvertFrom-Json

        if (-not $jsonObject.documents -or $jsonObject.documents.Count -lt 1) {
            throw 'API レスポンスに documents[0] が存在しません。'
        }

        $document = $jsonObject.documents[0]

        return [pscustomobject]@{
            RawResponse      = $RawResponse
            UploadFileName   = ''
            Summary          = $jsonObject.summary
            Document         = $document
            Rename           = [bool]$document.rename
            NewFileName      = [string]$document.new_filename
            Notes            = [string]$document.notes
            OriginalFileName = [string]$document.original_filename
        }
    }

    #------------------------------------------------------------
    # API アップロード時に使用する英数字のみのファイル名を生成する
    #------------------------------------------------------------
    function New-AsciiUploadFileNameLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$OriginalPath
        )

        $extension = [System.IO.Path]::GetExtension($OriginalPath)
        if ([string]::IsNullOrWhiteSpace($extension)) {
            $extension = '.bin'
        }

        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $token = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)

        return ('upload_{0}_{1}{2}' -f $timestamp, $token, $extension.ToLowerInvariant())
    }

    #------------------------------------------------------------
    # 命名支援 API を呼び出して JSON 結果を取得する
    #------------------------------------------------------------
    function Invoke-NamingApiRequestLocal {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$FilePath,

            [Parameter(Mandatory = $true)]
            [uri]$ApiUrl,

            [Parameter(Mandatory = $true)]
            [string]$ApiKeyValue,

            [Parameter(Mandatory = $true)]
            [int]$TimeoutSeconds
        )

        Add-Type -AssemblyName System.Net.Http

        $httpClient = [System.Net.Http.HttpClient]::new()
        try {
            $httpClient.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
            $httpClient.DefaultRequestHeaders.Accept.Clear()
            $null = $httpClient.DefaultRequestHeaders.Accept.Add(
                [System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json')
            )
            $null = $httpClient.DefaultRequestHeaders.Add('x-api-key', $ApiKeyValue)

            $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
            $memoryStream = [System.IO.MemoryStream]::new($fileBytes)
            try {
                $streamContent = [System.Net.Http.StreamContent]::new($memoryStream)
                $streamContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/octet-stream')

                $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
                try {
                    $uploadFileName = New-AsciiUploadFileNameLocal -OriginalPath $FilePath
                    $multipartContent.Add($streamContent, 'file', $uploadFileName)

                    $response = $httpClient.PostAsync($ApiUrl, $multipartContent).GetAwaiter().GetResult()
                    $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                    if (-not $response.IsSuccessStatusCode) {
                        throw ('HTTP {0} {1} body={2}' -f [int]$response.StatusCode, $response.ReasonPhrase, $rawResponse)
                    }

                    if ([string]::IsNullOrWhiteSpace($rawResponse)) {
                        throw 'API レスポンス本文が空です。'
                    }

                    $parsedResponse = ConvertFrom-NamingApiResponseLocal -RawResponse $rawResponse
                    $parsedResponse | Add-Member -NotePropertyName 'UploadFileName' -NotePropertyValue $uploadFileName -Force
                    return $parsedResponse
                }
                finally {
                    $multipartContent.Dispose()
                }
            }
            finally {
                $memoryStream.Dispose()
            }
        }
        finally {
            $httpClient.Dispose()
        }
    }

    $inputFolder = [string]$Shared.InputFolder
    $outputFolder = [string]$Shared.OutputFolder
    $apiUrl = [uri]$Shared.ApiUrl
    $apiKey = [string]$Shared.ApiKey
    $timeout = [int]$Shared.Timeout
    $copyNonRenamedFiles = [bool]$Shared.CopyNonRenamedFiles
    $organizeSourceFilesAfterCopy = [bool]$Shared.OrganizeSourceFilesAfterCopy
    $excludedRootFolder = [string]$Shared.ExcludedRootFolder
    $logFilePath = [string]$Shared.LogFilePath
    $logMutexName = [string]$Shared.LogMutexName
    $whatIfEnabled = [bool]$Shared.WhatIfEnabled

    $relativePath = Get-RelativePathLocal -BasePath $inputFolder -ChildPath $File.FullName
    $relativeDirectory = Split-Path -Parent $relativePath

    $destinationDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
        $outputFolder
    }
    else {
        Join-Path -Path $outputFolder -ChildPath $relativeDirectory
    }

    $excludedDestinationDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
        $excludedRootFolder
    }
    else {
        Join-Path -Path $excludedRootFolder -ChildPath $relativeDirectory
    }

    Write-WorkerLogEntry -Level INFO -Message ('START file={0}' -f $File.FullName) -LogPath $logFilePath -MutexName $logMutexName

    try {
        $apiResult = Invoke-NamingApiRequestLocal -FilePath $File.FullName -ApiUrl $apiUrl -ApiKeyValue $apiKey -TimeoutSeconds $timeout

        Write-WorkerLogEntry -Level INFO -Message ('API_UPLOAD file={0} upload_name={1}' -f $File.FullName, $apiResult.UploadFileName) -LogPath $logFilePath -MutexName $logMutexName
        Write-WorkerLogEntry -Level INFO -Message ('API_RESPONSE file={0} json={1}' -f $File.FullName, $apiResult.RawResponse) -LogPath $logFilePath -MutexName $logMutexName

        $document = $apiResult.Document
        $extractedDate = if ($document.extracted -and $document.extracted.date) { [string]$document.extracted.date } else { '' }
        $extractedVendor = if ($document.extracted -and $document.extracted.vendor) { [string]$document.extracted.vendor } else { '' }
        $extractedAmount = if ($document.extracted -and $document.extracted.amount) { [string]$document.extracted.amount } else { '' }

        Write-WorkerLogEntry -Level INFO -Message (
            'EXTRACTED file={0} rename={1} new_filename={2} date={3} vendor={4} amount={5} notes={6}' -f
            $File.FullName, $apiResult.Rename, $apiResult.NewFileName, $extractedDate, $extractedVendor, $extractedAmount, $apiResult.Notes
        ) -LogPath $logFilePath -MutexName $logMutexName

        if ($apiResult.Rename -and -not [string]::IsNullOrWhiteSpace($apiResult.NewFileName)) {
            $safeFileName = ConvertTo-SafeFileNameLocal -Name $apiResult.NewFileName

            if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($safeFileName))) {
                $safeFileName = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName) + $File.Extension
            }

            New-DirectoryIfNeededLocal -Path $destinationDirectory
            $destinationPath = Get-UniqueDestinationPathLocal -DestinationDirectory $destinationDirectory -FileName $safeFileName

            if (-not $whatIfEnabled) {
                Copy-FilePreservingTimestampLocal -SourcePath $File.FullName -DestinationPath $destinationPath
            }

            Write-WorkerLogEntry -Level INFO -Message ('RENAMED src={0} dst={1} whatif={2}' -f $File.FullName, $destinationPath, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName

            $sourceDeleted = $false
            if ($organizeSourceFilesAfterCopy) {
                if (-not $whatIfEnabled) {
                    Remove-SourceFileIfExistsLocal -Path $File.FullName
                }

                $sourceDeleted = $true
                Write-WorkerLogEntry -Level INFO -Message ('SOURCE_DELETED src={0} whatif={1}' -f $File.FullName, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName
            }

            return [pscustomobject]@{
                SourcePath          = $File.FullName
                SourceName          = $File.Name
                ActionType          = 'Renamed'
                DestinationPath     = $destinationPath
                DestinationFileName = [System.IO.Path]::GetFileName($destinationPath)
                Rename              = $apiResult.Rename
                NewFileName         = $apiResult.NewFileName
                Notes               = $apiResult.Notes
                UploadFileName      = $apiResult.UploadFileName
                RawResponse         = $apiResult.RawResponse
                SourceDeleted       = $sourceDeleted
                MovedToExcluded     = $false
                ErrorMessage        = ''
            }
        }

        if ($copyNonRenamedFiles) {
            New-DirectoryIfNeededLocal -Path $destinationDirectory
            $destinationPath = Get-UniqueDestinationPathLocal -DestinationDirectory $destinationDirectory -FileName $File.Name

            if (-not $whatIfEnabled) {
                Copy-FilePreservingTimestampLocal -SourcePath $File.FullName -DestinationPath $destinationPath
            }

            Write-WorkerLogEntry -Level INFO -Message ('COPIED_ORIGINAL src={0} dst={1} whatif={2}' -f $File.FullName, $destinationPath, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName

            $sourceDeleted = $false
            if ($organizeSourceFilesAfterCopy) {
                if (-not $whatIfEnabled) {
                    Remove-SourceFileIfExistsLocal -Path $File.FullName
                }

                $sourceDeleted = $true
                Write-WorkerLogEntry -Level INFO -Message ('SOURCE_DELETED src={0} whatif={1}' -f $File.FullName, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName
            }

            return [pscustomobject]@{
                SourcePath          = $File.FullName
                SourceName          = $File.Name
                ActionType          = 'CopiedWithoutRename'
                DestinationPath     = $destinationPath
                DestinationFileName = [System.IO.Path]::GetFileName($destinationPath)
                Rename              = $apiResult.Rename
                NewFileName         = $apiResult.NewFileName
                Notes               = $apiResult.Notes
                UploadFileName      = $apiResult.UploadFileName
                RawResponse         = $apiResult.RawResponse
                SourceDeleted       = $sourceDeleted
                MovedToExcluded     = $false
                ErrorMessage        = ''
            }
        }

        Write-WorkerLogEntry -Level INFO -Message ('SKIPPED_NONRENAMED file={0} notes={1}' -f $File.FullName, $apiResult.Notes) -LogPath $logFilePath -MutexName $logMutexName

        $movedToExcluded = $false
        $excludedDestinationPath = ''

        if ($organizeSourceFilesAfterCopy) {
            New-DirectoryIfNeededLocal -Path $excludedDestinationDirectory
            $excludedDestinationPath = Get-UniqueDestinationPathLocal -DestinationDirectory $excludedDestinationDirectory -FileName $File.Name

            if (-not $whatIfEnabled) {
                Move-FileEnsuringDirectoryLocal -SourcePath $File.FullName -DestinationPath $excludedDestinationPath
            }

            $movedToExcluded = $true
            Write-WorkerLogEntry -Level INFO -Message ('MOVED_TO_EXCLUDED src={0} dst={1} whatif={2}' -f $File.FullName, $excludedDestinationPath, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName
        }

        return [pscustomobject]@{
            SourcePath          = $File.FullName
            SourceName          = $File.Name
            ActionType          = 'SkippedNonRenamed'
            DestinationPath     = $excludedDestinationPath
            DestinationFileName = if ([string]::IsNullOrWhiteSpace($excludedDestinationPath)) { '' } else { [System.IO.Path]::GetFileName($excludedDestinationPath) }
            Rename              = $apiResult.Rename
            NewFileName         = $apiResult.NewFileName
            Notes               = $apiResult.Notes
            UploadFileName      = $apiResult.UploadFileName
            RawResponse         = $apiResult.RawResponse
            SourceDeleted       = $false
            MovedToExcluded     = $movedToExcluded
            ErrorMessage        = ''
        }
    }
    catch {
        Write-WorkerLogEntry -Level ERROR -Message ('FAILED file={0} error={1}' -f $File.FullName, $_.Exception.Message) -LogPath $logFilePath -MutexName $logMutexName

        $movedToExcluded = $false
        $excludedDestinationPath = ''

        if ($organizeSourceFilesAfterCopy -and (Test-Path -LiteralPath $File.FullName)) {
            try {
                New-DirectoryIfNeededLocal -Path $excludedDestinationDirectory
                $excludedDestinationPath = Get-UniqueDestinationPathLocal -DestinationDirectory $excludedDestinationDirectory -FileName $File.Name

                if (-not $whatIfEnabled) {
                    Move-FileEnsuringDirectoryLocal -SourcePath $File.FullName -DestinationPath $excludedDestinationPath
                }

                $movedToExcluded = $true
                Write-WorkerLogEntry -Level INFO -Message ('MOVED_TO_EXCLUDED_AFTER_ERROR src={0} dst={1} whatif={2}' -f $File.FullName, $excludedDestinationPath, $whatIfEnabled) -LogPath $logFilePath -MutexName $logMutexName
            }
            catch {
                Write-WorkerLogEntry -Level ERROR -Message ('FAILED_TO_MOVE_EXCLUDED file={0} error={1}' -f $File.FullName, $_.Exception.Message) -LogPath $logFilePath -MutexName $logMutexName
            }
        }

        return [pscustomobject]@{
            SourcePath          = $File.FullName
            SourceName          = $File.Name
            ActionType          = 'Error'
            DestinationPath     = $excludedDestinationPath
            DestinationFileName = if ([string]::IsNullOrWhiteSpace($excludedDestinationPath)) { '' } else { [System.IO.Path]::GetFileName($excludedDestinationPath) }
            Rename              = $false
            NewFileName         = ''
            Notes               = ''
            UploadFileName      = ''
            RawResponse         = ''
            SourceDeleted       = $false
            MovedToExcluded     = $movedToExcluded
            ErrorMessage        = $_.Exception.Message
        }
    }
}

# 並列ワーカーへ渡す共有情報を組み立てる
$sharedArguments = @{
    InputFolder                  = $resolvedContext.InputFolder
    OutputFolder                 = $resolvedContext.OutputFolder
    ApiUrl                       = $resolvedContext.ApiUrl.AbsoluteUri
    ApiKey                       = $ApiKey
    Timeout                      = $Timeout
    CopyNonRenamedFiles          = $CopyNonRenamedFiles.IsPresent
    OrganizeSourceFilesAfterCopy = $OrganizeSourceFilesAfterCopy.IsPresent
    ExcludedRootFolder           = $excludedRootFolder
    LogFilePath                  = $script:LogFilePath
    LogMutexName                 = $script:LogMutexName
    WhatIfEnabled                = [bool]$WhatIfPreference
}

# 全対象ファイルを並列実行する
$parallelResults = Invoke-ParallelTaskCollection `
    -Items $targetFiles `
    -ThrottleLimit $Parallelism `
    -WorkerScript $parallelWorker `
    -SharedArguments $sharedArguments

if ($OrganizeSourceFilesAfterCopy.IsPresent) {
    $removedEmptyDirectoryCount = Remove-EmptySubdirectories -RootPath $resolvedContext.InputFolder -WhatIfEnabled ([bool]$WhatIfPreference)
    Write-LogEntry -Level INFO -Message ('RemovedEmptySourceDirectories={0}' -f $removedEmptyDirectoryCount)
}
else {
    $removedEmptyDirectoryCount = 0
}

# 並列実行結果を集計する
foreach ($result in $parallelResults | Sort-Object -Property SourcePath) {
    $stats.Total++

    switch ($result.ActionType) {
        'Renamed' {
            $stats.Renamed++
        }

        'CopiedWithoutRename' {
            $stats.CopiedWithoutRename++
        }

        'SkippedNonRenamed' {
            $stats.SkippedNonRenamed++
        }

        'Error' {
            $stats.Errors++
        }

        default {
            $stats.Errors++
            Write-LogEntry -Level WARN -Message ('UNKNOWN_RESULT file={0} action={1}' -f $result.SourcePath, $result.ActionType)
        }
    }

    if ($result.SourceDeleted) {
        $stats.DeletedSource++
    }

    if ($result.MovedToExcluded) {
        $stats.MovedToExcluded++
    }
}

$stats.RemovedEmptySourceDirectories = $removedEmptyDirectoryCount

$summaryObject = [pscustomobject]$stats
$summaryJson = $summaryObject | ConvertTo-Json -Depth 10 -Compress

Write-LogEntry -Level INFO -Message ('SUMMARY json={0}' -f $summaryJson)

if ($PassThru.IsPresent) {
    $summaryObject
}