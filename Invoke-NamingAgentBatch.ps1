<#
.SYNOPSIS
電子帳簿保存法向けの命名支援 API を呼び出し、抽出結果に基づいてファイルをリネームしてコピーします。

.DESCRIPTION
InputFolder 配下の .pdf / .txt / .csv ファイルを再帰的に走査し、各ファイルを命名支援 API に送信します。
API が返す documents[0].new_filename を使用して OutputFolder 配下へコピーします。

OutputFolder には InputFolder と同じサブフォルダー構造を再現します。
rename=false または new_filename が空の場合は、CopyNonRenamedFiles の指定に応じて
元の名前のままコピーするか、コピーせずにスキップします。

ログには API の生レスポンス JSON と、抽出結果・コピー結果・エラー情報を記録します。

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

.PARAMETER CopyNonRenamedFiles
API が rename=false を返した場合でも、元のファイル名のまま OutputFolder にコピーします。
省略時はコピーしません。

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
    -CopyNonRenamedFiles `
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
    [switch]$CopyNonRenamedFiles,

    [Parameter()]
    [string]$LogFilePath,

    [Parameter()]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 許可対象の拡張子一覧
$script:AllowedExtensions = @('.pdf', '.txt', '.csv')

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

    Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8

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

$resolvedContext = Resolve-ExecutionContext -InputFolderPath $InputFolder -OutputFolderPath $OutputFolder -ApiBaseUrlValue $ApiBaseUrl
$script:LogFilePath = Initialize-LogFile -OutputFolderPath $resolvedContext.OutputFolder -RequestedLogFilePath $LogFilePath

Write-LogEntry -Level INFO -Message ('InputFolder={0}' -f $resolvedContext.InputFolder)
Write-LogEntry -Level INFO -Message ('OutputFolder={0}' -f $resolvedContext.OutputFolder)
Write-LogEntry -Level INFO -Message ('ApiUrl={0}' -f $resolvedContext.ApiUrl.AbsoluteUri)
Write-LogEntry -Level INFO -Message ('Timeout={0}' -f $Timeout)
Write-LogEntry -Level INFO -Message ('CopyNonRenamedFiles={0}' -f $CopyNonRenamedFiles.IsPresent)
Write-LogEntry -Level INFO -Message ('LogFilePath={0}' -f $script:LogFilePath)

# 処理対象ファイルを再帰的に収集する
$targetFiles = Get-ChildItem -LiteralPath $resolvedContext.InputFolder -File -Recurse |
    Where-Object { Test-SupportedExtension -Extension $_.Extension } |
    Sort-Object -Property FullName

Write-LogEntry -Level INFO -Message ('TargetFiles={0}' -f $targetFiles.Count)

# 処理サマリーを保持する
$stats = [ordered]@{
    Total               = 0
    Renamed             = 0
    CopiedWithoutRename = 0
    SkippedNonRenamed   = 0
    Errors              = 0
}

for ($index = 0; $index -lt $targetFiles.Count; $index++) {
    $file = $targetFiles[$index]
    $stats.Total++

    $relativePath = Get-RelativePath -BasePath $resolvedContext.InputFolder -ChildPath $file.FullName
    $relativeDirectory = Split-Path -Parent $relativePath
    $destinationDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
        $resolvedContext.OutputFolder
    }
    else {
        Join-Path -Path $resolvedContext.OutputFolder -ChildPath $relativeDirectory
    }

    $progressPercent = if ($targetFiles.Count -eq 0) { 0 } else { [int](($index + 1) / $targetFiles.Count * 100) }
    Write-Progress -Activity '命名支援 API によるファイル処理' -Status $file.FullName -PercentComplete $progressPercent

    Write-LogEntry -Level INFO -Message ('START file={0}' -f $file.FullName)

    try {
        $apiResult = Invoke-NamingApiRequest -FilePath $file.FullName -ApiUrl $resolvedContext.ApiUrl -ApiKeyValue $ApiKey -TimeoutSeconds $Timeout

        Write-LogEntry -Level INFO -Message ('API_UPLOAD file={0} upload_name={1}' -f $file.FullName, $apiResult.UploadFileName)
        Write-LogEntry -Level INFO -Message ('API_RESPONSE file={0} json={1}' -f $file.FullName, $apiResult.RawResponse)

        $document = $apiResult.Document
        $extractedDate = if ($document.extracted -and $document.extracted.date) { [string]$document.extracted.date } else { '' }
        $extractedVendor = if ($document.extracted -and $document.extracted.vendor) { [string]$document.extracted.vendor } else { '' }
        $extractedAmount = if ($document.extracted -and $document.extracted.amount) { [string]$document.extracted.amount } else { '' }

        Write-LogEntry -Level INFO -Message (
            'EXTRACTED file={0} rename={1} new_filename={2} date={3} vendor={4} amount={5} notes={6}' -f
            $file.FullName, $apiResult.Rename, $apiResult.NewFileName, $extractedDate, $extractedVendor, $extractedAmount, $apiResult.Notes
        )

        if ($apiResult.Rename -and -not [string]::IsNullOrWhiteSpace($apiResult.NewFileName)) {
            $safeFileName = ConvertTo-SafeFileName -Name $apiResult.NewFileName

            if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($safeFileName))) {
                $safeFileName = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName) + $file.Extension
            }

            New-DirectoryIfNeeded -Path $destinationDirectory
            $destinationPath = Get-UniqueDestinationPath -DestinationDirectory $destinationDirectory -FileName $safeFileName

            if ($PSCmdlet.ShouldProcess($destinationPath, 'リネーム後ファイルをコピーします')) {
                Copy-FilePreservingTimestamp -SourcePath $file.FullName -DestinationPath $destinationPath
                $stats.Renamed++
                Write-LogEntry -Level INFO -Message ('RENAMED src={0} dst={1}' -f $file.FullName, $destinationPath)
            }
        }
        elseif ($CopyNonRenamedFiles.IsPresent) {
            New-DirectoryIfNeeded -Path $destinationDirectory
            $destinationPath = Get-UniqueDestinationPath -DestinationDirectory $destinationDirectory -FileName $file.Name

            if ($PSCmdlet.ShouldProcess($destinationPath, '元のファイル名のままコピーします')) {
                Copy-FilePreservingTimestamp -SourcePath $file.FullName -DestinationPath $destinationPath
                $stats.CopiedWithoutRename++
                Write-LogEntry -Level INFO -Message ('COPIED_ORIGINAL src={0} dst={1}' -f $file.FullName, $destinationPath)
            }
        }
        else {
            $stats.SkippedNonRenamed++
            Write-LogEntry -Level INFO -Message ('SKIPPED_NONRENAMED file={0} notes={1}' -f $file.FullName, $apiResult.Notes)
        }
    }
    catch {
        $stats.Errors++
        Write-LogEntry -Level ERROR -Message ('FAILED file={0} error={1}' -f $file.FullName, $_.Exception.Message)
    }
}

Write-Progress -Activity '命名支援 API によるファイル処理' -Completed

$summaryObject = [pscustomobject]$stats
$summaryJson = $summaryObject | ConvertTo-Json -Depth 10 -Compress

Write-LogEntry -Level INFO -Message ('SUMMARY json={0}' -f $summaryJson)

if ($PassThru.IsPresent) {
    $summaryObject
}
