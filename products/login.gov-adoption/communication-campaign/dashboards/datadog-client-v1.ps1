# datadog-client-v1-merged.ps1
# Retains the original script's behavior and output naming, while adding:
# - Per-day checkpointing (.csv + .done marker)
# - Resume mode to skip completed days
# - Robust JSON construction (no hand-escaped JSON strings)
#
# Backwards-compatible invocation (same as OG):
#   pwsh ./datadog-client-v1-merged.ps1 "C:\path\to\output" week 1
#
# New optional flags:
#   -Resume               Skip any day that already has a .done marker
#   -StartDate/-EndDate   Override the computed window (dates, inclusive start / exclusive end)
#   -PruneOldCheckpoints  Keep only the most recent checkpoint folder for this period+window

param(
  [Parameter(Position=0, Mandatory=$true)][string]$OutputPath,
  [Parameter(Position=1, Mandatory=$true)][ValidateSet("week","month")][string]$Period,
  [Parameter(Position=2, Mandatory=$true)][int]$PeriodsBack,

  [datetime]$StartDate,
  [datetime]$EndDate,

  [switch]$Resume,
  [switch]$PruneOldCheckpoints,

  [switch]$MergeOnly,
  [switch]$Workbench

)

# -------------------------------
# OG constants (retain as-is)
# -------------------------------
$apiKey = 'UPDATE'
$applicationKey = 'UPDATE'

# Query retained from OG
$queryName = "mobileUsers"
$query = "service:vets-api AND (@message_content:*SignInController*callback OR @message_content:*SignInController*refresh*)"
$limit = 1000

# -------------------------------
# Window calculation (retain OG logic, but allow overrides)
# -------------------------------
Write-Host "$Period $PeriodsBack"

if ($Period -eq "week") {
    $fromWeek = (Get-Date).AddDays(-7 * $PeriodsBack)
    $fromDay = $fromWeek.AddDays(0 - $fromWeek.DayOfWeek)
    $toDay = $fromDay.AddDays(7)
} elseif ($Period -eq "month") {
    $fromMonth = (Get-Date).AddMonths(-$PeriodsBack)
    $toMonth = (Get-Date).AddMonths(-($PeriodsBack-1))
    $fromDay = $fromMonth.AddDays(0 - $fromMonth.Day + 1)
    $toDay = $toMonth.AddDays(0 - $toMonth.Day + 1)
} else {
    Write-Host "time period not supported $Period"
    Exit 1
}

# Optional window override (inclusive start, exclusive end)
if ($PSBoundParameters.ContainsKey('StartDate') -and $PSBoundParameters.ContainsKey('EndDate')) {
    $fromDay = $StartDate.Date
    $toDay   = $EndDate.Date
}

$fromIso = $fromDay.ToString("yyyy-MM-ddT00:00-00:00")
$toIso   = $toDay.ToString("yyyy-MM-ddT00:00-00:00")

$windowFromShort   = $fromDay.ToString("yyyyMMdd")
$windowToShort     = $toDay.ToString("yyyyMMdd")
$windowFromSlashes = $fromDay.ToString("MM/dd/yyyy")
$windowToSlashes   = $toDay.ToString("MM/dd/yyyy")

Write-Host "$fromIso $toIso"

function Get-ParentProcessName {
  try {
    $currentProcess = Get-WmiObject Win32_Process -Filter "ProcessId = $PID"
    $parent = Get-WmiObject Win32_Process -Filter "ProcessId = $($currentProcess.ParentProcessId)"
    if (-not $parent -or -not $parent.Name) { return $null }
    return $parent.Name.Split('.')[0]
  } catch {
    return $null
  }
}

$parentProcess = Get-ParentProcessName
if ($parentProcess -eq "WindowsTerminal") { $parentProcess = "shell" }

$isWorkbench = $Workbench -or ($parentProcess -eq "Workbench")
Write-Host "Parent Process: $parentProcess"
Write-Host "Workbench mode: $isWorkbench"

# -------------------------------
# Output naming (retain OG)
# -------------------------------
$fileName = "$queryName-$($Period)-$($windowFromShort)-$($windowToShort)"
$fileNameLatest = "$queryName-$($Period)"

$outfile = Join-Path $OutputPath "$fileName.csv"
$outfileLatest = Join-Path $OutputPath "$fileNameLatest-latest.csv"
$logFile = Join-Path $OutputPath "$fileNameLatest-log.txt"


# Workbench short-circuit must run BEFORE checkpoints/deletes/api calls

if ($isWorkbench) {
  Write-Host "Workbench run detected (or -Workbench provided). This script will NOT run Datadog pulls in Workbench." -ForegroundColor Yellow

  if ((Test-Path $outfile) -and (Test-Path $outfileLatest)) {
    $file1 = Get-Item $outfile
    $file2 = Get-Item $outfileLatest

    Write-Host "Found $($file1.Name) size=$($file1.Length); $($file2.Name) size=$($file2.Length)"

    if ($file1.Length -gt 0 -and $file2.Length -gt 0 -and $file1.Length -eq $file2.Length) {
      Write-Host "Files exist and sizes match. Exiting 0 so Workbench can upload." -ForegroundColor Green
      Exit 0
    }

    Write-Host "Files exist but are empty or size mismatch. Run from shell first to regenerate." -ForegroundColor Red
    Exit 1
  }

  Write-Host "One or both expected outputs are missing. Run from shell first to generate outputs, then rerun Workbench to upload." -ForegroundColor Red
  Exit 1
}


# -------------------------------
# Checkpoint folder strategy (new)
# -------------------------------
$checkpointRoot = Join-Path $OutputPath "checkpoints"
$checkpointFolder = Join-Path $checkpointRoot $fileName
New-Item -ItemType Directory -Force -Path $checkpointFolder | Out-Null

if ($PruneOldCheckpoints -and (Test-Path $checkpointRoot)) {
  Get-ChildItem $checkpointRoot -Directory |
    Where-Object { $_.FullName -ne $checkpointFolder } |
    ForEach-Object {
      try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch {}
    }
}

# If final outputs exist, remove them so the run is deterministic (retain OG behavior)
if (Test-Path $outfile) { Remove-Item $outfile -Force }
if (Test-Path $outfileLatest) { Remove-Item $outfileLatest -Force }

# -------------------------------
# Datadog request wiring (retain OG endpoint/headers)
# -------------------------------
$uri = "https://api.ddog-gov.com/api/v2/logs/events/search"
$headers = @{
    "DD-API-KEY" = $apiKey
    "DD-APPLICATION-KEY" = $applicationKey
}



# -------------------------------
# Helpers
# -------------------------------
function Process-Response {
    param(
      [Parameter(Mandatory=$true)]$json,
      [Parameter(Mandatory=$true)][hashtable]$icns,
      [Parameter(Mandatory=$true)][hashtable]$globals
    )

    $json.data | ForEach-Object {
        $icn = $_.attributes.attributes.payload.icn
        $csp = $_.attributes.attributes.payload.type
        $ial = $_.attributes.attributes.payload.ial

        if (($ial -eq 1) -or (-not $icn) -or (-not $csp)) {
            # ignore IAL1 and malformed records (retain OG intent)
        } else {
            $icncsp = "icn$($icn)csp$($csp)"
            $timestamp = $_.attributes.timestamp

            if ($icncsp) {
                $globals['totalFound']++
                $globals['lastTimestamp'] = $timestamp
                if (-not $icns.ContainsKey($icncsp)) {
                    $globals['totalUnique']++
                    $icns[$icncsp] = @{}
                    $icns[$icncsp]['icn'] = $icn
                    $icns[$icncsp]['csp'] = $csp
                }
            }
        }
    }
}

function Invoke-DdLogSearch {
  param(
    [Parameter(Mandatory=$true)][datetime]$DayStart,
    [Parameter(Mandatory=$true)][datetime]$DayEnd,
    [Parameter(Mandatory=$true)][hashtable]$icns,
    [Parameter(Mandatory=$true)][hashtable]$globals
  )

  $from = $DayStart.ToString("yyyy-MM-ddT00:00-00:00")
  $to   = $DayEnd.ToString("yyyy-MM-ddT00:00-00:00")

  $bodyObj = @{
    filter = @{
      indexes = @("*")
      query = $query
      from  = $from
      to    = $to
      storage_tier = "flex"
    }
    page = @{
      limit = $limit
    }
    sort = "timestamp"
  }

  $cursor = $null

  Write-Host "Calling DataDog query: $queryName. From: $from To: $to" -ForegroundColor Green
  "Calling DataDog query: $queryName. From: $from To: $to" >> $logFile

  while ($true) {
    if ($cursor) {
      $bodyObj.page.cursor = $cursor
    } else {
      if ($bodyObj.page.ContainsKey('cursor')) { $bodyObj.page.Remove('cursor') }
    }

    $body = $bodyObj | ConvertTo-Json -Depth 10

    $attempt = 0
    $maxAttempts = 3
    $response = $null

    while ($attempt -lt $maxAttempts) {
      try {
        $attempt++
        $response = Invoke-WebRequest -Uri $uri `
          -Method Post `
          -Headers $headers `
          -ContentType "application/json" `
          -Body $body
        break
      } catch {
        $err = $Error[0].ToString()
        Write-Warning $err
        $err >> $logFile

        # Simple retry on known transient timeouts (retain OG note + make resilient)
        if ($err -match "request_timeout|deadline_exceeded|timed out|timeout") {
          Start-Sleep -Seconds ([Math]::Min(30, (5 * $attempt)))
          if ($attempt -lt $maxAttempts) { continue }
        }

        throw
      }
    }

    $responseJSON = ConvertFrom-Json -InputObject $response.Content
    Process-Response -json $responseJSON -icns $icns -globals $globals

    $now = (Get-Date).ToString('T')
    $log = "$now Found: $($globals['totalFound']) Unique: $($globals['totalUnique']) LastTs: $($globals['lastTimestamp'])"
    Write-Host $log
    $log >> $logFile

    $cursor = $responseJSON.meta.page.after
    if (-not $cursor) { break }
  }
}

function Merge-DailyCheckpoints {
  param(
    [Parameter(Mandatory=$true)][string[]]$DailyFiles,
    [Parameter(Mandatory=$true)][string]$OutFile,
    [Parameter(Mandatory=$true)][string]$WindowFromSlashes,
    [Parameter(Mandatory=$true)][string]$WindowToSlashes,
    [int]$ProgressEveryNRows = 50000
  )

  $dedup = @{}  # key: "icn|csp" => first-seen normalized row
  $totalRows = 0
  $filesTotal = $DailyFiles.Count
  $filesDone = 0
  $lastReportAt = 0

  foreach ($f in $DailyFiles) {
    $filesDone++

    Write-Progress -Activity "Merging daily checkpoints" `
      -Status "File $filesDone / $filesTotal: $(Split-Path $f -Leaf)" `
      -PercentComplete ([int](($filesDone / $filesTotal) * 100))

    if (-not (Test-Path $f)) { continue }

    Import-Csv $f | ForEach-Object {
      $totalRows++

      $key = "$($_.icn)|$($_.csp)"
      if (-not $dedup.ContainsKey($key)) {
        $dedup[$key] = [pscustomobject]@{
          fromDate = $WindowFromSlashes
          toDate   = $WindowToSlashes
          icn      = $_.icn
          csp      = $_.csp
        }
      }

      if ($totalRows -ge ($lastReportAt + $ProgressEveryNRows)) {
        $lastReportAt = $totalRows
        Write-Host ("Merge progress: rows read={0:n0}, unique keys={1:n0}" -f $totalRows, $dedup.Count)
      }
    }
  }

  Write-Progress -Activity "Merging daily checkpoints" -Completed

  $dedup.Values |
    Sort-Object icn, csp |
    Export-Csv $OutFile -NoTypeInformation

  Write-Host ("Merge complete: rows read={0:n0}, unique={1:n0}, out={2}" -f $totalRows, $dedup.Count, $OutFile) -ForegroundColor Green
}

# -------------------------------
# MERGE-ONLY MODE (new)
# -------------------------------
if ($MergeOnly) {
  Write-Host "MergeOnly enabled: skipping Datadog pulls; merging existing checkpoints..." -ForegroundColor Yellow

  $dailyFiles = Get-ChildItem $checkpointFolder -Filter "$queryName-$Period-*.csv" -File |
    Sort-Object Name |
    Select-Object -ExpandProperty FullName

  if (-not $dailyFiles -or $dailyFiles.Count -eq 0) {
    Write-Warning "No daily checkpoint CSVs found in: $checkpointFolder"
    Exit 1
  }

  Merge-DailyCheckpoints `
    -DailyFiles $dailyFiles `
    -OutFile $outfile `
    -WindowFromSlashes $windowFromSlashes `
    -WindowToSlashes $windowToSlashes

  Copy-Item $outfile -Destination $outfileLatest -Force
  Write-Host "Done. Final: $outfileLatest" -ForegroundColor Green
  Exit 0
}

# -------------------------------
# PER-DAY LOOP (new)
# -------------------------------
Write-Host "Checkpoint folder: $checkpointFolder" -ForegroundColor Cyan
"Checkpoint folder: $checkpointFolder" >> $logFile

$day = $fromDay.Date
$dailyFiles = New-Object System.Collections.Generic.List[string]

while ($day -lt $toDay.Date) {
  $dayStart = $day
  $dayEnd = $day.AddDays(1)

  $dayShort = $dayStart.ToString("yyyyMMdd")
  $dayFromSlashes = $dayStart.ToString("MM/dd/yyyy")
  $dayToSlashes = $dayEnd.ToString("MM/dd/yyyy")

  $dailyCsv = Join-Path $checkpointFolder "$queryName-$Period-$dayShort.csv"
  $doneMarker = Join-Path $checkpointFolder "$queryName-$Period-$dayShort.done"

  if ($Resume -and (Test-Path $doneMarker) -and (Test-Path $dailyCsv)) {
    Write-Host "Skipping completed day $dayShort (resume)" -ForegroundColor Yellow
    $dailyFiles.Add($dailyCsv) | Out-Null
    $day = $day.AddDays(1)
    continue
  }

  Write-Host "Processing day $dayShort..." -ForegroundColor Green

  # Day-scoped hashtables (so a single bad day doesn't poison the whole run)
  $icnsDay = @{}
  $globalsDay = @{
      'totalFound' = 0
      'totalUnique' = 0
      'lastTimestamp' = ''
  }

  try {
    Invoke-DdLogSearch -DayStart $dayStart -DayEnd $dayEnd -icns $icnsDay -globals $globalsDay
  } catch {
    Write-Warning "Failed day $dayShort. Leaving checkpoints intact for resume. Error: $($Error[0])"
    "Failed day $dayShort. Error: $($Error[0])" >> $logFile
    Exit 1
  }

  # Write daily CSV (unique by icn+csp within the day)
  $icnsDay.Values |
    Select-Object @{Name='fromDate';Expression={$dayFromSlashes}},
                  @{Name='toDate';Expression={$dayToSlashes}},
                  icn,
                  csp |
    Export-Csv $dailyCsv -NoTypeInformation

  New-Item -ItemType File -Force -Path $doneMarker | Out-Null

  $dailyFiles.Add($dailyCsv) | Out-Null
  $day = $day.AddDays(1)
}

if ($dailyFiles.Count -eq 0) {
  Write-Warning "No daily files produced; exiting."
  Exit 1
}

# -------------------------------
# MERGE DAILY -> FINAL (new, but keeps OG final shape)
# -------------------------------
Write-Host "Merging daily checkpoints into final outputs..." -ForegroundColor Green

$dailyFiles = $dailyFiles.ToArray()  # ensure string[]

Merge-DailyCheckpoints `
  -DailyFiles $dailyFiles `
  -OutFile $outfile `
  -WindowFromSlashes $windowFromSlashes `
  -WindowToSlashes $windowToSlashes


Copy-Item $outfile -Destination $outfileLatest -Force

Write-Host "Done. Final: $outfileLatest" -ForegroundColor Green
Exit 0
