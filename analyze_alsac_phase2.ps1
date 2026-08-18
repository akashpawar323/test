param(
    [Parameter(Position=0)]
    [string]$FolderPath = (Get-Location).Path
)

<#
ALSAC Phase-2 LinearB CSV Analyzer (PowerShell)
-----------------------------------------------
Reads the Phase-2 CSV exports from one folder and generates aggregated,
screenshot-friendly reports WITHOUT uploading any company data anywhere.

Expected files (all optional; script works with whatever is present):
  01_alsac_overall_baseline.csv
  02_alsac_ai_coding.csv
  03_alsac_ai_code_review.csv
  04_alsac_agentic_pr.csv
  05_alsac_flow_metrics.csv
  06_alsac_developer_metrics.csv
  07_alsac_team_metrics.csv
  08_alsac_repository_metrics.csv

Outputs:
  ALSAC_ANALYSIS_SUMMARY.md
  ALSAC_SCREENSHOT_SUMMARY.txt
  ALSAC_ANALYSIS_SUMMARY.json
  ALSAC_SCHEMA_AUDIT.txt

Usage:
  .\analyze_alsac_phase2.ps1
  .\analyze_alsac_phase2.ps1 "C:\path\to\phase_2_alsac_data"

If execution policy blocks the script:
  powershell -ExecutionPolicy Bypass -File .\analyze_alsac_phase2.ps1
#>

$ErrorActionPreference = "Stop"

$TargetStart = [datetime]"2026-02-01"
$TargetEnd   = [datetime]"2026-05-31"

$ExpectedFiles = @(
    "01_alsac_overall_baseline.csv",
    "02_alsac_ai_coding.csv",
    "03_alsac_ai_code_review.csv",
    "04_alsac_agentic_pr.csv",
    "05_alsac_flow_metrics.csv",
    "06_alsac_developer_metrics.csv",
    "07_alsac_team_metrics.csv",
    "08_alsac_repository_metrics.csv"
)

$Benchmarks = [ordered]@{
    "AI coding-assisted PR %" = [ordered]@{ Elite = 54.0; Good = 35.0; Fair = 20.0 }
    "AI code-review PR %" = [ordered]@{ Elite = 57.0; Good = 26.0; Fair = 8.0 }
    "Agentic PR %" = [ordered]@{ Elite = 4.7; Good = 1.1; Fair = 0.1 }
    "PR merge rate / dev / week" = [ordered]@{ Elite = 2.6; Good = 1.9; Fair = 1.3 }
    "PR yield, all PRs %" = [ordered]@{ Elite = 90.0; Good = 86.0; Fair = 81.0 }
    "PR yield, agentic PRs %" = [ordered]@{ Elite = 79.0; Good = 58.0; Fair = 37.0 }
}

function Get-FieldValue {
    param(
        [Parameter(Mandatory=$true)]$Row,
        [Parameter(Mandatory=$true)][string]$Column
    )
    $prop = $Row.PSObject.Properties[$Column]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Convert-ToNumber {
    param($Value)
    if ($null -eq $Value) { return $null }
    $s = ([string]$Value).Trim().Replace(",", "")
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    if ($s.ToLowerInvariant() -in @("none","null","nan","n/a")) { return $null }

    $n = 0.0
    if ([double]::TryParse(
        $s,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$n
    )) {
        return $n
    }
    return $null
}

function Convert-ToDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $s = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }

    $formats = @("M/d/yyyy","MM/dd/yyyy","yyyy-MM-dd","M/d/yy","MM/dd/yy")
    foreach ($fmt in $formats) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParseExact(
            $s,
            $fmt,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$dt
        )) {
            return $dt
        }
    }

    $fallback = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$fallback)) {
        return $fallback
    }
    return $null
}

function Get-RowsInWindow {
    param(
        [Parameter(Mandatory=$true)]$Rows,
        [bool]$FullIntervalsOnly = $false
    )

    $out = @()
    foreach ($r in $Rows) {
        $after = Convert-ToDate (Get-FieldValue $r "after")
        $before = Convert-ToDate (Get-FieldValue $r "before")
        if ($null -eq $after -or $null -eq $before) { continue }
        if ($after -lt $TargetStart -or $before -gt $TargetEnd) { continue }
        if ($FullIntervalsOnly -and $after -ge $before) { continue }
        $out += $r
    }
    return $out
}

function Get-Sum {
    param($Rows, [string]$Column)
    $sum = 0.0
    foreach ($r in $Rows) {
        $v = Convert-ToNumber (Get-FieldValue $r $Column)
        if ($null -ne $v) { $sum += $v }
    }
    return $sum
}

function Get-Mean {
    param($Rows, [string]$Column)
    $vals = @()
    foreach ($r in $Rows) {
        $v = Convert-ToNumber (Get-FieldValue $r $Column)
        if ($null -ne $v) { $vals += $v }
    }
    if ($vals.Count -eq 0) { return $null }
    return ($vals | Measure-Object -Average).Average
}

function Get-Median {
    param($Rows, [string]$Column)
    $vals = @()
    foreach ($r in $Rows) {
        $v = Convert-ToNumber (Get-FieldValue $r $Column)
        if ($null -ne $v) { $vals += $v }
    }
    if ($vals.Count -eq 0) { return $null }
    $sorted = @($vals | Sort-Object)
    $n = $sorted.Count
    if ($n % 2 -eq 1) {
        return [double]$sorted[[int]($n / 2)]
    }
    return ([double]$sorted[$n/2 - 1] + [double]$sorted[$n/2]) / 2.0
}

function Safe-Divide {
    param($A, $B)
    if ($null -eq $A -or $null -eq $B) { return $null }
    if ([double]$B -eq 0) { return $null }
    return ([double]$A / [double]$B)
}

function Format-Number {
    param($Value, [int]$Digits = 1)
    if ($null -eq $Value) { return "N/A" }
    $d = [double]$Value
    if ([math]::Abs($d - [math]::Round($d)) -lt 0.000000001) {
        return ([math]::Round($d)).ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $d.ToString("N$Digits", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-Percent {
    param($Value, [int]$Digits = 1)
    if ($null -eq $Value) { return "N/A" }
    return ("{0:F$Digits}%" -f [double]$Value)
}

function Test-Column {
    param($Columns, [string]$Column)
    return ($Columns -contains $Column)
}

function Get-GroupColumn {
    param($Columns, [string]$FileName)

    $pref = @()
    if ($FileName -like "*ai_coding*") {
        $pref += "coding_assistant"
    } elseif ($FileName -like "*ai_code_review*") {
        $pref += "ai_review"
    } elseif ($FileName -like "*agentic*") {
        $pref += "agentic_pr"
    } elseif ($FileName -like "*developer*") {
        $pref += @("unified_user_id","unified_user","user_id")
    } elseif ($FileName -like "*team*") {
        $pref += @("team_id","team")
    } elseif ($FileName -like "*repository*") {
        $pref += @("repository_id","repository")
    }

    $pref += @(
        "coding_assistant","ai_review","agentic_pr",
        "unified_user_id","unified_user","user_id",
        "team_id","team","repository_id","repository","organization_id"
    )

    foreach ($c in $pref) {
        if ($Columns -contains $c) { return $c }
    }

    foreach ($c in $Columns) {
        $lc = $c.ToLowerInvariant()
        if ($c -notin @("after","before") -and
            ($lc.EndsWith("_id") -or $lc.Contains("assistant") -or $lc.Contains("review") -or $lc.Contains("agent"))) {
            return $c
        }
    }
    return $null
}

function Test-ManualGroup {
    param($Value)
    $s = ([string]$Value).Trim().ToLowerInvariant()
    return ($s -in @("manual","without ai","no ai","none","non-ai","non_ai"))
}

function Get-BenchmarkTier {
    param($Value, [string]$MetricName)
    if ($null -eq $Value -or -not $Benchmarks.Contains($MetricName)) { return "N/A" }
    $b = $Benchmarks[$MetricName]
    if ([double]$Value -ge [double]$b.Elite) { return "Elite or above" }
    if ([double]$Value -ge [double]$b.Good)  { return "Good" }
    if ([double]$Value -ge [double]$b.Fair)  { return "Fair" }
    return "Needs Focus"
}

function ConvertTo-MarkdownTable {
    param(
        [string[]]$Headers,
        [object[]]$Rows
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $safeHeaders = $Headers | ForEach-Object { ([string]$_).Replace("|","\|") }
    $lines.Add("| " + ($safeHeaders -join " | ") + " |")
    $lines.Add("| " + ((1..$Headers.Count | ForEach-Object { "---" }) -join " | ") + " |")
    foreach ($row in $Rows) {
        $cells = @($row | ForEach-Object { ([string]$_).Replace("|","\|") })
        $lines.Add("| " + ($cells -join " | ") + " |")
    }
    return ($lines -join [Environment]::NewLine)
}

function Analyze-Baseline {
    param($Rows, $Columns)

    $allWindow = @(Get-RowsInWindow $Rows $false)
    $full = @(Get-RowsInWindow $Rows $true)

    $metricNames = @(
        "pr.new","pr.merged","pr.reviewed","pr.reviews",
        "pr.merged.without.review.count","commit.total.count",
        "commit.total_changes","commit.activity.rework.count",
        "commit.code_churn.rework","contributor.coding_days",
        "releases.count"
    )

    $metrics = [ordered]@{}
    foreach ($c in $metricNames) {
        if (Test-Column $Columns $c) {
            $metrics[$c] = [ordered]@{
                all_window_total = Get-Sum $allWindow $c
                full_interval_total = Get-Sum $full $c
                full_interval_weekly_avg = Get-Mean $full $c
            }
        }
    }

    $opened = if ($metrics.Contains("pr.new")) { $metrics["pr.new"].full_interval_total } else { $null }
    $merged = if ($metrics.Contains("pr.merged")) { $metrics["pr.merged"].full_interval_total } else { $null }
    $mwr = if ($metrics.Contains("pr.merged.without.review.count")) { $metrics["pr.merged.without.review.count"].full_interval_total } else { $null }

    $mergeRatio = Safe-Divide $merged $opened
    $woReview = Safe-Divide $mwr $merged

    return [ordered]@{
        rows_all_window = $allWindow.Count
        rows_full_intervals = $full.Count
        metrics = $metrics
        raw_merge_open_ratio_pct = if ($null -ne $mergeRatio) { $mergeRatio * 100 } else { $null }
        merged_without_review_share_pct = if ($null -ne $woReview) { $woReview * 100 } else { $null }
    }
}

function Analyze-Grouped {
    param($Rows, $Columns, [string]$FileName)

    $full = @(Get-RowsInWindow $Rows $true)
    $groupCol = Get-GroupColumn $Columns $FileName
    if ($null -eq $groupCol) {
        return [ordered]@{
            group_column = $null
            groups = [ordered]@{}
            warning = "Could not detect grouping column."
        }
    }

    $groupBuckets = [ordered]@{}
    foreach ($r in $full) {
        $g = [string](Get-FieldValue $r $groupCol)
        $g = $g.Trim()
        if (-not $groupBuckets.Contains($g)) {
            $groupBuckets[$g] = New-Object System.Collections.ArrayList
        }
        [void]$groupBuckets[$g].Add($r)
    }

    $outGroups = [ordered]@{}
    foreach ($g in $groupBuckets.Keys) {
        $gr = @($groupBuckets[$g])
        $opened = if (Test-Column $Columns "pr.new") { Get-Sum $gr "pr.new" } else { $null }
        $merged = if (Test-Column $Columns "pr.merged") { Get-Sum $gr "pr.merged" } else { $null }
        $ratio = Safe-Divide $merged $opened

        $obj = [ordered]@{
            rows = $gr.Count
            "pr.new" = $opened
            "pr.merged" = $merged
            raw_merge_open_ratio_pct = if ($null -ne $ratio) { $ratio * 100 } else { $null }
        }

        foreach ($c in @(
            "pr.reviewed","pr.reviews","pr.maturity_ratio",
            "pr.merged.without.review.count","commit.total.count",
            "commit.total_changes","commit.activity.rework.count",
            "commit.code_churn.rework","contributor.coding_days","releases.count"
        )) {
            if (Test-Column $Columns $c) {
                if ($c -like "*ratio*") {
                    $obj[$c] = Get-Mean $gr $c
                } else {
                    $obj[$c] = Get-Sum $gr $c
                }
            }
        }

        $outGroups[$g] = $obj
    }

    $totalOpened = 0.0
    $manualOpened = 0.0
    foreach ($g in $outGroups.Keys) {
        $v = $outGroups[$g]["pr.new"]
        if ($null -ne $v) {
            $totalOpened += [double]$v
            if (Test-ManualGroup $g) { $manualOpened += [double]$v }
        }
    }
    $nonManualOpened = $totalOpened - $manualOpened
    $share = Safe-Divide $nonManualOpened $totalOpened

    return [ordered]@{
        group_column = $groupCol
        groups = $outGroups
        grouped_pr_open_total = $totalOpened
        manual_pr_open_total = $manualOpened
        nonmanual_pr_open_total = $nonManualOpened
        nonmanual_share_of_grouped_opened_pct = if ($null -ne $share) { $share * 100 } else { $null }
        warning = "Workflow tool groups may overlap if a PR used more than one AI tool. The non-manual share is diagnostic, not automatically the official adoption percentage."
    }
}

function Analyze-Developers {
    param($Rows, $Columns)

    $full = @(Get-RowsInWindow $Rows $true)
    $userCol = Get-GroupColumn $Columns "developer"
    if ($null -eq $userCol) {
        return [ordered]@{ warning = "Could not detect unified user column." }
    }

    $users = @{}
    $activeUserWeeks = 0
    $mergedTotal = 0.0
    $openedTotal = 0.0

    foreach ($r in $full) {
        $uid = ([string](Get-FieldValue $r $userCol)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($uid)) { $users[$uid] = $true }

        $codingDays = Convert-ToNumber (Get-FieldValue $r "contributor.coding_days")
        $opened = Convert-ToNumber (Get-FieldValue $r "pr.new")
        $merged = Convert-ToNumber (Get-FieldValue $r "pr.merged")
        $commits = Convert-ToNumber (Get-FieldValue $r "commit.total.count")
        $changes = Convert-ToNumber (Get-FieldValue $r "commit.total_changes")

        if ($null -eq $opened) { $opened = 0 }
        if ($null -eq $merged) { $merged = 0 }
        if ($null -eq $commits) { $commits = 0 }
        if ($null -eq $changes) { $changes = 0 }

        $isActive = (($null -ne $codingDays -and $codingDays -gt 0) -or
                     $opened -gt 0 -or $merged -gt 0 -or $commits -gt 0 -or $changes -gt 0)

        if (-not [string]::IsNullOrWhiteSpace($uid) -and $isActive) {
            $activeUserWeeks++
        }

        $openedTotal += $opened
        $mergedTotal += $merged
    }

    $perDevWeek = Safe-Divide $mergedTotal $activeUserWeeks
    $gap = if ($null -ne $perDevWeek) { 2.6 - $perDevWeek } else { $null }

    return [ordered]@{
        user_column = $userCol
        unique_users = $users.Keys.Count
        active_user_weeks = $activeUserWeeks
        pr_open_total = $openedTotal
        pr_merged_total = $mergedTotal
        pr_merged_per_active_dev_week = $perDevWeek
        benchmark_tier = Get-BenchmarkTier $perDevWeek "PR merge rate / dev / week"
        gap_to_elite = $gap
    }
}

function Analyze-Flow {
    param($Rows, $Columns)
    $full = @(Get-RowsInWindow $Rows $true)
    $result = [ordered]@{}

    foreach ($c in $Columns) {
        if ($c -in @("after","before")) { continue }
        if ($c -like "*_id") { continue }

        if ($c -match "branch\.time_to_pr|branch\.time_to_review|branch\.review_time|branch\.time_to_merge|branch\.time_to_prod|branch\.computed\.cycle_time") {
            $mean = Get-Mean $full $c
            $median = Get-Median $full $c
            if ($null -ne $mean -or $null -ne $median) {
                $result[$c] = [ordered]@{
                    weekly_mean_minutes = $mean
                    weekly_median_minutes = $median
                }
            }
        }
    }
    return $result
}

function Get-TopEntities {
    param($Rows, $Columns, [string]$FileName, [int]$TopN = 10)

    $full = @(Get-RowsInWindow $Rows $true)
    $groupCol = Get-GroupColumn $Columns $FileName
    if ($null -eq $groupCol) {
        return [ordered]@{ group_column = $null; items = @() }
    }

    $agg = @{}
    foreach ($r in $full) {
        $g = ([string](Get-FieldValue $r $groupCol)).Trim()
        if ([string]::IsNullOrWhiteSpace($g)) { continue }

        if (-not $agg.ContainsKey($g)) {
            $agg[$g] = [ordered]@{
                entity = $g
                "pr.new" = 0.0
                "pr.merged" = 0.0
                "commit.total_changes" = 0.0
                "releases.count" = 0.0
                "contributor.coding_days" = 0.0
            }
        }

        foreach ($c in @("pr.new","pr.merged","commit.total_changes","releases.count","contributor.coding_days")) {
            if (Test-Column $Columns $c) {
                $v = Convert-ToNumber (Get-FieldValue $r $c)
                if ($null -ne $v) { $agg[$g][$c] += $v }
            }
        }
    }

    $items = @()
    foreach ($g in $agg.Keys) {
        $obj = $agg[$g]
        $ratio = Safe-Divide $obj["pr.merged"] $obj["pr.new"]
        $items += [pscustomobject]@{
            entity = $obj.entity
            "pr.new" = $obj["pr.new"]
            "pr.merged" = $obj["pr.merged"]
            "commit.total_changes" = $obj["commit.total_changes"]
            "releases.count" = $obj["releases.count"]
            "contributor.coding_days" = $obj["contributor.coding_days"]
            raw_merge_open_ratio_pct = if ($null -ne $ratio) { $ratio * 100 } else { $null }
        }
    }

    $items = @($items | Sort-Object -Property @{Expression={ $_.'pr.merged' }; Descending=$true} | Select-Object -First $TopN)

    return [ordered]@{
        group_column = $groupCol
        items = $items
    }
}

# Resolve folder
$Folder = (Resolve-Path -LiteralPath $FolderPath).Path

# Load available files
$FileData = [ordered]@{}
foreach ($fname in $ExpectedFiles) {
    $p = Join-Path $Folder $fname
    if (Test-Path -LiteralPath $p) {
        $rows = @(Import-Csv -LiteralPath $p)
        $columns = @()
        if ($rows.Count -gt 0) {
            $columns = @($rows[0].PSObject.Properties.Name)
        }
        $FileData[$fname] = [ordered]@{
            rows = $rows
            columns = $columns
        }
    }
}

if ($FileData.Count -eq 0) {
    throw "No expected Phase-2 CSV files found in: $Folder"
}

$Analysis = [ordered]@{
    benchmark_window = [ordered]@{
        start = $TargetStart.ToString("yyyy-MM-dd")
        end = $TargetEnd.ToString("yyyy-MM-dd")
    }
    files_found = @($FileData.Keys)
    files_missing = @($ExpectedFiles | Where-Object { -not $FileData.Contains($_) })
    benchmarks = $Benchmarks
}

if ($FileData.Contains("01_alsac_overall_baseline.csv")) {
    $d = $FileData["01_alsac_overall_baseline.csv"]
    $Analysis["overall_baseline"] = Analyze-Baseline $d.rows $d.columns
}

foreach ($pair in @(
    @("02_alsac_ai_coding.csv","ai_coding"),
    @("03_alsac_ai_code_review.csv","ai_code_review"),
    @("04_alsac_agentic_pr.csv","agentic_pr")
)) {
    $fname = $pair[0]
    $key = $pair[1]
    if ($FileData.Contains($fname)) {
        $d = $FileData[$fname]
        $Analysis[$key] = Analyze-Grouped $d.rows $d.columns $fname
    }
}

if ($FileData.Contains("05_alsac_flow_metrics.csv")) {
    $d = $FileData["05_alsac_flow_metrics.csv"]
    $Analysis["flow_metrics"] = Analyze-Flow $d.rows $d.columns
}

if ($FileData.Contains("06_alsac_developer_metrics.csv")) {
    $d = $FileData["06_alsac_developer_metrics.csv"]
    $Analysis["developer_metrics"] = Analyze-Developers $d.rows $d.columns
}

if ($FileData.Contains("07_alsac_team_metrics.csv")) {
    $d = $FileData["07_alsac_team_metrics.csv"]
    $Analysis["top_teams"] = Get-TopEntities $d.rows $d.columns "team" 10
}

if ($FileData.Contains("08_alsac_repository_metrics.csv")) {
    $d = $FileData["08_alsac_repository_metrics.csv"]
    $Analysis["top_repositories"] = Get-TopEntities $d.rows $d.columns "repository" 10
}

# JSON
$JsonPath = Join-Path $Folder "ALSAC_ANALYSIS_SUMMARY.json"
$Analysis | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $JsonPath -Encoding UTF8

# Schema audit
$Audit = New-Object System.Collections.Generic.List[string]
foreach ($fname in $ExpectedFiles) {
    if (-not $FileData.Contains($fname)) {
        $Audit.Add("MISSING: $fname")
        continue
    }
    $d = $FileData[$fname]
    $win = @(Get-RowsInWindow $d.rows $false)
    $full = @(Get-RowsInWindow $d.rows $true)
    $Audit.Add("OK: $fname")
    $Audit.Add("  rows: $($d.rows.Count)")
    $Audit.Add("  columns ($($d.columns.Count)): $($d.columns -join ', ')")
    $Audit.Add("  rows inside Feb 1-May 31: $($win.Count)")
    $Audit.Add("  full-interval rows used for weekly analysis: $($full.Count)")
}
$AuditPath = Join-Path $Folder "ALSAC_SCHEMA_AUDIT.txt"
$Audit | Set-Content -LiteralPath $AuditPath -Encoding UTF8

# Markdown report
$Md = New-Object System.Collections.Generic.List[string]
$Md.Add("# ALSAC Phase-2 LinearB Analysis")
$Md.Add("")
$Md.Add("Benchmark window: **2026-02-01 to 2026-05-31**")
$Md.Add("")
$Md.Add("> IMPORTANT: ``merged/opened`` ratios below are diagnostics only. They are **not** LinearB's official 30-day PR Yield unless a true 30-day yield metric is available.")
$Md.Add("")

if ($Analysis.Contains("overall_baseline")) {
    $b = $Analysis["overall_baseline"]
    $m = $b.metrics
    $Md.Add("## 1. Overall Baseline")

    $rows = @()
    foreach ($pair in @(
        @("pr.new","PRs opened"),
        @("pr.merged","PRs merged"),
        @("pr.reviewed","PRs reviewed"),
        @("pr.reviews","Review events"),
        @("pr.merged.without.review.count","Merged without review"),
        @("commit.total.count","Commits"),
        @("commit.total_changes","Code changes"),
        @("contributor.coding_days","Coding days"),
        @("releases.count","Releases")
    )) {
        $key = $pair[0]; $label = $pair[1]
        if ($m.Contains($key)) {
            $rows += ,@(
                $label,
                (Format-Number $m[$key].full_interval_total 0),
                (Format-Number $m[$key].full_interval_weekly_avg 1)
            )
        }
    }
    $Md.Add((ConvertTo-MarkdownTable @("Metric","Full-week total","Avg / full week") $rows))
    $Md.Add("")
    $Md.Add("- Raw merged/opened diagnostic: **$(Format-Percent $b.raw_merge_open_ratio_pct)**")
    $Md.Add("- Merged-without-review share: **$(Format-Percent $b.merged_without_review_share_pct)**")
    $Md.Add("")
}

function Add-WorkflowMarkdownSection {
    param([string]$Title, [string]$Key)

    if (-not $Analysis.Contains($Key)) { return }
    $info = $Analysis[$Key]
    $Md.Add("## $Title")
    $Md.Add("Grouping column: ``$($info.group_column)``")

    $rows = @()
    foreach ($g in $info.groups.Keys) {
        $v = $info.groups[$g]
        $rows += ,@(
            $g,
            (Format-Number $v["pr.new"] 0),
            (Format-Number $v["pr.merged"] 0),
            (Format-Percent $v.raw_merge_open_ratio_pct),
            (Format-Number $v["commit.total_changes"] 0)
        )
    }

    $Md.Add((ConvertTo-MarkdownTable @("Group","PR Opened","PR Merged","Merged/Open diagnostic","Code changes") $rows))
    $Md.Add("")
    $Md.Add("- Non-manual share of grouped opened PRs (diagnostic): **$(Format-Percent $info.nonmanual_share_of_grouped_opened_pct)**")
    $Md.Add("- Note: $($info.warning)")
    $Md.Add("")
}

Add-WorkflowMarkdownSection "2. AI Coding Breakdown" "ai_coding"
Add-WorkflowMarkdownSection "3. AI Code Review Breakdown" "ai_code_review"
Add-WorkflowMarkdownSection "4. Agentic PR Breakdown" "agentic_pr"

if ($Analysis.Contains("developer_metrics")) {
    $dev = $Analysis["developer_metrics"]
    $Md.Add("## 5. Developer-Normalized Engineering Leverage")
    $Md.Add("- Unique users in export: **$(Format-Number $dev.unique_users 0)**")
    $Md.Add("- Active developer-weeks: **$(Format-Number $dev.active_user_weeks 0)**")
    $Md.Add("- Merged PRs per active developer-week: **$(Format-Number $dev.pr_merged_per_active_dev_week 3)**")
    $Md.Add("- LinearB benchmark tier (approx.): **$($dev.benchmark_tier)**")
    $Md.Add("- Gap to Elite 2.6 PR/dev/week: **$(Format-Number $dev.gap_to_elite 3)**")
    $Md.Add("")
}

if ($Analysis.Contains("flow_metrics")) {
    $flow = $Analysis["flow_metrics"]
    $Md.Add("## 6. Flow Metrics")
    $rows = @()
    foreach ($k in $flow.Keys) {
        $v = $flow[$k]
        $rows += ,@(
            $k,
            (Format-Number $v.weekly_mean_minutes 1),
            (Format-Number $v.weekly_median_minutes 1)
        )
    }
    $Md.Add((ConvertTo-MarkdownTable @("Metric","Mean of weekly values (min)","Median of weekly values (min)") $rows))
    $Md.Add("")
    $Md.Add("> These are summaries of weekly p50/p75/etc. values, not recomputed whole-period percentiles.")
    $Md.Add("")
}

foreach ($pair in @(
    @("7. Top Teams by Merged PRs","top_teams"),
    @("8. Top Repositories by Merged PRs","top_repositories")
)) {
    $title = $pair[0]; $key = $pair[1]
    if ($Analysis.Contains($key) -and $Analysis[$key].items.Count -gt 0) {
        $Md.Add("## $title")
        $rows = @()
        foreach ($item in $Analysis[$key].items) {
            $rows += ,@(
                $item.entity,
                (Format-Number $item.'pr.new' 0),
                (Format-Number $item.'pr.merged' 0),
                (Format-Percent $item.raw_merge_open_ratio_pct),
                (Format-Number $item.'commit.total_changes' 0)
            )
        }
        $Md.Add((ConvertTo-MarkdownTable @("ID/Group","PR Opened","PR Merged","Merged/Open diagnostic","Code changes") $rows))
        $Md.Add("")
    }
}

$Md.Add("## 9. LinearB Reference Benchmarks")
$benchRows = @()
foreach ($metric in $Benchmarks.Keys) {
    $b = $Benchmarks[$metric]
    $benchRows += ,@($metric, $b.Elite, $b.Good, $b.Fair)
}
$Md.Add((ConvertTo-MarkdownTable @("Metric","Elite","Good","Fair") $benchRows))
$Md.Add("")

$Md.Add("## 10. Still Needed for Final Task")
$Md.Add("- Exact 30-day PR Yield (all PRs), if LinearB exposes it in another report/export.")
$Md.Add("- Exact agentic 30-day PR Yield.")
$Md.Add("- Official AI-assisted PR %, AI-review PR %, and agentic PR % if workflow groups overlap.")
$Md.Add("- % developers by Very High / High / Moderate AI usage band.")
$Md.Add("- % merged code lines written by AI.")
$Md.Add("- Token spend / developer / month.")
$Md.Add("- AI license/platform spend and loaded people cost for ROI.")

$MdPath = Join-Path $Folder "ALSAC_ANALYSIS_SUMMARY.md"
$Md | Set-Content -LiteralPath $MdPath -Encoding UTF8

# Compact screenshot-friendly report
$Ss = New-Object System.Collections.Generic.List[string]
$Ss.Add("=" * 72)
$Ss.Add("ALSAC PHASE-2 ANALYSIS - SCREENSHOT SUMMARY")
$Ss.Add("=" * 72)
$Ss.Add("Window: 2026-02-01 to 2026-05-31")
$Ss.Add("")

if ($Analysis.Contains("overall_baseline")) {
    $b = $Analysis["overall_baseline"]
    $m = $b.metrics
    $Ss.Add("[OVERALL BASELINE - FULL INTERVALS]")
    foreach ($pair in @(
        @("pr.new","PRs opened"),
        @("pr.merged","PRs merged"),
        @("pr.reviewed","PRs reviewed"),
        @("pr.reviews","Review events"),
        @("pr.merged.without.review.count","Merged without review"),
        @("commit.total_changes","Code changes"),
        @("releases.count","Releases")
    )) {
        $key = $pair[0]; $label = $pair[1]
        if ($m.Contains($key)) {
            $Ss.Add(("{0,-28}: {1}" -f $label, (Format-Number $m[$key].full_interval_total 0)))
        }
    }
    $Ss.Add(("{0,-28}: {1}" -f "Merged/open diagnostic", (Format-Percent $b.raw_merge_open_ratio_pct)))
    $Ss.Add(("{0,-28}: {1}" -f "Merged w/o review share", (Format-Percent $b.merged_without_review_share_pct)))
    $Ss.Add("")
}

if ($Analysis.Contains("developer_metrics")) {
    $dev = $Analysis["developer_metrics"]
    $Ss.Add("[DEVELOPER-NORMALIZED]")
    $Ss.Add(("{0,-28}: {1}" -f "Unique users", (Format-Number $dev.unique_users 0)))
    $Ss.Add(("{0,-28}: {1}" -f "Active developer-weeks", (Format-Number $dev.active_user_weeks 0)))
    $Ss.Add(("{0,-28}: {1}" -f "Merged PR / active dev-week", (Format-Number $dev.pr_merged_per_active_dev_week 3)))
    $Ss.Add(("{0,-28}: {1}" -f "Benchmark tier", $dev.benchmark_tier))
    $Ss.Add("")
}

foreach ($pair in @(
    @("AI CODING","ai_coding"),
    @("AI CODE REVIEW","ai_code_review"),
    @("AGENTIC PR","agentic_pr")
)) {
    $title = $pair[0]; $key = $pair[1]
    if ($Analysis.Contains($key)) {
        $info = $Analysis[$key]
        $Ss.Add("[$title]")
        $Ss.Add("Group column: $($info.group_column)")
        foreach ($g in $info.groups.Keys) {
            $v = $info.groups[$g]
            $Ss.Add(
                "$g : opened=$(Format-Number $v["pr.new"] 0), " +
                "merged=$(Format-Number $v["pr.merged"] 0), " +
                "merge/open=$(Format-Percent $v.raw_merge_open_ratio_pct)"
            )
        }
        $Ss.Add("Non-manual grouped share (diagnostic): $(Format-Percent $info.nonmanual_share_of_grouped_opened_pct)")
        $Ss.Add("")
    }
}

$Ss.Add("IMPORTANT:")
$Ss.Add("- merged/open is NOT official 30-day PR Yield.")
$Ss.Add("- workflow tool groups may overlap; grouped AI share is diagnostic.")
$Ss.Add("- final ROI still needs AI spend + people cost.")
$Ss.Add("=" * 72)

$SsPath = Join-Path $Folder "ALSAC_SCREENSHOT_SUMMARY.txt"
$Ss | Set-Content -LiteralPath $SsPath -Encoding UTF8

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Folder: $Folder"
Write-Host "Created:"
Write-Host "  - ALSAC_ANALYSIS_SUMMARY.md"
Write-Host "  - ALSAC_SCREENSHOT_SUMMARY.txt"
Write-Host "  - ALSAC_ANALYSIS_SUMMARY.json"
Write-Host "  - ALSAC_SCHEMA_AUDIT.txt"
Write-Host ""
Write-Host "Open ALSAC_SCREENSHOT_SUMMARY.txt first - it is designed for easy screenshots." -ForegroundColor Cyan
