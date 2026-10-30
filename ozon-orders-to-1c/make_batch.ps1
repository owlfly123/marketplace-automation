# make_batch.ps1
# Batch version for many A files (CSV or XLSX) in a folder.
# For each A:
#   - Build B.xlsx with columns:
#       1) $BUYER_INN (text)
#       2) $BUYER_KPP (text)
#       3) A.col3        (integer if only digits, else text)
#       4) A.col7        (integer if only digits, else text)
#   - Skip first 3 rows of A (rows 1..3)
#   - Output filename = A!G2 (R2C7) as text, sanitized, ".xlsx"
#   - Save next to A
# Handles CSV with ; , or TAB; UTF-8 and CP1251.
# Writes summary at the end.

# ---------- SETTINGS ----------
# INN / KPP of the buyer as required by the 1C import format (demo values)
$BUYER_INN = "0000000000 "
$BUYER_KPP = "000000000"
# ------------------------------

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null

function Select-InputFolder {
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select a folder with A files (CSV/XLSX)"
    $fbd.ShowNewFolderButton = $false
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $fbd.SelectedPath }
    else { throw "No folder selected." }
}

function Detect-Delimiter { param([string]$firstLine)
    $cSem = ($firstLine -split ';').Count
    $cCom = ($firstLine -split ',').Count
    $cTab = ($firstLine -split "`t").Count
    $max = [Math]::Max($cSem, [Math]::Max($cCom, $cTab))
    if     ($max -eq $cSem) { ';' }
    elseif ($max -eq $cCom) { ',' }
    else                    { "`t" }
}

function Read-Csv { param([string]$Path)
    $rows = @()
    $encs = @([System.Text.Encoding]::UTF8, [System.Text.Encoding]::GetEncoding(1251))
    $ok = $false
    foreach ($enc in $encs) {
        try {
            $sr = New-Object System.IO.StreamReader($Path, $enc, $true)
            $first = $sr.ReadLine(); if ($null -eq $first) { $sr.Close(); break }
            $delim = Detect-Delimiter -firstLine $first; $sr.Close()
            $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, $enc, $true)
            $p.HasFieldsEnclosedInQuotes = $true
            $p.SetDelimiters($delim)
            while (-not $p.EndOfData) { $rows += ,($p.ReadFields()) }
            $p.Close(); $ok = $true; break
        } catch {}
    }
    if (-not $ok) { throw "Failed to read CSV." }
    return ,$rows
}

function Read-Xlsx { param([string]$Path)
    try { $excel = New-Object -ComObject Excel.Application } catch { throw "Excel COM not available to read XLSX." }
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($Path, $false, $true)
    $ws = $wb.Worksheets.Item(1)

    $used = $ws.UsedRange
    $rowsCount = $used.Rows.Count
    $colsCount = $used.Columns.Count

    $rows = @()
    for ($r=1; $r -le $rowsCount; $r++) {
        $arr = @()
        for ($c=1; $c -le $colsCount; $c++) {
            $text = [string]$ws.Cells.Item($r, $c).Text
            if ($null -eq $text) { $text = "" }
            $arr += ,$text
        }
        $rows += ,$arr
    }

    $g2 = [string]$ws.Cells.Item(2,7).Text

    $wb.Close($false)
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)   | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)| Out-Null

    return @{ Rows = $rows; G2 = $g2 }
}

function Read-A { param([string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -eq ".xlsx") { return Read-Xlsx -Path $Path }
    else { return @{ Rows = (Read-Csv -Path $Path); G2 = $null } }
}

function Sanitize-FileName { param([string]$name)
    if ($null -eq $name) { $name = "" }
    $name = $name.Trim()
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $cls = [regex]::Escape([string]::Join("", $invalid))
    $name = [regex]::Replace($name, "[" + $cls + "]", "")
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "B_" + (Get-Date).ToString("yyyyMMdd_HHmmss") }
    return $name
}

function Ensure-UniquePath { param([string]$path)
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $dir  = [System.IO.Path]::GetDirectoryName($path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $ext  = [System.IO.Path]::GetExtension($path)
    $i = 1
    while ($true) {
        $p = Join-Path $dir ("{0} ({1}){2}" -f $name, $i, $ext)
        if (-not (Test-Path -LiteralPath $p)) { return $p }
        $i++
    }
}

function Build-BRows { param($Rows)
    # skip first 3 rows (1..3) -> start at index 3 (0-based)
    $start = 3
    $out = @()
    for ($i=$start; $i -lt $Rows.Count; $i++) {
        $r = $Rows[$i]; if ($null -eq $r) { continue }
        $a3 = ""; if ($r.Count -ge 3) { $a3 = [string]$r[2] }
        $a7 = ""; if ($r.Count -ge 7) { $a7 = [string]$r[6] }
        $out += ,@($BUYER_INN, $BUYER_KPP, $a3, $a7)
    }
    return $out
}

function Write-Xlsx { param($Rows, $Path)
    try { $excel = New-Object -ComObject Excel.Application } catch { throw "Excel COM not available to write XLSX." }
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    $rowIdx = 1
    foreach ($r in $Rows) {
        for ($i=0; $i -lt $r.Count; $i++) {
            $val  = [string]$r[$i]
            $cell = $ws.Cells.Item($rowIdx, $i+1)
            if ($i -eq 0 -or $i -eq 1) {
                # fixed text columns
                $cell.NumberFormat = "@"
                $cell.Value2 = "'" + $val
            } elseif ($i -eq 2 -or $i -eq 3) {
                # prefer integer if digits-only (strip spaces)
                $clean = ($val -replace '\s','')
                if ($clean -match '^\d+$') {
                    $cell.Value2 = [int64]$clean
                    $cell.NumberFormat = "0"
                } else {
                    $cell.NumberFormat = "@"
                    $cell.Value2 = $val
                }
            }
        }
        $rowIdx++
    }

    $wb.SaveAs($Path, 51)  # xlOpenXMLWorkbook
    $wb.Close($false)
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)   | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)   | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)| Out-Null
}

# ---------- MAIN ----------
try {
    $folder = Select-InputFolder
    $files = @()
    $files += Get-ChildItem -LiteralPath $folder -Filter *.xlsx -File | ForEach-Object { $_.FullName }
    $files += Get-ChildItem -LiteralPath $folder -Filter *.csv  -File | ForEach-Object { $_.FullName }
    if ($files.Count -eq 0) { throw "No CSV/XLSX files found in the selected folder." }

    $ok = 0; $fail = 0; $log = @()

    foreach ($f in $files) {
        try {
            $res   = Read-A -Path $f
            $rowsA = $res.Rows
            if ($rowsA.Count -eq 0) { throw "A is empty." }

            # filename from G2 (preferred), else from rows[1][6] if present
            $outName = $res.G2
            if ([string]::IsNullOrWhiteSpace($outName) -and $rowsA.Count -ge 2 -and $rowsA[1].Count -ge 7) {
                $outName = [string]$rowsA[1][6]
            }
            $outName = Sanitize-FileName -name $outName
            $outPath = Join-Path ([System.IO.Path]::GetDirectoryName($f)) ($outName + ".xlsx")
            $outPath = Ensure-UniquePath -path $outPath

            $rowsB = Build-BRows -Rows $rowsA
            if ($rowsB.Count -eq 0) { throw "No rows to write after skipping first 3 rows." }

            Write-Xlsx -Rows $rowsB -Path $outPath
            $ok++
            $log += "OK: $([System.IO.Path]::GetFileName($f)) -> $([System.IO.Path]::GetFileName($outPath))"
        }
        catch {
            $fail++
            $log += "FAIL: $([System.IO.Path]::GetFileName($f)) -> $($_.Exception.Message)"
        }
    }

    $msg = "Done. Success: $ok, Failed: $fail.`n`n" + ($log -join "`n")
    [System.Windows.Forms.MessageBox]::Show($msg, "Batch result") | Out-Null
}
catch {
    [System.Windows.Forms.MessageBox]::Show(("Error: {0}" -f $_.Exception.Message), "Error") | Out-Null
}
