$ErrorActionPreference = "Stop"
$base = "c:\Users\dream\OneDrive\Desktop\클로두\99.zion\끝말잇기"
$src  = Join-Path $base "국어 기초 어휘 선정 및 어휘 등급화 목록 전체.xlsx"

# ---------- 1. Read source via Excel COM (whole-range) ----------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($src)
$ws = $wb.Worksheets.Item("전체(1~5등급), 40,000개")
$used = $ws.UsedRange
$rows = $used.Rows.Count
# columns: 1=등급, 2=어휘, 4=품사, 7=의미
$rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($rows,7))
$data = $rng.Value2
$wb.Close($false); $excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rng) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Host "Read rows: $rows"

# ---------- 2. Hangul helpers ----------
$YVOWELS = @(2,3,6,7,12,17,20)  # ㅑㅒㅕㅖㅛㅠㅣ -> trigger ㄹ/ㄴ to ㅇ
function Get-Decomp([char]$ch){
  $code = [int]$ch
  if($code -lt 0xAC00 -or $code -gt 0xD7A3){ return $null }
  $s = $code - 0xAC00
  return @([math]::Floor($s/588), [math]::Floor(($s%588)/28), ($s%28))
}
function New-Syll($onset,$nuc,$coda){
  return [char][int](0xAC00 + ($onset*588) + ($nuc*28) + $coda)
}
function Get-NextStarts([char]$ch){
  $d = Get-Decomp $ch
  if($d -eq $null){ return @([string]$ch) }
  $res = New-Object System.Collections.Generic.List[string]
  $res.Add([string]$ch)
  $onset=$d[0]; $nuc=$d[1]; $coda=$d[2]
  if($onset -eq 5){ # ㄹ
    if($YVOWELS -contains $nuc){ $res.Add([string](New-Syll 11 $nuc $coda)) }
    else { $res.Add([string](New-Syll 2 $nuc $coda)) }
  } elseif($onset -eq 2){ # ㄴ
    if($YVOWELS -contains $nuc){ $res.Add([string](New-Syll 11 $nuc $coda)) }
  }
  return $res.ToArray()
}
$reHangul = [regex]'^[가-힣]+$'

# clean definition: drop 「n」 sense markers, take first sense, collapse, trim
function Clean-Def([string]$raw){
  if([string]::IsNullOrWhiteSpace($raw)){ return "" }
  $t = $raw -replace "\r"," " -replace "\n"," "
  # if multiple senses 「1」「2」.. keep only the first sense text
  $m = [regex]::Match($t, "「1」(.+?)(「2」|$)")
  if($m.Success){ $t = $m.Groups[1].Value } else { $t = ($t -replace "「\d+」"," ") }
  $t = $t -replace "「[^」]*」"," "        # remove any remaining 「..」 tags
  $t = $t -replace "\s+"," "
  $t = $t.Trim()
  if($t.Length -gt 120){ $t = $t.Substring(0,120).Trim() + "…" }
  return $t
}

# ---------- 3. Build noun pool, dedup by word -> min level (+ keep def) ----------
$wordLevel = @{}   # word -> min level
$wordDef   = @{}   # word -> def (from the lowest-level row seen)
for($r=2; $r -le $rows; $r++){
  $pos = [string]$data[$r,4]
  if($pos -ne "명사"){ continue }
  $w = [string]$data[$r,2]
  if([string]::IsNullOrWhiteSpace($w)){ continue }
  $w = $w.Trim()
  if($w.Length -lt 2){ continue }
  if(-not $reHangul.IsMatch($w)){ continue }
  $lv = [int]([string]$data[$r,1]).Substring(0,1)
  $def = Clean-Def ([string]$data[$r,7])
  if($wordLevel.ContainsKey($w)){
    if($lv -lt $wordLevel[$w]){ $wordLevel[$w] = $lv; $wordDef[$w] = $def }
    elseif([string]::IsNullOrEmpty($wordDef[$w]) -and $def -ne ""){ $wordDef[$w] = $def }
  } else { $wordLevel[$w] = $lv; $wordDef[$w] = $def }
}
Write-Host "Unique 2+ syllable Hangul nouns: $($wordLevel.Count)"

# ---------- 4. Iterative dead-end removal (두음법칙 반영) ----------
$MIN_FOLLOW = 2
$first = @{}; $nextCand = @{}
foreach($w in $wordLevel.Keys){
  $first[$w] = $w[0]
  $nextCand[$w] = Get-NextStarts $w[$w.Length-1]
}
$alive = New-Object System.Collections.Generic.HashSet[string]
foreach($w in $wordLevel.Keys){ [void]$alive.Add($w) }

$iter = 0
while($true){
  $iter++
  $startCount = @{}
  foreach($w in $alive){
    $f = [string]$first[$w]
    if($startCount.ContainsKey($f)){ $startCount[$f]++ } else { $startCount[$f]=1 }
  }
  $toRemove = New-Object System.Collections.Generic.List[string]
  foreach($w in $alive){
    $cands = $nextCand[$w]; $follow = 0; $fw = [string]$first[$w]
    foreach($c in $cands){ if($startCount.ContainsKey($c)){ $follow += $startCount[$c] } }
    if($cands -contains $fw){ $follow -= 1 }
    if($follow -lt $MIN_FOLLOW){ $toRemove.Add($w) }
  }
  if($toRemove.Count -eq 0){ break }
  foreach($w in $toRemove){ [void]$alive.Remove($w) }
  Write-Host ("Iter {0}: removed {1}, remaining {2}" -f $iter,$toRemove.Count,$alive.Count)
  if($iter -gt 30){ break }
}
Write-Host "Final playable nouns: $($alive.Count)"

# ---------- 5. Emit words.json (sorted by level then word) ----------
$list = @($alive) | Sort-Object @{Expression={$wordLevel[$_]}}, @{Expression={$_}}
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('{"words":[')
$cnt = 0; $withDef = 0
foreach($w in $list){
  if($cnt -gt 0){ [void]$sb.Append(',') }
  $d = $wordDef[$w]; if($d -eq $null){ $d = "" }
  if($d -ne ""){ $withDef++ }
  $we = $w -replace '\\','\\\\' -replace '"','\"'
  $de = $d -replace '\\','\\\\' -replace '"','\"'
  [void]$sb.Append('{"w":"'); [void]$sb.Append($we)
  [void]$sb.Append('","lv":'); [void]$sb.Append($wordLevel[$w])
  [void]$sb.Append(',"def":"'); [void]$sb.Append($de); [void]$sb.Append('"}')
  $cnt++
}
[void]$sb.Append(']}')
$out = Join-Path $base "words.json"
[System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
Write-Host ("Wrote words.json: {0} words, {1} with def -> {2}" -f $cnt,$withDef,$out)
