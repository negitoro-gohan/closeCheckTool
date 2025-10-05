$folder = "C:\work\search\Program\Program\ProgramSource"
$files = Get-ChildItem $folder -Recurse -Filter *.vb

$withCloseCount = 0
$withoutCloseCount = 0

foreach ($file in $files) {
    $content = Get-Content $file.FullName
    $methodName = ""
    $inMethod = $false
    $opened = $false
    $closed = $false
    $lineNum = 0
    $isSub = $false

    foreach ($line in $content) {
        $lineNum++

        # Function or Sub の検出
        if ($line -match "Function\s+(\w+)") {
            $methodName = $matches[2]
            $inMethod = $true
            $isSub = $false
            $opened = $false
            $closed = $false
            $startLine = $lineNum

            #Write-Output "検出: $($file.FullName) の行 $lineNum に Function [$methodName]"
        }
        elseif ($line -match "Sub\s+(\w+)") {
            $methodName = $matches[2]
            $inMethod = $true
            $isSub = $true
            $opened = $false
            $closed = $false
            $startLine = $lineNum

            #Write-Output "検出: $($file.FullName) の行 $lineNum に Sub [$methodName]"
        }

        # Function / Sub 終了
        elseif ($line -match "^\s*End\s+(Function|Sub)") {
            if ($inMethod -and $opened) {
                if ($closed) {
                    #Write-Output "✅ $($file.FullName) : $methodName (行:$startLine～$lineNum) に OpenConnection / CloseConnection 両方あり"
                    $withCloseCount++
                }
                else {
                    Write-Output "⚠ $($file.FullName) : $methodName (行:$startLine～$lineNum) に OpenConnection あり / CloseConnection なし"
                    $withoutCloseCount++
                }
            }
            $inMethod = $false
        }

        if ($inMethod) {
            # コメントを除外して検索
            $codeLine = $line -replace "'.*$", ""

            # OpenConnection 検出
            if ($codeLine -match "OpenConnection") {
                $opened = $true
            }

            # CloseConnection 検出
            if ($codeLine -match "CloseConnection") {
                $closed = $true
            }
        }
    }
}

# 集計結果
Write-Output ""
Write-Output "----------------------------------------"
Write-Output "✅ Open + Close 両方あり件数：$withCloseCount"
Write-Output "⚠ Openのみ（Closeなし）件数　：$withoutCloseCount"
Write-Output "----------------------------------------"
pause
