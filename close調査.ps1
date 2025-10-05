$folder = "C:\work\search\Program\Program\ProgramSource"
#$folder = "C:\work\search\test"
$files = Get-ChildItem $folder -Recurse -Filter *.vb

foreach ($file in $files) {
    $content = Get-Content $file.FullName
    $methodName = ""
    $inMethod = $false
    $opened = $false
    $closed = $false
    $lineNum = 0

    foreach ($line in $content) {
        $lineNum++

        # Function名の検出
        if ($line -match "Function\s+(\w+)") {
            $methodName = $matches[1]
            $inMethod = $true
            $opened = $false
            $closed = $false
            $startLine = $lineNum

            # Function検出ログ出力
            #Write-Output "検出: $($file.FullName) の行 $lineNum に Function [$methodName]"
        }

        # Functionの終了
        elseif ($line -match "End Function") {
            if ($inMethod -and $opened -and -not $closed) {
                Write-Output " $($file.FullName) : $methodName (行:$startLine～$lineNum) に OpenConnection あり / CloseConnection なし"
            }
            $inMethod = $false
        }

        if ($inMethod) {
            # コメントを除外して検索
            $codeLine = $line -replace "'.*$", ""

            if ($codeLine -match "(?:\.|\s)OpenConnection\s*\(") {
                $opened = $true
            }
            if ($codeLine -match "(?:\.|\s)CloseConnection\s*\(") {
                $closed = $true
            }
        }
    }
}
