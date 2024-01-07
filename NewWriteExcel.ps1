# コマンドライン引数を取得
$psPath = $args[0]
$sheetName = $args[1]

# ExcelとCSVファイルを検索
$xlsFiles = Get-ChildItem -Path $psPath -Filter "*.xls*"
$csvFile = Get-ChildItem -Path $psPath -Filter "*.csv" | Select-Object -First 1

# CSVファイルからデータを読み込む
$csvData = Import-Csv -Path $csvFile.FullName

# Excelファイルを処理
$xlsFiles | ForEach-Object {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($_.FullName)
    $worksheet = $workbook.Worksheets | Where-Object { $_.Name -eq $sheetName }

    if ($worksheet -ne $null) {
        $csvData | ForEach-Object {
            $r = $_.row
            $c = $_.column
            $v = $_.value
            $worksheet.Cells.Item($r, $c).Value2 = $v
        }
        $workbook.Save()
        $workbook.Close()
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# ガベージコレクタを呼び出してCOMオブジェクトをクリーンアップ
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# ユーザーに完了を通知
[Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
[System.Windows.Forms.MessageBox]::Show("Finished.")