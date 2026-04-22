# QR फोल्डर path
$savePath = "C:\Users\laksh\Downloads\Qr"
if (!(Test-Path $savePath)) {
    New-Item -ItemType Directory -Path $savePath | Out-Null
}

# Excel COM object बनाओ
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Excel फाइल path (अपनी फाइल का path डालो)
$workbook = $excel.Workbooks.Open("C:\Users\laksh\Desktop\qr.xlsx")
$sheet = $workbook.Sheets.Item(1)  # जरूरत के हिसाब से sheet बदलो

# Last row निकालो
$lastRow = $sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row  # xlUp = -4162

# Loop कर के QR सेव करो
for ($i = 2; $i -le $lastRow; $i++) {
    $qrContent = $sheet.Cells($i, 1).Text
    $fileName = $sheet.Cells($i, 2).Text

    if ($qrContent -and $fileName) {
        $encodedContent = [System.Web.HttpUtility]::UrlEncode($qrContent)
        $url = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=$encodedContent"
        $outputFile = Join-Path $savePath "$fileName.png"

        Invoke-WebRequest -Uri $url -OutFile $outputFile
        Write-Host "Saved: $outputFile"
    }
}

# Cleanup
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null