Add-Type -AssemblyName System.Web

# === ПЪТИЩА ===
$excelPath = "C:\TestSignatures\users.xlsx"
$logoPath  = "C:\TestSignatures\logo.png"
$signDir   = "C:\TestSignatures\Signatures"
$logFile   = Join-Path $signDir "signature_log.txt"

# === ФУНКЦИЯ ЗА ВГРАЖДАНЕ НА ЛОГО ===

function Get-Base64Image {
    param ([string]$imagePath)
    $bytes = [System.IO.File]::ReadAllBytes($imagePath)
    $base64 = [System.Convert]::ToBase64String($bytes)
    return "data:image/png;base64,$base64"
}


# === ПОДГОТОВКА НА ПАПКА И ЛОГ ===

if (!(Test-Path $signDir)) {
    New-Item -Path $signDir -ItemType Directory | Out-Null
}
if (Test-Path $logFile) { Remove-Item $logFile -Force }
"=== Лог от изпълнението ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) ===" | Out-File -FilePath $logFile

$logoBase64 = Get-Base64Image -imagePath $logoPath


# === ЗАРЕЖДАНЕ НА EXCEL ===

$excel    = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)
$sheet    = $workbook.Sheets.Item(1)

$row = 2
$totalUsers = 0
$totalErrors = 0
$totalSignatures = 0

while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
    $email = $sheet.Cells.Item($row, 1).Value().Trim()
    $username = $email.Split("@")[0]
    $web = $sheet.Cells.Item($row, 26).Value()

    if (-not $email) {
        Add-Content $logFile "Ред ${row}: Пропуснат - липсва имейл."
        $totalErrors++
        $row++
        continue
    }

    $langs = @{
        "BG" = @{
            name        = $sheet.Cells.Item($row, 2).Value()
            job         = $sheet.Cells.Item($row, 5).Value()
            group       = $sheet.Cells.Item($row, 8).Value()
            dept        = $sheet.Cells.Item($row, 11).Value()
            directorate = $sheet.Cells.Item($row, 14).Value()
            company     = $sheet.Cells.Item($row, 17).Value()
            phone       = $sheet.Cells.Item($row, 20).Value()
            mob         = $sheet.Cells.Item($row, 23).Value()
            eko         = $sheet.Cells.Item($row, 27).Value()
            law         = $sheet.Cells.Item($row, 30).Value()
        }
        "RU" = @{
            name        = $sheet.Cells.Item($row, 3).Value()
            job         = $sheet.Cells.Item($row, 6).Value()
            group       = $sheet.Cells.Item($row, 9).Value()
            dept        = $sheet.Cells.Item($row, 12).Value()
            directorate = $sheet.Cells.Item($row, 15).Value()
            company     = $sheet.Cells.Item($row, 18).Value()
            phone       = $sheet.Cells.Item($row, 21).Value()
            mob         = $sheet.Cells.Item($row, 24).Value()
            eko         = $sheet.Cells.Item($row, 28).Value()
            law         = $sheet.Cells.Item($row, 31).Value()
        }
        "EN" = @{
            name        = $sheet.Cells.Item($row, 4).Value()
            job         = $sheet.Cells.Item($row, 7).Value()
            group       = $sheet.Cells.Item($row, 10).Value()
            dept        = $sheet.Cells.Item($row, 13).Value()
            directorate = $sheet.Cells.Item($row, 16).Value()
            company     = $sheet.Cells.Item($row, 19).Value()
            phone       = $sheet.Cells.Item($row, 22).Value()
            mob         = $sheet.Cells.Item($row, 25).Value()
            eko         = $sheet.Cells.Item($row, 29).Value()
            law         = $sheet.Cells.Item($row, 32).Value()
        }
    }

    foreach ($lang in $langs.Keys) {
        $data = $langs[$lang]
        $sigName = "$username`_$lang"
        $userDir = "$signDir\$username"

        if (!(Test-Path $userDir)) {
            New-Item -ItemType Directory -Path $userDir | Out-Null
        }

        $labels = @{
            "BG" = @{ phone = "тел.:"; mobile = "моб.:"; email = "мейл"; website = "website" }
            "RU" = @{ phone = "тел.:"; mobile = "моб.:"; email = "почта"; website = "website" }
            "EN" = @{ phone = "tel.:"; mobile = "mob.:"; email = "email"; website = "website" }
        }

        $lbl = $labels[$lang]
        $lines = @()
        if ($data.name)        { $lines += "<b>$([System.Web.HttpUtility]::HtmlEncode($data.name))</b><br>" }
        if ($data.job)         { $lines += "$([System.Web.HttpUtility]::HtmlEncode($data.job))<br>" }
        if ($data.group)       { $lines += "$([System.Web.HttpUtility]::HtmlEncode($data.group))<br>" }
        if ($data.dept)        { $lines += "$([System.Web.HttpUtility]::HtmlEncode($data.dept))<br>" }
        if ($data.directorate) { $lines += "$([System.Web.HttpUtility]::HtmlEncode($data.directorate))<br>" }
        if ($data.company)     { $lines += "$([System.Web.HttpUtility]::HtmlEncode($data.company))<br>" }
        if ($data.phone)       { $lines += "<b>$($lbl.phone)</b> $([System.Web.HttpUtility]::HtmlEncode($data.phone))<br>" }
        if ($data.mob)         { $lines += "<b>$($lbl.mobile)</b> $([System.Web.HttpUtility]::HtmlEncode($data.mob))<br>" }
        if ($email)            { $lines += "<b>$($lbl.email):</b> <a href='mailto:$email'>$email</a><br>" }
        if ($web)              { $lines += "<b>$($lbl.website):</b> <a href='http://$web'>$web</a><br>" }
        if ($data.eko)         { $lines += "<br><span style='font-size:9pt; color:green;'>$([System.Web.HttpUtility]::HtmlEncode($data.eko))</span>" }
        if ($data.law)         { $lines += "<br><br><span style='font-size:9pt; font-style:italic; color:black; letter-spacing: 0.8pt;'>$([System.Web.HttpUtility]::HtmlEncode($data.law))</span>" }

        $html = @"
<table cellpadding='5' style='font-family: Calibri; font-size: 11pt; color: black; text-align: left;'>
  <tr>
    <td style='vertical-align: top;'>
      <img src='$logoBase64' width='80' style='display:block;'>
    </td>
    <td>
      $($lines -join "`n")
    </td>
  </tr>
</table>
"@

        Set-Content -Path "$userDir\$sigName.htm" -Value $html -Encoding UTF8

        $txt = ""
        if ($data.name)        { $txt += "$($data.name)`n" }
        if ($data.job)         { $txt += "$($data.job)`n" }
        if ($data.group)       { $txt += "$($data.group)`n" }
        if ($data.dept)        { $txt += "$($data.dept)`n" }
        if ($data.directorate) { $txt += "$($data.directorate)`n" }
        if ($data.company)     { $txt += "$($data.company)`n" }
        if ($data.phone)       { $txt += "$($lbl.phone): $($data.phone)`n" }
        if ($data.mob)         { $txt += "$($lbl.mobile): $($data.mob)`n" }
        if ($email)            { $txt += "$email`n" }
        if ($web)              { $txt += "$web`n" }
        if ($data.eko)         { $txt += "`n$($data.eko)" }
        if ($data.law)         { $txt += "`n`n$($data.law)" }

        Set-Content -Path "$userDir\$sigName.txt" -Value $txt -Encoding UTF8
        Set-Content -Path "$userDir\$sigName.rtf" -Value $txt -Encoding UTF8

        Add-Content $logFile "Създаден подпис: $sigName"
        $totalSignatures++
    }

    $totalUsers++
    $row++
}

# === ЗАТВАРЯНЕ НА EXCEL ===
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# === РЕЗЮМЕ ===
Add-Content $logFile "`n=== Обобщение ==="
Add-Content $logFile "Обработени потребители: $totalUsers"
Add-Content $logFile "Общо подписи: $totalSignatures"
Add-Content $logFile "Грешки: $totalErrors"

Write-Host "`nГотово. Създадени са $totalSignatures подписа за $totalUsers потребители. Виж лога в: $logFile" -ForegroundColor Green
