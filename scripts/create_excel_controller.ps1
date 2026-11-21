$ErrorActionPreference = "Stop"

$projectWindowsPath = "\\wsl.localhost\Ubuntu\home\egors\MRO_hakk"
$outputPath = Join-Path $projectWindowsPath "MRO_Control.xlsm"
$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$propertyName = "AccessVBOM"
$propertyExisted = $false
$previousValue = $null
$excel = $null
$workbook = $null

if (Test-Path $securityKey) {
    $property = Get-ItemProperty -Path $securityKey -ErrorAction SilentlyContinue
    if ($null -ne $property -and $property.PSObject.Properties.Name -contains $propertyName) {
        $propertyExisted = $true
        $previousValue = $property.$propertyName
    }
} else {
    New-Item -Path $securityKey -Force | Out-Null
}

try {
    Set-ItemProperty -Path $securityKey -Name $propertyName -Type DWord -Value 1

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "MRO Control"
    $excel.ActiveWindow.DisplayGridlines = $false

    $sheet.Range("B2:F2").Merge()
    $sheet.Range("B2").Value2 = "Magnetic MRO test"
    $sheet.Range("B2").Font.Name = "Aptos"
    $sheet.Range("B2").Font.Size = 18
    $sheet.Range("B2").Font.Bold = $true
    $sheet.Range("B2").Font.Color = 0x4E1F0F

    $sheet.Range("B4:F4").Merge()
    $sheet.Range("B4").Value2 = "1. Asenda input-kaustas kolm Exceli faili uute sama nimega failidega."
    $sheet.Range("B5:F5").Merge()
    $sheet.Range("B5").Value2 = "2. Sulge eelmine raport Excelis."
    $sheet.Range("B6:F6").Merge()
    $sheet.Range("B6").Value2 = "3. Vajuta allolevat nuppu ja oota teadet VALMIS."

    $sheet.Range("B4:F6").Font.Name = "Aptos"
    $sheet.Range("B4:F6").Font.Size = 11
    $sheet.Range("B4:F6").RowHeight = 22

    $sheet.Range("B10").Value2 = "Olek"
    $sheet.Range("C10:F10").Merge()
    $sheet.Range("C10").Value2 = "Valmis käivitamiseks"
    $sheet.Range("B10:F10").Font.Name = "Aptos"
    $sheet.Range("B10:F10").Font.Size = 11
    $sheet.Range("B10").Font.Bold = $true
    $sheet.Range("B10:F10").Interior.Color = 0xE2F0D9

    $sheet.Columns.Item("A").ColumnWidth = 3
    $sheet.Columns.Item("B").ColumnWidth = 18
    $sheet.Columns.Item("C").ColumnWidth = 18
    $sheet.Columns.Item("D").ColumnWidth = 18
    $sheet.Columns.Item("E").ColumnWidth = 18
    $sheet.Columns.Item("F").ColumnWidth = 18

    $button = $sheet.Shapes.AddShape(5, 100, 165, 360, 64)
    $button.Name = "UpdateReportButton"
    $button.TextFrame2.TextRange.Text = "UUENDA RAPORT"
    $button.TextFrame2.TextRange.Font.Name = "Aptos"
    $button.TextFrame2.TextRange.Font.Size = 20
    $button.TextFrame2.TextRange.Font.Bold = $true
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    $button.Fill.ForeColor.RGB = 0x78431F
    $button.Line.ForeColor.RGB = 0x78431F
    $button.OnAction = "RunMRO"

    $macroCode = @'
Option Explicit

Public Sub RunMRO()
    Dim shellObject As Object
    Dim command As String
    Dim exitCode As Long
    Dim controlSheet As Worksheet

    On Error GoTo Failed
    Set controlSheet = ThisWorkbook.Worksheets("MRO Control")
    controlSheet.Range("C10").Value = "Töötlemine käib..."
    controlSheet.Range("B10:F10").Interior.Color = RGB(255, 242, 204)
    DoEvents

    command = "wsl.exe -d Ubuntu --cd /home/egors/MRO_hakk -- .venv/bin/python -B scripts/run_pipeline.py --config config.json --fresh --run-tests"
    Set shellObject = CreateObject("WScript.Shell")
    exitCode = shellObject.Run(command, 1, True)

    If exitCode <> 0 Then
        controlSheet.Range("C10").Value = "Viga. Vaata terminali teadet."
        controlSheet.Range("B10:F10").Interior.Color = RGB(244, 204, 204)
        MsgBox "Raporti uuendamine ebaõnnestus. Sulge avatud raport ja kontrolli input-faile.", vbCritical, "Magnetic MRO"
        Exit Sub
    End If

    controlSheet.Range("C10").Value = "Valmis. Raport on uuendatud."
    controlSheet.Range("B10:F10").Interior.Color = RGB(226, 240, 217)
    Shell "explorer.exe ""\\wsl.localhost\Ubuntu\home\egors\MRO_hakk\excel_files\output""", vbNormalFocus
    MsgBox "Raport on uuendatud. Output-kaustas on üks uus põhiraport.", vbInformation, "Magnetic MRO"
    Exit Sub

Failed:
    If Not controlSheet Is Nothing Then
        controlSheet.Range("C10").Value = "Viga: " & Err.Description
        controlSheet.Range("B10:F10").Interior.Color = RGB(244, 204, 204)
    End If
    MsgBox "Viga: " & Err.Description, vbCritical, "Magnetic MRO"
End Sub
'@

    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = "MROController"
    $module.CodeModule.AddFromString($macroCode)

    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }
    $workbook.SaveAs($outputPath, 52)
    $workbook.Close($true)
    $workbook = $null
    Write-Output "CREATED=$outputPath"
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }
    if ($propertyExisted) {
        Set-ItemProperty -Path $securityKey -Name $propertyName -Type DWord -Value $previousValue
    } else {
        Remove-ItemProperty -Path $securityKey -Name $propertyName -ErrorAction SilentlyContinue
    }
}
