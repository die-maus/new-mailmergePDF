function new-mailmergePDF {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true)][string]$DocumentName,
        [parameter(Mandatory = $true)][string]$DocumentOutPath,
        [parameter(Mandatory = $true)][string]$DataSourceName,
        [parameter(Mandatory = $true)][string]$SheetName,
        [parameter(Mandatory = $true)][string]$DataColumnOutFilename
    )

    <#
        .SYNOPSIS
            Word Serienbrief ausführen
        .DESCRIPTION
            Erstellt für jeden Datensatz in einer Datenquelle eines Word Serienbriefes eine einzelne PDF Datei
        .PARAMETER  DocumentName
            Dateiname des Word Serienbriefes
        .PARAMETER DocumentOutPath
            Pfad für die Ergebnisdateien
         .PARAMETER  DataSourceName
            Dateinname der Excel Datenquelle
        .PARAMETER  SheetName
            Name des Excel Tabellenblattes mit den Daten
        .PARAMETER  DataColumnOutFilename
            Name der Spalte mit dem Dateinamen der PDF Dateien
        .EXAMPLE
            new-mailmergePDF -DocumentName '.\TEST_SB\Serienbrief_ps\Serienbrief_Test.docx' -DataName '.\TEST_SB\Serienbrief_ps\Serienbrief_Daten.xlsx' -SheetName 'Test 1' -DataColumnOutFilename 'Dateiname' -Verbose
    #>

    $word = new-object -com Word.Application
    $word.Visible = $true
    #$word.displayalerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

    Write-Verbose "Serienbrief:           $DocumentName"
    Write-Verbose "Datenquelle:           $DataSourceName"
    Write-Verbose "Tabellenblatt:         $SheetName"
    Write-Verbose "Spalte für Dateinamen: $DataColumnOutFilename"

    $doc = $word.Documents.Open($DocumentName, $false, $true)

    $default = [Type]::Missing
    $doc.MailMerge.OpenDataSource( $DataSourceName, $default, $default, $default, $default, $default, $default, $default, $default, $default, $default, $default, ("SELECT * FROM [{0}$]" -f $Sheetname))
    $doc.MailMerge.MainDocumentType = [Microsoft.Office.Interop.Word.WdMailMergeMainDocType]::wdFormLetters

    for ($i = 1; $i -le $doc.MailMerge.DataSource.RecordCount; $i++) {

        $doc.MailMerge.Destination = [Microsoft.Office.Interop.Word.WdMailMergeDestination]::WdSendToNewDocument
        $doc.MailMerge.DataSource.FirstRecord = $i
        $doc.MailMerge.DataSource.LastRecord = $i
        $doc.MailMerge.DataSource.ActiveRecord = $i

        $doc.Mailmerge.Execute()

        $strFileName = [string] (Join-Path $DocumentOutPath $doc.MailMerge.DataSource.DataFields($DataColumnOutFilename).Value)
        Write-Verbose ("  {0}/{1}: {2}" -f $i, $doc.MailMerge.DataSource.RecordCount, $strFileName)

        $word.ActiveDocument.ExportAsFixedFormat($strFileName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF, $false, [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint)
        $word.ActiveDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    }

    $doc.close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    $word.Quit()

    [gc]::collect() 
    [gc]::WaitForPendingFinalizers()
}

# in dem Ordner "Serienbrief" liegt eine Word Serienbriefdatei .docx und eine Excel Datenquelle .xlsx - die PDF Ergebnisdokumente werden im Ordner pdf_out abgelegt

# Names des Excel Tabellenblattes der Datenquelle - wenn nicht gefüllt, dann wird das erste Datenblatt genutzt
$strFirstSheetName = ""

if ( !$strFirstSheetName ) {
    # das erste Excel Tabellenblatt wird für die Daten genutzt
    $ExcelFile = ([string[]](Get-ChildItem (Join-Path $PSScriptRoot "Serienbrief") -Filter *.xlsx).FullName)[0]
    $excel = new-object -comobject Excel.Application
    $workbook = $excel.Workbooks.Open($ExcelFile)
    $strFirstSheetName = $workbook.sheets[1].Name
    #$workbook.sheets | ForEach-Object { Write-Host ("Tabellenblatt {0}: {1}" -f $_.Index, $_.name) }
    $excel.Quit()
}

new-mailmergePDF -DocumentName ([string[]](Get-ChildItem (Join-Path $PSScriptRoot "Serienbrief") -Filter *.docx).FullName)[0] -DocumentOutPath (Join-Path $PSScriptRoot "pdf_out") -DataSourceName ([string[]](Get-ChildItem (Join-Path $PSScriptRoot "Serienbrief") -Filter *.xlsx).FullName)[0] -SheetName $strFirstSheetName -DataColumnOutFilename 'Dateiname' -Verbose