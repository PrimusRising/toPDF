# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Get invocation path
#$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path

# Create a PowerPoint object
$ppt_app = New-Object -ComObject PowerPoint.Application
$word_app = New-Object -ComObject Word.Application
# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    $opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    $document.SaveAs($pdf_filename, $opt)
    # Close PowerPoint file
    $document.Close()
}

#Quit and release ComObject
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)

#Word to Pdf
Get-ChildItem -Path $curr_path -Filter *.doc? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in Word
    $document = $word_app.Documents.Open($_.FullName)
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source doc file, replace $curr_path with $_.DirectoryName
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    # Close doc file
    $document.Close()
}

#Quit and release ComObject
$word_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word_app)

#Exel to PDF works 
$xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]
$objExcel = New-Object -ComObject excel.application
$objExcel.visible = $False
Get-ChildItem -Path $curr_path -Recurse -Filter *.xls? | ForEach-Object {
    Write-Host "Processing" $_.BaseName "..."
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $workbook = $objExcel.workbooks.open($_.fullname, 3)
    $workbook.Saved = $true
    #"saving $pdf_filename”
    #$workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)
    $workbook.WorkSheets.Item(1).ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdf_filename)
    $objExcel.Workbooks.close()
}

#Quit and release ComObject
$objExcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

# Convert picture files to PDF
Get-ChildItem -Path $curr_path -Include *.png, *.jpg, *.jpeg, *.gif, *.bmp | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    $image = [System.Drawing.Image]::FromFile($_.FullName)
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $image.Save($pdf_filename, [System.Drawing.Imaging.ImageFormat]::Pdf)
    $image.Dispose()
}


