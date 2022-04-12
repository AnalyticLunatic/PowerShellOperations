# Script:             Burst File Contents into 1 Worksheet Workbooks
#
# Script Purpose:     Given a folder location, loop through all files (txt, xls, xlsx) and split each worksheet/file content 
#                     into an individual xlsx workbook file. This ensures each workbook has a single worksheet with the same
#                     worksheet name, making it easy for Power BI to pull in contents from multiple files for modeling.
#
# Author:             James Scurlock 
#-----------------------------------------------------------------------------
# Instructions:
#-----------------------------------------------------------------------------
# Call the PowerShell Script or copy and Execute in PowerShell ISE. You will be prompted for several folder paths:
#
# 1. Folder location to store processing log file:                            C:\Users\{YourUser}\Downloads\BurstFiles\ProcessingLogs
# 2. Folder location of files you are wanting to process:                     C:\Users\{YourUser}\Downloads\BurstFiles
# 3. Folder location where Processed original files should be moved to:       C:\Users\{YourUser}\Downloads\BurstFiles\OriginalFiles
# 4. Folder location where files with Error in processing should be moved to: C:\Users\{YourUser}\Downloads\BurstFiles\ProcessingError
#-----------------------------------------------------------------------------
# Various Functions
#-----------------------------------------------------------------------------
$processedFilesCount = 0

function SplitExcelWorksheetsApart {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({Test-Path -Path $_ -PathType Container})]
        [Alias('Path')]
        [string]$SourceFolder
    )
    $processedOriginalFilesDirectory = Read-Host -Prompt 'Enter folder location to store processed original files:' 
    $errorFilesDirectory = Read-Host -Prompt 'Enter folder location to store any files that Error during processing (Ex. Password Protected):'
    #$burstedFilesDirectory = Read-Host -Prompt 'Enter file location to store bursted files'

    # Remove any files existing already in the processed directory from past runs
    Remove-Item $processedOriginalFilesDirectory\*.*

    # Get a list of all files in provided directory
    $allFiles = Get-ChildItem -Path $SourceFolder -File # Can Filter for a particular type if desired # -Filter '*.xlsx' -File

    $totalCount = ($allFiles | Measure-Object).Count
    write-Output "`nBEGIN PROCESS: Total files to process for splitting worksheets into workbooks: $totalCount `n" 
    
    foreach ($excelFile in $allFiles) {
        $processedFilesCount++
        Write-Output "`n~Processing file $processedFilesCount of $totalCount"
        Write-Output "-------------------------------------------------------------------------------------"

        SplitWorkbook $excelFile.fullname -output_type "xlsx" 

        # Move original file into a [Processed] Folder
        Move-Item -Path $excelFile.fullname -Destination $processedOriginalFilesDirectory
        Write-Output "  `nSource file moved to: $processedOriginalFilesDirectory"
    }

    Write-Host "`n-------------------------------------------------------------------------------------`nPROCESS COMPLETE: $processedFilesCount of $totalCount files processed." -ForegroundColor Green
    $errorFiles = Get-ChildItem -Path $errorFilesDirectory -File # Can Filter for a particular type if desired # -Filter '*.xlsx' -File
    $errorFileCount = ($errorFiles | Measure-Object).Count
    Write-Host "`n-------------------------------------------------------------------------------------`nFiles Skipped due to errors in processing: $errorFileCount." -ForegroundColor Yellow
}
#-----------------------------------------------------------------------------
# Figures out and returns the 'XlFileFormat Enumeration' ID for the specified format.
# http://msdn.microsoft.com/en-us/library/office/bb241279%28v=office.12%29.aspx 
# NOTE: The code being used for 'xls' is actually a 'text' type, but it seemed
# to work the best for splitting the worksheets into separate Excel files.
function GetOutputFileFormatID 
{ Param([string]$fomat_name) 

    $Result = 0 

    switch($fomat_name) 
    { 
        "csv" {$Result = 6} 
        "txt" {$Result = 20} 
        "xls" {$Result = 21} 
        "html" {$Result = 44} 
        default {$Result = 51} 
    } 
    
    return $Result 
}
#-----------------------------------------------------------------------------# 
# Purpose: Extract all of the worksheets from an Excel file into separate files.
#	-output_type: The filetype to save the Worksheets as can be csv, txt, xls, html.
function SplitWorkbook {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,Position=0)] 
        [string]$filepath,

        [Parameter(Mandatory=$true,Position=1)] 
        [ValidateSet("csv","txt","xls","html","xlsx")] 
        [string]$output_type
    )

    $Excel = New-Object -ComObject "Excel.Application" 
    $Excel.Visible = $false       # Runs Excel in the background. 
    $Excel.DisplayAlerts = $false # Supress alert messages. 

    # Attempt to open the workbook by providing a dummy password - if it IS  
    # password protected and fails, catch the exception and move file for review.
    try {
        $Workbook = $Excel.Workbooks.open($filepath, 0, 0, 5, "dummypassword") 
    }
    catch {
        Write-Host "$excelFile - ERROR encountered during processing. File is possibly password protected - moving file for later review." -ForegroundColor Yellow
        #$Workbook.Close() 
        $Excel.Quit()
        Move-Item -Path $excelFile.fullname -Destination $errorFilesDirectory
        continue
    }

    # Loop through the Workbook and extract each Worksheet in the specified file type. 
    if ($Workbook.Worksheets.Count -gt 0) { 

        #Strip off the Excel extension. 
        $WorkbookName = $filepath -replace ".xlsx", ""    #Post 2007 extension
        $WorkbookName = $WorkbookName -replace ".xls", "" #Pre 2007 extension 
        
        $FileFormat = GetOutputFileFormatID($output_type) 
    
        write-Output "Now processing: $WorkbookName"
        Write-Host -NoNewline "  Worksheets in file: " $Workbook.Worksheets.Count "`n" 
    
        $Worksheet = $Workbook.Worksheets.item(1) 

        foreach($Worksheet in $Workbook.Worksheets) 
        { 
            #In keeping with standard Excel practise, copying a worksheet to no destination will create a new workbook 
            #with that worksheet as the solitary worksheet and that workbook will be the ActiveWorkbook.
            $Worksheet.Copy()
            $ExtractedFileName = $WorkbookName + "__" + $Worksheet.Name + "." + $output_type
            $Excel.ActiveWorkbook.SaveAs($ExtractedFileName, $FileFormat)
            $Excel.ActiveWorkbook.Close()

            Write-Output "    Created file: $ExtractedFileName" 

            #Draft code for renaming all sheets while in loop 
            #Write-Host "Renaming all worksheets in '$excelFile.FullName'"
            #for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            #    $workbook.Sheets.Item($i).Name = ('Sheet'+ $i)
            #}

            #$Worksheet.SaveAs($ExtractedFileName, $FileFormat) 
            #write-Output "    Created file: $ExtractedFileName" 
        } 
    } 

    # Clean up & close the main Excel objects. 
    $Workbook.Close() 
    $Excel.Quit() 
}
#-----------------------------------------------------------------------------
# Script Beginning
#-----------------------------------------------------------------------------
cls # Clear Screen
$logFileDirectory = Read-Host -Prompt 'Enter folder location for Processing log file:'
$currentDateTime = Get-Date -Format "MM.dd.yyyy_HH.mm"
Start-Transcript -path "$($logFileDirectory)\ProcessLog_$currentDateTime.txt" 
$sourceFilesDirectory = Read-Host -Prompt "Enter folder location of files to process:"
SplitExcelWorksheetsApart $sourceFilesDirectory
Stop-Transcript
