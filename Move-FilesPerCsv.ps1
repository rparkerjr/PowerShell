<#
 .Synopsis
  Move files listed in a CSV to a single location. 

 .Description
  Uses a CSV created using the Map-Directory script to batch move documents to a single folder.
  Script will look for a "FullMap.csv" file in the same directory as the script itself. Creates 
  a log file indicating the status of each file listed in the CSV. 

 .Example
  # example here

#>

$importFile = ".\FullMap.csv"
$import = Import-Csv -Path $importFile -Delimiter ","
## Change the dump location as needed.
$dump = "C:\DUMP"
$runDate = Get-Date -Format "yyyyMMdd"
$logFile = ".\Moved Files " + $runDate + ".LOG"

$rowsToProcess = $import.Length
$rowsProcessed = 0
$rowsCopied = 0
$rowsDirectory = 0
$rowsErrored = 0

$outstring = (Get-Date -Format "yyyyMMdd HH:mm:ss") + " [ START ] Moving files listed in '" + $importFile + "'"
Add-Content -Path $logFile $outstring
foreach ($row in $import) {
    $path = $row.FullName
    $fname = Split-Path $row.FullName -Leaf
    $fulldump = Join-Path $dump $fname

    if (Test-Path -Path $path -PathType Leaf) {
        #Write-Host $fulldump
        if (Test-Path -Path $fulldump -PathType Leaf) {
            #Write-Host "duplicate"
            $fulldump += $rowsProcessed
        }
        ## Swap the comments to change this from copy to move
        ##Copy-Item -Path $path -Destination $fulldump
        Move-Item -Path $path -Destination $fulldump
        $outstring = (Get-Date -Format "yyyyMMdd HH:mm:ss") + " [SUCCESS] File moved: '" + $path + "'"
        #Write-Host $outstring
        Add-Content -Path $logFile $outstring
        $rowsCopied += 1
    } 
    elseif (Test-Path -Path $path -PathType Container) {
        ## Filepath is a directory, skipped
        $outstring = (Get-Date -Format "yyyyMMdd HH:mm:ss") + " [SKIPPED] Path is a folder: '" + $path + "'"
        #Write-Host $outstring
        Add-Content -Path $logFile $outstring
        $rowsDirectory += 1
    }
    else {
        $outstring = (Get-Date -Format "yyyyMMdd HH:mm:ss") + " [ ERROR ] None such file: '" + $path + "'"
        #Write-Host $outstring
        Add-Content -Path $logFile $outstring
        $rowsErrored += 1
    }
    $rowsProcessed += 1
}

$outstring = (Get-Date -Format "yyyyMMdd HH:mm:ss") + " [  END  ] Moved files to '" + $dump + "'"
Add-Content -Path $logFile $outstring

Write-Host("Rows to process :          " + $rowsToProcess)
Write-Host("Rows copied :              " + $rowsCopied)
Write-Host("Rows skipped (directory) : " + $rowsDirectory)
Write-Host("Rows errored :             " + $rowsErrored)
Write-Host("Rows processed :           " + $rowsProcessed)
