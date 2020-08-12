#
#
# SarahJanieC 08/10/2020
#
# Purpose: Convert excel files to csv from designated folder and append store name to save file. 
#          Move excel files to archive folder upon completion.


$CsvPath = [Environment]::GetFolderPath("Desktop") + '\fileloc\fileloc'
$archivePath = [Environment]::GetFolderPath("Desktop") + '\fileloc\fileloc'

### Gets the most recent CSV files in the folder specified by $CsvsPath
$Files = Get-ChildItem -Path $CsvPath| 
            Where-Object {$_.Extension -like ".xlsx" } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -first 10

function Csvcreate()
{
       #Open excel in the background
        $xlsx = New-Object -comobject excel.application
        $xlsx.DisplayAlerts = $False


        #Parse files and save copies as csv files + move xlsx files to archive
        foreach($File in $Files)
        {
            $csv = $xlsx.workbooks.open($File.fullname)
            $csv.saveas($CsvsPath + "\" + $Store + "_" + $File.BaseName + '.csv',6 )
            Move-Item -Path $File.FullName -Destination $archivePath
        }

        $xlsx.quit()
        $xlsx=$null
        [GC]::Collect()
}

    if($Files -eq $null)
    {
        Write-Output "No Excel Files Found in Path: $CsvsPath"
    }
    else
    {
        #Input store appendage for new filename
        Write-Output "What is the abbreviated Store Name?"
        $store = Read-Host

        Csvcreate
    }
