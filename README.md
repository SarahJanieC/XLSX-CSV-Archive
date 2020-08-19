# XLSX-CSV-Archive

This program serves two main purposes.
1. To convert xlsx files found in designated path to csv.
2. To move xlsx files to archive folder upon completetion. 



$CsvPath = [Environment]::GetFolderPath("Desktop") + '\filepath\filepath'
$archivePath = [Environment]::GetFolderPath("Desktop") + '\filepath\filepath\\archivepath'

#Gets the most recent CSV files in the folder specified by $CsvsPath
$Files = Get-ChildItem -Path $CsvPath| Where-Object {$_.Extension -like ".xlsx" } |Sort-Object LastWriteTime -Descending | Select-Object -first 10

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
        
        #Call function
        Csvcreate
}
