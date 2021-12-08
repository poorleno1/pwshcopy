param(
[bool] $Validate=$true,
[bool] $Process=$false
)

#Requires –Modules 7Zip4Powershell
#Installation of 7zip module: Install-Module -Name 7Zip4Powershell -Confirm:$false -Force


$drive = "C:\scripts\DDT"
$source_path = Join-Path $drive -ChildPath "LZ"
$destination_path = Join-Path $drive -ChildPath "Preservation"
$trigger_path = Join-Path $drive -ChildPath "TRIGGER"

$RequestFile = Join-Path $trigger_path -ChildPath "MSGM0018_BULK_REQUEST.csv"
$DeliveryReport = Join-Path $trigger_path -ChildPath "MSGM0018_DELIVERY_REPORT.csv"
#Delivery report contains one entry for dbdirid
$hash_algoritm = "SHA256"


function Create-Folder ($destination_path, $foldername)
{
   $p = Join-Path $destination_path -ChildPath $foldername
   if (!(Test-Path $p))
   {
       try
       {
           Write-Host "Creating a folder $destination_path\$foldername"
           New-Item -Path $destination_path -Name $foldername -ItemType Directory | Out-Null
       }
       catch 
       {
           Write-Host "Cannot create a folder $foldername in $destination_path"
           break
       }
   }
   else
   {
        Write-Host "Folder $destination_path\$foldername already created."
   }
}




function Set-FolderStructure ($username, $PackageSource, $destination_path)
{
   $Folders = [ordered]@{}
   
   Create-Folder $destination_path $username
   
   $UserNamePath = Join-Path $destination_path -ChildPath $username
   Create-Folder $UserNamePath $PackageSource

   $DataSourceFolder = Join-Path $UserNamePath -ChildPath $PackageSource
   Create-Folder $DataSourceFolder "Data"

   $DataFolder = Join-Path $DataSourceFolder -ChildPath "Data"
   Create-Folder $DataFolder "Evidence"

   $Folders.Add("UserNamePath",$UserNamePath)
   $Folders.Add("DataSourceFolder",$DataSourceFolder)
   $Folders.Add("DataFolder",$DataFolder)
   $Folders.Add("EvidenceFolder",$DataFolder+"\Evidence")

   return $Folders
}

function Copy-Files ($source, $destination)
{
    Write-Host "Copying $source to $destination"

    try
    {
        $c = Copy-Item -Path $source -Destination $destination -PassThru
    }
    catch 
    {
        Write-Host "ERROR: File copy from $source to $destination failed." 
    }
    finally
    {
        if ($c.count -gt 0)
        {
            Write-Host "Success. Copied $($c.count) file(s)."
        }
        else
        {
            Write-Host "Failed. Copied $($c.count) files."
        }
    }
    return $c.name
}

Function RoboCopy-Files {
    [CmdletBinding()]Param($source, $destination,$log_path,$Overwrite=$false,$WithMove=$false)

    Write-Host "Copying $source to $destination"
    $robocopy_exe = get-command robocopy.exe | select -ExpandProperty source
    if (!($robocopy_exe))
    {
        Write-Host "Robocopy is not installed."
        break
    }
    $robocopy_args = $(Split-Path $source -Parent) +" "+$destination+" " + $(Split-Path $source -Leaf) + " /R:5 /W:3 /LOG+:$log_path\robocopy.log /NP"
    
    $p = Start-Process $robocopy_exe -ArgumentList $robocopy_args -Wait -PassThru

    switch ($p.ExitCode)
    {
        0 {Write-Verbose "Robocopy: No change"}
        1 {Write-Verbose "Robocopy: OK Copy"}
        2 {Write-Verbose "Robocopy: Extra files"}
        3 {Write-Verbose "Robocopy: OK Copy + Extra files"}
        4 {Write-Verbose "Robocopy: Mismatches"}
        5 {Write-Verbose "Robocopy: OK Copy + Mismatches"}
        6 {Write-Verbose "Robocopy: Mismatches + Extra files"}
        7 {Write-Verbose "Robocopy: OK Copy + Mismatches + Extra files"}
        8 {Write-Verbose "Robocopy: Fail"}
        9 {Write-Verbose "Robocopy: OK Copy + Fail"}
        10 {Write-Verbose "Robocopy: Fail + Extra files"}
        11 {Write-Verbose "Robocopy: OK Copy + Fail + Extra files"}
        12 {Write-Verbose "Robocopy: Fail + Mismatches"}
        13 {Write-Verbose "Robocopy: OK Copy + Fail + Mismatches"}
        14 {Write-Verbose "Robocopy: Fail + Mismatches + Extra files"}
        15 {Write-Verbose "Robocopy: OK Copy + Fail + Mismatches + Extra files"}
        16 {Write-Verbose "Robocopy: ***Fatal Error***";break}
        Default {}
    }
}

function Check-Hash ($InputFile, $hash, $evidencePath, $Package)
{
    Write-Host "Checking hash of $InputFile against $hash"
    $PackageHash = Get-FileHash $InputFile -Algorithm $hash_algoritm
    $HashFileContentFileName =  Join-Path $evidencePath -ChildPath "$($Package)_HashLog.txt"
    if ($hash -eq $PackageHash.Hash)
    {
        Add-Content $HashFileContentFileName -Value "Hashes of $zipfileFullName matched!. $($PackageHash.Hash)"
    }
    else
    {
        Add-Content $HashFileContentFileName -Value "ERROR: Hashes of $zipfileFullName did not matched!. LogHash: $hash, calculated hash:$($PackageHash.Hash)"
    }

}

function Get-7ZipContent ($InputFile, $LogFile, $7zipPassword)
{
    Write-Host "Getting content of $InputFile"
    try
    {
        $zipFileContent = Get-7Zip $zipfileFullName -Password $7zipPassword | select -ExpandProperty filename
        Add-Content $zipFileContentFileName -Value $zipFileContent
    }
    catch
    {
        Write-Host "Error: Problem with getting 7Zip content." -ForegroundColor Red
    }
}


function Main-Process ($param1, $param2)
{
    
$DeliveryReportImported = Import-Csv $DeliveryReport

Import-Csv $RequestFile | ForEach-Object {
    $Package =  $_.filename
    $username = $_.filename.split("_")[0]
    $PackageSource  = $_.filename.split("_")[1]
    $PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
    $7zipPassword = $_.password

    $folders = Set-FolderStructure $username $PackageSource $destination_path

    $ParentPath = $null
    $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "Landing Zone Path"

    $source = join-path $ParentPath -ChildPath $Package
    RoboCopy-Files -source $source"*.txt" -destination $folders["EvidenceFolder"] -log_path $folders["EvidenceFolder"] -Verbose
    RoboCopy-Files -source $source".7z*" -destination $folders["DataFolder"] -log_path $folders["EvidenceFolder"] -Verbose

    $zipfile = $Package+".7z.001"
    $zipfileFullName =  join-path $folders["DataFolder"] -ChildPath $zipfile
    $zipFileContentFileName =  Join-Path $folders["EvidenceFolder"] -ChildPath "$($Package)_7ZipContent.txt"
    Get-7ZipContent -InputFile $zipfileFullName -LogFile $zipFileContentFileName $7zipPassword
    
    $hash = $null
    #$hash = Get-Content $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") | Select-String -pa "$Package"
    #Select-String -Path $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") -Pattern $Package.ToString()
    $hash =(Select-String -Path $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") -Pattern "SHA256 for data"| Select-Object -ExpandProperty Line).split(":")[1].Trim()
    Check-Hash -hash $hash -InputFile $(Join-Path $folders["DataFolder"] -ChildPath $zipfile) -evidencePath $folders["EvidenceFolder"] -Package $Package
    }
}

function Main-Validate ($param1, $param2)
{
    Write-Host "Validation."
    $DeliveryReportImported = Import-Csv $DeliveryReport

    Import-Csv $RequestFile | ForEach-Object {
        $Package =  $_.filename
        $username = $_.filename.split("_")[0]
        $PackageSource  = $_.filename.split("_")[1]
        $PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
        $7zipPassword = $_.password

        $ParentPath = $null
        $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "Landing Zone Path"
        $FileSize = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "filesize"
    
        $zipfile = $Package+".7z.001"
        $zipfileFullName = join-path $ParentPath -ChildPath $zipfile
        $passwordCorrect = $false
        try
        {
            Get-7Zip $zipfileFullName -Password $7zipPassword | select -First 1 -ErrorVariable stop | Out-Null
            if ($?)
            {
                $passwordCorrect = $true
            }
        }
        catch 
        {
            Write-Host "Zip file error"
        }

        $hash = $null
        $hash =(Select-String -Path $(Join-Path $ParentPath -ChildPath $Package"_hash.txt") -Pattern "SHA256 for data"| Select-Object -ExpandProperty Line).split(":")[1].Trim()

        #Creating input for csv rollup report
        
        [System.Collections.ArrayList]$report = @()
        $val = [pscustomobject]@{
            "DBDIRID"=$_.DBDirID; 
            "SMTP"=$_.SMTP;
            "MinDate" = $_.filename.split("_")[-2];
            "MaxDate" = $_.filename.split("_")[-1];
            "Total Messages" =$null;
            "Signed Messages" =$null;
            "Encrypted Messages" = $null;
            "Normal Messages" = $null;
            "File Name" = Split-Path $_."Mailfile Path" -Leaf;
            "File size" = $FileSize;
            "File Hash" = $hash;
            "Package Name" = $_.filename;
            "Package Size" = $null;
            "Package Hash" = $null;
            "Password Verified" = $passwordCorrect

            }
        
        $report.Add($val)

        
        $reportName = $Package+"_rollup.csv"
        $report | Export-Csv -Path $(Join-Path $ParentPath -ChildPath $reportName) -NoTypeInformation
     
    }

}


if ($Validate -and $Process)
{
     Main-Validate
     Main-Process
}
elseif ($Validate)
{
     Main-Validate
     
}
elseif ($Process)
{
    Main-Process
}
else
{
    Write-Host "Choose either Validate or Process or both."
}
#Script logic:
# 1. Validate files that are on the list for copying.
#    a) verify hash of file that will be copied
#    b) open a zipped with password - verify that password is working
#    c) get the content of the zipped file
# 2. Write the result of verification to LZ folder keeping the file name convention: (0000001_NOTESNJIDE_08Sep2000_20Nov2016_rollup.csv)
# example rollout report:
#DBDIRID,SMTP,MinDate,MaxDate,Total Messages,Signed Messages,Encrypted Messages,Normal Messages,File Name,File size,File Hash,Package Name,Package Size,Package Hash,Password Verified
#12345,john.doe@db.com,1/1/2010,12/7/2014,,,,,doejoh.nsf,112645411,D4DE725B600D348470598E719633F1BD60FC04B0F4811031979517FE222DE9C3,12345_NOTESNJIEMEA_01Jan2010-07Dec2014.7z,2649554453,8DFADCDF0A50FB0A7FE5B5292F4735EC3F9CB7929201AFECA80B216A2503D3EF
# 3.Create a destination folder structure and copy files to destination.
