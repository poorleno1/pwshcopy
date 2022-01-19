param(
[bool] $Validate=$true,
[bool] $Process=$true,
[parameter()][validateset('Notes','EV')]
[string] $DataSourceType = 'EV'
)

#
#
# THIS IS POC code 
#
#

#Requires –Modules 7Zip4Powershell
#Installation of 7zip module: Install-Module -Name 7Zip4Powershell -Confirm:$false -Force


$drive = "C:\scripts\DDT"
$source_path = Join-Path $drive -ChildPath "LZ"
$destination_path = Join-Path $drive -ChildPath "Preservation"
$trigger_path = Join-Path $drive -ChildPath "TRIGGER"

#This part is for Lotus Notes, use proper datasource parameter
$RequestFile = Join-Path $trigger_path -ChildPath "MSGM0018_BULK_REQUEST.csv"
$DeliveryReport = Join-Path $trigger_path -ChildPath "MSGM0018_DELIVERY_REPORT.csv"

#This part is for EV
#$RequestFile = Join-Path $trigger_path -ChildPath "ctp_ev_bulk_request_20211008.csv"
#$DeliveryReport = Join-Path $trigger_path -ChildPath "EV_Delivery_Report_IDS2_POC.csv"


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


function Build-ReportName ($rn, $ParentPath)
{
    
    #This function will build up to 10 rollup reports based on input file name and path
    $reportPath = Join-Path $ParentPath -ChildPath $rn".csv"
    $r = $rn
    $found = $false
    0..10 | % {

        if ($_ -eq 0)
        {
            $r1 = $r+"_rollup.csv"
        }
        else
        {
            $r1= $r + "_rollup_$_.csv"
        }
        $ret = $null
        $reportPath = Join-Path $ParentPath -ChildPath $r1
        if (-not $found)
        {
            if (-not (Test-Path $reportPath))
            {
                $ret = $r1
                $found = $true
                $ret
            }
        }
    }
}

function Main-Process ($param1, $param2)
{
    
#This function based on $RequestFile does following:
# 1. Creates a folder stucture

$DeliveryReportImported = Import-Csv $DeliveryReport

Import-Csv $RequestFile | ForEach-Object {
    $Package =  $_.filename
    $username = $_.filename.split("_")[0]
    $PackageSource  = $_.filename.split("_")[1]
    $PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
    $7zipPassword = $_.password

    $folders = Set-FolderStructure $username $PackageSource $destination_path

    $ParentPath = $null
    if ($DataSourceType -eq "Notes")
    {
        $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "Landing Zone Path"
    }
    elseif ($DataSourceType -eq "EV")
    {
        $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.custodian_email | select -ExpandProperty "Landing Zone Path"
    }
    else
    {

    }
    $source = join-path $ParentPath -ChildPath $Package
    RoboCopy-Files -source $source"*.txt" -destination $folders["EvidenceFolder"] -log_path $folders["EvidenceFolder"] -Verbose
    RoboCopy-Files -source $source".7z*" -destination $folders["DataFolder"] -log_path $folders["EvidenceFolder"] -Verbose

    $zipfile = $Package+".7z.001"
    $zipfileFullName = $null
    $zipfileFullName =  join-path $folders["DataFolder"] -ChildPath $zipfile
    $zipFileContentFileName =  Join-Path $folders["EvidenceFolder"] -ChildPath "$($Package)_7ZipContent.txt"
    Get-7ZipContent -InputFile $zipfileFullName -LogFile $zipFileContentFileName $7zipPassword
    
    $hash = $null
    $hashPath = $null
    $hashPath = Join-Path $ParentPath -ChildPath $Package"_hash.txt"
    if ($hashPath -and $(Test-Path $hashPath))
    {
        $hash =(Select-String -Path $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") -Pattern "SHA256 for data"| Select-Object -ExpandProperty Line).split(":")[1].Trim()
    }
    else
    {
        Write-Warning "Hash file is missing: $hashPath. Continuing."
    }
    #$hash = Get-Content $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") | Select-String -pa "$Package"
    #Select-String -Path $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt") -Pattern $Package.ToString()
    
    Check-Hash -hash $hash -InputFile $(Join-Path $folders["DataFolder"] -ChildPath $zipfile) -evidencePath $folders["EvidenceFolder"] -Package $Package
    }
}

function Main-Validate ($param1, $param2)
{
    #This function creates a validation report in $ParentPath (property "Landing Zone Path" in $DeliveryReportImported)
    
    Write-Host "Validation."
    $DeliveryReportImported = Import-Csv $DeliveryReport

    Import-Csv $RequestFile | ForEach-Object {
        $Package =  $_.filename
        $username = $_.filename.split("_")[0]
        $PackageSource  = $_.filename.split("_")[1]
        $PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
        $7zipPassword = $_.password

        $ParentPath = $null
        if ($DataSourceType -eq "Notes")
        {
            $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "Landing Zone Path"
            $FileSize = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "filesize" -ErrorAction SilentlyContinue
        }
        elseif ($DataSourceType -eq "EV")
        {
            $ParentPath = $DeliveryReportImported | ? SMTP -eq $_.custodian_email | select -ExpandProperty "Landing Zone Path"
            $FileSize = $DeliveryReportImported | ? SMTP -eq $_.custodian_email | select -ExpandProperty "filesize" -ErrorAction SilentlyContinue
        }
        else
        {

        }
        
        if (-not $ParentPath)
        {
            Write-Host "Cannot get `"Landing Zone Path`" value. Check value in $RequestFile for user $username."
            break
        }
        
        
        
        if (-not $FileSize)
        {
            Write-Warning "Cannot get `"filesize`" value. Check value in $RequestFile for user $username."
            #break
        }
        
        $zipfileFullName = $null
        $zipfile = $Package+".7z.001"
        try
        {
            $zipfileFullName = join-path $ParentPath -ChildPath $zipfile -ErrorAction Stop
        }
        catch
        {
            Write-Host "Cannot create a variable with full path to ZIP file."
            break
        }
        
        
        $passwordCorrect = $false
        if ($zipfileFullName -and $(Test-path $zipfileFullName))
        {
            # This will error out in case of wrong password
            Get-7Zip $zipfileFullName -Password $7zipPassword | select -First 1 -ErrorAction stop | Out-Null
            if ($?)
            {
                $passwordCorrect = $true
            }
        }
        else
        {
            Write-Host "Missing path to ZIP file $zipfile for user $username or File does not exists. Exiting"
            break
        }
        

        $hash = $null
        $hashPath = $null
        $hashPath = Join-Path $ParentPath -ChildPath $Package"_hash.txt"
        if ($hashPath -and $(Test-Path $hashPath))
        {
            $hash =(Select-String -Path $hashPath -Pattern "SHA256 for data"| Select-Object -ExpandProperty Line).split(":")[1].Trim()    
        }
        else
        {
            Write-Warning "Hash file is missing: $hashPath. Continuing."
        }

        

        #Creating input for csv rollup report
        
        $val = $null
        [System.Collections.ArrayList]$report = @()
        if ($DataSourceType -eq "Notes")
        {
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
                "File size" = [int]($FileSize/1kb);
                "File Hash" = $hash;
                "Package Name" = $_.filename;
                "Package Size" = $null;
                "Package Hash" = $null;
                "Password Verified" = $passwordCorrect
                }
        }
        elseif ($DataSourceType -eq "EV")
        {
               $val = [pscustomobject]@{
                "DBDIRID"=$_.dbdirid; 
                "SMTP"=$_.custodian_email;
                "MinDate" = $_.filename.Split("_")[-1].Split("-")[-2];
                "MaxDate" = $_.filename.Split("_")[-1].Split("-")[-1];
                "Total Messages" =$null;
                "Signed Messages" =$null;
                "Encrypted Messages" = $null;
                "Normal Messages" = $null;
                "File Name" = Split-Path $_.filename -Leaf;
                "File size" = [int]($FileSize/1kb);
                "File Hash" = $hash;
                "Package Name" = $_.filename;
                "Package Size" = $null;
                "Package Hash" = $null;
                "Password Verified" = $passwordCorrect
                }
        }

        if ($val)
        {
            $report.Add($val)
        }
        else
        {
            Write-host "Nothing to be added to rollup report for user $username"
            break
        }

        #$reportName = $Package+"_rollup.csv"
        $reportName = Build-ReportName $($Package) $ParentPath
        Write-Host "Saving a validation report to $(Join-Path $ParentPath -ChildPath $reportName)"
        $report |  convertto-csv -NoTypeInformation -Delimiter "," | % {$_ -replace '"',''} | Out-File $(Join-Path $ParentPath -ChildPath $reportName)
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
