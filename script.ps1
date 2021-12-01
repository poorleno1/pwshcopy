#Requires –Modules 7Zip4Powershell
#Installation of 7zip module: Install-Module -Name 7Zip4Powershell -Confirm:$false -Force

$drive = "C:\scripts\DDT"
$source_path = Join-Path $drive -ChildPath "LZ"
$destination_path = Join-Path $drive -ChildPath "Preservation"
$trigger_path = Join-Path $drive -ChildPath "TRIGGER"

$RequestFile = Join-Path $trigger_path -ChildPath "MSGM0018_BULK_REQUEST.csv"
$DeliveryReport = Join-Path $trigger_path -ChildPath "MSGM0018_DELIVERY_REPORT.csv"
#Delivery report contains one entry for dbdirid
$hash_algoritm = "MD5"


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
       finally
       {
            Write-Host "Success."
       }
   }
   else
   {
        Write-Host "Folder $destination_path\$foldername already created."
        $success = $true
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

function Check-Hash ($File, $hash)
{
    
}


$DeliveryReportImported = Import-Csv $DeliveryReport

Import-Csv $RequestFile | ForEach-Object {
$Package =  $_.filename
$username = $_.filename.split("_")[0]
$PackageSource  = $_.filename.split("_")[1]
$PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
$7zipPassword = $_.passwords

$folders = Set-FolderStructure $username $PackageSource $destination_path

$ParentPath = $null
$ParentPath = $DeliveryReportImported | ? SMTP -eq $_.smtp | select -ExpandProperty "Landing Zone Path"

$source = join-path $ParentPath -ChildPath $Package

$zipfile = Copy-Files $source".7z*" $folders["DataFolder"]
$zipfileFullName =  join-path $folders["DataFolder"] -ChildPath $zipfile
$zipFileContent = Get-7Zip $zipfileFullName -Password $7zipPassword | select -ExpandProperty filename
$zipFileContentFileName =  Join-Path $folders["EvidenceFolder"] -ChildPath "$($Package)_7ZipContent.txt"
Add-Content $zipFileContentFileName -Value $zipFileContent
$evidencefile = Copy-Files $source"*.txt" $folders["EvidenceFolder"]

$hash = Get-Content $(Join-Path $folders["EvidenceFolder"] -ChildPath $Package"_hash.txt")
$PackageHash = Get-FileHash $(Join-Path $folders["DataFolder"] -ChildPath $zipfile) -Algorithm $hash_algoritm
$HashFileContentFileName =  Join-Path $folders["EvidenceFolder"] -ChildPath "$($Package)_HashLog.txt"
if ($hash -eq $PackageHash.Hash)
{
    Add-Content $HashFileContentFileName -Value "Hashes of $zipfileFullName matched!. $($PackageHash.Hash)"
}
else
{
    Add-Content $HashFileContentFileName -Value "ERROR: Hashes of $zipfileFullName did not matched!. LogHash: $hash, calculated hash:$($PackageHash.Hash)"
}





}