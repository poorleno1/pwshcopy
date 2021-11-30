$drive = "C:\DDT"
$source_path = Join-Path $drive -ChildPath "LZ"
$destination_path = Join-Path $drive -ChildPath "Preservation"
$trigger_path = Join-Path $drive -ChildPath "TRIGGER"

$RequestFile = Join-Path $trigger_path -ChildPath "MSGM0018_BULK_REQUEST.csv"
$DeliveryReport = Join-Path $trigger_path -ChildPath "MSGM0018_DELIVERY_REPORT.csv"
#Delivery report contains one entry for dbdirid


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
   Create-Folder $UserNamePath "DataSource"
   
   $PackageSourceFolder = Join-Path $UserNamePath -ChildPath "DataSource"
   Create-Folder $PackageSourceFolder $PackageSource
   Create-Folder $PackageSourceFolder "Data"
   
   $DataFolder = Join-Path $PackageSourceFolder -ChildPath "Data"
   Create-Folder $DataFolder "Evidence"

   $Folders.Add("UserNamePath",$UserNamePath)
   $Folders.Add("PackageSourceFolder",$PackageSourceFolder)
   $Folders.Add("PackageSourceFolder",$PackageSourceFolder)
   $Folders.Add("DataFolder",$DataFolder)
   $Folders.Add("EvidenceFolder",$DataFolder+"\Evidence")

   return $Folders
}



function Set-FolderStructure2 ($username, $PackageSource, $destination_path)
{
   $Folders = [ordered]@{}
   
   
   Create-Folder $destination_path $username
   
   $DatasourcePath = Join-Path $destination_path -ChildPath $username
   Create-Folder $DatasourcePath "DataSource"
   
   $PackageSourceFolder = Join-Path $DatasourcePath -ChildPath "DataSource"
   Create-Folder $PackageSourceFolder $PackageSource
   Create-Folder $PackageSourceFolder "Data"
   
   $DataFolder = Join-Path $PackageSourceFolder -ChildPath "Data"
   Create-Folder $DataFolder "Evidence"

   $Folders.Add("DatasourcePath",$DatasourcePath)
   $Folders.Add("PackageSourceFolder",$PackageSourceFolder)
   $Folders.Add("PackageSourceFolder",$PackageSourceFolder)
   $Folders.Add("DataFolder",$DataFolder)
   $Folders.Add("EvidenceFolder",$DataFolder+"\Evidence")

   return $Folders
}

function Copy-Files ($param1, $param2)
{
    
}


Import-Csv $RequestFile | ForEach-Object {
$Package =  $_.filename
$username = $_.filename.split("_")[0]
$PackageSource  = $_.filename.split("_")[1]
$PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
$7zipPassword = $_.password


}

Set-FolderStructure $username $PackageSource $destination_path

Import-Csv $DeliveryReport 