$drive = "C:\DDT"
$source_path = Join-Path $drive -ChildPath "LZ"
$destination_path = Join-Path $drive -ChildPath "Preservation"
$trigger_path = Join-Path $drive -ChildPath "TRIGGER"

$RequestFile = Join-Path $trigger_path -ChildPath "MSGM0018_BULK_REQUEST.csv"
$DeliveryReport = Join-Path $trigger_path -ChildPath "MSGM0018_DELIVERY_REPORT.csv"
#Delivery report contains one entry for dbdirid


Import-Csv $RequestFile | ForEach-Object {
$Package =  $_.filename
$username = $_.filename.split("_")[0]
$PackageSource  = $_.filename.split("_")[1]
$PointOfInterest = $_.filename.split("_")[2]+"_"+$_.filename.split("_")[3]
$7zipPassword = $_.password
}



Import-Csv $DeliveryReport 