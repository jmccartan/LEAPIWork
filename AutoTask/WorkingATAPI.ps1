

# incomplete - the bottom section repeats... a wip 

Function GenerateQueryXML($type, $field, $op, $arg)
{
return "<queryxml><entity>$type</entity><query><field>$field<expression op=`"$op`">$arg</expression></field></query></queryxml>"
}

#start with region 1 to connect 

$URI = "https://webservices" + $ATRegion + ".Autotask.net/atservices/1.5/atws.wsdl"
 
$credential = Get-Credential -Message "Enter AutoTask Login Information"

$atws = New-WebServiceProxy -URI $URI -Credential $credential
Write-Host $URI
$zoneInfo = $atws.getZoneInfo($credential.username)
# update URL based on user name 
$URI = $zoneInfo.URL 
Write-Host $URI
# recreate teh atws object using new 
$atws = New-WebServiceProxy -URI $URI -Credential $credential


