# Username and password for Autotask
# The username must begin with a backslash, otherwise Windows will add a domain
# element that Autotask do not understand.
# ======================================================================== 
# ~~~~!!!! for script to work - need to update username and password below
# ~~~~!!!! username should be inform of    \user@mail.com  ""
# best practice would be to either set up a service account to investigate whether AT supports an API key 
 
$username = "insert username here"
$password = ConvertTo-SecureString "insert password here" -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential($username, $password) 

# Web services URL for Autotask (updated to point at region 5 for LE)
$atws = New-WebServiceProxy -URI https://webservices5.Autotask.net/atservices/1.5/atws.wsdl -Credential $credentials
 
#$atws = New-WebServiceProxy -URI $ATurl -Credential $ATcredentials
#$zoneInfo = $atws.getZoneInfo($ATusername)
#$ATurl = $zoneInfo.URL.Replace(".asmx",".wsdl")
#$atws = New-WebServiceProxy -URI $ATurl -Credential $ATcredentials
  
$entity= $atws.getFieldInfo("Account")
  
foreach ($picklist in $entity) {
    $picklist | select Name,Label,Description | ft
    foreach ($values in $picklist.PicklistValues) { $values | select Label,Value,IsActive }
}
 
$output= $atws.getThresholdAndUsageInfo()
$output.EntityReturnInfoResults.message
