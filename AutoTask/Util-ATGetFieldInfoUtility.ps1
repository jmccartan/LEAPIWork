
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



$EntityType = 'Ticket'
$FieldName = 'QueueId'
$TicketInfo = $atws.GetFieldInfo($EntityType)
$QueueIdInfo = $TicketInfo | Where-Object {$_.Name -eq $FieldName}
$QueueIdInfo.PicklistValues | Select value,label

$EntityType = 'Note'
$FieldName = 'Publish'
$TicketInfo = $atws.GetFieldInfo($EntityType)
$QueueIdInfo = $TicketInfo | Where-Object {$_.Name -eq $FieldName}
$QueueIdInfo.PicklistValues | Select value,label