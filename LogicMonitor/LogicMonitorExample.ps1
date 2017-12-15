# example from http://www.techcolumnist.com/wp/2016/02/22/powershell-connect-to-logicmonitors-rest-api/
$user = "username"
$pass= "P@ssw0rd"
 
#get epoch time for current and x hours before
$date1 = Get-Date -Date "01/01/1970"
#get start time
$date2 = (Get-Date).AddHours(4)
$epochStart= (New-TimeSpan -Start $date1 -End $date2).TotalSeconds
#get end time
$date2 = (Get-Date).AddHours(5)
$epochEnd= (New-TimeSpan -Start $date1 -End $date2).TotalSeconds
#round the time to not have decimals
$epochStart= [math]::Round($epochStart)
$epochEnd= [math]::Round($epochEnd)
 
$filter = "_all~update" #check LM documentation on filters
$fields = "username,happenedOnLocal,description"
#build uri for access logs
$uri = "https://{account}.logicmonitor.com/santaba/rest/setting/accesslogs?sort=-happenedOn&filter=$filter,happenedOn>:$epochStart&fields=$fields"
#build base64Auth for the header
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user,$pass)))
#get the events
$events = Invoke-RestMethod -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -Uri $uri
$events #display events that were gathered