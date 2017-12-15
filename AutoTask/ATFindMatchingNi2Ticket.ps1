# look for matching Ni2 tickets in AutoTask 
# some work to do yet - get the EVT # on each row for the output (see rhapsody api call)

# New-ATWSQuery include
. .\AT-New-ATWSQuery.ps1
# end external reference
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

# open file - sourced from AutoTask - looking for Ni2 in the Title - export to CSV
$ticketsToCheck = Import-Csv c:\Ni2Tickets.csv | Select-Object "Ticket Number", "Ticket Title", "Account"

# begin checking for multiple tickets - look for EVT- and get 

# assumptions for EVT tickets in AT 
$Ni2Prefix = "EVT-"
$lenOfNi2ID = 12

$completeList = @()

# process list from CSV - looking for multiple tickets with Ni2 ID
foreach ($ticket in $ticketsToCheck) {
    $inTitle = $ticket.'Ticket Title'
    $Ni2ID = $inTitle.Substring($inTitle.IndexOf($Ni2Prefix), $lenOfNi2ID).Trim()
    $ticketQuery = New-ATWSQuery "Ticket" "Title" "Contains" $Ni2ID
    $Result = $atws.query($ticketQuery)
    $foundCount = $Result.EntityResults.Length    
    Write-Host $Ni2ID " Count found " $Result.EntityResults.Length    
    foreach ($found in $Result.EntityResults) {
        $found | Add-Member -MemberType NoteProperty -Name GroupCount -Value $foundCount
        # add the EVT ID
        $found | Add-Member -MemberType NoteProperty -Name EVTID -Value $Ni2ID
        # look up status type - this should be broken into a shared function (or maybe use a lookup against the API)
        $statusText = switch ($found.Status)
        {
            1 {"New"}
            5 {"Complete"}
            7 {"Waiting Customer Customer"}
            8 {"In Progress Progress"}
            9 {"Waiting Materials Materials"}
            10 {"Dispatched"}
            11 {"Escalate"}
            12 {"Waiting Vendor Vendor"}
            13 {"Waiting Approval Approval"}
            14 {"On Hold Hold"}
            15 {"Scheduled"}
            16 {"Approved by by PRM"}
            17 {"Managed App App Delivery"}
            18 {"Cancelled"}    
        }
        $found | Add-Member -MemberType NoteProperty -Name StatusText -Value $statusText
        # look up resource name?
        if ($found.AssignedResourceID.Length -gt 0) {
            $resourceQuery = New-ATWSQuery "Resource" "id" "Equals" $found.AssignedResourceID
            $resource = $atws.query($resourceQuery)            
            if ($resource.EntityResults.Length -gt 0) {
                $resName = $resource.EntityResults.LastName + ", " + $resource.EntityResults.FirstName            
                $found | Add-Member -MemberType NoteProperty -Name ResourceName -Value $resName
            }
        }
    }
    $completeList += $Result
    

    # $Result.EntityResults. | Get-Member
    #$Result.EntityResults | Select-Object $Ni2ID, TicketNumber, Title, Status, AssignedResourceID | Format-Table
    #$Result.EntityResults | Select-Object $Ni2ID, TicketNumber, Title, Status, AssignedResourceID | Export-Csv "c:\temp\duptickets.csv"
  

    Write-Host "========================================="
}

Write-Host "===========done=============================="
Write-Host "Completed Tickets" 

#ISSUE - the evt isn't on each line of the output 
#$completeList.EntityResults | Select-Object EVTID, TicketNumber, Title, Status, StatusText, AssignedResourceID, ResourceName | Export-Csv "c:\temp\duptickets.csv"
$completeList.EntityResults | Select-Object GroupCount, EVTID, TicketNumber, Title, Status, StatusText, AssignedResourceID, ResourceName | ConvertTo-Csv -NoTypeInformation | Select-Object | Set-Content -Path "$(get-date -f MM-dd-yyyy-hh-mm)_EVTDups.csv"

# $Result.EntityResults | Select-Object $Ni2ID, TicketNumber, Title, Status, AssignedResourceID | Export-Csv "c:\temp\duptickets.csv"




#Write-host "Count of Tickets: " $($Result.EntityResults).Count
# Print any returned entities to console:
#$Result.EntityResults | Select-Object AccountID, Status, TicketNumber
# Import-Csv c:\scripts\test.txt | Sort-Object score | Select-Object -first 5
# from https://technet.microsoft.com/en-us/library/ee176900.aspx
