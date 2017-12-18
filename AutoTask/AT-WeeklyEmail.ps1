# weekly email of open and unassigned autotask tickets - 9/15/17
# needs autotask credentials and o365 credentials updated to function 
# idea: add a list of open ticket counts by user to email (ranking)


# internal function to build the XML query 
function New-ATWSQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$UDF,
        [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
        [String[]]$QueryText
 
    )
    # List of allowed operators in QueryXML
    $Operators = @('and', 'or', 'begin')
 
    # List of all allowed conditions in QueryXML
    $Conditions = @('Equals', 'NotEqual', 'GreaterThan', 'LessThan', 'GreaterThanorEquals', 'LessThanOrEqual', 
        'BeginsWith', 'EndsWith', 'Contains', 'IsNotNull', 'IsNull', 'IsThisDay', 'Like', 'NotLike', 'SoundsLike')
 
    $NoValueNeeded = @('IsNotNull', 'IsNull', 'IsThisDay')
 
    # Create an XML document object. Only used to create XML elements.
    $xml = New-Object XML
 
    # Create base element and add a single Entity definition to it.
    $queryxml = $xml.CreateElement('queryxml')
    $entityxml = $xml.CreateElement('entity')
    $queryxml.AppendChild($entityxml) | Out-Null
 
    # Entity is the first element of the querytext
    $Entityxml.InnerText = $QueryText[0]
 
    # Create an XML element for the query tag.
    # It will contain all conditions.
    $Query = $xml.CreateElement('query')
    $queryxml.AppendChild($Query) | Out-Null
 
    # Set generic pointer $Node to the query tag
    $Node = $Query
 
    # Create an index pointer that starts on the second element
    # of the querytext array
    For ($i = 1; $i -lt $QueryText.Count; $i++) {
        Switch ($QueryText[$i]) {
            # Check for operators
            {$Operators -contains $_} {
                # Element is an operator. Add a condition tag with
                # attribute 'operator' set to the value of element
                $Condition = $xml.CreateElement('condition')
                If ($_ -eq 'begin') {
                    # Add nested condition
                    $Node.AppendChild($Condition) | Out-Null
                    $Node = $Condition
                    $Condition = $xml.CreateElement('condition')
                }
                If ('or', 'and' -contains $_) {
                    $Condition.SetAttribute('operator', $_)
                }
 
                # Append condition to current $Node
                $Node.AppendChild($Condition) | Out-Null
 
                # Set condition tag as current $Node. Next field tag
                # should be nested inside the condition tag.
                $Node = $Condition
                Break
            }
            # End a nested condition
            'end' {
                $Node = $Node.ParentNode
                Break
            }
            # Check for a condition
            {$Conditions -contains $_} {
                # Element is a condition. Add an expression tag with
                # attribute 'op' set to the value of element
                $Expression = $xml.CreateElement('expression')
                $Expression.SetAttribute('op', $_)
 
                # Append condition to current $Node
                $Node.AppendChild($Expression) | Out-Null
 
                # Not all conditions need a value. 
                If ($NoValueNeeded -notcontains $_) {
                    # Increase pointer and add next element as 
                    # Value to expression
                    $i++
                    $Expression.InnerText = $QueryText[$i]
                }
 
                # An expression closes a field tag. The next
                # element refers to the next level up.
                $Node = $Node.ParentNode
 
                # If the parentnode is a conditiontag we need
                # to go one more step up
                If ($Node.Name -eq 'condition') {
                    $Node = $Node.ParentNode
                }
                Break
            }
            # Everything that aren't an operator or a condition is treated
            # as a field.
            default {
                # Create a field tag, fill it with element
                # and add it to current Node
                $Field = $xml.CreateElement('field')
                $Field.InnerText = $QueryText[$i]
                $Node.AppendChild($Field) | Out-Null
 
                # If UDF is set we must add an attribute to the field
                # tag. But only once!
                If ($UDF) {
                    $Field.SetAttribute('udf', 'true')
                    # Only the first field can be UDF
                    $UDF = $false
                }
 
                # The field tag is now the current Node
                $Node = $Field
            }
        }
    }
 
    # Return formatted XML as text
    $queryxml.OuterXml
 
} 

# internal function to set ticket type based on AT types
function ATTicketType  {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$statusLookUp
    )
    switch ($statusLookUp) {
        1 {"New"}
        5 {"Complete"}
        7 {"Waiting Customer"}
        8 {"In Progress"}
        9 {"Waiting Materials"}
        10 {"Dispatched"}
        11 {"Escalate"}
        12 {"Waiting Vendor"}
        13 {"Waiting Approval"}
        14 {"On Hold Hold"}
        15 {"Scheduled"}
        16 {"Approved by by PRM"}
        17 {"Managed App Delivery"}
        18 {"Cancelled"}   
        default {"Ticket Type Not Found - " + $statusLookUp } 
    }
}

function CalcDeltaPercentage  {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$thisWeekCount
    )
    try {
        # read in last week's count  -- ASSUMPTION - THE CSV FILE IS THERE AND LIVES IN CURRENT DIRECTORY
        $lastWeek = (Get-Content LastWeekCount.csv)
        $returnValue = "<br /><br />Total Open Ticket Count Last Week = " + $lastWeek + "<br /><br />"
        $percentChange = [math]::Round(((($thisWeekCount - $lastWeek) / $lastWeek) * 100))
        if ($percentChange -lt 0) {
            $percentChange = $percentChange * -1
            $returnValue = $returnValue +  "This is a decrease of " + $percentChange + "% from last week. Going in the right direction!"
        }
        else {
            if ($percentChange -eq 0) {
                $returnValue = $returnValue + "No change in open ticket count from last week."
            }
            else {
               $returnValue = $returnValue + "This is an <strong>increase of " + $percentChange + "%</strong> from last week.&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://www.youtube.com/watch?v=-Wwg4wFRb3E'>C'mon man</a>  -  move that needle."
            }   
        }
        # return the value 
        $returnValue 
    }
    catch {
        # not able to report week over week change - bag it     
        $returnValue = ""     
    }
}

################################
# begin inline script execution
################################

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



# get total of open tickets 
Write-Host "Querying for open ticket count..."
# 5 is complete, 15 is scheduled, 18 is cancelled 
$ticketQuery = New-ATWSQuery "Ticket" "Status" "NotEqual" "5" "and" "Status" "NotEqual" "15" "and" "Status" "NotEqual" "18"    
$ticketResult = $atws.query($ticketQuery)
$totalOpenTicketCount = $ticketResult.EntityResults.Count
write-host "Total open tickets =  " $totalOpenTicketCount  
$openTicketsForNextWeek = $totalOpenTicketCount

#get get unassigned tickets
Write-Host "Querying tickets unassinged ..."
# 5 is complete, 15 is scheduled, 18 is cancelled 
$ticketQuery = New-ATWSQuery "Ticket" "AssignedResourceID" "ISNull" "and" "Status" "NotEqual" "5" "and" "Status" "NotEqual" "15" "and" "Status" "NotEqual" "18"    
$ticketResult = $atws.query($ticketQuery)
$unassignedCount = $ticketResult.EntityResults.Count
write-host "Retrieved " $unassignedCount  " tickets"
#$ticketResult.EntityResults | Select-Object *
foreach ($currTicket in $ticketResult.EntityResults) {
    $ticketTypeUnassigned = ATTicketType($currTicket.Status)        
    $currTicket| Add-Member -MemberType NoteProperty -Name StatusText -Value $ticketTypeUnassigned
}    

#todo - check for presence and clear out folder ? - 11/8/17 - finding that the file gets overwritten?
$dateForFileName = Get-Date -Format M-dd-yyyy
$fileNameAll = "c:\temp\ATUnassigned_" + $dateForFileName + ".csv"
$ticketResult.EntityResults | Select-Object TicketNumber, Title, statusText, CreateDate, LastActivityDate, Description | ConvertTo-Csv -NoTypeInformation | Select-Object | Set-Content -Path $fileNameAll   

#---------------
#get active resources
$userQuery = New-ATWSQuery "Resource" "Active" "Equals" "True"
$Result = $atws.query($userQuery)
$ATUsers = $Result.EntityResults | Select-Object id, Email, @{Name = "Fullname"; Expression = {$_.FirstName + " " + $_.LastName}} 

# query for status <> complete tickets for each user 
foreach ($userCurr in $ATUsers) {
    # query for open tickets
    Write-Host "Querying tickets for " $userCurr.Fullname " .... "
    $ticketQuery = New-ATWSQuery "Ticket" "AssignedResourceID" "Equals" $userCurr.id "and" "Status" "NotEqual" "5" "and" "Status" "NotEqual" "15" "and" "Status" "NotEqual" "18"        
    $ticketResult = $atws.query($ticketQuery)
    $userTicketCount = $ticketResult.EntityResults.Count
    $sendEmailTo = $userCurr.Email
    write-host "Retrieved " $userTicketCount " tickets"
    #$ticketResult.EntityResults | Select-Object *
    foreach ($currTicket in $ticketResult.EntityResults) {
        $currTicket | Add-Member -MemberType NoteProperty -Name TicketAssignedTo -Value $userCurr.Fullname
        $thisUser = $userCurr.Fullname
        $ticketStatusCurr = ATTicketType($currTicket.Status)
        $currTicket | Add-Member -MemberType NoteProperty -Name StatusText -Value $ticketStatusCurr
    }    
    $fileName = "c:\temp\" + $thisUser + " - " + $(get-date -Format MM-dd-yyyy-hh-mm) + " - ATTickets.csv"    
    # write the file
    $ticketResult.EntityResults | Select-Object TicketNumber, Title, statusText, CreateDate, LastActivityDate, TicketAssignedTo, Description | ConvertTo-Csv -NoTypeInformation | Select-Object | Set-Content -Path $fileName   
    # send email 
    if ($userTicketCount -gt 0) {    
        $subject = "Autotask Tickets - " + (get-date -format MM-dd-yyyy)
        $toName = $thisUser
        if ($unassignedCount -gt 0) {
            $unassignedCSV = $fileNameAll           
            $unassignedMessage = "<br/><br/><hr size=1/><br/>Also attached is the list of <strong>" + $unassignedCount + "</strong> open AT tickets. Please take a look at the unassigned list and assign yourself if appropriate."
        }
        else {
            $unassignedCSV = ""
            $unassignedMessage = "<br/><br/><hr size=1/><br/><strong>No unassgined tickets were found in Autotask - Nice!</strong>"           
        }
        $userCSV = $fileName
        
        #build up body of email 
        $totalOpenTicketText = "<br/><br/><hr size=1><br/>Total Open Ticket Count this week = " + $totalOpenTicketCount 
        $deltaText = CalcDeltaPercentage($totalOpenTicketCount)
        $totalOpenTicketText = $totalOpenTicketText + "<br /><br />" + $deltaText +  "<br /><br />" 


        if ($totalOpenTicketCount -lt 100){
            $goalText = "<br/><br/><strong>We made it under the century mark! Keep it going!</strong>"
        }
        else {
            $goalText = "<br/><br/>Let's see if we can get the total count under 3 digits?!"
        }
        $totalOpenTicketText = $totalOpenTicketText + $goalText
        
        
        $body = "<H3>Dear " + $toName + "<br/><br/>Your list of <strong>" + $userTicketCount + "</strong> open Autotask tickets is attached.<br/> <br/>Please review the attached and complete any closed tickets (or update status to reflect current state) in Autotask." + $unassignedMessage + $totalOpenTicketText + "<br/><br/><hr size=1><br/>Thanks!<br/><br/>Signed, Your Friendly Ticket Nanny :)</H3>"
        $usernameO365 = "insert o365 username here user@lightege.com or service account"
        $passwordO365 = ConvertTo-SecureString "insert user pw here in plain text" -AsPlainText -Force
        $credO365 = New-Object System.Management.Automation.PSCredential($usernameO365, $passwordO365)
    
        # NEED TO UNCOMMENT THE SEND EMAIL LINES BELOW FOR THE MAIL TO BE SENT - ALSO NEED TO UPDATE THE USERNAME@
        if ($unassignedCount -gt 0) {
            # Send-MailMessage -To $sendEmailTo -Cc "5418adcb.lightedge0.onmicrosoft.com@amer.teams.ms" -from USER@lightedge.com -Subject $subject -Body $body -BodyAsHtml -Attachments $unassignedCSV,$userCSV  -smtpserver smtp.office365.com -usessl -Credential $credO365 -Port 587 
        }
        else {
            # Send-MailMessage -To $sendEmailTo -Cc "5418adcb.lightedge0.onmicrosoft.com@amer.teams.ms" -from USER@lightedge.com -Subject $subject -Body $body -BodyAsHtml -Attachments $userCSV  -smtpserver smtp.office365.com -usessl -Credential $credO365 -Port 587 
            
        }
        # send mail line commented out for testing 
        # Send-MailMessage -To $sendEmailTo -Cc "5418adcb.lightedge0.onmicrosoft.com@amer.teams.ms" -from USER@lightedge.com -Subject $subject -Body $body -BodyAsHtml -Attachments $unassignedCSV,$userCSV  -smtpserver smtp.office365.com -usessl -Credential $credO365 -Port 587 
        Write-Host "send email here to " $sendEmailTo
        Write-Host "subject " $subject
        Write-Host "salutation " $toName
        Write-Host "attach 1 " $unassignedCSV
        Write-Host "attach 2 " $userCSV
        $body 
        Write-Host "---------------------------------------"
        Write-Host "<br /><br /> "
        Write-Host " "
        Write-Host " "
    }
}
# update the file for next week 
$openTicketsForNextWeek | Out-File "LastWeekCount.csv"
    
