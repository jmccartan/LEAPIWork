# internal function to build the XML query 
function New-ATWSQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$UDF,
        [Parameter(Mandatory=$true,ValueFromRemainingArguments=$true)]
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
    For ($i = 1; $i -lt $QueryText.Count; $i++)
    {
        Switch ($QueryText[$i])
        {
            # Check for operators
            {$Operators -contains $_}
                {
                    # Element is an operator. Add a condition tag with
                    # attribute 'operator' set to the value of element
                    $Condition = $xml.CreateElement('condition')
                    If ($_ -eq 'begin')
                    {
                        # Add nested condition
                        $Node.AppendChild($Condition) | Out-Null
                        $Node = $Condition
                        $Condition = $xml.CreateElement('condition')
                    }
                    If ('or','and' -contains $_)
                    {
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
            'end'
                {
                    $Node = $Node.ParentNode
                    Break
                }
           # Check for a condition
            {$Conditions -contains $_} 
                {
                    # Element is a condition. Add an expression tag with
                    # attribute 'op' set to the value of element
                    $Expression = $xml.CreateElement('expression')
                    $Expression.SetAttribute('op', $_)
 
                    # Append condition to current $Node
                    $Node.AppendChild($Expression) | Out-Null
 
                    # Not all conditions need a value. 
                    If ($NoValueNeeded -notcontains $_)
                    {
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
                    If ($Node.Name -eq 'condition')
                    {
                        $Node = $Node.ParentNode
                    }
                    Break
                }
            # Everything that aren't an operator or a condition is treated
            # as a field.
            default
                {
                    # Create a field tag, fill it with element
                    # and add it to current Node
                    $Field = $xml.CreateElement('field')
                    $Field.InnerText = $QueryText[$i]
                    $Node.AppendChild($Field) | Out-Null
 
                    # If UDF is set we must add an attribute to the field
                    # tag. But only once!
                    If ($UDF)
                    {
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


#Import-Csv .\TicketList.csv 
# should check for TicketID 



$ticketQuery = New-ATWSQuery "Ticket" "TicketNumber" "Equals" "T20170831.0006" 
$Result = $atws.query($ticketQuery)

foreach ($ticket in $Result.EntityResults) {
#    $ticket.Status = 18
#    $atws.update($ticket)
    #$ticket | Get-Member
    # --------------------------
    # create a note
    $noteToCreate = New-Object  Microsoft.PowerShell.Commands.NewWebserviceProxy.AutogeneratedTypes.WebServiceProxy1k_net_atservices_1_5_atws_wsdl.TicketNote
    $noteToCreate.id = 0
    $noteToCreate.TicketID = $ticket.ProblemTicketId
    $noteToCreate.Title = "Ticket Update"
    $noteToCreate.Description = "Adding a note to ticket via API call."
    $noteToCreate.NoteType = 13
    $noteToCreate.Publish = 1
    $atws.create($noteToCreate)
    <# If ($noteToCreate.Errors -ne $null)
    {
        $message = $QueryResults.Errors.Message
        Write-Error "AutoTask API Error: $message`nQuery: $queryxml"
        Break
    } #>
else {}

}


Write-host "Count of Tickets: " $($Result.EntityResults).Count

# Print any returned entities to console:
$Result.EntityResults | Select-Object AccountID, Status, TicketNumber

# Import-Csv c:\scripts\test.txt | Sort-Object score | Select-Object -first 5
# from https://technet.microsoft.com/en-us/library/ee176900.aspx
