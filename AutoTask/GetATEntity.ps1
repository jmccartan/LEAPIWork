# how to install https://msdn.microsoft.com/en-us/library/dd878350(v=vs.85).aspx
function Get-ATEntity {
    <# 
    .SYNOPSIS
    Get-ATEntity queries AutoTask for the supplied values, Changes User Defined Fields into Object Members, and returns the objects
    
    .DESCRIPTION
    Get-ATEntity can be used with a single parameter to return all AT Entities of that type. This function returns objects that have the User Defined Fields added as members in order to make it easier to work with them in PowerShell. This means that the objects returned by this script cannot be updated back to AT unless these additional members are stripped back out. The Set-ATEntity function is designed to do this. 
    
    See https://www.autotask.net/help/content/Userguides/T_WebServicesAPIv1_5.pdf for information on the AT API
    
    .PARAMETER Entity
    The AT Entity being queried. Example: Account, Ticket, InstalledProduct
    
    .PARAMETER Field
    The AT entity field being checked. Example AccountName, id, PrimaryResource
    
    .PARAMETER Expression
    Must be one of:
    Equals
    NotEqual
    GreaterThan
    LessThan
    GreaterThanorEquals
    LessThanOrEquals
    BeginsWith
    EndsWith
    Contains
    IsNotNull
    IsNull
    IsThisDay
    Like
    NotLike
    SoundsLike
    
    .PARAMETER Value
    The Value to compare against.
    
    .PARAMETER UDF
    Switch to indicate that the field being queried is a User Defined Field. Necessary for AutoTask Query
    
    .EXAMPLE
    Get-ATEntity Account
    Get-ATEntity Account AccountName Equals 'A+ Widget Company'
    Get-ATEntity InstalledProduct 'Primary Domain Name' Equals apluswidgets.com -UDF
    #>
    [cmdletbinding()]
    Param (
        [Parameter(
            Position=0,
            Mandatory=$True)]
            [string]$Entity,
        [Parameter(
            Position=1,
            ParameterSetName='WithFields')]
            [string]$Field,
        [Parameter(
            Position=2,
            ParameterSetName='WithFields')]
            [ValidateSet("Equals",
                    "NotEqual",
                    "GreaterThan",
                    "LessThan",
                    "GreaterThanorEquals",
                    "LessThanOrEquals",
                    "BeginsWith",
                    "EndsWith",
                    "Contains",
                    "IsNotNull",
                    "IsNull",
                    "IsThisDay",
                    "Like",
                    "NotLike",
                    "SoundsLike")] 
            [string]$Expression,
        [Parameter(
            Position=3,
            ParameterSetName='WithFields')]
            [string]$Valueinit,
        [Parameter(
            Position=4,
            ParameterSetName='WithFields')]
            [switch]$UDF
    )
    
        $creds = Get-ATCredentials
        $AutoTask = New-WebServiceProxy https://webservices5.autotask.net/atservices/1.5/atws.wsdl -NameSpace ATWS -Credential ($creds)
        if ($Valueinit -ne ""){$Value = $Valueinit -replace "&","&amp;" -replace "'","&apos;"}
    
        function GenerateQuery{
            [cmdletbinding()]
            Param($entity,$field,$expression,$Value,$id=0,[boolean]$udf) 
            Write-Debug "In GenerateQuery"
            write-debug "Parameters: $entity, $field, $expression, $Value"
            if (($field -eq $Null) -and ($expression -eq $null) -and ($value -eq $null)){
                return [string]"<queryxml><entity>$Entity</entity><query><field>id<expression op='greaterthan'>$id</expression></field></query></queryxml>"
            }
            elseif (($field -ne $null) -and ($expression -ne $null) -and ($Value -ne $null)) {
                if ($udf){$fieldtag = "<field udf='true'>"} #need to add udf='true' if we're searching for a User Defined Field
                else {$fieldtag = '<field>'}
                return [string]"<queryxml><entity>$Entity</entity><query><condition>$fieldtag$Field<expression op=`'$expression`'>$Value</expression></field></condition><condition><field>id<expression op='GreaterThan'>$ID</expression></field></condition></query></queryxml>"
            }
            else {Write-error "Incorrect number of parameters passed"}
        }
        
        [int]$length = 500
        [int]$id = 0
    
        #AutoTaski API only retrieves 500 results at a time, so we have to loop through, starting each loop with the id where the last loop stopped
        While ($length -eq 500){
            write-debug "Parameters: Entity = $entity, Field = $field, Expression = $expression, Value = $Value"
            Write-Debug "Enter While loop. ID = $ID, Length = $length"
            if (($field -eq '') -and ($expression -eq '') -and ($Value -eq $null)){
                $queryxml = GenerateQuery -Entity $Entity -id $id
                }
            elseif (($field -ne $null) -and ($expression -ne $null) -and ($Value -ne $null)){
                $queryxml = GenerateQuery -Entity $Entity -field $Field -expression $Expression -Value $value -id $id $udf
                }
            write-debug "Query = $queryxml`nStarting Query"
            $QueryResults = $AutoTask.query($QueryXML)
            If ($QueryResults.Errors -ne $null)
            {
                $message = $QueryResults.Errors.Message
                Write-Error "AutoTask API Error: $message`nQuery: $queryxml"
                Break
            }
    
            Else {
                foreach ($result in $QueryResults.EntityResults){
                    $UDFs = $result.UserDefinedFields
                    foreach ($userdefinedfield in $UDFs){
                        $name = "UDF" + $userdefinedfield.name
                        [string]$UDFvalue = $userdefinedfield.Value -replace "`"","" -split "`n"
                        $result | Add-Member -NotePropertyName $name -NotePropertyValue $UDFValue   
                    }
                    [array]$results += $result
                }
                $length = $QueryResults.EntityResults.Length
                if ($QueryResults.EntityResults.Length -gt 0){ $id = $results[-1].id}
                Write-Debug "End of While loop. ID = $ID, Length = $length"
                }
            }
            Return $results
           
    }
    

    function Get-ATCredentials{
        $PWord = ConvertTo-SecureString -String "PASSWORD" -AsPlainText -Force
        $creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList USER@lightedge.com, $PWord
        return $creds
    }
    
    function Set-ATEntity {
    <# 
    .SYNOPSIS
    Set-ATEntity takes an object that has been generated by Get-ATEntity and updates AutoTask with the object's current properties
    
    .DESCRIPTION
    This scripts converts the 'UDFUser Defined Field Name' properties that are added to an AT object by the Get-ATEntity script back to the
    format expected by Autotask's update() method. This script expects an object produced by Get-ATEntity
    
    See https://www.autotask.net/help/content/Userguides/T_WebServicesAPIv1_5.pdf for information on the AT API
    
    .PARAMETER id
    The AT Entity id of the AutoTask object being updated. This is read from the input object
    
    .EXAMPLE
    $atentity | Set-ATEntity
    Set ATObject $atentity
    #>
    
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipelineByPropertyName=$true)][string]$id
     )
    BEGIN{
    
        $creds = Get-ATCredentials
        
        $AutoTask = New-WebServiceProxy https://webservices5.autotask.net/atservices/1.5/atws.wsdl -Credential ($creds)
        $NameSpace = $AutoTask.GetType().Namespace
        $UserDefinedField = "$NameSpace.UserDefinedField"
        $InstalledProduct = "$NameSpace.InstalledProduct"
      
        $userid = (Get-ATEntity Resource UserName Equals ($creds.UserName).Split("@")[0]).id
    
        $zone = $AutoTask.getZoneInfo($creds.UserName)
        $AutoTask.Url = $Zone.URL
    
        Function getatobject { #gets a fresh copy of the object from AT, so that we're able to manipulate it
        [cmdletbinding()]
        Param (
            [Parameter(
                Position=0,
                Mandatory=$True)]
                [string]$Entity,
            [Parameter(
                Position=1,
                ParameterSetName='WithFields')]
                [string]$Field,
            [Parameter(
                Position=2,
                ParameterSetName='WithFields')]
                [ValidateSet("Equals",
                        "NotEqual",
                        "GreaterThan",
                        "LessThan",
                        "GreaterThanorEquals",
                        "LessThanOrEquals",
                        "BeginsWith",
                        "EndsWith",
                        "Contains",
                        "IsNotNull",
                        "IsNull",
                        "IsThisDay",
                        "Like",
                        "NotLike",
                        "SoundsLike")] 
                [string]$Expression,
            [Parameter(
                Position=3,
                ParameterSetName='WithFields')]
                [string]$Valueinit,
            [Parameter(
                Position=4,
                ParameterSetName='WithFields')]
                [switch]$UDF
        )
            #$DebugPreference = Continue
       if ($Valueinit -ne ""){$Value = $Valueinit -replace "&","&amp;" -replace "'","&apos;"}
    
            function GenerateQuery{
                [cmdletbinding()]
                Param($entity,$field,$expression,$Value,$id=0,[boolean]$udf) 
                Write-Debug "In GenerateQuery"
                write-debug "Parameters: $entity, $field, $expression, $Value"
                if (($field -eq $Null) -and ($expression -eq $null) -and ($value -eq $null)){
                    return [string]"<queryxml><entity>$Entity</entity><query><field>id<expression op='greaterthan'>$id</expression></field></query></queryxml>"
                }
                elseif (($field -ne $null) -and ($expression -ne $null) -and ($Value -ne $null)) {
                    if ($udf){$fieldtag = "<field udf='true'>"} #need to add udf='true' if we're searching for a User Defined Field
                    else {$fieldtag = '<field>'}
                    return [string]"<queryxml><entity>$Entity</entity><query><condition>$fieldtag$Field<expression op=`'$expression`'>$Value</expression></field></condition><condition><field>id<expression op='GreaterThan'>$ID</expression></field></condition></query></queryxml>"
                }
                else {Write-error "Incorrect number of parameters passed"}
            }
        
            [int]$length = 500
            [int]$id = 0
    
            #AutoTaski API only retrieves 500 results at a time, so we have to loop through, starting each loop with the id where the last loop stopped
            While ($length -eq 500){
                write-debug "Parameters: Entity = $entity, Field = $field, Expression = $expression, Value = $Value"
                Write-Debug "Enter While loop. ID = $ID, Length = $length"
                if (($field -eq '') -and ($expression -eq '') -and ($Value -eq $null)){
                    $queryxml = GenerateQuery -Entity $Entity -id $id
                    }
                elseif (($field -ne $null) -and ($expression -ne $null) -and ($Value -ne $null)){
                    $queryxml = GenerateQuery -Entity $Entity -field $Field -expression $Expression -Value $value -id $id $udf
                    }
                write-debug "Query = $queryxml`nStarting Query"
                $QueryResults = $AutoTask.query($QueryXML)
                If ($QueryResults.Errors -ne $null)
                {
                    $message = $QueryResults.Errors.Message
                    Write-Error "AutoTask API Error: $message`nQuery: $queryxml"
                    Break
                }
    
                Else {
                    foreach ($result in $QueryResults.EntityResults){
                        [array]$results += $result
                    }
                    $length = $QueryResults.EntityResults.Length
                    if ($QueryResults.EntityResults.Length -gt 0){ $id = $results[-1].id}
                    Write-Debug "End of While loop. ID = $ID, Length = $length"
                    }
                }
                Return $results
        }
    
    }
    PROCESS{
    
    
        # Get a "fresh" copy of the AT Entity straight from the API
        $ATObject = getatobject $_.GetType().Name id Equals $id
    
        # Set the properties of the newly returned entity to match the properties of the passed entity
        $properties = $ATObject | gm | where{($_.MemberType -eq 'Property') -and ($_.Name -ne 'UserDefinedFields')} | select Name -ExpandProperty Name 
        foreach ($Property in $properties){
            $ATObject."$Property" = $_."$Property"
        }
        
        # Process the UDFFieldName fields of the passed entity back into AutoTask UserDefinedField Name/Value pairs
        $UDFs = $_ | gm | where Name -Match UDF | select Name -ExpandProperty Name 
        $NewUDFs = @()
        foreach ($UDF in $UDFs){
            $newUDF = New-Object "$NameSpace.UserDefinedField"
            $newUDF.Name = $UDF -replace "UDF",""
            $newUDF.Value = $_.$udf
            $newUDFs += $newUDF
        }
        $ATObject.UserDefinedFields = $NewUDFs
    
        # Update the object we got from AT, pass through any errors
        $updateobject = $AutoTask.update($ATObject)
        If ($updateobject.Errors -ne $null)
                {
                    $message = $QueryResults.Errors.Message
                    Write-Error "AutoTask API Error: $message`nQuery: $queryxml"
                    Break
                }
        else {}
        
    }    
    END{}   
    }
    