# IT Glue API calls to get account list 
$URI = "https://api.itglue.com"
$accessToken = "<<<<<need to update this with the access token for script to work>>>>>>>>>>>>>>>>>"
$headerToken = @{"x-api-key" = $accessToken}
$action = "/organizations?sort=-id"   # sort=-id returns newest to oldest records
$pageSize = 50
$pageNumber = 1
$paging = "&configurations?page[size]=" + $pageSize + "&page[number]=" + $pageNumber
$URIfull = $URI + $action + $paging

$allAccounts = @()

do 
{
    Write-Host "Getting Page Number... " $pageNumber    
    $OrgList = Invoke-RestMethod -Uri $URIfull -Method Get -ContentType application/vnd.api+json -Headers $headerToken
    $allAccounts += $OrgList
    $pageNumber ++    
    $paging = "&configurations?page[size]=" + $pageSize + "&page[number]=" + $pageNumber
    $URIfull = $URI + $action + $paging
} 
until ($OrgList.data.Count -eq 0)

Write-Host "Records Retrieved... " $allAccounts.data.count
# show the first 10 rows of the output 
$allAccounts.data.attributes | Select-Object name, short-name, id, created-at, updated-at, organization-status-id,organization-type-id    -first 10 | Format-Table

# file will be written to the local path where the script is executing - could preface with drive/path if needed 
$FullListOutputFile = "ITGlueAccountList.csv"
$allAccounts.data.attributes | Select-Object ITGlueID, name, short-name, created-at, updated-at,organization-status-id,organization-type-id   | export-csv $FullListOutputFile
Write-Host "File Output Complete"




#  strucutre of object outlined below with 3 get-member results 


# PS C:\WINDOWS\system32> $OrgList.data.attributes | get-member
# Name                     MemberType   Definition                                
# ----                     ----------   ----------                                
# Equals                   Method       bool Equals(System.Object obj)            
# GetHashCode              Method       int GetHashCode()                         
# GetType                  Method       type GetType()                            
# ToString                 Method       string ToString()                         
# alert                    NoteProperty object alert=null                         
# created-at               NoteProperty string created-at=2017-03-10T16:06:30.000Z
# description              NoteProperty object description=null                   
# name                     NoteProperty string name=Gill Studios Inc.             
# organization-status-id   NoteProperty int organization-status-id=8757           
# organization-status-name NoteProperty string organization-status-name=Active    
# organization-type-id     NoteProperty int organization-type-id=17790            
# organization-type-name   NoteProperty string organization-type-name=Customer    
# primary                  NoteProperty bool primary=False                        
# quick-notes              NoteProperty object quick-notes=null                   
# short-name               NoteProperty object short-name=null                    
# updated-at               NoteProperty string updated-at=2017-03-10T16:06:30.000Z

# PS C:\WINDOWS\system32> $OrgList.meta | Get-Member
# Name         MemberType   Definition                    
# ----         ----------   ----------                    
# Equals       Method       bool Equals(System.Object obj)
# GetHashCode  Method       int GetHashCode()             
# GetType      Method       type GetType()                
# ToString     Method       string ToString()             
# current-page NoteProperty int current-page=1            
# next-page    NoteProperty int next-page=2               
# prev-page    NoteProperty object prev-page=null         
# total-count  NoteProperty int total-count=2327          
# total-pages  NoteProperty int total-pages=47            


# PS C:\WINDOWS\system32> $OrgList.data | get-member
#    TypeName: System.Management.Automation.PSCustomObject
# Name        MemberType   Definition                                            
# ----        ----------   ----------                                            
# Equals      Method       bool Equals(System.Object obj)                        
# GetHashCode Method       int GetHashCode()                                     
# GetType     Method       type GetType()                                        
# ToString    Method       string ToString()                                     
# attributes  NoteProperty System.Management.Automation.PSCustomObject attribu...
# id          NoteProperty string id=1774816                                     
# type        NoteProperty string type=organizations                             

