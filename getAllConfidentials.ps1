
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 

function Get-ConnectionDetails
{
    
   
    Write-Host "Use User with admin credentials"
    $userName = Read-Host -Prompt 'Admin account'
    Write-Host "If MFA is enabled use app password!"
    $secPassword = Read-Host -Prompt 'Password'-AsSecureString
    
    $path = Read-Host -Prompt 'Define the path to save CSV file. Example: C:\PowerShell'
      
    #Checking type of run
    Write-Host "Do you want run full script?"
    $full = Read-Host "Yes or No"

        while("yes","no","y","n" -notcontains $full)
        {
	    $full = Read-Host "Yes or No"
        }
    Write-Host "Do you want save data to list or CSV file?"
    $type = Read-Host "List or CSV"
    
            while("list","csv" -notcontains $type)
            {
            $type = Read-Host "List or CSV"
            } 
    if ($type -Contains "list") {
        Write-Hotst "------Saving data to list selected------"
        $desSiteUrl = Read-Host -Prompt 'Provide destination site URL'
        $createList = Read-Host -Prompt 'Provide destination list Title'
        $CSV = $false
    }    
    else{
        $CSV = $true
    }    

        If($full -Contains "y")
        { 
            Write-Host "------Full scan selected------"
            Write-Host "Provide your tenant admin Url"
            Write-Host "Example: DomainName-admin.sharepoint.com"
            $tenantUrl = Read-Host -Prompt 'Tenant URL: '
    
            #Getting the connection
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $secPassword
            Connect-SPOService -Url $tenantUrl -Credential $cred
    
    #Getting site collections list
    $sites =(Get-SPOSite).Url
    
    foreach ($siteUrl in $sites){
        Get-SPOAllWeb -Username $userName -secPassword $secPassword -siteUrl $siteUrl -path $path
    }
        Write-Host "-----Saving of list and sites finished----"
        if($CSV){
           Get-ListsDetails -userName $userName -secPassword $secPassword -path $path
        }
        else
        {
            Get-ListItems -siteUrl $siteUrl -userName $userName -secPassword $secPassword -listTitle $listTitle -path $path  -createList $createList -desSiteUrl $desSiteUrl
        }

        } 
        else
        {
            Write-Host "------Single list scan selected------"
            $SiteUrl = Read-Host -Prompt 'Provide site URL'
            $listTitle = Read-Host -Prompt 'Provide list Title'
            if($CSV){
            Get-ListItems -siteUrl $siteUrl -userName $userName -secPassword $secPassword -listTitle $listTitle -path $path 
            }
            else{
            Get-ListItems -siteUrl $siteUrl -userName $userName -secPassword $secPassword -listTitle $listTitle -path $path  -createList $createList  -desSiteUrl $desSiteUrl
            }
            
        } 
}


function Get-SPOAllWeb
{
   param (
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateNotNullOrEmpty()]	[string]$Username,
    [Parameter(Mandatory=$true,Position=2)]
	[ValidateNotNullOrEmpty()] [SecureString]$secPassword,
    [Parameter(Mandatory=$true,Position=3)]
    [ValidateNotNullOrEmpty()]	[string]$siteUrl,
    [Parameter(Mandatory=$true,Position=4)]
	[ValidateNotNullOrEmpty()]	[string]$path
        )
    $ctx = Create-SPOContext  -siteUrl $siteUrl -userName $userName -secPassword  $secPassword      

  $ctx.Load($ctx.Web.Webs)
  $ctx.Load($ctx.Web)
  try{
    $ctx.ExecuteQuery()

    Write-Host $ctx.Web.Url
   Get-AllLists -ctx $ctx -path $path
 
  if($ctx.Web.Webs.Count -ne 0)
  {
    Write-Host "Number of subsites in current site" $ctx.Web.Webs.Count
  foreach ($web in $ctx.Web.Webs)
  {
      Get-SPOAllWeb -Username $Username -secPassword $secPassword -siteUrl $web.Url -path $path
  }}
  else{
      Write-Host "No subsite"
  }
}
catch{
    Write-Output $_.Exception.GetType().FullName, $_.Exception.Message
}
}

function Get-AllLists
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()] $ctx,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string]$path
       
    )
    $path = $path+"\ListItems.csv"
    #Get All Lists
     $Lists = $ctx.web.Lists
     $web = $ctx.web.Url
     $ctx.Load($Lists)
     $ctx.ExecuteQuery()
     Write-Host "Ammount of the list in site"$Lists.Count
  
     #Iterate through each list in a site  
     ForEach($List in $Lists)
     {
         #Get the List Name
         $ListName = $List.Title
                  
        $results = @()
        $details = @{ListName=$ListName
        Site=$web} 
                              
        $results += New-Object PSObject -Property $details | export-csv -Path $path -NoTypeInformation -Append
        
       }
              
     
     
}

function Get-ListsDetails
{
    [CmdletBinding()]
    param
    (
        
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [string] $userName,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [SecureString] $secPassword,
        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()] [string] $path,
        [Parameter(Mandatory=$false, Position=3)]
        [ValidateNotNullOrEmpty()] [string] $createList
    )    
    $csvFile = $path+"\ListItems.csv"
    $table = Import-Csv $csvFile
    
    Foreach($row in $table){
        $siteUrl = $row.site
        $listTitle = $row.ListName
        if($oldSiteUrl -ne $siteUrl)
        {
            Write-Host "Current site" $siteUrl
        }
        Write-Host "Looking for items in: " $listTitle
        if(-not ([string]::IsNullOrEmpty($createList)) ){
        Get-ListItems -siteUrl $siteUrl -Username $userName -secPassword $secPassword -listTitle $listTitle -path $path
        }
        else{
        Get-ListItems -siteUrl $siteUrl -userName $userName -secPassword $secPassword -listTitle $listTitle -path $path  -createList $createList 
        }
        $oldSiteUrl = $siteUrl
    }
}

function Get-ListItems
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [string] $siteUrl,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string] $userName,
        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()] [SecureString] $secPassword,
        [Parameter(Mandatory=$true, Position=3)]
        [ValidateNotNullOrEmpty()] [string] $listTitle,
        [Parameter(Mandatory=$true, Position=4)]
        [ValidateNotNullOrEmpty()] [string] $path,
        [Parameter(Mandatory=$false, Position=5)]
        [string] $createList,
        [Parameter(Mandatory=$false, Position=5)]
        [string] $desSiteUrl

    )
        $path = $path+"\Items.csv" 
        $ctx = Create-SPOContext  -siteUrl $siteUrl -userName $userName -secPassword  $secPassword   
         
        #Get the List
        $List = $ctx.web.Lists.GetByTitle($listTitle)
        $qry = New-Object Microsoft.SharePoint.Client.CamlQuery
        $qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Confidential'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>"
        $ListItems = $List.GetItems($qry)
        $ctx.Load($ListItems)
    try{
        
        $ctx.ExecuteQuery()  
        
    foreach ($item in $ListItems)  {
        $results = @()
        if(-not ([string]::IsNullOrEmpty($item["File_x0020_Type"])))
        {
            Write-Host "Document: " $item["FileLeafRef"] "URL " $item["FileRef"] -ForegroundColor Green
            
            $details = @{
                            ItemName=$item["FileLeafRef"]
                            ItemURL=$item["FileRef"] 
                            ItemList=$listTitle
                       }
            if(-not ([string]::IsNullOrEmpty($createList)) )
            {          
            Create-ListItem -desSiteUrl $desSiteUrl -userName $userName -secPassword  $secPassword -listTitle $createList -itemTitle $item["FileLeafRef"] -itemUrl $item["FileRef"] 
             }     
        }
        else {
             
            $itemUrl = $item["FileDirRef"]+"/DispForm.aspx?ID="+$item["ID"]
            Write-Host "Item title: "  $item["Title"]  "URL"  $itemUrl -ForegroundColor Green
            $details = @{
                            ItemName=$item["Title"]
                            ItemURL=$itemUrl 
                            ItemList=$listTitle
                        }
            }
            if(-not ([string]::IsNullOrEmpty($createList)) )
            {          
            Create-ListItem -desSiteUrl $desSiteUrl -userName $userName -secPassword  $secPassword -listTitle $createList -itemTitle $item["Title"] -itemUrl $itemUrl 
             } 

         if([string]::IsNullOrEmpty($createList)) {                     
         $results += New-Object PSObject -Property $details | export-csv -Path $path -NoTypeInformation -Append  
        }  
    }
    }
    catch{
        Write-Output $_.Exception.GetType().FullName, $_.Exception.Message
    }

}

function Create-SPOContext
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [string] $siteUrl,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string] $userName,
        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()] [SecureString] $secPassword
        
    )

    $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $secPassword)     

    return $ctx
}

function Create-ListItem
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [string] $desSiteUrl,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string] $userName,
        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()] [SecureString] $secPassword,
        [Parameter(Mandatory=$true, Position=3)]
        [ValidateNotNullOrEmpty()] [string] $listTitle,
        [Parameter(Mandatory=$true, Position=4)]
        [ValidateNotNullOrEmpty()] [string] $itemTitle,
        [Parameter(Mandatory=$true, Position=5)]
        [ValidateNotNullOrEmpty()] [string] $itemUrl
        
    )

    $ctx = Create-SPOContext -siteUrl $desSiteUrl -userName $userName -secPassword  $secPassword  
    
    try{  
        $lists = $ctx.web.Lists  
        $list = $lists.GetByTitle($listTitle)  
        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation  
        $listItem = $list.AddItem($listItemInfo)  
        $listItem["Title"] = $itemTitle  
        $listItem["ItemURL"] = $itemUrl + ", Link to item"
        $listItem.Update()      
        $ctx.load($list)      
        $ctx.executeQuery()  
        Write-Host "Item Added with ID - " $listItem.Id      
    }  
    catch{  
        write-host "$($_.Exception.Message)" -foregroundcolor red  
    } 
}

Get-ConnectionDetails
