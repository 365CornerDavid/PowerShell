
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
     
        If($full -Contains "y")
        { 
            Write-Host "Full scan selected:"
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
    Write-Host "Saving of list and sites finished"
        Get-ListsDetails -userName $userName -secPassword $secPassword -path $path
        } 
        else
        {
            Write-Host "Single list scan selected:"
            $SiteUrl = Read-Host -Prompt 'Provide site URL'
            $listTitle = Read-Host -Prompt 'Provide list Title'
            Get-ListItems -siteUrl $siteUrl -userName $userName -secPassword $secPassword -listTitle $listTitle -path $path 
            
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
          
 $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
  $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $secPassword)
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
        [ValidateNotNullOrEmpty()] [string] $path
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

        Get-ListItems -siteUrl $siteUrl -Username $userName -secPassword $secPassword -listTitle $listTitle -path $path
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
        [ValidateNotNullOrEmpty()] [string] $path
    )
        $path = $path+"\Items.csv"    
        $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
        $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $secPassword)      
        
        
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
                               
         $results += New-Object PSObject -Property $details | export-csv -Path $path -NoTypeInformation -Append    
    }
    }
    catch{
        Write-Output $_.Exception.GetType().FullName, $_.Exception.Message
    }

}

Get-ConnectionDetails
