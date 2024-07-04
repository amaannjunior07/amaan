#Credentials
$ClientId = "abd8279c-e1f4-46d6-9e45-f153cbbf9ff0"
$ClientSecret = "WxkUVyEeh5tzY/B/bEK+fIsIxvHi2ixCkysAuz7w+LQ="

<#The code for applying site templates for Collaboration & Long Term Storage Template starts from here#>

#Application Insights details
#Instrumentation Key = sharepoint-dev 
$InstrumentationKey = "fdea3cd9-23da-4d8d-9f60-cf2d2c9e6b48"
$Telclient = New-Object "Microsoft.ApplicationInsights.TelemetryClient"
$Telclient.InstrumentationKey = $InstrumentationKey

#variables used to log information or exception to App Insights
$ApplicationName = "DWToolsAppUAT"
$RunbookName = "DWToolsProvisioning"
$EnvironmentName = "IT-Staging"

#variables
$SiteURL = "https://exelixis.sharepoint.com/sites/DWPDepartmentTest"
#UAT SiteURL
$DWToolSiteURL = "https://exelixis.sharepoint.com/sites/DWToolsAppUAT"
#Set Page Title
$Title = "Welcome to $ReqName"
#Set Page Banner Image
$BannerImageUrl = "/sites/ExelixisCDN/Images1/High Resolution images/DepartmentSite.png"
$RequestID = 2419

#Function for Sending email on exception
function ExceptionEmail {     
                      
    Connect-PnPOnline -Url $DWToolSiteURL -ClientId $ClientId -ClientSecret $ClientSecret -WarningAction Ignore
    $Subject = "SiteTemplateAutomation Job Exception"
    $JobOwner = "aabdul@exelixis.com"
    $HtmlBody = @"
<table cellspacing="0" width="600" border="1" style="font-family: Calibri;">
<tbody>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Request Name</td>
<td style="width: 70%; padding: 5px;">$ReqName</td>
</tr>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Environment Name</td>
<td style="width: 70%; padding: 5px;">$EnvironmentName</td>
</tr>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Runbook Name</td>
<td style="width: 70%; padding: 5px;">$RunbookName</td>
</tr>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Exception</td>
<td style="width: 50%; padding: 5px;">$ErrorMessage</td>
</tr>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Site URL</td>
<td style="width: 70%; padding: 5px;">$SiteURL</td>
</tr>
<tr>
<td style="font-weight: bold; width: 30%; padding: 5px;">Template Name</td>
<td style="width: 70%; padding: 5px;">$Template Site Template</td>
</tr>
</tbody>
</table>
<p style="font-family: Calibri;">
Please check the run history of the Job and take corrective actions.
</p>
"@
    Send-PnPMail -To $JobOwner -Subject $Subject -Body $Htmlbody
}

#Site Template function
function SiteTemplateAutomation {
    param (
        [Parameter(Mandatory = $true)]
        [string] $SiteURL,
        [Parameter(Mandatory = $true)]
        [string] $ReqName,
        [Parameter(Mandatory = $true)]
        [bool] $bool
    )
    
    try {
        # Connect to SharePoint using client ID and client secret
        Connect-PnPOnline -Url $SiteURL -ClientId $ClientId -ClientSecret $ClientSecret -WarningAction Ignore
        Write-Output "Connection success"

        # Apply Exelixis theme
        Set-PnPWebTheme -Theme "Exelixis Theme"
        Write-Output "Applied Exelixis Theme"
        #Logging in traces to Application Insights
        $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-Apply Exelixis Theme success:$($ReqName)") 
        $TelClient.Flush()
         
        # Set Page Title,name,header
        Set-PnPClientSidePage -Identity $PageName -Name "Home" -Title $Title -HeaderType Custom -ServerRelativeImageUrl $BannerImageUrl -TranslateX 13.5 -TranslateY 17.0 
        Write-Output "Set Page Title, Page name and Page banner"
        #Logging in traces to Application Insights
        $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-Set Page Title,Name,Banner success:$($ReqName)") 
        $TelClient.Flush()

        # Disable social bar on home page
        Set-PnPSite -Identity $SiteURL -SocialBarOnSitePagesDisabled $bool
        Write-Output "Social Bar on Site Pages is disabled/enabled based on true/false values"
        #Logging in traces to Application Insights
        $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-Enable/Disable Social Bar on Site Pages success:$($ReqName)") 
        $TelClient.Flush()
        
        $PageName = "Home"
        
        $page = Get-PnPClientSidePage -Identity $PageName
        # Save the page as template to be reused later-on
        $page.Save($PageName)
        Write-Output "Save $PageName success"
        #Logging in traces to Application Insights
        $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-Save $PageName success:$($ReqName)") 
        $TelClient.Flush()
    }
    catch {
        $Action = "Apply Site Template"
        $ErrorMessage = $_.Exception.Message
        Write-Output "Error: $ErrorMessage"
        #Logging in exceptions to Application Insights
        $TelClient.TrackException("ERROR-$($ApplicationName)-$($RunbookName)-$($Action)-$($ErrorMessage):$($ReqName)") 
        $TelClient.Flush()
        ExceptionEmail
        
    }
}

#Specify the list name
$listNameRequest = "DWToolsRequests"

#Define query
$Query = "
<View>
  <Query>
    <Where>
      <And>
        <Or>
          <Eq>
            <FieldRef Name='Template'/>
            <Value Type='Lookup'>Collaboration</Value>
          </Eq>
          <Eq>
            <FieldRef Name='Template'/>
            <Value Type='Lookup'>Long Term Storage</Value>
          </Eq>
        </Or>
        <And>
          <Eq>
            <FieldRef Name='Status'/>
            <Value Type='Choice'>Approved</Value>
          </Eq>
          <IsNotNull>
            <FieldRef Name='WorkloadID'/>
          </IsNotNull>
          <Eq>
            <FieldRef Name='ID'/>
            <Value Type='Integer'>$RequestID</Value>
          </Eq>
        </And>
      </And>
    </Where>
  </Query>
</View>
"


try {
    
    Connect-PnPOnline -Url $DWToolSiteURL -ClientId $ClientId -ClientSecret $ClientSecret -WarningAction Ignore

    # Retrieve items from the SharePoint list
    $ListItems = Get-PnPListItem -List $listNameRequest -Query $Query

    # Check if any items were retrieved
    if ($ListItems.Count -lt 1) {
        Write-Output "No Records are available"
    
    }
    else {
        # Loop through the retrieved items
        foreach ($Item in $ListItems) {
            $SPSiteURL = $Item["URL"]
            $ReqName = $Item["Title"]
            $Template = $Item["Template"].LookupValue
            
            if ($Template -eq "Collaboration") {
                Write-Output "SP Site Template: $Template"
                SiteTemplateAutomation -SiteURL $SPSiteURL -SiteTitle $ReqName -bool $true
                #Logging in traces to Application Insights
                $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-Collaboration SiteTemplate Automation Completed:$($RequestName)") 
                $TelClient.Flush()
        
            }
            else {
                Write-Output "SP Site Template: $Template"
                SiteTemplateAutomation -SiteURL $SPSiteURL -SiteTitle $ReqName -bool $false
                #Logging in traces to Application Insights
                $TelClient.TrackTrace("INFO-$($ApplicationName)-$($RunbookName)-LongTermStorage SiteTemplate Automation Completed:$($RequestName)") 
                $TelClient.Flush()
            }
        
        }
    }

}
catch {
    $Action = ""
    $ErrorMessage = $_.Exception.Message
    Write-Output "Error: $ErrorMessage"
    #Logging in exceptions to Application Insights
    $TelClient.TrackException("ERROR-$($ApplicationName)-$($RunbookName)-$($Action)-$($ErrorMessage)") 
    $TelClient.Flush()
    ExceptionEmail
}
<#DWToolsSiteTemplate Automation code ends here#>