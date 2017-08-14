function Set-SSRSSubscription
{

  <#.SYNOPSIS
  run this to update StartDateTime or EndDate of a subscription
  .Description
 This function will update the subscription passed and update the datetime of either the EndDate or startdate
  .EXAMPLE
  set-RSsubscription -subscriptionid [guid] -EndDate "1/1/2020"
  updates the guid supplied to EndDate of 1-1-2020
  .EXAMPLE
  get-SSRSSubscripton -Path "\reports\areAwesome" -Owner "warren" | set-SSRSSubscripton -EndDate "1/1/2020"
  updates all Subscription in folder reports\areAwesome by Owner warren to end date of 1/1/2020
  .PARAMETER Proxy
  .EXAMPLE
  get-SSRSSubscripton -Path "\reports*" -WildCardSearch $true | set-SSRSSubscripton -stardate "1/1/2017"
  updates all Subscription with Path starting with \reportins, to start date of 1/1/2017
  This is assigned to the global Proxy variable that's created in the get-SSRSWebProxy -SSRSServername
  You can also create a seperate Proxy and manually assign
  .PARAMETER StartDateTime
  specify the StartDateTime for the Subscription - any valid datetime format
  .PARAMETER EndDate
  Specify EndDate for subscripion - any valid datetime format
#>

  [cmdletbinding(SupportsShouldProcess = $True)]
  param
  (
    [parameter(Mandatory = $false, Position = 0)]
    $Proxy = $script:Proxy,
    [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName)]
    [string]$SubscriptionID,
    [Parameter(Mandatory = $false, position = 2)]
    [datetime]$StartDateTime,
    [parameter(mandatory = $false, position = 3)]
    [datetime]$EndDate

  )

  BEGIN
  {
      
  }
  PROCESS
  {
    <#
    [ref] is a type accelerator for [System.Management.Automation.PSReference]
    need to assign a null value to create the variable when declared as psreference
    #>
    
    [ref]$extSettings = $null
    [ref]$Description = $null
    [ref]$activestate = $null
    [ref]$status = $null
    [ref]$eventtype = $null
    [ref]$matchdata = $null
    [ref]$ParamValues = $null
   
    #can also simply disable/enable subscriptions
    #$Proxy.DisableSubscription($SubscriptionID)
    #$Proxy.EnableSubscription($SubscriptionID)
    Try
    {
      $Proxy.GetSubscriptionProperties($Subscriptionid, $extSettings, $Description, $activestate, $status, $eventtype, $matchdata, $paramvalues) | Out-Null
    }
    Catch
    {
      Write-Error $error[0].Exception
    }
    #create new xml variable to add/change nodes
    [xml]$xmlmatch = $matchdata.Value 
    # $scrpt:xmltest = $xmlmatch

    #update StartDateTime if variable is not null
    if ($StartDateTime)
    {
      try
      {
        $xmlmatch.ScheduleDefinition.StartDateTime.InnerText = $StartDateTime
        Write-verbose "StartDateTime updated to $StartDateTime"
      }
      catch
      {
        Write-Error $error[0].Exception
      }
      
    }
    if ($EndDate)
    {
      #check to see if end date exists as a node
      $EndExists = $xmlmatch.SelectNodes("//*") | Select-Object name | Where-Object name -eq "EndDate"

      #if EndDate doesn't exist in nodes create it under scheduledefinition parent
      if ($EndExists -eq $null)
      {
        $child = $xmlmatch.CreateElement("EndDate")
        $child.InnerText = $EndDate
        try
        {
          $xmlmatch.ScheduleDefinition.AppendChild($child)
        }
        catch
        {
          Write-Error $error[0].Exception          
        }
      }
      else
      {
        try
        {
          $xmlmatch.ScheduleDefinition.EndDate.InnerText = $EndDate
        }
        catch
        {
          Write-Error $error[0].Exception
        }
         
      } 
      
      
    }

    #update the subscription if either variable has value
    if ($StartDateTime -ne $null -or $EndDate -ne $null)
    {
      Try
      {
        if ($PSCmdlet.ShouldProcess($SubscriptionID))
        {
          $Proxy.SetSubscriptionProperties($subscriptionID, $extSettings.Value, $Description.value, $eventtype.Value, $xmlmatch.OuterXml, $ParamValues.value)
        }
        
      }
      Catch
      {
        Write-Error $error[0].Exception
      }
    }

  }
}


function Get-SSRSSubscription
{

  <#.SYNOPSIS
  get lists of Subscription based on folder or author or Description
  .Description
 get lists of Subscription based on folder or author or Description. Get-SSRSWebProxy must already be run to set the global Proxy 
 variable
  .EXAMPLE
  Get-SSRSSubscription -Path "\foldername\" 
  gets all Subscription based on Path only 
  .EXAMPLE
  Get-SSRSSubscripitions -Path "\foldername\" -author "warren" -Description "test report"
  gets all Subscription from a specific folder for only user warren and test report Description
  .EXAMPLE
  Get-SSRSSubscription -Path "\foldername\" 
  gets all Subscription based on Path only 
  .PARAMETER Proxy
  this is set by the Get-SSRSWebProxy function as a global variable since its a webProxy
  .PARAMETER folder
  Enter name of the folder if you want all of the Subscription in that folder
  .PARAMETER Owner
  The creator of the report can also be use to predicate Subscription returned. get all Subscription based on Owner, or folder + Owner.
  .PARAMETER Description
  Description used in report can also be used to predicate Subscription returned
#>

  [cmdletbinding()]
  param
  ([parameter(Mandatory = $false, Position = 0)]
    $Proxy = $script:Proxy,
    [parameter(Mandatory = $false, Position = 1)]
    [string]$Path,
    [parameter(Mandatory = $false, Position = 2)]
    [string]$Owner,
    [parameter(mandatory = $false, Position = 3)]
    [string]$Description,
    [parameter(mandatory = $false, position = 4)]
    [Switch]$WildCardSearch
  )

  BEGIN
  {
    
  }
  PROCESS
  {

    $Subscriptions = $Proxy.listSubscriptions("/")
    #Path search    
    if ($Path)
    {
      if ($WildCardSearch)
      {
        $Subscriptions = $Subscriptions | Where-Object Path -like "$Path*"
      }
      else
      {   
        $Subscriptions = $Subscriptions | Where-Object Path -eq $Path
      }
    }
    #Owner search
    if ($Owner)
    {
      $Subscriptions = $subscriptions | Where-Object Owner -eq $Owner
    }

    #Description search
    if ($Description)
    {
      $Subscriptions = $Subscriptions | Where-Object Description -eq $Description
    }
    $Subscriptions
  }

}



function Set-SSRSWebProxy
{

  <#.SYNOPSIS
  Create SSRS webProxy script variable to use with the get and set Subscription
  .Description
 This will create the script scope Proxy that both the get and set Subscription need
  .EXAMPLE
  Get-SSRSWebProxy -SSRSServerName SSRS-1-YODA
  This creates a SSRS webProxy to SSRS-1-YODA to run the SOAP API calls
  .PARAMETER SSRSServerName
  The SSRS server name
#>

  [cmdletbinding(SupportsShouldProcess = $True)]
  param
  ([parameter(Mandatory = $True, Position = 0)]
    $SSRSServerName
  )

  BEGIN
  {

  }
  PROCESS
  {

    $reportServerurl = "http://$SSRSServerName/Reportserver/ReportService2010.asmx?wsdl"

    try
    {
      #Create Proxy
      if ($PSCmdlet.ShouldProcess($reportServerurl))
      {
        $Script:Proxy = New-WebServiceProxy -Uri $reportserverurl -UseDefaultCredential -Namespace SSRS -ErrorAction stop
      }
    }
    catch
    {
      Write-Error $error[0].Exception
    }
 
    
  }

}

