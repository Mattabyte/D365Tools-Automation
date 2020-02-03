<#
This tool can enable and disable CRM plugins in a given CRM instance. (As Referenced by an Instance URL)

To Enable all plugins and workflows use the -Enable switch
To Disable all plugins and workflows use the -Disable switch

EXAMPLE USAGE:
./disablePluginsWorkflows.ps1 -Username admin -Password pass123 -DynamicsInstanceUrl https://someorgname.crm.dynamics.com -Disable -SaveState

This tool can also save/load states of plugins/workflows using the -SaveState and -Loadstate <path> switches.
State is saved as a JSON file in either a location specified, or the directory the script is run from.
#>


Param (
    [Parameter(mandatory=$true)]
    [string] $Username = "",
    [Parameter(mandatory=$true)]
    [string] $Password = "",
    [Parameter(mandatory=$true)]
    [string] $DynamicsInstanceUrl = "",
    [Parameter(mandatory=$false)]
    [switch] $Enable,
    [Parameter(mandatory=$false)]
    [switch] $Disable,
    [Parameter(mandatory=$false)]
    [switch] $SaveState,
    [Parameter(mandatory=$false)]
    [string] $SaveStateFile = "$PSScriptRoot\CRM-PluginWorkflowState.json",
    [Parameter(mandatory=$false)]
    [string] $LoadState
)

Import-Module Microsoft.Xrm.Data.Powershell

function EnablePluginStep {
    param
    (
        [Guid] $StepId    
    )

    Set-CrmRecordState -conn $CrmConnection -EntityLogicalName sdkmessageprocessingstep -Id $StepId `
    -StateCode Enabled -StatusCode Enabled
}

function DisablePluginStep {
    param
    (
        [Guid] $StepId    
    )

    Set-CrmRecordState -conn $CrmConnection -EntityLogicalName sdkmessageprocessingstep -Id $StepId `
    -StateCode Disabled -StatusCode Disabled
}

function Get-Workflows {
    param
    (
        [string]$State   
    )
#### WORKFLOWS ####
[xml] $workflowsFetchQuery=@"
<fetch mapping="logical" >
  <entity name="workflow">
    <all-attributes/>
        <filter type="and">
            <condition attribute="type" operator="eq" value="1" />
            <condition attribute="rendererobjecttypecode" operator="null" />
        <filter type="or">
            <condition attribute="category" operator="eq" value="0" />
        <filter type="and">
            <condition attribute="category" operator="eq" value="1" />
            <condition attribute="languagecode" operator="eq-userlanguage" />
        </filter>
            <condition attribute="category" operator="eq" value="3" />
            <condition attribute="category" operator="eq" value="4" />
        </filter>
        </filter>
  </entity>
</fetch>
"@

    $workflowsfetchexpr = New-Object Microsoft.Xrm.Sdk.Query.FetchExpression($workflowsFetchQuery.InnerXml)
    Write-Host "Fetching workflows..."
    $Workflows = $CrmConnection.RetrieveMultiple($workflowsfetchexpr)

    # Save the state of these workflows
    $ActiveWorkflows = ($Workflows.Entities | ? {$_.FormattedValues.Get_Item("statecode") -eq "Activated"})
    $DraftWorkflows = ($Workflows.Entities | ? {$_.FormattedValues.Get_Item("statecode") -eq "Draft"})
    $SystemState.Workflows.Enabled = ($ActiveWorkflows.Attributes | ? {$_.Key -eq "workflowid"}).value
    $SystemState.Workflows.Disabled = ($DraftWorkflows.Attributes | ? {$_.Key -eq "workflowid"}).value
    Write-Host "Found $($Workflows.Entities.count) workflows."
    switch($State){
        Active {return $ActiveWorkflows}
        Draft {return $DraftWorkflows}
    }
}

function Get-Plugins {
    param
    (
        [string]$State   
    )
    #### PLUGINS ####
    Write-Host "Retrieving plugins..."
    $assemblies = Get-CrmRecords -conn $CrmConnection -EntityLogicalName pluginassembly `
    -FilterAttribute customizationlevel -FilterOperator eq -FilterValue 1 `
    -Fields * -WarningAction SilentlyContinue

    # Instantiate a List to contain Steps/Already disabled ones
    $enabledsteps = New-Object System.Collections.Generic.List[object]
    $disabledsteps = New-Object System.Collections.Generic.List[object]

    # Loop all assemblies to get steps. 
    foreach($assembly in $assemblies.CrmRecords)
    {
        Write-Host 'Getting Steps for' $assembly.name

        # Get all registered steps for the assembly
        $sdkmessages = Get-CrmSdkMessageProcessingStepsForPluginAssembly `
        -conn $CrmConnection -PluginAssemblyName $assembly.name -WarningAction SilentlyContinue

        # Add only enabled step to the list
        foreach($enabledStep in ($sdkmessages | ? {$_.statecode -eq 'Enabled'}))
        {
            $enabledsteps.Add($enabledStep)
            $SystemState.Plugins.Enabled += $enabledStep.sdkmessageprocessingstepid
        }
        # Store disabled step to the list
        foreach($disabledstep in ($sdkmessages | ? {$_.statecode -eq 'Disabled'}))
        {
            $disabledsteps.Add($disabledstep)
            $SystemState.Plugins.Disabled += $disabledstep.sdkmessageprocessingstepid
        }
    }
    switch($State){
        Enabled {return $enabledsteps}
        Disabled {return $disabledsteps}
    }
}

#Prepare Build Agent (PROXY FIX)
[System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#Run Setup Security Protocal for TLS 1.2 - Required for CDS\XRM 9.x + 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null


$sourceconnection = "AuthType=Office365;Url=$DynamicsInstanceUrl;UserName=$Username;Password=$Password;RequireNewInstance=true;Timeout=00:05:00"
$CrmConnection = Get-CrmConnection -ConnectionString $sourceconnection -Verbose

if(!$CrmConnection.IsReady){
    throw 'Connection to CRM could not be established.'
} else {
    Write-Host "Connected to " $CrmConnection.ConnectedOrgFriendlyName
}


## STATE ##
#Build the state object
$SystemState = [ordered]@{
    Plugins=[ordered]@{
        Enabled=@()
        Disabled=@()
    }
    Workflows=[ordered]@{
        Enabled=@()
        Disabled=@()
    }
}


if ($Disable){
    #Disable all Active workflows
    $Workflows = Get-Workflows -State Active
    foreach ($Workflow in $Workflows){
        Write-Progress -Activity "Disabling Workflows" -Status "$($Workflows.IndexOf($Workflow))/$($Workflows.Count)" -CurrentOperation "$($Workflow.Attributes["name"])" 
        try {
            Set-CrmRecordState -conn $CrmConnection -EntityLogicalName $Workflow.LogicalName -Id $Workflow.Id.Guid -StateCode Draft -StatusCode Draft
        } Catch {
            Write-Host "Couldn't disable workflow: " $Workflow.Attributes["name"] " -- " $_.Exception.Message
        }
    }

    # Disable all enabled plugin steps.
    $Pluginsteps = Get-Plugins -State Enabled
    foreach($step in $Pluginsteps)
    {
        Write-Progress -Activity "Disabling Plugins" -Status "$($Pluginsteps.IndexOf($step))/$($Pluginsteps.Count)" -CurrentOperation "$($step.name)" 
        try{
            DisablePluginStep -StepId $step.sdkmessageprocessingstepid
        } catch {
            Write-Host "Couldn't disable:  " $step.name " -- " $_.Exception.Message
        }
    }
}

if ($Enable){
    #Enable all draft workflows
    $Workflows = Get-Workflows -State Draft
    foreach ($Workflow in $Workflows){
        Write-Progress -Activity "Enabling Workflows" -Status "$($Workflows.IndexOf($Workflow))/$($Workflows.Count)" -CurrentOperation "$($Workflow.Attributes["name"])" 
        try {
            Set-CrmRecordState -conn $CrmConnection -EntityLogicalName $Workflow.LogicalName -Id $Workflow.Id.Guid -StateCode Activated -StatusCode Activated
        } Catch {
            Write-Host "Couldn't enable workflow: " $Workflow.Attributes["name"] " -- " $_.Exception.Message
        }
    }

    # Enable all disabled steps.
    $Pluginsteps = Get-Plugins -State Disabled
    foreach($step in $Pluginsteps)
    {
        Write-Progress -Activity "Enabling Plugins" -Status "$($Pluginsteps.IndexOf($step))/$($Pluginsteps.Count)" -CurrentOperation "$($step.name)" 
        try{
            EnablePluginStep -StepId $step.sdkmessageprocessingstepid
        } catch {
            Write-Host "Couldn't enable:  " $step.name " -- " $_.Exception.Message
        }
    }
}

if ($SaveState){
    Get-Plugins
    Get-Workflows
    $SystemState | ConvertTo-Json | Out-File $SaveStateFile
    Write-Host "System state saved as $SaveStateFile"
}

if ($LoadState) {
    if(![System.IO.File]::Exists($LoadState)){
        throw 'No LoadState file specified to load from (path doesnt exist).'
    }
    $RestoreState = Get-Content $LoadState | ConvertFrom-Json
    

    #Enable workflows From State
    foreach ($Workflow in $RestoreState.Workflows.Enabled){
        Write-Progress -Activity "Enabling Workflows" -Status "$($RestoreState.Workflows.Enabled.IndexOf($Workflow))/$($RestoreState.Workflows.Enabled.Count)"
        try {
            Set-CrmRecordState -conn $CrmConnection -EntityLogicalName workflow -Id $Workflow -StateCode Activated -StatusCode Activated
        } Catch {
            Write-Host "Couldn't enable workflow: " $Workflow " -- " $_.Exception.Message
        }
    }

    #Enable plugins From State
    foreach($step in $RestoreState.Plugins.Enabled)
    {
        Write-Progress -Activity "Enabling Plugins" -Status "$($RestoreState.Plugins.Enabled.IndexOf($step))/$($RestoreState.Plugins.Enabled.Count)"
        try{
            EnablePluginStep -StepId $step
        } catch {
            Write-Host "Couldn't enable:  " $step " -- " $_.Exception.Message
        }
    }

    #Disable workflows from State
    foreach ($Workflow in $RestoreState.Workflows.Disabled){
        Write-Progress -Activity "Disabling Workflows" -Status "$($RestoreState.Workflows.Disabled.IndexOf($workflow))/$($RestoreState.Workflows.Disabled.Count)"  
        try {
            Set-CrmRecordState -conn $CrmConnection -EntityLogicalName workflow -Id $Workflow -StateCode Draft -StatusCode Draft
        } Catch {
            Write-Host "Couldn't disable workflow: " $Workflow " -- " $_.Exception.Message
        }
    }

    #Disable plugins from State
    foreach($step in $RestoreState.Plugins.Disabled)
    {
        Write-Progress -Activity "Disabling Plugins" -Status "$($RestoreState.Plugins.Disabled.IndexOf($step))/$($RestoreState.Plugins.Disabled.Count)" 
        try{
            DisablePluginStep -StepId $step
        } catch {
            Write-Host "Couldn't disable:  " $step " -- " $_.Exception.Message
        }
    }
}

Write-Host "Complete!"