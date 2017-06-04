$currentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Debug -Message "CurrentPath: $currentPath"

# Load Common Code
Import-Module $currentPath\..\..\xSQLServerHelper.psm1 -Verbose:$false -ErrorAction Stop

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [parameter(Mandatory = $true)]
        [System.String]
        $InstanceName,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLServer,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLInstanceName,

        [parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $SQLAdminCredential
    )

    if(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName -ErrorAction SilentlyContinue)
    {
        $InstanceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName).$InstanceName
        $SQLVersion = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\Setup" -Name "Version").Version).Split(".")[0]
        $RSConfig = Invoke-Command -ComputerName . -Credential $SQLAdminCredential -ScriptBlock {
            $SQLVersion = $args[0]
            $InstanceName = $args[1]
            $RSConfig = Get-WmiObject -Class MSReportServer_ConfigurationSetting -Namespace "root\Microsoft\SQLServer\ReportServer\RS_$InstanceName\v$SQLVersion\Admin"
            $RSConfig
        } -ArgumentList @($SQLVersion,$InstanceName)
        if($RSConfig.DatabaseServerName.Contains("\"))
        {
            $RSSQLServer = $RSConfig.DatabaseServerName.Split("\")[0]
            $RSSQLInstanceName = $RSConfig.DatabaseServerName.Split("\")[1]
        }
        else
        {
            $RSSQLServer = $RSConfig.DatabaseServerName
            $RSSQLInstanceName = "MSSQLSERVER"
        }
        $IsInitialized = $RSConfig.IsInitialized
    }
    else
    {  
        throw New-TerminatingError -ErrorType SSRSNotFound -FormatArgs @($InstanceName) -ErrorCategory ObjectNotFound
    }

    $returnValue = @{
        InstanceName = $InstanceName
        RSSQLServer = $RSSQLServer
        RSSQLInstanceName = $RSSQLInstanceName
        IsInitialized = $IsInitialized
    }

    $returnValue
}


function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true)]
        [System.String]
        $InstanceName,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLServer,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLInstanceName,

        [parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $SQLAdminCredential
    )

    if(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName -ErrorAction SilentlyContinue)
    {
        Invoke-Command -ComputerName . -Credential $SQLAdminCredential -ScriptBlock {
            # this is a separate PS session, need to load Common Code again
            Import-Module $using:currentPath\..\..\xSQLServerHelper.psm1 -Verbose -ErrorAction Stop
            # smart import of the SQL module
            Import-SQLPSModule

            $InstanceName = $args[0]
            $RSSQLServer = $args[1]
            $RSSQLInstanceName = $args[2]
            $InstanceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName).$InstanceName
            $SQLVersion = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\Setup" -Name "Version").Version).Split(".")[0]
            if($InstanceName -eq "MSSQLSERVER")
            {
                $RSServiceName = "ReportServer"
                $RSVirtualDirectory = "ReportServer"
                $RMVirtualDirectory = "Reports"
                $RSDatabase = "ReportServer"
            }
            else
            {
                $RSServiceName = "ReportServer`$$InstanceName"
                $RSVirtualDirectory = "ReportServer_$InstanceName"
                $RMVirtualDirectory = "Reports_$InstanceName"
                $RSDatabase = "ReportServer`$$InstanceName"
            }
            if($RSSQLInstanceName -eq "MSSQLSERVER")
            {
                $RSConnection = "$RSSQLServer"
            }
            else
            {
                $RSConnection = "$RSSQLServer\$RSSQLInstanceName"
            }
            $Language = (Get-WMIObject -Class Win32_OperatingSystem -Namespace root/cimv2 -ErrorAction SilentlyContinue).OSLanguage
            $RSConfig = Get-WmiObject -Class MSReportServer_ConfigurationSetting -Namespace "root\Microsoft\SQLServer\ReportServer\RS_$InstanceName\v$SQLVersion\Admin"
            if($RSConfig.VirtualDirectoryReportServer -ne $RSVirtualDirectory)
            {
                $null = $RSConfig.SetVirtualDirectory("ReportServerWebService",$RSVirtualDirectory,$Language)
                $null = $RSConfig.ReserveURL("ReportServerWebService","http://+:80",$Language)
            }
            if($RSConfig.VirtualDirectoryReportManager -ne $RMVirtualDirectory)
            {
                $null = $RSConfig.SetVirtualDirectory("ReportManager",$RMVirtualDirectory,$Language)
                $null = $RSConfig.ReserveURL("ReportManager","http://+:80",$Language)
            }
            $RSCreateScript = $RSConfig.GenerateDatabaseCreationScript($RSDatabase,$Language,$false)

            # Determine RS service account
            $RSSvcAccountUsername = (Get-WmiObject -Class Win32_Service | Where-Object {$_.Name -eq $RSServiceName}).StartName
            $RSRightsScript = $RSConfig.GenerateDatabaseRightsScript($RSSvcAccountUsername,$RSDatabase,$false,$true)

            Invoke-Sqlcmd -ServerInstance $RSConnection -Query $RSCreateScript.Script
            Invoke-Sqlcmd -ServerInstance $RSConnection -Query $RSRightsScript.Script
            $RSConfig.SetDatabaseConnection($RSConnection,$RSDatabase,2,"","")
            $RSConfig.InitializeReportServer($RSConfig.InstallationID)

        } -ArgumentList @($InstanceName,$RSSQLServer,$RSSQLInstanceName)
    }

    if(!(Test-TargetResource @PSBoundParameters))
    {
        throw New-TerminatingError -ErrorType TestFailedAfterSet -ErrorCategory InvalidResult
    }
}


function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [parameter(Mandatory = $true)]
        [System.String]
        $InstanceName,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLServer,

        [parameter(Mandatory = $true)]
        [System.String]
        $RSSQLInstanceName,

        [parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $SQLAdminCredential
    )

    $result = (Get-TargetResource @PSBoundParameters).IsInitialized
    
    $result
}


Export-ModuleMember -Function *-TargetResource
