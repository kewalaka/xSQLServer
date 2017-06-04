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

        [parameter()]
        [System.String]
        $ReportServerVirtualDir,

        [parameter()]
        [System.String]
        $ReportsVirtualDir,

        [parameter()]
        [System.String[]]
        $ReportServerReservedUrl,

        [parameter()]
        [System.String[]]
        $ReportsReservedUrl
    )

    if(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName -ErrorAction SilentlyContinue)
    {
        $InstanceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName).$InstanceName
        $SQLVersion = [int]((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\Setup" -Name "Version").Version).Split(".")[0]

        $RSConfig = Get-WmiObject -Class MSReportServer_ConfigurationSetting -Namespace "root\Microsoft\SQLServer\ReportServer\RS_$InstanceName\v$SQLVersion\Admin"
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

        if($IsInitialized)
        {
            # SSRS Web Portal application name changed in SQL Server 2016
            # https://docs.microsoft.com/en-us/sql/reporting-services/breaking-changes-in-sql-server-reporting-services-in-sql-server-2016
            $RMApplicationName = if ($SQLVersion -ge 13) { 'ReportServerWebApp' } else { 'ReportManager'}
            
            $ReportServerVirtualDir = $RSConfig.VirtualDirectoryReportServer
            $ReportsVirtualDir = $RSConfig.VirtualDirectoryReportManager

            $reservedUrls = $RSConfig.ListReservedUrls()

            $ReportServerReservedUrl = @()
            $ReportsReservedUrl = @()

            for($i = 0; $i -lt $reservedUrls.Application.Count; ++$i)
            {
                if($reservedUrls.Application[$i] -eq "ReportServerWebService") { $ReportServerReservedUrl += $reservedUrls.UrlString[$i] }
                if($reservedUrls.Application[$i] -eq "$RMApplicationName") { $ReportsReservedUrl += $reservedUrls.UrlString[$i] }
            }
        }
    }
    else
    {  
        throw New-TerminatingError -ErrorType SSRSNotFound -FormatArgs @($InstanceName) -ErrorCategory ObjectNotFound
    }

    $returnValue = @{
        InstanceName = $InstanceName
        RSSQLServer = $RSSQLServer
        RSSQLInstanceName = $RSSQLInstanceName
        ReportServerVirtualDir = $ReportServerVirtualDir
        ReportsVirtualDir = $ReportsVirtualDir
        ReportServerReservedUrl = $ReportServerReservedUrl
        ReportsReservedUrl = $ReportsReservedUrl
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

        [parameter()]
        [System.String]
        $ReportServerVirtualDir,

        [parameter()]
        [System.String]
        $ReportsVirtualDir,

        [parameter()]
        [System.String[]]
        $ReportServerReservedUrl,

        [parameter()]
        [System.String[]]
        $ReportsReservedUrl
    )

    if(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName -ErrorAction SilentlyContinue)
    {
        $InstanceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\RS" -Name $InstanceName).$InstanceName
        $SQLVersion = [int]((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\Setup" -Name "Version").Version).Split(".")[0]
        if($InstanceName -eq "MSSQLSERVER")
        {
            $RSServiceName = "ReportServer"
            if([string]::IsNullOrEmpty($ReportServerVirtualDir)) { $ReportServerVirtualDir = "ReportServer" }
            if([string]::IsNullOrEmpty($ReportsVirtualDir)) { $ReportsVirtualDir = "Reports" }
            $RSDatabase = "ReportServer"
        }
        else
        {
            $RSServiceName = "ReportServer`$$InstanceName"
            if([string]::IsNullOrEmpty($ReportServerVirtualDir)) { $ReportServerVirtualDir = "ReportServer_$InstanceName" }
            if([string]::IsNullOrEmpty($ReportsVirtualDir)) { $ReportsVirtualDir = "Reports_$InstanceName" }
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

        # SSRS Web Portal application name changed in SQL Server 2016
        # https://docs.microsoft.com/en-us/sql/reporting-services/breaking-changes-in-sql-server-reporting-services-in-sql-server-2016
        $RMApplicationName = if ($SQLVersion -ge 13) { 'ReportServerWebApp' } else { 'ReportManager'}

        if(!$RSConfig.IsInitialized)
        {
            New-VerboseMessage -Message "Initializing Reporting Services on $RSSQLServer\$RSSQLInstanceName."

            if($ReportServerReservedUrl -eq $null) { $ReportServerReservedUrl = @("http://+:80") }
            if($ReportsReservedUrl -eq $null) { $ReportsReservedUrl = @("http://+:80") }

            if($RSConfig.VirtualDirectoryReportServer -ne $ReportServerVirtualDir)
            {
                New-VerboseMessage -Message "Setting report server virtual directory on $RSSQLServer\$RSSQLInstanceName to $ReportServerVirtualDir."
                $null = $RSConfig.SetVirtualDirectory("ReportServerWebService",$ReportServerVirtualDir,$Language)
                $ReportServerReservedUrl | ForEach-Object {
                    New-VerboseMessage -Message "Adding report server URL reservation on $RSSQLServer\$RSSQLInstanceName`: $_."
                    $null = $RSConfig.ReserveURL("ReportServerWebService",$_,$Language)
                }
            }
            if($RSConfig.VirtualDirectoryReportManager -ne $ReportsVirtualDir)
            {
                New-VerboseMessage -Message "Setting reports virtual directory on $RSSQLServer\$RSSQLInstanceName to $ReportServerVirtualDir."
                $null = $RSConfig.SetVirtualDirectory($RMApplicationName,$ReportsVirtualDir,$Language)
                $ReportsReservedUrl | ForEach-Object {
                    New-VerboseMessage -Message "Adding reports URL reservation on $RSSQLServer\$RSSQLInstanceName`: $_."
                    $null = $RSConfig.ReserveURL($RMApplicationName,$_,$Language)
                }
            }
            $RSCreateScript = $RSConfig.GenerateDatabaseCreationScript($RSDatabase,$Language,$false)

            # Determine RS service account
            $RSSvcAccountUsername = (Get-WmiObject -Class Win32_Service | Where-Object {$_.Name -eq $RSServiceName}).StartName
            $RSRightsScript = $RSConfig.GenerateDatabaseRightsScript($RSSvcAccountUsername,$RSDatabase,$false,$true)

            # smart import of the SQL module
            Import-SQLPSModule
            Invoke-Sqlcmd -ServerInstance $RSConnection -Query $RSCreateScript.Script
            Invoke-Sqlcmd -ServerInstance $RSConnection -Query $RSRightsScript.Script

            $null = $RSConfig.SetDatabaseConnection($RSConnection,$RSDatabase,2,"","")
            $null = $RSConfig.InitializeReportServer($RSConfig.InstallationID)
        }
        else
        {
            $currentConfig = Get-TargetResource @PSBoundParameters

            if(![string]::IsNullOrEmpty($ReportServerVirtualDir) -and ($ReportServerVirtualDir -ne $currentConfig.ReportServerVirtualDir))
            {
                New-VerboseMessage -Message "Setting report server virtual directory on $RSSQLServer\$RSSQLInstanceName to $ReportServerVirtualDir."

                # to change a virtual directory, we first need to remove all URL reservations, 
                # change the virtual directory and re-add URL reservations
                $currentConfig.ReportServerReservedUrl | ForEach-Object { $null = $RSConfig.RemoveURL("ReportServerWebService",$_,$Language) }
                $RSConfig.SetVirtualDirectory("ReportServerWebService",$ReportServerVirtualDir,$Language)
                $currentConfig.ReportServerReservedUrl | ForEach-Object { $null = $RSConfig.ReserveURL("ReportServerWebService",$_,$Language) }
            }
            
            if(![string]::IsNullOrEmpty($ReportsVirtualDir) -and ($ReportsVirtualDir -ne $currentConfig.ReportsVirtualDir))
            { 
                New-VerboseMessage -Message "Setting reports virtual directory on $RSSQLServer\$RSSQLInstanceName to $ReportServerVirtualDir."

                # to change a virtual directory, we first need to remove all URL reservations, 
                # change the virtual directory and re-add URL reservations
                $currentConfig.ReportsReservedUrl | ForEach-Object { $null = $RSConfig.RemoveURL($RMApplicationName,$_,$Language) }
                $RSConfig.SetVirtualDirectory($RMApplicationName,$ReportsVirtualDir,$Language)
                $currentConfig.ReportsReservedUrl | ForEach-Object { $null = $RSConfig.ReserveURL($RMApplicationName,$_,$Language) }
            }

            if(($ReportServerReservedUrl -ne $null) -and ((Compare-Object -ReferenceObject $currentConfig.ReportServerReservedUrl -DifferenceObject $ReportServerReservedUrl) -ne $null))
            {
                $currentConfig.ReportServerReservedUrl | ForEach-Object {
                    $null = $RSConfig.RemoveURL("ReportServerWebService",$_,$Language)
                }

                $ReportServerReservedUrl | ForEach-Object {
                    New-VerboseMessage -Message "Adding report server URL reservation on $RSSQLServer\$RSSQLInstanceName`: $_."
                    $null = $RSConfig.ReserveURL("ReportServerWebService",$_,$Language)
                }
            }

            if(($ReportsReservedUrl -ne $null) -and ((Compare-Object -ReferenceObject $currentConfig.ReportsReservedUrl -DifferenceObject $ReportsReservedUrl) -ne $null))
            {
                $currentConfig.ReportsReservedUrl | ForEach-Object {
                    $null = $RSConfig.RemoveURL($RMApplicationName,$_,$Language)
                }

                $ReportsReservedUrl | ForEach-Object {
                    New-VerboseMessage -Message "Adding reports URL reservation on $RSSQLServer\$RSSQLInstanceName`: $_."
                    $null = $RSConfig.ReserveURL($RMApplicationName,$_,$Language)
                }
            }
        }
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

        [parameter()]
        [System.String]
        $ReportServerVirtualDir,

        [parameter()]
        [System.String]
        $ReportsVirtualDir,

        [parameter()]
        [System.String[]]
        $ReportServerReservedUrl,

        [parameter()]
        [System.String[]]
        $ReportsReservedUrl
    )

    $result = $true

    $currentConfig = Get-TargetResource @PSBoundParameters
    
    if(!$currentConfig.IsInitialized)
    { 
        New-VerboseMessage -Message "Reporting services $RSSQLServer\$RSSQLInstanceName are not initialized."
        $result = $false 
    }

    if(![string]::IsNullOrEmpty($ReportServerVirtualDir) -and ($ReportServerVirtualDir -ne $currentConfig.ReportServerVirtualDir)) 
    { 
        New-VerboseMessage -Message "Report server virtual directory on $RSSQLServer\$RSSQLInstanceName is $($currentConfig.ReportServerVirtualDir), should be $ReportServerVirtualDir."
        $result = $false 
    }

    if(![string]::IsNullOrEmpty($ReportsVirtualDir) -and ($ReportsVirtualDir -ne $currentConfig.ReportsVirtualDir))
    { 
        New-VerboseMessage -Message "Reports virtual directory on $RSSQLServer\$RSSQLInstanceName is $($currentConfig.ReportsVirtualDir), should be $ReportsVirtualDir."
        $result = $false 
    }

    if(($ReportServerReservedUrl -ne $null) -and ((Compare-Object -ReferenceObject $currentConfig.ReportServerReservedUrl -DifferenceObject $ReportServerReservedUrl) -ne $null)) 
    { 
        New-VerboseMessage -Message "Report server reserved URLs on $RSSQLServer\$RSSQLInstanceName are $($currentConfig.ReportServerReservedUrl -join ', '), should be $($ReportServerReservedUrl -join ', ')."
        $result = $false 
    }

    if(($ReportsReservedUrl -ne $null) -and ((Compare-Object -ReferenceObject $currentConfig.ReportsReservedUrl -DifferenceObject $ReportsReservedUrl) -ne $null))
    { 
        New-VerboseMessage -Message "Reports reserved URLs on $RSSQLServer\$RSSQLInstanceName are $($currentConfig.ReportsReservedUrl -join ', ')), should be $($ReportsReservedUrl -join ', ')."
        $result = $false 
    }

    $result
}


Export-ModuleMember -Function *-TargetResource
