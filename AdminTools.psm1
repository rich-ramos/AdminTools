Function Get-MachineInfo {

<#
.SYNOPSIS
    Get-MachineInfo presents gathered Commom Information Model(CIM) information in a PSCustom Object.
.DESCRIPTION
    Get-MachineInfo retrieves Common Information Model(CIM) objects through a open CimSession and creates a PSCustom Object to reperesent the retrived data.

    NOTE: Currently, this function uses remote connections only through the WS-Man protocol. Thus, older operating systems that
    do not communicate through the WS-Man protocol and instead rely on the Distributed Component Object Model(DCOM) will not work.
    However, this feature will be added in a future update.
.PARAMETER ComputerName
    The name of the computer(s) you want to query.
.EXAMPLE
    PS C:\> Get-MachineInfo -ComputerName DC01.lab.com

    ComputerName          : DC01
    Manufacturer          : innotek GmbH
    OSVersion             : 10.0.17763
    Processor             : Intel(R) Core(TM) i7-8750H CPU @ 2.20GHz
    FreePercent           : 50
    RAM                   : 1
    Drive                 : C:
    OSName                : Microsoft Windows Server 2019 Standard
    FreeSpace             : 12
    BIOSVersion           : VirtualBox
    BIOSSerial            : 0
    Model                 : VirtualBox
    DiskSize              : 24
    OSArchitecture        : 64-bit
    Domain                : lab.com
    ProcessorAddressWidth : 64


    Retrieve CIM information from DC01 and repersent it in a PSCustom Object.
.EXAMPLE
    PS C:\> SERVER01, SERVER02, SERVER03 | Get-MachineInfo

    Pipes SERVER01, SERVER02, SERVER03 to Get-MachineInfo and creates a PSCustom Object for each computer.
.EXAMPLE
    PS C:\> ComputerNameList.csv | Get-MachineInfo

    Get-MachineInfo's parameter -ComputerName can bind pipe line information through its parameter name.
    Thus, if ComputerNameList.csv contains a column named "ComputerName" with computer names listed below
    the column, then the names of the computers listed will come through the pipe line and powershell will bind
    them to the -ComputerName parameter and output a PSCustom Object for each one.
.INPUTS
    System.String
    You can pipe an array or list of computer names to this function
.OUTPUTS
    System.Management.Automation.PSCustomObject
.NOTES
    Last Updated: October 08,2019
    Version     : 1.0
.LINK
    Get-CimInstance
    New-CimSessionOption
    about_Remote
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string[]]$ComputerName
    )

    Begin {
        Write-Verbose "Starting $($Myinvocation.MyCommand)"
        Write-Verbose "Bound Parameters"
        Write-Verbose ($PSBoundParameters | Out-String)
    } #End Begin 
    
    Process {
        $cimSessionList = New-Object -TypeName System.Collections.Generic.List[Microsoft.Management.Infrastructure.CimSession]
        foreach ($computer in $ComputerName) {
            Write-Verbose "Creating CimSession to $computer"
            $cimSess = New-CimSession -ComputerName $computer
            $cimSessionList.Add($cimSess)
        }

        foreach ($cimSession in $cimSessionList) {
            Write-Verbose "Opening CimSession to $cimSession.ComputerName"
            $cs_params = @{'Class'='Win32_ComputerSystem'
                           'CimSession'=$cimSession}
            $cs = Get-CimInstance @cs_params

            $os_params = @{'Class'='Win32_OperatingSystem'
                           'CimSession'=$cimSession}
            $os = Get-CimInstance @os_params

            $systemDrive = $os.SystemDrive
            $disk_params = @{'Class'='Win32_LogicalDisk'
                             'Filter'="DeviceID='$systemDrive'"
                             'CimSession'=$cimSession}
            $disk = Get-CimInstance @disk_params

            $bios_params = @{'Class'='Win32_BIOS'
                             'CimSession'=$cimSession}
            $bios = Get-CimInstance @bios_params

            $proc_params = @{'Class'='Win32_Processor'
                             'CimSession'=$cimSession}
            $proc = Get-CimInstance @proc_params

            $props = @{'ComputerName'=$cs.Name
                       'Manufacturer'=$cs.Manufacturer
                       'Model'=$cs.Model
                       'Domain'=$cs.Domain
                       'RAM'="{0:N2}" -f ($cs.TotalPhysicalMemory / 1GB) -as [int]
                       'Drive'=$disk.DeviceID
                       'DiskSize'="{0:N2}" -f ($disk.Size / 1GB) -as [int]
                       'FreeSpace'="{0:N2}" -f ($disk.FreeSpace / 1GB) -as [int]
                       'FreePercent'= "{0:N2}" -f ($disk.FreeSpace / $disk.Size * 100) -as [int]
                       'BIOSVersion'=$bios.SMBIOSBIOSVersion
                       'BIOSSerial'=$bios.SerialNumber
                       'OSArchitecture'=$os.OSArchitecture
                       'OSName'=$os.Caption
                       'OSVersion'=$os.Version
                       'Processor'=$proc.Name
                       'ProcessorAddressWidth'=$proc.AddressWidth}

            $obj = New-Object -TypeName PSObject -Property $props
            Write-Output $obj
        }
    } #End Process 
    
    End {  
        foreach ($cimSession in $cimSessionList) {
            Write-Verbose "Removing CimSession to $($cimSession.ComputerName)"
            Remove-CimSession -CimSession $cimSession
        }
        Write-Verbose "Ending $($Myinvocation.MyCommand)"
    } #End End
} #End Function 

Function Save-MachineInfoToSqlDatabase {
    
<#
.SYNOPSIS
    Save-MachineInfoToSqlDatabase saves data gathered from Get-MachineInfo into a Sql database.
.DESCRIPTION
    Save-MachineInfoToSqlDatabase takes in System.Objects piped in via the pipeline from Get-MachineInfo.
    First, it sees if the data already exist on the database. If so, it deletes the current data and makes
    the table available for the new record which is inserted into the table.
.PARAMETER InputObject
    System.Object generated from Get-MachineInfo
.PARAMETER SqlInstance
    The instance of where the Sql database is hosted.
.PARAMETER SqlDatabaseName
    Name of the database located on the sql instance.
.PARAMETER SqlTableName
    Name of the table where data from InputObject will be saved.
.EXAMPLE
    PS C:\> Get-MachineInfo -ComputerName SERVER01 | Save-MachineInfoToSqlDatabase -SqlInstance Sql01\Sql2 -DatabaseName AdminData -SqlTableName ComputerInfo

    Gathers machine info from SERVER01 in the form of a System.Object which is piped over to Save-MachineInfoToSqlDatabase.
    Data from the Get-MachineInfo function is saved to the Sql table "ComputerInfo" on the "AdminData" database hosted on the "Sql01\Sql2" instance.
    The -InputeObject parameter binds the object to the parameter by value. Thus, it is not needed to explicitly define the parameter.
.EXAMPLE
    PS C:\> $machineInfo = Get-MachineInfo -ComputerName SERVER01, SERVER02
    PS C:\> Save-MachineInfoToSqlDatabase -InputObject $machineInfo -SqlInstance Sql01\Sql2 -DatabaseName AdminData -SqlTableName ComputerInfo

    Stores data gathered from Get-MachineInfo into a variable called $machineInfo. This variable is then passed to the -InputObject parameter of
    Save-MachineInfoToSqlDatabase which iterates through each object in the variable and saves it to the AdminData database.
.INPUTS
    System.Object
.OUTPUTS
.NOTES
    Last Updated: October 08,2019
    Version     : 1.0
.LINK
    Get-MachineInfo
#>

    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        [System.Object[]]$InputObject,

        [Parameter()]
        [string]$SqlInstance,

        [Parameter()]
        [string]$SqlDatabaseName,

        [Parameter()]
        [string]$SqlTableName
    )

    Begin {
        Write-Verbose "Starting $($Myinvocation.MyCommand)"
        Write-Verbose "Bound Parameters"
        Write-Verbose ($PSBoundParameters | Out-String)

        $sqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connectionString = "Server=$SqlInstance; DataBase=$SqlDatabaseName; Trusted_Connection=True"
        $sqlConnection.ConnectionString = $connectionString
        Write-Verbose "Connection string: $connectionString"

        $sqlConnection.Open() | Out-Null
        Write-Verbose "Opening connection to Sql Instance $($sqlConnection.DataSource)"
    } #End Begin 
    
    Process {  
        $sqlCommand = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $sqlCommand.Connection = $sqlConnection

        foreach ($object in $InputObject) {
            $query = "DELETE FROM $SqlTableName WHERE ComputerName = '$($object.ComputerName)' AND Drive = '$($object.Drive)'"
            $sqlCommand.CommandText = $query
            try {
                $sqlCommand.ExecuteNonQuery() | Out-Null
            }
            catch {
                $msg = "Query Failed: $($_.Exception.Message)"
                Write-Output $msg
            } #End try\catch

            $query = "INSERT INTO $SqlTableName (ComputerName,
                                                Manufacturer,
                                                OSVersion,
                                                Processor,
                                                FreePercent,
                                                RAM,
                                                Drive,
                                                OSName,
                                                FreeSpace,
                                                BIOSVersion,
                                                BIOSSerial,
                                                Model,
                                                DiskSize,
                                                OSArchitecture,
                                                Domain,
                                                ProcessorAddressWidth)
                                                
                                                VALUES ('$($object.ComputerName)',
                                                        '$($object.Manufacturer)',
                                                        '$($object.OSVersion)',
                                                        '$($object.Processor)',
                                                        '$($object.FreePercent)',
                                                        '$($object.RAM)',
                                                        '$($object.Drive)',
                                                        '$($object.OSName)',
                                                        '$($object.FreeSpace)',
                                                        '$($object.BIOSVersion)',
                                                        '$($object.BIOSSerial)',
                                                        '$($object.Model)',
                                                        '$($object.DiskSize)',
                                                        '$($object.OSArchitecture)',
                                                        '$($object.Domain)',
                                                        '$($object.ProcessorAddressWidth)')"

            $sqlCommand.CommandText = $query
            try {
                $sqlCommand.ExecuteNonQuery() | Out-Null
            }
            catch {
                $msg = "Query Failed: $($_.Exception.Message)"
                Write-Output $msg
            } #End try\catch
        } #foreach ($object)
    } #End Process 
    
    End {
        Write-Verbose "Closing connection to $($sqlConnection.DataSource)"
        $sqlConnection.Close() | Out-Null
        Write-Verbose "Ending $($Myinvocation.MyCommand)"
    } #End End

} #End Function 

Function Get-SqlTableData {

<#
.SYNOPSIS
    Get-SqlTableData queries data from an Sql database and outputs it as a System.Data.DataRow object.
.DESCRIPTION
    Get-SqlTableData connects to an instance of Sql server and a Sql database. Then, a query is performed that
    selects all the data defined in the -SqlTable parameter. This data is outputed as a System.Data.DataRow foreach object.
    The user can use this object to filter through or pass as arguments to another function.
.PARAMETER SqlInstance
    Name of the Sql instance to connect to that host the database and table you want to query.
.PARAMETER SqlDatabase
    Name of the Sql database that is hosted by the instance defined in -SqlInstance.
.PARAMETER SqlTable
    Name of the Sql table that you want to query.
.EXAMPLE
    PS C:\> Get-SqlTableData -SqlInstance Sql01\Sql2 -SqlDatabase AdminData -SqlTable ComputerInfo

    Id                    : 1004
    ComputerName          : ADMINPC01
    Drive                 : C:        
    DiskSize              : 29
    FreeSpace             : 5
    FreePercent           : 16
    Manufacturer          : innotek GmbH
    Model                 : VirtualBox
    OSName                : Microsoft Windows 10 Pro
    OSVersion             : 10.0.18362
    OSArchitecture        : 64-bit
    Processor             : Intel(R) Core(TM) i7-8750H CPU @ 2.20GHz
    ProcessorAddressWidth : 64
    BIOSSerial            : 0
    BIOSVersion           : VirtualBox
    RAM                   : 2
    Domain                : lab.com

    Returns a System.Data.DataRow object that holds the data queried from the ComputerInfo table.
.INPUTS
.OUTPUTS
    System.Data.DataRow
.NOTES
    Last Updated: October 9,2019
    Version     : 1.0
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage="Enter a Sql instance to connect to.")]
        [string]$SqlInstance,

        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage="Enter a Sql database that is hosted on the instance defined in -SqlInstance.")]
        [string]$SqlDatabase,

        [Parameter(Mandatory=$true,
                   Position=2,
                   HelpMessage="Enter a Sql table to query.")]
        [string]$SqlTable
    )

    Begin {  
        Write-Verbose "Staring $($Myinvocation.MyCommand)"
        Write-Verbose "Bound Parameters"
        Write-Verbose ($PSBoundParameters | Out-String)

        $sqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connectionString = "Server=$SqlInstance; Database=$SqlDatabase; Trusted_Connection=True"
        $sqlConnection.ConnectionString = $connectionString
        Write-Verbose "Connection string: $connectionString"

        $sqlConnection.Open() | Out-Null
        Write-Verbose "Opening connection to $($sqlConnection.DataSource)"
    } #End Begin 
    
    Process {  
        $sqlCommand = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $sqlCommand.Connection = $sqlConnection

        $sqlQuery = "SELECT * FROM $SqlTable"
        Write-Verbose "Excecuting $sqlQuery"
        $sqlCommand.CommandText = $sqlQuery
        $reader = $sqlCommand.ExecuteReader()

        $dataTable = New-Object -TypeName System.Data.DataTable
        $dataTable.Load($reader)

        Write-Output $dataTable
    } #End Process 
    
    End {
        Write-Verbose "Closing connection to $($Myinvocation.MyCommand)"
        $sqlConnection.Close()
        Write-Verbose "Ending $($sqlConnection.DataSource)"
    } #End End
} #End Function

Function Test-PCConnection {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Enter a computer name to test the connection to.")]
        [string[]]$ComputerName
    )
    foreach ($computer in $ComputerName) {
        if (Test-Connection -ComputerName $computer -Quiet) {
            Write-Verbose "Pinging $computer successful"
            try {
                Get-CimInstance -ClassName Win32_BIOS -ComputerName $ComputerName | Out-Null
            }
            catch {
                $($_.Exception.Message)
            }
            Write-Output $computer
        }
        else {
            Write-Verbose "Ping to $computer failed"
        }
    }
} #End Function