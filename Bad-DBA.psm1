###Internal Helper Functions
function Invoke-SafeSqlCommand {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [hashtable]$Parameters = @{},
        
        [switch]$NonQuery
    )
    
    $Connection = $null
    Try {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()
        
        $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)
        
        foreach ($param in $Parameters.GetEnumerator()) {
            $Command.Parameters.AddWithValue($param.Key, $param.Value) | Out-Null
        }
        
        if ($NonQuery) {
            $result = $Command.ExecuteNonQuery()
        } else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Command)
            $Dataset = New-Object System.Data.DataSet
            $Adapter.Fill($DataSet) | Out-Null
            $result = $DataSet.Tables[0]
        }
        
        return $result
    }
    Catch {
        Write-Error "SQL command failed: $_"
        Throw $_
    }
    Finally {
        if ($Connection -and $Connection.State -eq 'Open') {
            $Connection.Close()
        }
        if ($Connection) {
            $Connection.Dispose()
        }
    }
}

function Write-SqlStatus {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Operation,
        
        [Parameter(Mandatory=$true)]
        [string]$Database,
        
        [ValidateSet('Start','Complete','Failed')]
        [string]$Status
    )
    
    $messages = @{
        'Start' = "Starting ${Operation} on [${Database}]"
        'Complete' = "${Operation} complete on [${Database}]"
        'Failed' = "${Operation} failed on [${Database}]"
    }
    
    switch ($Status) {
        'Start' { Write-Verbose $messages[$Status] }
        'Complete' { Write-Verbose $messages[$Status] }
        'Failed' { Write-Error $messages[$Status] }
    }
}

function Invoke-SimpleSqlQuery {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [switch]$UseMaster
    )
    
    $cs = if ($UseMaster) { $ConnectionString.Master } else { $ConnectionString }
    Invoke-Sql -ConnectionString $cs -Query $Query
}

function Test-SqlIdentifier {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Value,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet('Username','Role','Schema')]
        [string]$Type
    )
    
    $patterns = @{
        'Username' = '^[a-zA-Z0-9_\-\.@\\]+$'
        'Role' = '^[a-zA-Z0-9_]+$'
        'Schema' = '^[a-zA-Z0-9_]+$'
    }
    
    $messages = @{
        'Username' = "Only alphanumeric characters, underscores, hyphens, dots, @ and backslash are allowed"
        'Role' = "Only alphanumeric characters and underscores are allowed"
        'Schema' = "Only alphanumeric characters and underscores are allowed"
    }
    
    if ($Value -notmatch $patterns[$Type]) {
        throw "Invalid ${Type} format: ${Value}. $($messages[$Type])"
    }
}

function New-BackupScript {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Database,
        
        [Parameter(Mandatory=$true)]
        [string]$BackupPath,
        
        [ValidateSet('FULL','LOG')]
        [string]$BackupType = 'FULL',
        
        [switch]$CopyOnly
    )
    
    $backupCommand = if ($BackupType -eq 'FULL') { 'BACKUP DATABASE' } else { 'BACKUP LOG' }
    $backupName = "${Database}_$(Get-Date -Format yyyyMMdd)-${BackupType}"
    
    $script = "${backupCommand} [${Database}] TO DISK = N'${BackupPath}' WITH NOFORMAT, NOINIT, NAME = N'${backupName}', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
    
    if ($CopyOnly) {
        $script += ", COPY_ONLY"
    }
    
    return $script
}

###Exportable Functions
function New-ConnectionString() {
    [CmdletBinding(DefaultParameterSetName='WindowsAuth')]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$DbServer,

        [string]$DbInstance,

        [Parameter(Mandatory=$true)]
        [string]$DbName,

        [int]$Port,
        
        [Parameter(ParameterSetName='SqlAuth', Mandatory=$true)]
        [PSCredential]$Credential,
        
        # Deprecated parameters for backward compatibility
        [Parameter(ParameterSetName='LegacyAuth')]
        [string]$DbUser,
        
        [Parameter(ParameterSetName='LegacyAuth')]
        [string]$DbPassword
    )

    $ConnectionString = New-Object system.data.sqlclient.sqlconnectionstringbuilder
    #Enum Keys --> $ConnectionString.Keys

    if ([string]::IsNullOrWhiteSpace($DbInstance)) {
        $ConnectionString['Data Source'] = $DbServer
    } else {
        $ConnectionString['Data Source'] = ($DbServer + "\" + $DbInstance)
    }

    if ($Port) {
        $ConnectionString['Data Source'] += ",$Port"
    }

    $ConnectionString['Initial Catalog'] = $DBName

    if ($Credential) {
        # SQL Authentication with PSCredential (secure)
        $ConnectionString['User ID'] = $Credential.UserName
        $ConnectionString['Password'] = $Credential.GetNetworkCredential().Password
    }
    elseif ($DBUser -and $DBPassword) {
        # Legacy plain-text authentication (deprecated)
        Write-Warning "Plain-text password parameters are deprecated. Use -Credential parameter with PSCredential for secure authentication."
        $ConnectionString['User ID'] = $DBUser
        $ConnectionString['Password'] = $DBPassword
    }
    else {
        # Windows Authentication (default)
        $ConnectionString['Integrated Security'] = $true
    }

    $ConnectionString | Add-Member -MemberType NoteProperty -Name Instance -Value $DbInstance

    #Adding connectionstring for [master]
    $MasterCS = New-Object system.data.sqlclient.sqlconnectionstringbuilder $ConnectionString
    $MasterCS['Initial Catalog'] = 'master'
    $ConnectionString | Add-Member -MemberType NoteProperty -Name Master -Value $MasterCS

    return $ConnectionString
}

function Invoke-Sql() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
		[psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [string]$Query,

        [int]$Timeout = 0,
        
        [switch]$AsDataTable
      )

    Begin {
        $Connection = $null
        Try {
			$Connection = New-Object system.data.SqlClient.SQLConnection($ConnectionString)
			$Command = New-Object system.data.sqlclient.sqlcommand($Query,$Connection)
            $Command.CommandTimeout = $Timeout
			$Connection.Open()
		}

        Catch {
            Write-Error "Failed to open connection to $($ConnectionString.DataSource): $_"
            if ($Connection) {
                $Connection.Close()
                $Connection.Dispose()
            }
            Throw $_
        }
    }

    Process {
        Try {
			$Adapter = New-Object System.Data.sqlclient.sqlDataAdapter $Command
			$Dataset = New-Object System.Data.DataSet
			$Adapter.Fill($DataSet) | Out-Null
		}
		
        Catch {
            Write-Error "Query failed: $Query - $_"
            Throw $_
        }
    }

    End {
        Try {
            if ($Connection -and $Connection.State -eq 'Open') {
                $Connection.Close()
            }
        }
        Finally {
            if ($Adapter)    { $Adapter.Dispose() }
            if ($Command)    { $Command.Dispose() }
            if ($Connection) { $Connection.Dispose() }
        }

        # Return results
        if ($DataSet -and $DataSet.Tables.Count -gt 0 -and $DataSet.Tables[0].Rows.Count -gt 0) {
            if ($AsDataTable) {
                # Legacy behavior: return DataTable
                return $DataSet.Tables[0]
            }
            else {
                # Modern behavior: return PSCustomObject array for pipeline compatibility
                $DataSet.Tables[0] | ForEach-Object {
                    $row = $_
                    $obj = [PSCustomObject]@{}
                    foreach ($column in $DataSet.Tables[0].Columns) {
                        $obj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $row[$column.ColumnName] -Force
                    }
                    $obj
                }
            }
        }
        else {
            # No results
            return $null
        }
    }
}

function Invoke-AllSqlScripts() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [psobject[]]$SqlPath,

        [int]$Timeout = 0,

        [switch]$Force
    )

    Begin {
        $LastLocation = Get-Location
        $DbManifest = Get-DbManifest -ConnectionString $ConnectionString
    }

    Process {
        Try {
            $defaultProperties = @('Name','IsComplete','Failed','HasData')
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

            [psobject[]]$ObjColl = Get-ChildItem -Path $SqlPath -Filter *.sql -Recurse | % {
                $Obj = New-Object -TypeName PsObject
                $Obj | Add-Member -MemberType NoteProperty -Name File -Value $_
                $Obj | Add-Member -MemberType NoteProperty -Name DbManifest -Value $DbManifest
                $Obj | Add-Member -MemberType NoteProperty -Name ConnectionString -Value $ConnectionString
                $Obj | Add-Member -MemberType NoteProperty -Name Timeout -Value $Timeout
                $Obj | Add-Member -MemberType NoteProperty -Name Force -Value $Force
                $Obj | Add-Member -MemberType ScriptMethod -Name RunSql -Value { Invoke-SqlRunner }
                $Obj | Add-Member -MemberType NoteProperty -Name IsComplete -Value $false
                $Obj | Add-Member -MemberType ScriptProperty -Name HasData -Value { ($this.Output -ne $null) -or ($this.Error -ne $null) }
                $Obj | Add-Member -MemberType ScriptProperty -Name Md5Sum -Value { $this.File | Get-FileHash -Algorithm MD5 | Select-Object -ExpandProperty Hash }
                $Obj | Add-Member -MemberType ScriptProperty -Name RelativePath -Value { $this.File.Directory.Name }
                $Obj | Add-Member -MemberType ScriptProperty -Name NeedsToRun -Value { ($this.Md5Sum -notin $this.DbManifest) -or ($this.Force) }
                $Obj | Add-Member -MemberType NoteProperty -Name Failed -Value $false
                $Obj | Add-Member -MemberType NoteProperty -Name Output -Value $null
                $Obj | Add-Member -MemberType NoteProperty -Name Error -Value $null
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                $Obj | Add-Member -MemberType ScriptMethod -Name UpdateManifest -Value { Invoke-Sql -ConnectionString $this.ConnectionString -Query $($ExecutionContext.InvokeCommand.ExpandString($UpdateSqlManifest)) }

                $Obj.PSObject.TypeNames.Insert(0,'SQL.Scripts')
                $Obj | Add-Member MemberSet PSStandardMembers $PSStandardMembers

                $Obj
            }

            return $ObjColl
        }
            
        Catch {
            Throw $_
        }
    }

    End {
        Set-Location $LastLocation
    }
}

function Get-OpenTrans() {
    [CmdletBinding()]
    Param (
        [psobject]$ConnectionString
    )

    Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query "SELECT spid, status, hostname, program_name, loginame, open_tran, last_batch, cmd FROM sys.sysprocesses WHERE open_tran = 1"
}

function New-BackupFull() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [string]$BackupPath,

        [switch]$CopyOnly
    )

    Begin {
        Remove-Item -Path $BackupPath -ErrorAction SilentlyContinue -Force | Out-Null
	    $Script = New-BackupScript -Database $ConnectionString.InitialCatalog -BackupPath $BackupPath -BackupType 'FULL' -CopyOnly:$CopyOnly
    }

	Process {
        Try {
            Write-SqlStatus -Operation 'Full Backup' -Database $ConnectionString.InitialCatalog -Status 'Start'
            Invoke-Sql -ConnectionString $ConnectionString.Master -Query $Script
            Write-SqlStatus -Operation 'Full Backup' -Database $ConnectionString.InitialCatalog -Status 'Complete'
        }

        Catch {
            Write-SqlStatus -Operation 'Full Backup' -Database $ConnectionString.InitialCatalog -Status 'Failed'
            Throw $_
        }
    }
}

function New-BackupLog() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [string]$BackupPath,

        [switch]$CopyOnly
    )

    Begin {
        Remove-Item -Path $BackupPath -ErrorAction SilentlyContinue -Force | Out-Null
    	$Script = New-BackupScript -Database $ConnectionString.InitialCatalog -BackupPath $BackupPath -BackupType 'LOG' -CopyOnly:$CopyOnly
    }

    Process {
        Try {
            Write-SqlStatus -Operation 'Log Backup' -Database $ConnectionString.InitialCatalog -Status 'Start'
            Invoke-Sql -ConnectionString $ConnectionString.Master -Query $Script
            Write-SqlStatus -Operation 'Log Backup' -Database $ConnectionString.InitialCatalog -Status 'Complete'
        }

        Catch {
            Write-SqlStatus -Operation 'Log Backup' -Database $ConnectionString.InitialCatalog -Status 'Failed'
            Throw $_
        }
    }
}

function Restore-BackupFull() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$BackupPath,
        
        [Parameter(Mandatory=$true)]
        [psobject]$NameSpaces,

        [switch]$KillAll,

        [switch]$Replace,
        
        [switch]$WithRecovery
    )

	Begin {
        $RestoreTo = Get-SqlDefaultLocation $ConnectionString
        [string]$files = ""

        foreach ($NameSpace in $NameSpaces) {
		    Switch ($NameSpace.file_type) {
			    "mdf" { $files += ", MOVE N'$($NameSpace.name)' TO N'$($RestoreTo.DefaultFile)\$($ConnectionString.InitialCatalog)_$($NameSpace.name).mdf'" }
			    "ndf" { $files += ", MOVE N'$($NameSpace.name)' TO N'$($RestoreTo.DefaultFile)\$($ConnectionString.InitialCatalog)_$($NameSpace.name).ndf'" }
			    "ldf" { $files += ", MOVE N'$($NameSpace.name)' TO N'$($RestoreTo.DefaultLog)\$($ConnectionString.InitialCatalog)_$($NameSpace.name).ldf'" }
		    }
	    }

        $recoveryOption = if ($WithRecovery) { "RECOVERY" } else { "NORECOVERY" }
        [string]$Script = "RESTORE DATABASE [$($ConnectionString.InitialCatalog)] FROM DISK = N'${BackupPath}' WITH FILE = 1${files}, $recoveryOption, NOUNLOAD, STATS = 10"
    }

	Process {
        Try {
            if ($KillAll) {
                Close-AllSqlConnections -ConnectionString $ConnectionString
            }

            if ($Replace) {
                $Script += ", REPLACE"
            }

            Write-SqlStatus -Operation 'Full Restore' -Database $ConnectionString.InitialCatalog -Status 'Start'
            Invoke-Sql -ConnectionString $ConnectionString.Master -Query $Script
            Write-SqlStatus -Operation 'Full Restore' -Database $ConnectionString.InitialCatalog -Status 'Complete'
        }

        Catch {
            Write-SqlStatus -Operation 'Full Restore' -Database $ConnectionString.InitialCatalog -Status 'Failed'
            Throw $_
        }
    }
}

function Restore-BackupLog() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [string]$BackupPath,

        [switch]$KillAll
    )

	Begin {
        [string]$Script = "RESTORE LOG [$($ConnectionString.InitialCatalog)] FROM DISK = N'${BackupPath}' WITH FILE = 1, NORECOVERY, NOUNLOAD, STATS = 10"
    }

    Process {
        Try {
            if ($KillAll) {
                Close-AllSqlConnections -ConnectionString $ConnectionString
            }

            Write-SqlStatus -Operation 'Log Restore' -Database $ConnectionString.InitialCatalog -Status 'Start'
            Invoke-Sql -ConnectionString $ConnectionString.Master -Query $Script
            Write-SqlStatus -Operation 'Log Restore' -Database $ConnectionString.InitialCatalog -Status 'Complete'
        }

        Catch {
            Write-SqlStatus -Operation 'Log Restore' -Database $ConnectionString.InitialCatalog -Status 'Failed'
            Throw $_
        }
    }
}

function Get-SqlDefaultLocation() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

	Begin {
        [string]$Script = "declare @SmoDefaultFile nvarchar(512) exec master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'DefaultData', @SmoDefaultFile OUTPUT declare @SmoDefaultLog nvarchar(512) exec master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'DefaultLog', @SmoDefaultLog OUTPUT SELECT ISNULL(@SmoDefaultFile,N'') AS [DefaultFile], ISNULL(@SmoDefaultLog,N'') AS [DefaultLog]"
    }

    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

# Backward compatibility alias
Set-Alias -Name Get-DfltLoc -Value Get-SqlDefaultLocation

function Get-Namespaces() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

	Begin {
        [string]$Script = "SELECT LOWER(RIGHT(physical_name, 3)) AS file_type, name FROM SYS.DATABASE_FILES"
    }

    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script
    }
}

function Set-RecoveryModel() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [ValidateSet("SIMPLE","FULL","BULK_LOGGED")]
        [string]$Model
    )

	Begin {
        [string]$Script = "ALTER DATABASE [$($ConnectionString.InitialCatalog)] SET RECOVERY $model"
    }

    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function Get-Endpoint() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

	Begin {
        [string]$Script = "SELECT endpoint_id, name, type, type_desc, state, state_desc, protocol, protocol_desc FROM sys.database_mirroring_endpoints WITH (NOLOCK) WHERE name = 'HADR_ENDPOINT'"
    }

    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function New-Endpoint() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [int]$Port = 5022
    )

	Begin {
    	[string]$Script = "CREATE ENDPOINT [HADR_ENDPOINT] STATE = STARTED AS TCP (LISTENER_PORT = $Port, LISTENER_IP = ALL) FOR DATA_MIRRORING (ROLE = ALL, AUTHENTICATION = WINDOWS NEGOTIATE, ENCRYPTION = REQUIRED ALGORITHM AES)"
    }

    Process {
	    if (!(Get-Endpoint $ConnectionString)) { 
	        Write-Verbose "Creating HADR endpoint on port $Port"
	        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster 
	    }
	    else {
	        Write-Verbose "HADR endpoint already exists"
	    }
    }
}

function Get-DbSvcUser() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

	Begin {
    	[string]$Script = "DECLARE @SrvAccount varchar(100) set @SrvAccount ='' EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'SYSTEM\CurrentControlSet\Services\MSSQLSERVER', N'ObjectName', @SrvAccount OUTPUT, N'no_output' SELECT @SrvAccount as SQLAgent_ServiceAccount"
    }

    Process {
	    Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function Get-DbUser() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$UserName
    )

	Begin {
        Test-SqlIdentifier -Value $UserName -Type 'Username'
    	[string]$Script = "SELECT principal_id, name, type, type_desc, create_date, modify_date, is_disabled FROM sys.server_principals WHERE name = '${UserName}'"
    }

    Process {
	    Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function New-DbUser() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$UserName
    )

	Begin {
        Test-SqlIdentifier -Value $UserName -Type 'Username'
    }

    Process {
	    if (!(Get-DBUser -ConnectionString $ConnectionString -UserName $UserName)) {
            Try {
                # Escape brackets in identifier by doubling them
                $EscapedUserName = $UserName.Replace(']', ']]')
                $Query = "CREATE LOGIN [$EscapedUserName] FROM WINDOWS"
                
                Invoke-SafeSqlCommand -ConnectionString $ConnectionString.Master -Query $Query -NonQuery
                
                Write-Verbose "Created login: $UserName"
            }
            Catch {
                Write-Error "Failed to create login '$UserName': $_"
                Throw $_
            }
        }
    }
}

function Get-DbRoles() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

    Begin {
        [string]$Script = "SELECT [name] FROM sys.sysusers where issqlrole = 1"
    }

    Process {
        Invoke-Sql -ConnectionString $ConnectionString -Query $Script -ErrorAction:$ErrorActionPreference
    }
}

function Get-DbSchemas() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

    Begin {
        [string]$Script = "SELECT [name] FROM sys.schemas"
    }

    Process {
        Invoke-Sql -ConnectionString $ConnectionString -Query $Script -ErrorAction:$ErrorActionPreference
    }
}

function Add-DbUser() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$UserName,

        [ValidateScript({ $(Get-DbRoles -ConnectionString $ConnectionString).Name -contains $_ })]
        [string]$DbRole = "db_owner"
    )

    Begin {
        Test-SqlIdentifier -Value $UserName -Type 'Username'
        Test-SqlIdentifier -Value $DbRole -Type 'Role'
    }

    Process {
        Try {
            # Use parameterized sp_adduser call
            # Note: sp_adduser is deprecated but we're making it safe first, will modernize later
            $Query = "EXEC sp_adduser @loginname, @name_in_db, @grpname"
            $Params = @{
                '@loginname' = $UserName
                '@name_in_db' = $UserName
                '@grpname' = $DbRole
            }
            
            Invoke-SafeSqlCommand -ConnectionString $ConnectionString -Query $Query -Parameters $Params -NonQuery
            
            Write-Verbose "Added user '$UserName' to database with role '$DbRole'"
        }
        Catch {
            Write-Error "Failed to add user '$UserName' to database: $_"
            Throw $_
        }
    }
}

function Set-DbUserSchema() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$UserName,

        [ValidateScript({ $(Get-DbSchemas -ConnectionString $ConnectionString).Name -contains $_ })]
        [string]$DbSchema = "dbo"
    )

    Begin {
        Test-SqlIdentifier -Value $UserName -Type 'Username'
        Test-SqlIdentifier -Value $DbSchema -Type 'Schema'
    }

    Process {
        Try {
            # Escape brackets in identifiers by doubling them
            $EscapedUserName = $UserName.Replace(']', ']]')
            $EscapedDbSchema = $DbSchema.Replace(']', ']]')
            
            # Build safe query with escaped identifiers
            $Query = "ALTER USER [$EscapedUserName] WITH DEFAULT_SCHEMA = [$EscapedDbSchema]"
            
            Invoke-SafeSqlCommand -ConnectionString $ConnectionString -Query $Query -NonQuery
            
            Write-Verbose "Set default schema for user '$UserName' to '$DbSchema'"
        }
        Catch {
            Write-Error "Failed to set schema for user '$UserName': $_"
            Throw $_
        }
    }
}

function Set-EndpointACL() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory=$true)]
        [string]$UserName
    )

	Begin {
        Test-SqlIdentifier -Value $UserName -Type 'Username'
    	[string]$Script = "GRANT CONNECT ON ENDPOINT::HADR_ENDPOINT TO [${UserName}]"
    }
	
    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function Get-InstanceFQDN() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

	Begin {
    	[string]$Script = "DECLARE @Domain NVARCHAR(100) EXEC master.dbo.xp_regread 'HKEY_LOCAL_MACHINE', 'SYSTEM\CurrentControlSet\services\Tcpip\Parameters', N'Domain',@Domain OUTPUT SELECT Cast(SERVERPROPERTY('MachineName') as nvarchar) + '.' + @Domain AS FQDN"
    }

    Process {
	    Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function New-Mirror() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$SrcConnectionString,
        [Parameter(Mandatory=$true)]
        [psobject]$DestConnectionString,
        [int]$Port
    )

    Begin {
        $SrcFQDN = $(Get-InstanceFQDN -ConnectionString $SrcConnectionString).FQDN
        $DestFQDN = $(Get-InstanceFQDN -ConnectionString $DestConnectionString).FQDN
    }

    Process {
	    Try {
		    Write-Verbose "Creating database mirror between source and destination"
		    Invoke-Sql -ConnectionString $DestConnectionString.Master -Query "ALTER DATABASE [$($DestConnectionString.InitialCatalog)] SET PARTNER = 'TCP://${SrcFQDN}:${Port}'"
		    Invoke-Sql -ConnectionString $SrcConnectionString.Master -Query "ALTER DATABASE [$($SrcConnectionString.InitialCatalog)] SET PARTNER = 'TCP://${DestFQDN}:${Port}'"
		    return $true
	    }
	
        Catch {
		    return $false
	    }
    }
}

function Close-AllSqlConnections() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

    Begin {
        [string]$Script = @"
            Declare @spid int
            Select @spid = min(spid) from dbo.sysprocesses
            where dbid = db_id('$($ConnectionString.InitialCatalog)')
            While @spid Is Not Null
            Begin
                Execute ('Kill ' + @spid)
                Select @spid = min(spid) from dbo.sysprocesses
                where dbid = db_id('$($ConnectionString.InitialCatalog)') and spid > @spid
            End
"@
    }

    Process {
        Try {
            $SysDatabases = Get-SysDatabases -ConnectionString $ConnectionString | Select-Object -ExpandProperty Name

            if ($ConnectionString.InitialCatalog -in $SysDatabases) {
                Write-Verbose "Killing all SPIDs on [$($ConnectionString.InitialCatalog)]"
                Invoke-Sql -ConnectionString $ConnectionString.Master -Query $Script
                Write-Verbose "All SPIDs terminated on [$($ConnectionString.InitialCatalog)]"
            } else {
                Write-Warning "Database does not exist: [$($ConnectionString.InitialCatalog)]"
            }
        }

        Catch {
            Write-Error "Unable to kill SPIDs on [$($ConnectionString.InitialCatalog)]: $_"
            Throw $_
        }
    }
}

function Get-SysDatabases() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

    Begin {
        [string]$Script = "SELECT name, database_id, state_desc, recovery_model_desc, compatibility_level, create_date, is_read_only, collation_name FROM sys.databases"
    }

    Process {
        Invoke-SimpleSqlQuery -ConnectionString $ConnectionString -Query $Script -UseMaster
    }
}

function Invoke-DbShrink() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [ValidateSet('mdf','ndf','ldf')]
        [string[]]$Type
    )

    Begin {
        if ($Type) {
            $NameSpaces = Get-Namespaces -ConnectionString $ConnectionString | ? { $_.file_type -in $Type }
        } else {
            $NameSpaces = Get-Namespaces -ConnectionString $ConnectionString
        }
    }

    Process {
        foreach ($n in $NameSpaces) {
            Write-Verbose "Shrinking database file: [$($n.name)]"
            Invoke-Sql -ConnectionString $ConnectionString -Query "DBCC SHRINKFILE (N'$($n.name)' , 0, TRUNCATEONLY)"
        }
    }
}

function Get-SnapshotState() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Backup','Restore')]
        [string]$Phase
    )
    
    Begin {
        [string]$Script = "select session_id, blocking_session_id, db_name(database_id) as [Database], command, percent_complete, wait_type,wait_time, wait_resource, scheduler_id, Qry.text from sys.dm_exec_requests req cross apply sys.fn_get_sql(req.sql_handle) as Qry where req.session_id>50 and command = '$($Phase.ToUpper()) DATABASE'"
    }
    
    Process {
        Invoke-Sql -ConnectionString $ConnectionString -Query $Script
    }
}

###Helper Functions (Not Exported)
function Invoke-SqlRunner() {
    Try {
        Write-Verbose "Executing SQL script: $($this.File.FullName)"

        if ($this.NeedsToRun) {
            $sqlContent = Get-Content -Path $this.File.FullName -Raw -Encoding UTF8
            $batches = $sqlContent -split '(?im)^\s*GO\s*$' | Where-Object { $_.Trim() -ne '' }

            foreach ($batch in $batches) {
                Invoke-SafeSqlCommand -ConnectionString $this.ConnectionString -Query $batch -NonQuery | Out-Null
            }

            $State = "Done"
            Write-Verbose "SQL script completed successfully: $($this.File.Name)"
        } else {
            $State = "Skipped"
            Write-Verbose "SQL script skipped (already executed): $($this.File.Name)"
        }

        $this.IsComplete = $true
        $this.Failed     = $false
        $this.Error      = $null
    }
    Catch {
        $this.IsComplete = $false
        $this.Failed     = $true
        $this.Error      = $_.Exception.Message

        $State = "Error"
        Write-Error "SQL script execution failed: $($this.File.Name) - $_"
    }
    Finally {
        $this.UpdateManifest()
        Write-Information -MessageData "$($this.File.Name): $State" -InformationAction Continue
    }
}

function Get-DbManifest() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [psobject]$ConnectionString
    )

    Try {
        $result = Invoke-Sql -ConnectionString $ConnectionString -Query $GetSqlManifest
        if ($result) {
            return $result | Select-Object -ExpandProperty Md5Sum
        }
        return @()
    }

    Catch {
        Write-Verbose "Failed to retrieve manifest: $_"
        # Return empty array if manifest doesn't exist yet
        return @()
    }
}

function Export-DbSchema() {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [psobject]$ConnectionString,
        
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string]$ExportPath
    )

    Write-Verbose "Exporting database schema from [$($ConnectionString.InitialCatalog)]"
    
    # Export Tables
    Write-Verbose "Generating table DDL scripts..."
    $tableScript = @"
SELECT 
    '-- Table: [' + SCHEMA_NAME(t.schema_id) + '].[' + t.name + ']' + CHAR(13) + CHAR(10) +
    'IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[' + SCHEMA_NAME(t.schema_id) + '].[' + t.name + ']'') AND type = ''U'')' + CHAR(13) + CHAR(10) +
    'BEGIN' + CHAR(13) + CHAR(10) +
    'CREATE TABLE [' + SCHEMA_NAME(t.schema_id) + '].[' + t.name + '] (' + CHAR(13) + CHAR(10) +
    STUFF((
        SELECT CHAR(9) + ',[' + c.name + '] ' + 
               TYPE_NAME(c.user_type_id) +
               CASE 
                   WHEN TYPE_NAME(c.user_type_id) IN ('varchar','char','nvarchar','nchar') 
                   THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(CASE WHEN TYPE_NAME(c.user_type_id) LIKE 'n%' THEN c.max_length/2 ELSE c.max_length END AS VARCHAR) END + ')'
                   WHEN TYPE_NAME(c.user_type_id) IN ('decimal','numeric')
                   THEN '(' + CAST(c.precision AS VARCHAR) + ',' + CAST(c.scale AS VARCHAR) + ')'
                   WHEN TYPE_NAME(c.user_type_id) IN ('datetime2','time','datetimeoffset')
                   THEN '(' + CAST(c.scale AS VARCHAR) + ')'
                   ELSE ''
               END +
               CASE WHEN c.is_identity = 1 THEN ' IDENTITY(' + CAST(IDENT_SEED(SCHEMA_NAME(t.schema_id) + '.' + t.name) AS VARCHAR) + ',' + CAST(IDENT_INCR(SCHEMA_NAME(t.schema_id) + '.' + t.name) AS VARCHAR) + ')' ELSE '' END +
               CASE WHEN c.is_nullable = 0 THEN ' NOT NULL' ELSE ' NULL' END +
               CHAR(13) + CHAR(10)
        FROM sys.columns c
        WHERE c.object_id = t.object_id
        ORDER BY c.column_id
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') +
    ');' + CHAR(13) + CHAR(10) +
    'END' + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.tables t
WHERE t.is_ms_shipped = 0
ORDER BY SCHEMA_NAME(t.schema_id), t.name
"@
    
    $tableResults = Invoke-Sql -ConnectionString $ConnectionString -Query $tableScript -AsDataTable
    if ($tableResults) {
        $tableResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@Tables_Script.sql" -Encoding UTF8
        Write-Verbose "Table scripts exported to @Tables_Script.sql"
    }
    
    # Export Views
    Write-Verbose "Generating view DDL scripts..."
    $viewScript = @"
SELECT 
    '-- View: [' + SCHEMA_NAME(v.schema_id) + '].[' + v.name + ']' + CHAR(13) + CHAR(10) +
    'IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N''[' + SCHEMA_NAME(v.schema_id) + '].[' + v.name + ']''))' + CHAR(13) + CHAR(10) +
    'BEGIN' + CHAR(13) + CHAR(10) +
    'EXEC(''' + REPLACE(m.definition, '''', '''''') + ''')' + CHAR(13) + CHAR(10) +
    'END' + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.views v
INNER JOIN sys.sql_modules m ON v.object_id = m.object_id
WHERE v.is_ms_shipped = 0
ORDER BY SCHEMA_NAME(v.schema_id), v.name
"@
    
    $viewResults = Invoke-Sql -ConnectionString $ConnectionString -Query $viewScript -AsDataTable
    if ($viewResults) {
        $viewResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@Views_Script.sql" -Encoding UTF8
        Write-Verbose "View scripts exported to @Views_Script.sql"
    }
    
    # Export Stored Procedures
    Write-Verbose "Generating stored procedure DDL scripts..."
    $procScript = @"
SELECT 
    '-- Stored Procedure: [' + SCHEMA_NAME(p.schema_id) + '].[' + p.name + ']' + CHAR(13) + CHAR(10) +
    m.definition + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.procedures p
INNER JOIN sys.sql_modules m ON p.object_id = m.object_id
WHERE p.is_ms_shipped = 0
ORDER BY SCHEMA_NAME(p.schema_id), p.name
"@
    
    $procResults = Invoke-Sql -ConnectionString $ConnectionString -Query $procScript -AsDataTable
    if ($procResults) {
        $procResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@Procs_Script.sql" -Encoding UTF8
        Write-Verbose "Stored procedure scripts exported to @Procs_Script.sql"
    }
    
    # Export Functions
    Write-Verbose "Generating function DDL scripts..."
    $funcScript = @"
SELECT 
    '-- Function: [' + SCHEMA_NAME(o.schema_id) + '].[' + o.name + ']' + CHAR(13) + CHAR(10) +
    m.definition + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.objects o
INNER JOIN sys.sql_modules m ON o.object_id = m.object_id
WHERE o.type IN ('FN', 'IF', 'TF')  -- Scalar, Inline Table-Valued, Table-Valued
  AND o.is_ms_shipped = 0
ORDER BY SCHEMA_NAME(o.schema_id), o.name
"@
    
    $funcResults = Invoke-Sql -ConnectionString $ConnectionString -Query $funcScript -AsDataTable
    if ($funcResults) {
        $funcResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@Functions_Script.sql" -Encoding UTF8
        Write-Verbose "Function scripts exported to @Functions_Script.sql"
    }
    
    # Export Triggers
    Write-Verbose "Generating trigger DDL scripts..."
    $triggerScript = @"
SELECT 
    '-- Trigger: [' + SCHEMA_NAME(t.schema_id) + '].[' + tr.name + ']' + CHAR(13) + CHAR(10) +
    m.definition + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.triggers tr
INNER JOIN sys.sql_modules m ON tr.object_id = m.object_id
INNER JOIN sys.tables t ON tr.parent_id = t.object_id
WHERE tr.is_ms_shipped = 0
  AND tr.parent_class = 1  -- Object or column triggers
ORDER BY SCHEMA_NAME(t.schema_id), t.name, tr.name
"@
    
    $triggerResults = Invoke-Sql -ConnectionString $ConnectionString -Query $triggerScript -AsDataTable
    if ($triggerResults) {
        $triggerResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@Triggers_Script.sql" -Encoding UTF8
        Write-Verbose "Trigger scripts exported to @Triggers_Script.sql"
    }
    
    # Export Database Triggers
    Write-Verbose "Generating database trigger DDL scripts..."
    $dbTriggerScript = @"
SELECT 
    '-- Database Trigger: [' + tr.name + ']' + CHAR(13) + CHAR(10) +
    m.definition + CHAR(13) + CHAR(10) +
    'GO' + CHAR(13) + CHAR(10) AS DDL
FROM sys.triggers tr
INNER JOIN sys.sql_modules m ON tr.object_id = m.object_id
WHERE tr.is_ms_shipped = 0
  AND tr.parent_class = 0  -- Database triggers
ORDER BY tr.name
"@
    
    $dbTriggerResults = Invoke-Sql -ConnectionString $ConnectionString -Query $dbTriggerScript -AsDataTable
    if ($dbTriggerResults) {
        $dbTriggerResults | Select-Object -ExpandProperty DDL | Out-File "$ExportPath\@DBTriggers_Script.sql" -Encoding UTF8
        Write-Verbose "Database trigger scripts exported to @DBTriggers_Script.sql"
    }
    
    Write-Verbose "Schema export complete for [$($ConnectionString.InitialCatalog)]"
}

###SQL Scripts
$UpdateSqlManifest = @'
IF NOT EXISTS (SELECT * FROM sysobjects WHERE Name = N'SqlDeployManifest' AND XType = N'U')
BEGIN
CREATE TABLE SqlDeployManifest (
    ID int IDENTITY(1,1) PRIMARY KEY,
    ScriptName VARCHAR(64) NOT NULL,
    Succeeded BIT NOT NULL,
    LastRun DATETIME2 NULL DEFAULT GETDATE(),
    MD5Sum VARCHAR(64) NOT NULL,
    ScriptPath VARCHAR(255) NOT NULL,
	RunCount RowVersion
)
END

IF OBJECT_ID(N'[dbo].[trg_SqlDeployManifest_update]') IS NULL 
BEGIN
	EXEC('
		CREATE TRIGGER [dbo].[trg_SqlDeployManifest_update] ON [dbo].[SqlDeployManifest]
		AFTER INSERT, UPDATE
		AS
			UPDATE s set LastRun = GETDATE()
			FROM [dbo].[SqlDeployManifest] AS s
			INNER JOIN inserted AS i ON s.ID = i.ID
	')
END

IF NOT EXISTS (SELECT * FROM [dbo].[SqlDeployManifest] WHERE ScriptName = N'$($this.Name)' and ScriptPath = N'$($this.RelativePath)')
BEGIN
	INSERT INTO [dbo].[SqlDeployManifest] (ScriptName, Succeeded, MD5Sum, ScriptPath)
	VALUES (N'$($this.Name)', $([int]!$this.Failed), N'$($this.Md5Sum)', N'$($this.RelativePath)')
END
ELSE
BEGIN
	UPDATE [dbo].[SqlDeployManifest]
	SET Succeeded = $([int]!$this.Failed),
	MD5Sum = N'$($this.Md5Sum)'
	WHERE ScriptName = N'$($this.Name)'
	AND ScriptPath = N'$($this.RelativePath)'
END
'@

$GetSqlManifest = @'
SELECT Md5Sum
FROM [dbo].[SqlDeployManifest]
WHERE Succeeded = 1
'@
