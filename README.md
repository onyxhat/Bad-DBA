# Bad DBA PowerShell Module

Quick and dirty database administration PowerShell module with ZERO external dependencies. Including backup/restore operations, database mirroring, and automated deployment workflows.

---

## Table of Contents

- [Installation](#installation)
- [Getting Started](#getting-started)
- [Core Concepts](#core-concepts)
- [Function Reference](#function-reference)
  - [Connection Management](#connection-management)
  - [Query Execution](#query-execution)
  - [Backup & Restore](#backup--restore)
  - [Database Administration](#database-administration)
  - [Security & User Management](#security--user-management)
  - [High Availability](#high-availability)
  - [Schema Management](#schema-management)
- [Common Workflows](#common-workflows)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)

---

## Installation

### Manual Installation

```powershell
# Copy the module to your PowerShell modules directory
Copy-Item -Path ".\Bad-DBA" -Destination "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\" -Recurse

# Import the module
Import-Module Bad-DBA
```

### Import in Scripts

```powershell
Import-Module Bad-DBA -Force -DisableNameChecking
```

---

## Getting Started

### Basic Connection

```powershell
# Windows Authentication (Recommended)
$cs = New-ConnectionString -DbServer "SQLServer01" -DbName "MyDatabase"

# SQL Authentication with Secure Credentials
$cred = Get-Credential
$cs = New-ConnectionString -DbServer "SQLServer01" -DbName "MyDatabase" -Credential $cred

# Named Instance
$cs = New-ConnectionString -DbServer "SQLServer01" -DbInstance "PROD" -DbName "MyDatabase"

# Custom Port
$cs = New-ConnectionString -DbServer "SQLServer01" -DbName "MyDatabase" -Port 1433
```

### Simple Query

```powershell
$cs = New-ConnectionString -DbServer "SQLServer01" -DbName "MyDatabase"
$results = Invoke-Sql -ConnectionString $cs -Query "SELECT name, database_id FROM sys.databases"

# Results are returned as PSCustomObject array for easy filtering
$results | Where-Object { $_.database_id -gt 4 } | Format-Table
```

---

## Core Concepts

### Connection String Objects

The `New-ConnectionString` function returns a `SqlConnectionStringBuilder` object with added properties:

- **Standard Properties**: Data Source, Initial Catalog, User ID, Password, Integrated Security
- **`.Instance`**: The SQL Server instance name (if specified)
- **`.Master`**: A separate connection string pointing to the `master` database

```powershell
$cs = New-ConnectionString -DbServer "SQLServer01" -DbName "MyDatabase"

# Access properties
Write-Host $cs.DataSource          # SQLServer01
Write-Host $cs.InitialCatalog      # MyDatabase
Write-Host $cs.Master.InitialCatalog  # master
```

### Script Manifest Tracking

The module uses a `SqlDeployManifest` table to track executed SQL scripts via MD5 checksums. This enables **idempotent deployments** - scripts are only executed once unless changed or forced.

**Required Table Schema:**

```sql
CREATE TABLE SqlDeployManifest (
    ID int IDENTITY(1,1) PRIMARY KEY,
    ScriptName VARCHAR(64) NOT NULL,
    Succeeded BIT NOT NULL,
    LastRun DATETIME2 NULL DEFAULT GETDATE(),
    MD5Sum VARCHAR(64) NOT NULL,
    ScriptPath VARCHAR(255) NOT NULL,
    RunCount RowVersion
)
```

The table and an associated trigger are automatically created by `Invoke-AllSqlScripts` if they do not exist.

---

## Function Reference

### Connection Management

#### New-ConnectionString

Creates a SQL connection string with support for named instances, custom ports, and multiple authentication methods.

**Parameters:**
- `DbServer` (Required): SQL Server hostname
- `DbName` (Required): Database name
- `DbInstance` (Optional): Named instance
- `Port` (Optional): Custom port number
- `Credential` (Optional): `PSCredential` object for secure SQL authentication
- `DbUser` (Deprecated): Legacy SQL authentication username
- `DbPassword` (Deprecated): Legacy SQL authentication password

**Returns:** `SqlConnectionStringBuilder` with `.Instance` and `.Master` properties

**Example:**
```powershell
# Windows Auth
$cs = New-ConnectionString -DbServer "SQL01" -DbName "Prod"

# Secure SQL Auth
$cs = New-ConnectionString -DbServer "SQL01" -DbInstance "PROD" `
                           -DbName "Sales" -Credential (Get-Credential)
```

---

### Query Execution

#### Invoke-Sql

Executes a SQL query. By default, returns an array of `PSCustomObject` for better pipeline integration.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `Query` (Required): SQL query to execute
- `Timeout` (Optional): Command timeout in seconds (0 = infinite)
- `AsDataTable` (Optional): Return a `System.Data.DataTable` instead (legacy behavior)

**Returns:** `PSCustomObject[]` or `DataTable`

**Example:**
```powershell
$cs = New-ConnectionString -DbServer "SQL01" -DbName "Prod"
$data = Invoke-Sql -ConnectionString $cs -Query "SELECT TOP 10 name FROM sys.objects"
$data | ForEach-Object { Write-Host "Found object: $($_.name)" }
```

#### Invoke-AllSqlScripts

Executes all `.sql` files in a directory tree with MD5-based change tracking.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `SqlPath` (Required): Path to directory containing .sql files
- `Timeout` (Optional): Command timeout in seconds (0 = infinite)
- `Force` (Optional): Re-run all scripts even if unchanged

**Returns:** Array of `SQL.Scripts` objects with properties:
- `Name`: Script filename
- `IsComplete`: Execution finished successfully
- `Failed`: Execution failed
- `HasData`: Output or error data present
- `NeedsToRun`: Script needs execution (changed or forced)
- `Md5Sum`: MD5 hash of the script file
- `RelativePath`: Directory containing the script

**Example:**
```powershell
$cs = New-ConnectionString -DbServer "SQL01" -DbName "Prod"

# Run all changed scripts
$results = Invoke-AllSqlScripts -ConnectionString $cs -SqlPath ".\migrations"

# Show execution summary
$results | Where-Object { $_.NeedsToRun } | Format-Table Name, IsComplete, Failed
```

---

### Backup & Restore

#### New-BackupFull

Creates a full database backup.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `BackupPath` (Required): Full path to backup file (.bak)
- `CopyOnly` (Optional): Create copy-only backup (doesn't break log chain)

#### New-BackupLog

Creates a transaction log backup.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `BackupPath` (Required): Full path to backup file (.trn)
- `CopyOnly` (Optional): Create copy-only backup

#### Restore-BackupFull

Restores a full database backup with automatic file relocation.

**Parameters:**
- `ConnectionString` (Required): Target connection string
- `BackupPath` (Required): Path to backup file
- `NameSpaces` (Required): File mapping (from `Get-Namespaces`)
- `KillAll` (Optional): Terminate active connections before restore
- `Replace` (Optional): Overwrite existing database
- `WithRecovery` (Optional): Restore WITH RECOVERY (default: NORECOVERY)

**Example:**
```powershell
$src = New-ConnectionString -DbServer "SQL01" -DbName "ProdSource"
$dst = New-ConnectionString -DbServer "SQL02" -DbName "ProdCopy"

# Get source file layout
$namespaces = Get-Namespaces -ConnectionString $src

# Restore and bring online
Restore-BackupFull -ConnectionString $dst -BackupPath "C:\Backups\Prod.bak" `
                   -NameSpaces $namespaces -KillAll -Replace -WithRecovery
```

#### Restore-BackupLog

Restores a transaction log backup in NORECOVERY mode.

**Parameters:**
- `ConnectionString` (Required): Target connection string
- `BackupPath` (Required): Path to log backup file
- `KillAll` (Optional): Terminate active connections before restore

---

### Database Administration

#### Get-Namespaces

Retrieves database file logical names and physical paths.

**Returns:** Array of objects with `name` and `file_type` (mdf, ndf, ldf)

#### Get-SqlDefaultLocation

Gets default SQL Server data and log file locations. (Alias: `Get-DfltLoc`)

**Returns:** Object with `DefaultFile` and `DefaultLog` directory paths.

#### Set-RecoveryModel

Changes database recovery model.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `Model` (Required): "SIMPLE", "FULL", or "BULK_LOGGED"

#### Get-OpenTrans

Lists sessions with open transactions.

**Returns:** Detailed information about active transactions from `sys.sysprocesses`.

#### Close-AllSqlConnections

Kills all active connections to a database. **⚠️ Disruptive operation!**

#### Invoke-DbShrink

Shrinks database files to reclaim unused space.

**Parameters:**
- `Type` (Optional): Target specific file types ("mdf", "ndf", "ldf")

#### Get-SysDatabases

Lists all databases on the SQL Server instance.

#### Get-SnapshotState

Tracks progress of active backup or restore operations.

**Parameters:**
- `Phase` (Required): "Backup" or "Restore"

**Example:**
```powershell
Get-SnapshotState -ConnectionString $cs -Phase "Restore" | Format-Table command, percent_complete, wait_type
```

---

### Security & User Management

#### Get-DbSvcUser

Retrieves the SQL Server service account name.

#### Get-DbUser

Checks for the existence of a SQL Server login.

**Parameters:**
- `UserName` (Required): Login name to check

#### New-DbUser

Creates a new SQL Server login from a Windows account.

**Parameters:**
- `UserName` (Required): Windows account name (e.g., "DOMAIN\User")

#### Add-DbUser

Adds a SQL Server login to a database.

**Parameters:**
- `UserName` (Required): Login name
- `DbRole` (Optional): Database role to assign (default: "db_owner")

#### Set-DbUserSchema

Changes a user's default schema.

**Parameters:**
- `UserName` (Required): Database user name
- `DbSchema` (Required): New default schema (e.g., "dbo")

#### Get-DbRoles

Lists all database roles in the current database.

#### Get-DbSchemas

Lists all schemas in the current database.

---

### High Availability

#### Get-Endpoint

Lists database mirroring endpoints (specifically searches for 'HADR_ENDPOINT').

#### New-Endpoint

Creates a database mirroring endpoint.

**Parameters:**
- `Port` (Optional): TCP port number (default: 5022)

#### Set-EndpointACL

Grants CONNECT permission on the mirroring endpoint to a specific user.

#### Get-InstanceFQDN

Retrieves the fully qualified domain name of the SQL Server instance.

#### New-Mirror

Establishes database mirroring partnership between two servers.

**Returns:** Boolean (true = success)

---

### Schema Management

#### Get-DbManifest

Retrieves the MD5 checksums of all successfully executed scripts.

#### Export-DbSchema

Exports database schema to SQL script files. This implementation uses native T-SQL and does **not** require SMO assemblies.

**Parameters:**
- `ConnectionString` (Required): Connection string object
- `ExportPath` (Required): Directory path for output files

**Output Files:**
- `@Tables_Script.sql`
- `@Views_Script.sql`
- `@Procs_Script.sql`
- `@Functions_Script.sql`
- `@Triggers_Script.sql`
- `@DBTriggers_Script.sql`

---

## Common Workflows

### Database Snapshot Workflow

```powershell
Import-Module Bad-DBA

# Define source and destination
$src = New-ConnectionString -DbServer "SQL01" -DbName "Production"
$dst = New-ConnectionString -DbServer "SQL02" -DbName "Production_Snapshot"

# Get file layout from source
$namespaces = Get-Namespaces -ConnectionString $src

# Create backup
$backupPath = "\\FileServer\Backups\Production_Snapshot.bak"
New-BackupFull -ConnectionString $src -BackupPath $backupPath -CopyOnly

# Restore to destination and bring online
Restore-BackupFull -ConnectionString $dst -BackupPath $backupPath `
                   -NameSpaces $namespaces -KillAll -Replace -WithRecovery

Write-Host "Snapshot complete."
```

### Automated Migration Deployment

```powershell
Import-Module Bad-DBA

$cs = New-ConnectionString -DbServer "SQL01" -DbName "Prod"

# Build SQL deployment object(s)
$run_all = New-SqlDeployment -ConnectionString $cs -SqlPath ".\always_run\" -Force
$scripts = New-SqlDeployment -ConnectionString $cs -SqlPath ".\migration_scripts\"

# Execute deployment(s)
$run_all.Execute()
$scripts.Execute()

# Check results
$results = $run_all + $scripts

if ($results.Failed -contains $true) {
    Write-Error "Deployment failed."
    $results | ? { $_.Failed } | Format-Table Name, Error
} else {
    Write-Host "Deployment successful."
}
```

---

## Troubleshooting

### Connection Issues

**Problem:** "Login failed for user"

**Solution:**
Use the `-Credential` parameter for secure SQL authentication, or ensure your Windows account has appropriate permissions for Windows Authentication.

```powershell
$cs = New-ConnectionString -DbServer "SQL01" -DbName "Prod" -Credential (Get-Credential)
Invoke-Sql -ConnectionString $cs -Query "SELECT 1"
```

### Script Execution Not Idempotent

**Problem:** Scripts execute every time despite no changes

**Solution:**
Verify the `SqlDeployManifest` table exists and contains the correct MD5 hashes. The module automatically creates this table on the first run of `Invoke-AllSqlScripts`.

### Schema Export Failures

**Problem:** Export-DbSchema fails or misses objects.

**Solution:**
Ensure the account running the script has `VIEW DEFINITION` permissions on the database. This version of the module uses native metadata views (`sys.sql_modules`) and does not require SMO.
