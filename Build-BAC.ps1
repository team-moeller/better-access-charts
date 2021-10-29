# --------------------------------------------------------------------
# --- Read all files in the sources subdir 
# --- (exported from an Access Database)
# --- and rebuilds the Access Database (.accdb)
# --- Created By K.D. Gundermann 11/2021
# --------------------------------------------------------------------
Param(
    [Parameter(Mandatory=$true)]
    [string]$dbName
)

# Constant Declarations
$acNewDatabaseFormatAccess2007 = 12
$acCmdCompileAndSaveAllModules = 126
$acCmdSaveAllModules = 280

$acQuery = 1
$acForm = 2
$acClass = 2
$acMacro = 4
$acModule = 5

$acStructureAndData = 1

$guidADODB = "{2A75196C-D9EB-4129-B803-931327F72D5C}"

# Variables Declarations
$dbPath = (Get-Item .).FullName + "\"
$appAccess = $null
$vbComponents = $null

# Private Functions
function Import-Table {
    param ( $file )
    $appAccess.ImportXML($file, $acStructureAndData)
}
function Import-Query {
    param ( $file )
    $appAccess.LoadFromText($acQuery, (Get-Item $file).BaseName, $file)
}

function Import-Form {
    param ( $file )
    $appAccess.LoadFromText($acForm, (Get-Item $file).BaseName, $file)
}

function Import-Macro {
    param ( $file )
    $appAccess.LoadFromText($acMacro, (Get-Item $file).BaseName, $file)
}

function Import-Module {
    param ( $file )

    $mod = $vbComponents.Import($file.FullName)
    # $mod.Activate()          
    # $appAccess.DoCmd.Save($acModule, $file.BaseName)
}
function Import-Class {
    param ( $file )
    $mod = $vbComponents.Import($file.FullName)  
    # $mod.Activate()          
    # $appAccess.DoCmd.Save($acClass, $file.BaseName)
}

if (-not $dbName.EndsWith(".accdb")) {
    $dbName = $dbName + ".accdb"
}

if (Get-Item -Path ($dbPath + $dbName) -ErrorAction Ignore ) {
    Remove-Item -Path ($dbPath + $dbName) -ErrorAction Inquire
}
$appAccess = New-Object -ComObject Access.Application
$appAccess.Visible = $true
$appAccess.NewCurrentDatabase($dbPath + $dbName, $acNewDatabaseFormatAccess2007)
$vbComponents = $appAccess.VBE.ActiveVBProject.VBComponents

Set-Location -Path ($dbPath + "\Source")

Get-ChildItem ("Tables\*.tbl")  | ForEach-Object {Import-Table $_ }
Get-ChildItem ("Forms\*.frm")   | ForEach-Object {Import-Form $_ }
Get-ChildItem ("Macros\*.mcr")  | ForEach-Object {Import-Macro $_ }
Get-ChildItem ("Queries\*.qry") | ForEach-Object {Import-Query $_ }
Get-ChildItem ("Modules\*.bas") | ForEach-Object {Import-Module $_ }
Get-ChildItem ("Modules\*.cls") | ForEach-Object {Import-Class $_ }

$vbComponents = $null

$appAccess.References.AddFromGuid($guidADODB, 2, 8)

# $appAccess.DoCmd.RunCommand($acCmdSaveAllModules)
$appAccess.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
$appAccess.CloseCurrentDatabase()
$appAccess.Quit()

$appAccess = $null
