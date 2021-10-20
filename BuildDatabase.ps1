$acNewDatabaseFormatAccess2007 = 12
$acCmdCompileAndSaveAllModules = 126

$acClass = 2
$acForm = 2
$acModule = 5

$dbPath = (Get-Item .).FullName + "\"
$dbName = "BAC.accdb"
$appAccess = ""
function Import-File {
    param (
        $file
    )

    $type = -1
    $mod = ""
    if ($file.Extension -ceq ".bas") {
        Write-Host $file.Name        
        $mod = $vbComponents.Import($file.FullName)
        $type = $acModule
    }
    if ($file.Extension -ceq ".cls") {
        Write-Host $file.Name        
        $mod = $vbComponents.Import($file.FullName)        
        $type = $acClass
    }
    if ($type -gt 0) {
        $appAccess.DoCmd.Save($type, $file.BaseName)
    }
}

Remove-Item -Path ($dbPath + $dbName)
$appAccess = New-Object -ComObject Access.Application
$appAccess.Visible = $true
$appAccess.NewCurrentDatabase($dbPath + $dbName, $acNewDatabaseFormatAccess2007)
$vbComponents = $appAccess.VBE.ActiveVBProject.VBComponents

# $appAccess.LoadFromText(5, "cls_Better_Access_Chart", $dbPath + "cls_Better_Access_Chart.cls")
# $appAccess.VBE.ActiveVBProject.VBComponents.Import($dbPath + "cls_Better_Access_Chart.cls")

Get-ChildItem $dbPath | ForEach-Object {Import-File $_ }

$appAccess.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
#$appAccess.CloseCurrentDatabase()
#$appAccess.Quit()

