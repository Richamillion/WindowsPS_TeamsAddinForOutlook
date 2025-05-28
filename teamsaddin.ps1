function Get-MsiVersion {
    param ([string]$msiPath)

    $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    $database = $windowsInstaller.OpenDatabase($msiPath, 0)
    $query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
    $view = $database.OpenView($query)
    $view.Execute()
    $record = $view.Fetch()
    $version = $record.StringData(1)
    $view.Close()
    return $version
}

$teamsFolder = Get-ChildItem -Path "C:\Program Files\WindowsApps" -Recurse | Where-Object {$_.Name -like "MSTeams_*" -and $_.PSIsContainer} | Select-Object -ExpandProperty FullName
$addinPath = "$teamsFolder\MicrosoftTeamsMeetingAddinInstaller.msi"
$teamsAddinVersion = Get-MsiVersion -msiPath $addinPath
$targetPath = "C:\Program Files (x86)\Microsoft\TeamsMeetingAdd-in"
msiexec /i $addinPath ALLUSERS=1 /qn /norestart TARGETDIR="$targetPath\$teamsAddinVersion\"
