<#
This script checks if a known physical path is shared by any user in the EFT server.
It connects to the EFT server, retrieves the list of users,
then iterates through their home directories to find any subfolders that map to the specified physical path.
The results are collected in a custom object for further analysis or reporting.

TODO: Iterate into all subfolders of the home directory, not just the immediate children.
#>
$comUser = Read-Host -Prompt "Enter EFT Admin Username"
$comPass = Read-Host -Prompt "Enter EFT Admin Password" # PS 5.1 does not provide a good way to mask this input, so be cautious when entering in a screenshare. SecureString is not supported in PS 5.1 for COM objects. I believe PS 7.2 and later can handle this with -MaskInput or seure string conversion.
$Server = New-Object -ComObject "SFTPCOMInterface.CIServer"
$Server.Connect("localhost", 1100, $comUser, $comPass)
$Sites = $Server.Sites()
$Site = $Sites.Item(0)

$userList = $Site.GetUsers()

$knownPhysicalPath = "C:\InetPub\EFTRoot\MySite\Shared\"

$results = @()

foreach ($userName in $userList) {
    $userSettings = $Site.GetUserSettings($userName)
    $homeVfsPath = $userSettings.GetHomeDirString()

    $folderListRaw = $Site.GetFolderList($homeVfsPath)
    $subFolders = $folderListRaw -split "`r`n" | Where-Object { $_ -ne '' }

    foreach ($folder in $subFolders) {
        $childVfsPath = $homeVfsPath + $folder
        $physicalPath = $Site.GetPhysicalPath($childVfsPath)

        if ($physicalPath.TrimEnd('\') -ieq $knownPhysicalPath.TrimEnd('\')) {
            $results += [PSCustomObject]@{
                User     = $userName
                VfsPath  = $childVfsPath
                Physical = $physicalPath
            }
        }
    }
}

$results