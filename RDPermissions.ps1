#Requires -Version 3.0
#Requires -RunAsAdministrator

$global:perms = @()

function Show-CurrentPermissions()
{
    $global:perms = @(); $count = 0
    foreach($entry in (Get-WmiObject -Namespace "root\cimv2\terminalservices" -Query "SELECT * FROM Win32_TSAccount"))
    {
        $row = "" | Select Index,Connection,Group,Type,Permissions
        $row.Index = $count
        $row.Connection = $entry.TerminalName
        $row.Group = $entry.AccountName
        if(!$entry.PermissionsDenied)
        {
            $row.Type = "Allow"
            $row.Permissions = Convert-Permissions($entry.PermissionsAllowed)
        }
        else
        {
            $row.Type = "Deny"
            $row.Permissions = Convert-Permissions($entry.PermissionsDenied)
        }
        $global:perms += $row
        $count ++
    }
    $global:perms | Format-Table -AutoSize
}

function Convert-Permissions([int]$bits)
{
    $permissions = @{
        0x001 = "Query"
        0x002 = "Set"
        0x004 = "Logoff"
        0x010 = "Shadow"
        0x020 = "Logon"
        0x040 = "Reset"
        0x080 = "Message"
        0x100 = "Connect"
        0x200 = "Disconnect"
        0xF0008 = "Virtual Channels"
    }

    foreach($bitmask in $permissions.Keys | Sort-Object)
    {
        if(($bits -band $bitmask) -eq $bitmask)
        {
            $output += $permissions[$bitmask] + ", "
        }
    }
    if($output) { $output = $output.Substring(0, $output.Length-2) }
    return $output
}

function Add-Account
{
    Write-Host "Enter the name of the group you want to add: " -ForegroundColor Yellow -NoNewline; $account = Read-Host
    Write-Host "Enter the connection you want to add this group to (leave empty for all): " -ForegroundColor Yellow -NoNewline; $terminal = Read-Host

    foreach($object in (Get-WmiObject -Class "Win32_TSPermissionsSetting" -Namespace "root\cimv2\terminalservices"))
    {
        if(!$terminal -or $object.TerminalName -eq $terminal)
        {
            Invoke-WmiMethod -InputObject $object -Name "AddAccount" -ArgumentList $account,3
        }
    }

    Write-Host "The user/group has now been added, please add permissions!" -ForegroundColor Green
    Menu
}

function Delete-Account
{
    Write-Host "Select the index of the group you want to delete: " -ForegroundColor Yellow -NoNewline; $index = Read-Host

    $entry = Get-WmiObject -Namespace "root\cimv2\terminalservices" -Query ("SELECT * FROM Win32_TSAccount WHERE TerminalName='"+$perms[$index].Connection+"' AND AccountName='"+$perms[$index].Group+"'").Replace("\", "\\")
    Write-Host $entry -ForegroundColor Red
    $entry.Delete()

    Write-Host "The user/group has now been deleted from the permissions!" -ForegroundColor Green
    Menu
}

function Edit-Permissions
{
    Write-Host "Select the index, or enter the group name for multiple connections: " -ForegroundColor Yellow -NoNewline; $index = Read-Host
    Write-Host "[0] Query`n[1] Set`n[2] Logoff`n[3] Virtual Channels`n[4] Shadow`n[5] Logon`n[6] Reset`n[7] Message`n[8] Connect`n[9] Disconnect"
    Write-Host "Select what permission you want to edit: " -ForegroundColor Yellow -NoNewline; $permission = Read-Host
    Write-Host "Select [0] to deny or [1] to allow this permission: " -ForegroundColor Yellow -NoNewline; [int]$allow = Read-Host

    if($index -match "^\d+$") { $query = "SELECT * FROM Win32_TSAccount WHERE TerminalName='"+$perms[$index].Connection+"' AND AccountName='"+$perms[$index].Group+"'" }
    else { $query = "SELECT * FROM Win32_TSAccount WHERE AccountName='"+$index+"'" }

    foreach($object in Get-WmiObject -Namespace "root\cimv2\terminalservices" -Query $query.Replace("\", "\\"))
    {
        $object.ModifyPermissions($permission, $allow)
    }

    Write-Host "The permissions of this user/group are now changed!" -ForegroundColor Green
    Menu
}

function Restore-Permissions
{
    $object = Get-WmiObject -Class "Win32_TSPermissionsSetting" -Namespace "root\cimv2\terminalservices"
    foreach($entry in $object) { Invoke-WmiMethod -InputObject $entry -Name "RestoreDefaults" }

    Write-Host "All Remote Desktop permissions have been reset to default!" -ForegroundColor Green
    Menu
}

function Menu
{
    Show-CurrentPermissions
    Write-Host "Select an option; [A] Add a group, [D] Delete a group, [E] Edit permissions, [R] Restore to default: " -ForegroundColor Yellow -NoNewline; $option = Read-Host

    if($option -eq "A") { Add-Account }
    if($option -eq "D") { Delete-Account }
    if($option -eq "E") { Edit-Permissions }
    if($option -eq "R") { Restore-Permissions }
}

Menu