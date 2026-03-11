#Requires -Version 5.1
<#
.SYNOPSIS
    Audits NTFS permissions on a network share root and all subfolders.

.DESCRIPTION
    Recursively enumerates folders under a given UNC path and reports:
    - Identity (user or group)
    - Account type (local, domain user, domain group, built-in)
    - Access rights
    - Allow/Deny
    - Inheritance status
    - Whether the ACE is inherited or explicitly set

.PARAMETER RootPath
    UNC path to audit. Default: \\hkfiles\E

.PARAMETER OutputCsv
    Optional CSV export path. If omitted, outputs to console only.

.PARAMETER Depth
    Max folder recursion depth. 0 = root only. Default: unlimited (-1)

.PARAMETER SkipInherited
    If set, only reports explicitly-set ACEs (filters out inherited entries).

.PARAMETER IncludeFiles
    If set, also audits individual file ACLs (can be very slow on large shares).

.EXAMPLE
    .\Get-SharePermissionAudit.ps1 -RootPath "\\hkfiles\E" -OutputCsv "C:\Reports\HKFiles_Perms.csv"

.EXAMPLE
    .\Get-SharePermissionAudit.ps1 -RootPath "\\hkfiles\E" -SkipInherited -Depth 2

.NOTES
    - Run as a user with read access to the share and permission to read ACLs
    - For large shares, pipe output to CSV to avoid console overflow
    - Requires ActiveDirectory module for group membership resolution (optional)
#>

[CmdletBinding()]
param(
    [string]$RootPath       = "\\hkfiles\E",
    [string]$OutputCsv      = "",
    [int]$Depth             = -1,
    [switch]$SkipInherited,
    [switch]$IncludeFiles,
    [switch]$ExpandGroups       # Optionally resolve AD group members
)

#region --- Helper Functions ---

function Get-AccountType {
    param([string]$IdentityRef)
    switch -Regex ($IdentityRef) {
        '^BUILTIN\\'           { return "Built-in" }
        '^NT AUTHORITY\\'      { return "NT Authority" }
        '^CREATOR\s'           { return "Special" }
        '^S-1-'                { return "Unresolved SID" }
        default {
            # Try AD lookup if module is available
            if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
                $samAccount = $IdentityRef -replace '^.*\\', ''
                try {
                    $adUser = Get-ADUser -Identity $samAccount -ErrorAction Stop
                    return "Domain User"
                } catch {}
                try {
                    $adGroup = Get-ADGroup -Identity $samAccount -ErrorAction Stop
                    return "Domain Group ($($adGroup.GroupScope) $($adGroup.GroupCategory))"
                } catch {}
            }
            return "Domain Account"
        }
    }
}

function Get-FolderACL {
    param(
        [string]$FolderPath,
        [int]$CurrentDepth,
        [int]$MaxDepth
    )

    $results = @()

    # Get ACL
    try {
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
    } catch {
        Write-Warning "Cannot read ACL on: $FolderPath — $_"
        return $results
    }

    foreach ($ace in $acl.Access) {
        # Skip inherited if flag set
        if ($SkipInherited -and $ace.IsInherited) { continue }

        $identity   = $ace.IdentityReference.ToString()
        $acctType   = Get-AccountType -IdentityRef $identity
        $accessType = $ace.AccessControlType.ToString()   # Allow / Deny
        $rights     = $ace.FileSystemRights.ToString()
        $inherited  = $ace.IsInherited
        $propagation = $ace.PropagationFlags.ToString()
        $inheritance = $ace.InheritanceFlags.ToString()

        $results += [PSCustomObject]@{
            Path            = $FolderPath
            Identity        = $identity
            AccountType     = $acctType
            AccessType      = $accessType
            FileSystemRights = $rights
            IsInherited     = $inherited
            InheritanceFlags = $inheritance
            PropagationFlags = $propagation
            Owner           = $acl.Owner
        }
    }

    # Recurse into subfolders
    if ($MaxDepth -eq -1 -or $CurrentDepth -lt $MaxDepth) {
        try {
            $subFolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
        } catch {
            Write-Warning "Cannot enumerate: $FolderPath — $_"
            return $results
        }

        foreach ($sub in $subFolders) {
            $results += Get-FolderACL -FolderPath $sub.FullName -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
        }
    }

    # Optionally process files too
    if ($IncludeFiles) {
        try {
            $files = Get-ChildItem -Path $FolderPath -File -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                try {
                    $fileAcl = Get-Acl -Path $file.FullName -ErrorAction Stop
                    foreach ($ace in $fileAcl.Access) {
                        if ($SkipInherited -and $ace.IsInherited) { continue }
                        $results += [PSCustomObject]@{
                            Path            = $file.FullName
                            Identity        = $ace.IdentityReference.ToString()
                            AccountType     = Get-AccountType -IdentityRef $ace.IdentityReference.ToString()
                            AccessType      = $ace.AccessControlType.ToString()
                            FileSystemRights = $ace.FileSystemRights.ToString()
                            IsInherited     = $ace.IsInherited
                            InheritanceFlags = $ace.InheritanceFlags.ToString()
                            PropagationFlags = $ace.PropagationFlags.ToString()
                            Owner           = $fileAcl.Owner
                        }
                    }
                } catch {
                    Write-Warning "Cannot read ACL on file: $($file.FullName)"
                }
            }
        } catch {}
    }

    return $results
}

function Expand-ADGroupMembers {
    param([string]$GroupName)
    $samAccount = $GroupName -replace '^.*\\', ''
    try {
        $members = Get-ADGroupMember -Identity $samAccount -Recursive -ErrorAction Stop
        return ($members | Select-Object -ExpandProperty SamAccountName) -join "; "
    } catch {
        return "N/A"
    }
}

#endregion

#region --- Main Execution ---

Write-Host "`n[*] Starting NTFS Audit on: $RootPath" -ForegroundColor Cyan
Write-Host "[*] Depth limit: $(if ($Depth -eq -1) { 'Unlimited' } else { $Depth })"
Write-Host "[*] Skip inherited: $SkipInherited"
Write-Host "[*] Expand AD groups: $ExpandGroups`n"

$startTime = Get-Date

# Validate path
if (-not (Test-Path -Path $RootPath)) {
    Write-Error "Cannot access path: $RootPath"
    exit 1
}

# Run audit
$allResults = Get-FolderACL -FolderPath $RootPath -CurrentDepth 0 -MaxDepth $Depth

# Optionally expand group members
if ($ExpandGroups -and (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    Write-Host "[*] Expanding AD group memberships..." -ForegroundColor Yellow
    $allResults = $allResults | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name "GroupMembers" -Value (
            if ($_.AccountType -like "Domain Group*") {
                Expand-ADGroupMembers -GroupName $_.Identity
            } else { "N/A" }
        ) -PassThru
    }
}

$elapsed = (Get-Date) - $startTime
Write-Host "`n[+] Audit complete. $($allResults.Count) ACEs found across $(($allResults | Select-Object -ExpandProperty Path -Unique).Count) paths." -ForegroundColor Green
Write-Host "[+] Elapsed: $($elapsed.ToString('hh\:mm\:ss'))`n"

# Output
if ($OutputCsv -ne "") {
    try {
        $allResults | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
        Write-Host "[+] CSV exported to: $OutputCsv" -ForegroundColor Green
    } catch {
        Write-Error "Failed to write CSV: $_"
    }
} else {
    # Console output — group by path for readability
    $grouped = $allResults | Group-Object -Property Path
    foreach ($group in $grouped) {
        Write-Host "`n── $($group.Name)" -ForegroundColor Yellow
        $group.Group | Format-Table Identity, AccountType, AccessType, FileSystemRights, IsInherited -AutoSize
    }
}

#endregion
