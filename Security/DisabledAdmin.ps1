<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Disable Local Administrator Account
SYNOPSIS    : Disables the built-in local Administrator account
DESCRIPTION : Checks if the local Administrator account exists and is
              enabled. If enabled, the script disables the account to
              reduce local attack surface and improve security posture
==============================================================================
#>

try {
$Account = 'Administrator'
$isEnabled = (Get-LocalUser $Account -ErrorAction Stop).enabled
}
catch {
"No such account exists"
}
if ($isEnabled) {
Disable-LocalUser $Account
}
