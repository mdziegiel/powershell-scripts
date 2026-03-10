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