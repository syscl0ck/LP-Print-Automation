$action = New-ScheduledTaskAction -Execute "C:\Program Files(x86)\PalletizerAutoPrint\PalletizerAutoPrint.exe"
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -LogonType Interactive -RunLevel LeastPrivilege -UserId "isaac" #-User "HELIENE\Operator"
Register-ScheduledTask -TaskName "StartPalletizerAutoPrint" -Action $action -Trigger $trigger -Principal $principal
