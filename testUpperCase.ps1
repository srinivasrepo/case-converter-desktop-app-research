# [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
# [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


# [System.Windows.Forms.Sendkeys]::SendWait("%{TAB}")
# Start-Sleep -Seconds 3
# [System.Windows.Forms.SendKeys]::SendWait("^C")
# $value=Get-Clipboard
# Write-Output $value
# Write-Output "Copied to clipboard"
# $value=$value.ToUpper()
# Set-Clipboard -Value $value
# [System.Windows.Forms.SendKeys]::SendWait("^V")

# $oExec = $obj.Exec("clip")
# sleep -s 3
$obj = New-Object -com Wscript.Shell
$t = $obj.SendKeys("^c")
$value=Get-Clipboard
Write-Output $value
Write-Output $foo
Set-Clipboard -Value $value
$obj.SendKeys("^v")
sleep -s 10