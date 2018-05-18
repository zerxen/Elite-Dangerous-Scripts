param (
    [Parameter(ValueFromRemainingArguments=$true)]
    $paramPath
)

if ($paramPath) {

  $linkPath  = $paramPath.TrimEnd(".lnk")

  $WshShell = New-Object -ComObject WScript.Shell
  $Shortcut = $WshShell.CreateShortcut("$paramPath")
  $targetPath = $Shortcut.TargetPath

  New-Item -Force -Path $linkPath -ItemType Junction -Target $targetPath -ErrorAction SilentlyContinue

} else {

 $sendToPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\SendTo"

 $scriptPath = $MyInvocation.MyCommand.Definition 

 cp $scriptPath $sendToPath

 New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue

 ""
 "Usage: right-click a shortcut and click 'Send to' -> shortcut2junction.ps1"

 if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
   $arguments = "& '" + $MyInvocation.MyCommand.Definition + "'"
   start-process powershell -Verb runAs -ArgumentList $arguments
   break
 }

 $command = '"' + (Get-Command powershell.exe).Definition + '" -File "%1" %*'

 Set-ItemProperty -Path HKCR:\Microsoft.PowerShellScript.1\ShellEx\DropHandler -Name "(Default)" -Value "{60254CA5-953B-11CF-8C96-00AA00B8708C}"
 Set-ItemProperty -Path HKCR:\Microsoft.PowerShellScript.1\Shell\Open\Command  -Name "(Default)" -Value $command

 ""
 Read-Host "Press ENTER to exit"

}
