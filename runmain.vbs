'Create a WshShell Object
Set WshShell = Wscript.CreateObject("Wscript.Shell")

'Create a WshShortcut Object
Set oShellLink = WshShell.CreateShortcut("runmain.lnk")

'Set the Target Path for the shortcut
oShellLink.TargetPath = "c:\windows\system32\mshta.exe"

'Set the additional parameters for the shortcut
oShellLink.Arguments = "c:\users\public\cdp\dev\snapback\apps\HibernateShow\main.html"

'Save the shortcut
oShellLink.Save

'Clean up the WshShortcut Object
Set oShellLink = Nothing