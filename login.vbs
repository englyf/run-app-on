set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.run "runas /user:username %comspec%" 'Open command prompt
WScript.Sleep 1000
WshShell.SendKeys "password" 'send password 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000 'Open TM
WshShell.SendKeys Chr(34) + "path direction with application full name" + Chr(34)
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "exit" 'Close command prompt
WshShell.SendKeys "{ENTER}"
