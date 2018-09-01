set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.run "runas /user:username %comspec%" 'Open command prompt
  'username is the other account login name which you can use your account name to replace it
WScript.Sleep 1000
WshShell.SendKeys "password" 'send password
  'password is the other account login password which you should use your account password to replace it just like 'username'
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000 'Open TM
WshShell.SendKeys Chr(34) + "path direction with application full name" + Chr(34)
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "exit" 'Close command prompt
WshShell.SendKeys "{ENTER}"
