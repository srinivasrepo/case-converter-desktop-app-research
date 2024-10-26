Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "^%{TAB}"
WScript.Sleep 2000
WshShell.SendKeys "^{C}"

' Get clipboard text
Set objHTML = CreateObject("htmlfile")
text = objHTML.ParentWindow.ClipboardData.GetData("text")


Set oExec = WshShell.Exec("clip")
Set oIn = oExec.stdIn
oIn.WriteLine UCase(text)
oIn.Close
WScript.Sleep 2000
WshShell.SendKeys "^V"