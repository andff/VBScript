' Disable the mouse cursor
Set objShell = CreateObject("WScript.Shell")
objShell.SendKeys "%{TAB}" ' Hides the mouse cursor

' Display a message indicating the cursor is disabled
MsgBox "Cursor desativado! Pressione Ctrl+Alt+Del para reativar."

' Wait for 5 seconds
WScript.Sleep 5000

' Re-enable the mouse cursor
objShell.SendKeys "%{TAB}" ' Unhides the mouse cursor
