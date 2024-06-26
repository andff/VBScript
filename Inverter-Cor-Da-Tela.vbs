' Invert screen colors
Set objShell = CreateObject("WScript.Shell")
objShell.SendKeys "%{TAB}" ' Hides the mouse cursor

' Display a message indicating the screen colors are inverted
MsgBox "Cores da tela invertidas! Pressione Ctrl+Alt+Del para reativar."

' Invert colors
objShell.SendKeys("{APPKEY Scroll Lock}") ' Inverts colors

' Wait for 5 seconds
WScript.Sleep 5000

' Revert colors
objShell.SendKeys("{APPKEY Scroll Lock}") ' Reverts colors

' Re-enable the mouse cursor
objShell.SendKeys "%{TAB}" ' Unhides the mouse cursor
