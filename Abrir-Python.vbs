Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python your_python_script.py", 1 ' 1 waits for script to finish
