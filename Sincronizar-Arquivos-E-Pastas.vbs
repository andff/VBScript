' Set source and destination folders
strSourceFolder = "C:\SourceFolder" ' Replace with your source folder path
strDestFolder = "C:\DestinationFolder" ' Replace with your destination folder path

' Create an FSO object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get source and destination folders
Set objSourceFolder = objFSO.GetFolder(strSourceFolder)
Set objDestFolder = objFSO.GetFolder(strDestFolder)

' Synchronize files and folders
objSourceFolder.SyncToFolder objDestFolder, True

' Display success message
MsgBox "Sincronização de arquivos e pastas concluída com sucesso!"
