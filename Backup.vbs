' Set source folder for backup
strSourceFolder = "C:\SourceFolder" ' Replace with your source folder path

' Set backup folder
strBackupFolder = "C:\BackupFolder" ' Replace with your backup folder path

' Create an FSO object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get source folder
Set objSourceFolder = objFSO.GetFolder(strSourceFolder)

' Create backup folder if it doesn't exist
If Not objFSO.FolderExists(strBackupFolder) Then
    objFSO.CreateFolder(strBackupFolder)
End If

' Get current date and time for backup file name
strBackupFileName = Format(Now, "yyyyMMddHHmmss") & ".zip"

' Create a ZIP object for compression
Set objZip = CreateObject("Scripting.FileSystemObject").CreateObject("Scripting.Compression.Archive")

' Add files from the source folder to the ZIP object
objZip.AddItem strSourceFolder & "\*"

' Save the ZIP file in the backup folder
objZip.SaveAs strBackupFolder & "\" & strBackupFileName

' Display backup completion message
MsgBox "Backup concluído com sucesso para o arquivo: " & strBackupFileName

' Function to restore files from a backup
Function RestoreBackup(strBackupFile)
    ' Set destination folder for restoration
    strDestFolder = "C:\DestinationFolder" ' Replace with your destination folder path

    ' Create an FSO object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Get destination folder
    Set objDestFolder = objFSO.GetFolder(strDestFolder)

    ' Create a ZIP object for decompression
    Set objZip = CreateObject("Scripting.FileSystemObject").CreateObject("Scripting.Compression.Archive")

    ' Open the backup ZIP file
    objZip.Open strBackupFile

    ' Extract files from the ZIP object to the destination folder
    objZip.ExtractHere strDestFolder

    ' Display restoration completion message
    MsgBox "Restauração concluída com sucesso do arquivo: " & strBackupFileName
End Function

' Example usage of the RestoreBackup function
' Replace "C:\BackupFolder\myBackup.zip" with the actual path to your backup file
RestoreBackup "C:\BackupFolder\myBackup.zip"
