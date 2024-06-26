' Set the path to the MP3 file
strMusicFile = "C:\Path\To\Your\Music.mp3" ' Replace with the actual path to your MP3 file

' Create a WMP object
Set objWMPlayer = CreateObject("WMPlayer.OCX")

' Set the media item to the MP3 file
objWMPlayer.MediaURL = strMusicFile

' Play the music
objWMPlayer.Controls.Play

' Wait until the music finishes playing
Do While objWMPlayer.playState <> 11
    WScript.Sleep 100
Loop

' Close the WMP object
objWMPlayer.Quit
