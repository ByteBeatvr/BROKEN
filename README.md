Dim objShell
Dim appPath1, appPath2, appPath3, appPath4
Dim soundPath1, soundPath2, soundPath3, soundPath4
Dim wmpPath

' Path to the "New Folder" in Downloads
folderPath = "C:\Users\Administrator\Downloads\New Folder\"

' Paths to the applications you want to open
appPath1 = folderPath & "bouncingcircles.exe"
appPath2 = folderPath & "inv.exe"
appPath3 = folderPath & "mlt.exe"
appPath4 = folderPath & "gdihell.exe"

' Paths to the sound files
soundPath1 = folderPath & "html5bytebeat.wav"
soundPath2 = folderPath & "html5bytebeat (1).wav"
soundPath3 = folderPath & "html5bytebeat (2).wav"
soundPath4 = folderPath & "html5bytebeat (3).wav"

' Path to Windows Media Player (WMP)
wmpPath = "C:\Program Files (x86)\Windows Media Player\wmplayer.exe"

Set objShell = CreateObject("WScript.Shell")

' Repeat the process four times
For i = 1 To 4
    If i = 1 Then
        ' Open the first application
        objShell.Run appPath1
        WScript.Sleep 1000 ' Wait 1 second to ensure the app has opened

        ' Open Windows Media Player and play the first sound
        objShell.Run """" & wmpPath & """ """ & soundPath1 & """"
        WScript.Sleep 1000 ' Wait for WMP to open and start playing

        ' Wait for 1 minute (60000 ms)
        WScript.Sleep 60000

        ' Close the first application
        objShell.AppActivate("bouncingcircles") ' Change "bouncingcircles" to the actual window title of the app
        objShell.SendKeys "^w" ' This simulates pressing Ctrl+W to close the app
    ElseIf i = 2 Then
        ' Open the second application
        objShell.Run appPath2
        WScript.Sleep 1000 ' Wait 1 second to ensure the app has opened

        ' Open Windows Media Player and play the second sound
        objShell.Run """" & wmpPath & """ """ & soundPath2 & """"
        WScript.Sleep 1000 ' Wait for WMP to open and start playing

        ' Wait for 1 minute (60000 ms)
        WScript.Sleep 60000

        ' Close the second application
        objShell.AppActivate("inv") ' Change "inv" to the actual window title of the app
        objShell.SendKeys "^w" ' This simulates pressing Ctrl+W to close the app
    ElseIf i = 3 Then
        ' Open the third application
        objShell.Run appPath3
        WScript.Sleep 1000 ' Wait 1 second to ensure the app has opened

        ' Open Windows Media Player and play the third sound
        objShell.Run """" & wmpPath & """ """ & soundPath3 & """"
        WScript.Sleep 1000 ' Wait for WMP to open and start playing

        ' Wait for 1 minute (60000 ms)
        WScript.Sleep 60000

        ' Close the third application
        objShell.AppActivate("mlt") ' Change "mlt" to the actual window title of the app
        objShell.SendKeys "^w" ' This simulates pressing Ctrl+W to close the app
    ElseIf i = 4 Then
        ' Open the fourth application
        objShell.Run appPath4
        WScript.Sleep 1000 ' Wait 1 second to ensure the app has opened

        ' Open Windows Media Player and play the fourth sound
        objShell.Run """" & wmpPath & """ """ & soundPath4 & """"
        WScript.Sleep 1000 ' Wait for WMP to open and start playing

        ' Wait for 1 minute (60000 ms)
        WScript.Sleep 60000

        ' Close the fourth application
        objShell.AppActivate("gdihell") ' Change "gdihell" to the actual window title of the app
        objShell.SendKeys "^w" ' This simulates pressing Ctrl+W to close the app
    End If

    ' Wait 1 second before opening the next app
    WScript.Sleep 1000
Next

Set objShell = Nothing
