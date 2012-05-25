' Get status of services running on multiple computers, looking for custom
' (non-default) credentials
'
' Targets are retrieved from a list of target computers in a plain-text file, 
' one computer per line
' by "Waldo G" (gwaldo@gmail.com)

Const INPUT_FILE_NAME = "C:\path\to\computerlist.txt"
Const FOR_READING = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(INPUT_FILE_NAME, FOR_READING)
strComputers = objFile.ReadAll
objFile.Close
arrComputers = Split(strComputers, vbCrLf)
For Each strComputer In arrComputers
	Pingable = IsAlive(strComputer)
		If not Pingable Then	'EPIC FAIL
			WScript.Echo strComputer & " is not pingable."
		Else					'Server Pings; moving on with life
			WScript.Echo strComputer
			Set objWMIService = GetObject("winmgmts:" _
				& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colRunningServices = objWMIService.ExecQuery _
				("Select * from Win32_Service")
			For Each objService in colRunningServices
				If (lcase(objService.Startname)) <> "localsystem" Then 
					If (lcase(objService.StartName)) <> "nt authority\localservice" Then
						If (lcase(objService.StartName)) <> "nt authority\networkservice" Then
							Wscript.Echo objService.DisplayName  & "," & objService.StartName
						End If
					End If
				End If
			Next
		End If
Next




'======================================
'======================================
'===Putting the "Fun" in "Functions"===
'======================================
'======================================


'Function fnPing(strComputer)
Function IsAlive(strComputer)
	' by Phil Gordemer of ARH Associates
	' from http://www.tek-tips.com/viewthread.cfm?qid=1279504&page=3
	
	'--- Test to see if host or url alive through ping ---
	' Returns True if Host responds to ping
	'
	' Though there are other ways to ping a computer, Win2K,
	' XP and different versions of PING return different error
	' codes. So the only reliable way to see if the ping
	' was sucessful is to read the output of the ping
	' command and look for "TTL="
	'
	' strHost is a hostname or IP
	
	const OpenAsASCII = 0
	const FailIfNotExist = 0
	const ForReading =  1
	Dim objShell, objFSO, sTempFile, fFile
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sTempFile = objFSO.GetSpecialFolder(2).ShortPath & "\" & objFSO.GetTempName
	
	objShell.Run "%comspec% /c ping.exe -n 2 -w 500 " & strComputer & ">" & sTempFile, 0 , True
	Set fFile = objFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII)
	Select Case InStr(fFile.ReadAll, "TTL=")
		Case 0
			IsAlive = False
		Case Else
			IsAlive = True
	End Select
	fFile.Close
	objFSO.DeleteFile(sTempFile)
	Set objFSO = Nothing
	Set objShell = Nothing
End Function


'======================================
'======================================