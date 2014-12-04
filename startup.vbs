strComputer = "."
 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

Dim driveLetter

For Each objDrive in colDrives
	If objDrive.VolumeName = "CONTEXT" Then
		driveLetter = objDrive.DriveLetter & ":"
	End If 
Next

If len(driveLetter) <= 0 Then
    driveLetter = "C:"
End If
 
contextPath = driveLetter & "\context.ps1"

If objFSO.FileExists(contextPath) Then
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run("powershell -NonInteractive -NoProfile -NoLogo -ExecutionPolicy Unrestricted -file " & contextPath)
End If

