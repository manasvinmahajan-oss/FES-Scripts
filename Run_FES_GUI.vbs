Option Explicit

' ====================================================================
' FES Bids Manager - Silent Launcher (VBScript)
' Double-click this file to launch the GUI silently (no console window)
' ====================================================================

Dim objShell, objFSO, strScriptPath, strPythonPath, strGuiPath
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strGuiPath = strScriptPath & "\FES_Bids_Runner_PRODUCTION.py"

' Check if GUI file exists
If Not objFSO.FileExists(strGuiPath) Then
    MsgBox "Error: FES_Bids_Runner_PRODUCTION.py not found!" & vbCrLf & vbCrLf & _
           "Expected location: " & strGuiPath, vbCritical, "FES Bids Manager"
    WScript.Quit 1
End If

' Try to find Anaconda Python
Dim anacondaPaths(2)
anacondaPaths(0) = objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\anaconda3\pythonw.exe"
anacondaPaths(1) = "C:\Users\enrolment\AppData\Local\anaconda3\pythonw.exe"
anacondaPaths(2) = "C:\ProgramData\Anaconda3\pythonw.exe"

' Try each Anaconda path
Dim i, foundPython
foundPython = False

For i = 0 To UBound(anacondaPaths)
    If objFSO.FileExists(anacondaPaths(i)) Then
        strPythonPath = anacondaPaths(i)
        foundPython = True
        Exit For
    End If
Next

' If Anaconda not found, try system pythonw
If Not foundPython Then
    On Error Resume Next
    strPythonPath = objShell.Exec("where pythonw").StdOut.ReadLine
    If Err.Number = 0 And strPythonPath <> "" Then
        foundPython = True
    End If
    On Error GoTo 0
End If

' Launch GUI or show error
If foundPython Then
    ' Launch GUI silently (no console window)
    objShell.Run """" & strPythonPath & """ """ & strGuiPath & """", 0, False
Else
    MsgBox "Error: Could not find Python!" & vbCrLf & vbCrLf & _
           "Please ensure Anaconda or Python is installed." & vbCrLf & vbCrLf & _
           "Expected locations:" & vbCrLf & _
           "  - %USERPROFILE%\anaconda3" & vbCrLf & _
           "  - C:\Users\enrolment\AppData\Local\anaconda3" & vbCrLf & _
           "  - C:\ProgramData\Anaconda3", vbCritical, "FES Bids Manager"
    WScript.Quit 1
End If

' Cleanup
Set objShell = Nothing
Set objFSO = Nothing
