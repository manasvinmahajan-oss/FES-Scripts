Option Explicit

' ====================================================================
' FES Bids Manager - Universal Silent Launcher (VBScript)
' Works on any computer with Python/Anaconda installed
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

' Build list of possible Python locations
Dim pythonPaths(15)
Dim userProfile, localAppData
userProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
localAppData = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

' Common Anaconda locations (works for any user)
pythonPaths(0) = userProfile & "\anaconda3\pythonw.exe"
pythonPaths(1) = userProfile & "\Anaconda3\pythonw.exe"
pythonPaths(2) = localAppData & "\anaconda3\pythonw.exe"
pythonPaths(3) = localAppData & "\Continuum\anaconda3\pythonw.exe"
pythonPaths(4) = "C:\ProgramData\Anaconda3\pythonw.exe"
pythonPaths(5) = "C:\Anaconda3\pythonw.exe"
pythonPaths(6) = "C:\Python39\pythonw.exe"
pythonPaths(7) = "C:\Python310\pythonw.exe"
pythonPaths(8) = "C:\Python311\pythonw.exe"
pythonPaths(9) = "C:\Python312\pythonw.exe"
pythonPaths(10) = "C:\Python313\pythonw.exe"
pythonPaths(11) = userProfile & "\AppData\Local\Programs\Python\Python39\pythonw.exe"
pythonPaths(12) = userProfile & "\AppData\Local\Programs\Python\Python310\pythonw.exe"
pythonPaths(13) = userProfile & "\AppData\Local\Programs\Python\Python311\pythonw.exe"
pythonPaths(14) = userProfile & "\AppData\Local\Programs\Python\Python312\pythonw.exe"
pythonPaths(15) = userProfile & "\AppData\Local\Programs\Python\Python313\pythonw.exe"

' Try each path
Dim i, foundPython
foundPython = False

For i = 0 To UBound(pythonPaths)
    If objFSO.FileExists(pythonPaths(i)) Then
        strPythonPath = pythonPaths(i)
        foundPython = True
        Exit For
    End If
Next

' If not found in standard locations, try system PATH
If Not foundPython Then
    On Error Resume Next
    Dim objExec
    Set objExec = objShell.Exec("where pythonw")
    If Err.Number = 0 Then
        strPythonPath = objExec.StdOut.ReadLine
        If strPythonPath <> "" And objFSO.FileExists(strPythonPath) Then
            foundPython = True
        End If
    End If
    On Error GoTo 0
End If

' If still not found, try 'python' instead of 'pythonw'
If Not foundPython Then
    On Error Resume Next
    Set objExec = objShell.Exec("where python")
    If Err.Number = 0 Then
        strPythonPath = objExec.StdOut.ReadLine
        If strPythonPath <> "" And objFSO.FileExists(strPythonPath) Then
            ' Replace python.exe with pythonw.exe if it exists
            Dim pythonwPath
            pythonwPath = Replace(strPythonPath, "python.exe", "pythonw.exe")
            If objFSO.FileExists(pythonwPath) Then
                strPythonPath = pythonwPath
                foundPython = True
            Else
                ' Use python.exe if pythonw.exe doesn't exist
                foundPython = True
            End If
        End If
    End If
    On Error GoTo 0
End If

' Launch GUI or show error
If foundPython Then
    ' Launch GUI silently (window style 0 = hidden)
    objShell.Run """" & strPythonPath & """ """ & strGuiPath & """", 0, False
Else
    ' Show detailed error message
    MsgBox "Error: Could not find Python/Anaconda!" & vbCrLf & vbCrLf & _
           "Please ensure Python or Anaconda is installed." & vbCrLf & vbCrLf & _
           "Searched locations:" & vbCrLf & _
           "  • %USERPROFILE%\anaconda3" & vbCrLf & _
           "  • %LOCALAPPDATA%\anaconda3" & vbCrLf & _
           "  • C:\ProgramData\Anaconda3" & vbCrLf & _
           "  • C:\Python3X\" & vbCrLf & _
           "  • System PATH" & vbCrLf & vbCrLf & _
           "If Python IS installed, try using Run_FES_GUI_Universal.bat" & vbCrLf & _
           "to see detailed error messages.", vbCritical, "FES Bids Manager"
    WScript.Quit 1
End If

' Cleanup
Set objExec = Nothing
Set objShell = Nothing
Set objFSO = Nothing
