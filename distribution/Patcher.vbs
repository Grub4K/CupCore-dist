Option Explicit

If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
      , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

Dim strRegValue, strFolder, objFolder, strCupheadDir, strCupheadDataDir, arrPatches, CurrentPatch, blnUnpatching, strPatchMessage, strSaveLocation, intOKCancel, file, BinaryData, strMD5
Dim objWshShl : Set objWshShl = CreateObject("WScript.Shell")
Dim objShl : Set objShl = CreateObject("Shell.Application")
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objMD5:  Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Dim objStream : Set objStream = CreateObject("ADODB.Stream")
Dim objXML : Set objXML = CreateObject("MSXML2.DOMDocument")
Dim objElement : Set objElement = objXML.CreateElement("tmp")
' Chenge current directory hotfix
objWshShl.CurrentDirectory = objFso.GetParentFolderName(WScript.ScriptFullName)

' Find Cuphead location.
' If cannot find, let select manually
On Error Resume Next
strRegValue = objWshShl.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 268910\InstallLocation")
On Error GoTo 0
If len(strRegValue) = 0 or Err.Number <> 0 Then
    Set objFolder = objShl.BrowseForFolder(0,"Cuphead not found, please select location manually.",0,17)

    If objFolder is Nothing Then
        Wscript.Quit()
    Else
        If Not objFso.FolderExists(objFolder.Self.Path & "\Cuphead_Data\") Then
            patcherError "Cuphead_Data not found!"
        Else
            strCupheadDir = objFolder.Self.Path
        End If
    End If
Else
    strCupheadDir = strRegValue
End If
strCupheadDataDir = strCupheadDir & "\Cuphead_Data\"
' Got location



' Patching array
arrPatches = Array(_
    Array("Managed\", "Assembly-CSharp.dll"),_
    Array("", "sharedassets1.assets"),_
    Array("", "sharedassets3.assets"),_
    Array("", "sharedassets10.assets") _
)

If (NOT objFso.FileExists("data\xdelta3.exe")) Then
    patcherError "Could not locate xdelta3"
End If
' Check Assembly-CSharp.dll as significant file
If ( objFso.FileExists(strCupheadDataDir & "Managed\Assembly-CSharp.dll" & ".bak") ) Then
    blnUnpatching = True
    strPatchMessage = "un"
ElseIf ( objFso.FileExists(strCupheadDataDir & "Managed\Assembly-CSharp.dll") ) Then
    ' Check for Current Patch
    If verifyMd5("e39a8a234edb59c07087a829de4fac34", strCupheadDataDir & "Managed\Assembly-CSharp.dll") Then
        patcherError "Cuphead Current Patch detected! Please install the LEGACY version."
    End If
    blnUnpatching = False
    strPatchMessage = ""
Else
    patcherError "Could not locate ""Assembly-CSharp.dll""" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
End If

' Check for files
For each file in arrPatches
    ' Check delta files
    If NOT objFso.FileExists("data\" & file(1) & ".xdelta") Then
        patcherError "Could not locate """ & file(1) & ".xdelta"""
    End If
    ' Check Cuphead files
    CurrentPatch = strCupheadDataDir & file(0) & file(1)
    If blnUnpatching Then
        If NOT objFso.FileExists(CurrentPatch & ".bak") Then
            patcherError "Could not locate """ & file(1) & ".bak""" & vbCrLf & "Patching cannot be reverted" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
        End If
    Else
        If NOT objFso.FileExists(CurrentPatch) Then
            patcherError "Could not locate """ & file(1) & """" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
        End If
    End If
Next

'strSaveLocation = objWshShl.ExpandEnvironmentStrings("%APPDATA%") & "\Cuphead\"
'If NOT (objFso.FolderExists(strSaveLocation)) Then
'    patcherError "Could not locate Cuphead save files"
'End If

' Last check before patching
intOKCancel = MsgBox("Click OK to " & strPatchMessage & "patch" & vbCrLf & vbCrLf & "(" & strCupheadDir & ")", vbOKCancel, "CupCore Patcher")
if intOKCancel = 2 Then
    WScript.Quit()
End If

' Patching gets done here
for each file in arrPatches
    CurrentPatch = strCupheadDataDir & file(0) & file(1)
    ' If .bak was found we are unpatching
    If blnUnpatching Then
        objFso.DeleteFile CurrentPatch
        objFso.MoveFile CurrentPatch & ".bak", CurrentPatch
    Else
        objFso.MoveFile CurrentPatch, CurrentPatch & ".bak"
        objWshShl.Run "data\xdelta3 -d -s """ & CurrentPatch & ".bak"" ""data\" & file(1) & ".xdelta"" """ & CurrentPatch & """", 0, True
    End If
Next
' Backup save files
' Will be added in full release

' Done patching

MsgBox "Files " & strPatchMessage & "patched successfully", 32, "CupCore Patcher"
' Patcher End
WScript.Quit
''''''''''''

Function patcherError(message)
    MsgBox message, 16, "CupCore Patcher Error"
    WScript.Quit()
End Function


Function verifyMd5(hash, filepath)
	' get binary data
	objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile filepath
    BinaryData = objStream.Read
    objStream.Close

    ' generate hash
    objMD5.ComputeHash_2(BinaryData)

    objElement.DataType = "bin.hex"
    objElement.NodeTypedValue = objMD5.Hash
    strMD5 = objElement.text

    ' return if hash matches
    verifyMd5 = (strMD5 = hash)
End Function
