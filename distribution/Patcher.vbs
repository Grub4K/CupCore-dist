Option Explicit

If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
      , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

Dim strRegValue, strFolder, objFolder, strCupheadDir, strCupheadDataDir,_
    arrPatches, CurrentPatch, blnUnpatching, strPatchMessage, strSaveLocation,_
    strSaveFile, arrSaveFiles, arrSaveEndings, intOKCancel, file, BinaryData, strMD5
Dim objWshShl : Set objWshShl = CreateObject("WScript.Shell")
Dim objShl : Set objShl = CreateObject("Shell.Application")
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objMD5:  Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Dim objStream : Set objStream = CreateObject("ADODB.Stream")
Dim objXML : Set objXML = CreateObject("MSXML2.DOMDocument")
Dim objElement : Set objElement = objXML.CreateElement("tmp")
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
strSaveLocation = objWshShl.ExpandEnvironmentStrings("%APPDATA%") & "\Cuphead\"
' Got location



' Patching array
arrPatches = Array(_
    Array("Managed\", "Assembly-CSharp.dll",   "dc51ec25ceb570b88afc6df0ca1601a1"),_
    Array("",         "sharedassets1.assets",  "bbd44f4eb1b9dbf62a858c807c5933b6"),_
    Array("",         "sharedassets3.assets",  "cede5a9ee9e0af64057ba60dfec2a0ea"),_
    Array("",         "sharedassets10.assets", "ff35ae46a3b9219e6e643ec50a9cf0cb") _
)
' Save file base name array
arrSaveFiles = Array(_
    "cuphead_player_data_v1_slot_0.sav",_
    "cuphead_player_data_v1_slot_1.sav",_
    "cuphead_player_data_v1_slot_2.sav" _
)
' Check for xdelta3.exe
If (NOT objFso.FileExists("data\xdelta3.exe")) Then
    patcherError "Could not locate xdelta3"
End If
' Check Assembly-CSharp.dll as significant file
If ( objFso.FileExists(strCupheadDataDir & "Managed\Assembly-CSharp.dll" & ".bak") ) Then
    ' Set some values for later
    blnUnpatching = True
    strPatchMessage = "un"
    arrSaveEndings = Array(".core", ".bak")
ElseIf ( objFso.FileExists(strCupheadDataDir & "Managed\Assembly-CSharp.dll") ) Then
    ' Check for Current Patch
    If verifyMd5("e39a8a234edb59c07087a829de4fac34", strCupheadDataDir & "Managed\Assembly-CSharp.dll") Then
        patcherError "Cuphead Current Patch detected! Please install the LEGACY version."
    End If
    ' Set some values for later
    blnUnpatching = False
    strPatchMessage = ""
    arrSaveEndings = Array(".bak", ".core")
Else
    ' Lethal file is missing
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
        ' For unpatching we need the backup
        If NOT objFso.FileExists(CurrentPatch & ".bak") Then
            patcherError "Could not locate """ & file(1) & ".bak""" & vbCrLf & "Patching cannot be reverted" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
        End If
    Else
        ' Verifying md5, since xdelta will throw an error if file not matching
        If NOT verifyMd5( file(2), CurrentPatch ) Then
            patcherError "Could not verify """ & file(1) & """" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
        End If
    End If
    ' Lethal file not found, cannot patch
    If NOT objFso.FileExists(CurrentPatch) Then
        patcherError "Could not locate """ & file(1) & """" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
    End If
Next



' Last check before patching
intOKCancel = MsgBox("Click OK to " & strPatchMessage & "patch" & vbCrLf & vbCrLf & "(" & strCupheadDir & ")", vbOKCancel, "CupCore Patcher")
if intOKCancel = 2 Then
    WScript.Quit()
End If



' Patching gets done here
for each file in arrPatches
    CurrentPatch = strCupheadDataDir & file(0) & file(1)
    ' If flag is set we are unpatching
    If blnUnpatching Then
        objFso.DeleteFile CurrentPatch
        objFso.MoveFile CurrentPatch & ".bak", CurrentPatch
    Else
        ' Check for external manipulation
        If objFso.FileExists(CurrentPatch & ".bak") Then
            objFso.DeleteFile CurrentPatch & ".bak"
        End If
        objFso.MoveFile CurrentPatch, CurrentPatch & ".bak"
        objWshShl.Run "data\xdelta3 -d -s """ & CurrentPatch & ".bak"" ""data\" & file(1) & ".xdelta"" """ & CurrentPatch & """", 0, True
    End If
Next
' Backup save files
If objFso.FolderExists(strSaveLocation) Then
    For each file in arrSaveFiles
        strSaveFile = strSaveLocation & file
        ' Rename .sav to backups
        If objFso.FileExists(strSaveFile) Then
            objFso.MoveFile strSaveFile, strSaveFile & arrSaveEndings(0)
        End If
        ' Rename backups to .sav
        If objFso.FileExists(strSaveFile & arrSaveEndings(1)) Then
            objFso.MoveFile strSaveFile & arrSaveEndings(1), strSaveFile
        End If
    Next
Else
    MsgBox "Saves could not be located, backups were not created", 32, "CupCore Patcher"
End If
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
