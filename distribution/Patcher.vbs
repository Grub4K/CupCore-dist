Option Explicit

If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
      , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

Dim strRegValue, strFolder, objFolder, strCupheadDir, strCupheadDataDir, arrPatches, CurrentPatch, blnUnpatched, intOKCancel, file, BinaryData, strMD5
Dim objWshShl : Set objWshShl = CreateObject("WScript.Shell")
Dim objShl : Set objShl = CreateObject("Shell.Application")
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objMD5:  Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Dim objStream : Set objStream = CreateObject("ADODB.Stream")
Dim objXML : Set objXML = CreateObject("MSXML2.DOMDocument")
Dim objElement : Set objElement = objXML.CreateElement("tmp")

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
            MsgBox "Cuphead_Data not found!", 16, "CupCore Patcher Error"
            WScript.Quit()
        Else
            strCupheadDir = objFolder.Self.Path
        End If
    End If
Else
    strCupheadDir = strRegValue
End If
strCupheadDataDir = strCupheadDir & "\Cuphead_Data\"
' Got location

' Check for Current Patch (not working for now)
If verifyMd5("e39a8a234edb59c07087a829de4fac34", strCupheadDataDir & "Managed\Assembly-CSharp.dll") Then
    patcherError "Cuphead Current Patch detected! Please install the LEGACY version."
End If

' Last check before patching
intOKCancel = MsgBox("Click OK to patch Cuphead to: " & vbCrLf & vbCrLf & strCupheadDir, vbOKCancel, "CupCore Patcher")
if intOKCancel = 2 Then
    WScript.Quit()
End If

' Patching array
arrPatches = Array(_
    Array("Managed\", "Assembly-CSharp.dll"),_
    Array("", "sharedassets1.assets"),_
    Array("", "sharedassets3.assets"),_
    Array("", "sharedassets10.assets") _
)

' Patching gets done here
for each file in arrPatches
    CurrentPatch = strCupheadDataDir & file(0) & file(1)
    ' If .bak was found we are unpatching
    If (objFso.FileExists(CurrentPatch & ".bak")) Then
        blnUnpatched = True
        objFso.DeleteFile CurrentPatch
        objFso.MoveFile CurrentPatch & ".bak", CurrentPatch
    ElseIf (objFso.FileExists(CurrentPatch)) Then
        objFso.MoveFile CurrentPatch, CurrentPatch & ".bak"
        objWshShl.Run "xdelta3 -d -s """ & CurrentPatch & ".bak"" " & file(1) & ".xdelta """ & CurrentPatch & """", 0, True
    	blnUnpatched = False
    Else
        patcherError """" & file(1) & """ not found"
    End If
Next
' Done patching

if blnUnpatched = True Then
    MsgBox "Safely unpached", 32, "CupCore Patcher"
Else
    MsgBox "Files patched successfully", 32, "CupCore Patcher"
End If


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
