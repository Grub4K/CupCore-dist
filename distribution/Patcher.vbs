Option Explicit

Dim strRegValue, strFolder, objFolder, strCupheadDir, strCupheadDataDir, arrPatches, CurrentPatch, blnUnpatched, intOKCancel, file, BinaryData, strMD5
Dim objWshShl : Set objWshShl = CreateObject("WScript.Shell")
Dim objShl : Set objShl = CreateObject("Shell.Application")
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objMD5:  Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Dim objStream : Set objStream = CreateObject("ADODB.Stream")
Dim objXML : Set objXML = CreateObject("MSXML2.DOMDocument")
Dim objElement : Set objElement = objXML.CreateElement("tmp")

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

If Not objFso.FileExists(strCupheadDataDir & "Managed\Assembly-CSharp.dll") Then
	MsgBox "Assembly-CSharp.dll not found!", 16, "CupCore Patcher Error"
	WScript.Quit()
End If


If verifyMd5("e39a8a234edb59c07087a829de4fac34", strCupheadDataDir & "Managed\Assembly-CSharp.dll") = True Then 
	MsgBox "Cuphead Current Patch detected! Please install the LEGACY version.", 16, "CupCore Patcher Error"
	WScript.Quit()
	
ElseIf verifyMd5("dc51ec25ceb570b88afc6df0ca1601a1", strCupheadDataDir & "Managed\Assembly-CSharp.dll") = False Then 
	MsgBox "Invalid Cuphead Data Files! Please re-install the legacy version of Cuphead.", 16, "CupCore Patcher Error"
	WScript.Quit()
End If	

intOKCancel = MsgBox("Click OK to patch Cuphead to: " & vbCrLf & vbCrLf & strCupheadDir, vbOKCancel, "CupCore Patcher")

if intOKCancel = 2 Then
    WScript.Quit()
End If

' Patching array
arrPatches = Array(Array("Managed\", "Assembly-CSharp.dll"))

'
' Patching gets done here
'
for each file in arrPatches
    CurrentPatch = strCupheadDataDir & file(0) & file(1)
    
    If (objFso.FileExists(CurrentPatch & ".temp")) Then
        blnUnpatched = True
        objFso.DeleteFile CurrentPatch
        objFso.MoveFile CurrentPatch & ".temp", CurrentPatch
    Else
        objFso.MoveFile CurrentPatch, CurrentPatch & ".temp"
        objWshShl.Run "xdelta3 -d -s """ & CurrentPatch & ".temp"" " & file(1) & ".xdelta """ & CurrentPatch & """", 0, True
    	blnUnpatched = False
    End If
Next
' Done patching

if blnUnpatched = True Then
    MsgBox "Safely unpached", 32, "CupCore Patcher"
Else
    MsgBox "Files patched successfully", 32, "CupCore Patcher"
End If


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
    If strMD5 <> hash Then
    	verifyMd5 = False
    Else
    	verifyMd5 = True
    End If
End Function
