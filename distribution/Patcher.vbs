Option Explicit

If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
      , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

' Global constants for XML file
Const SETTINGS_PATH = "settings.xml"
Const CUPHEAD_PATH = "CupheadPath"
Const PATCHED = "PatchInstalled"
' Patching Array
Dim arrPatches : arrPatches = Array(_
    Array("Managed\Assembly-CSharp.dll",   "dc51ec25ceb570b88afc6df0ca1601a1"),_
    Array("sharedassets1.assets",          "bbd44f4eb1b9dbf62a858c807c5933b6"),_
	Array("sharedassets2.assets",          "ec7f96b925a643f2f2f35fb06436e781"),_
    Array("sharedassets3.assets",          "cede5a9ee9e0af64057ba60dfec2a0ea"),_
    Array("sharedassets10.assets",         "ff35ae46a3b9219e6e643ec50a9cf0cb"),_
	Array("sharedassets13.assets",          "d59795608681033dff90fada643f0f70"),_
	Array("sharedassets34.assets", "4ec8182652c5aa22b5582a1c1dd45ac5"),_
	Array("sharedassets17.assets",          "76cd1cd04e4a8ec5f12950a3f263aabe")_
)
' Save file base name array
Dim arrSaveFiles : arrSaveFiles = Array(_
    "cuphead_player_data_v1_slot_0.sav",_
    "cuphead_player_data_v1_slot_1.sav",_
    "cuphead_player_data_v1_slot_2.sav" _
)
' create Settings object
Dim Settings : Set Settings = (new XmlSettings)( SETTINGS_PATH )

Dim strRegValue, strFolder, objFolder, strPatchMessage, strSaveLocation, strMD5,_
    strSaveFile, arrSaveEndings, intOKCancel, file, CurrentFile, BinaryData
Dim objWshShl : Set objWshShl = CreateObject("WScript.Shell")
Dim objShl : Set objShl = CreateObject("Shell.Application")
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objMD5:  Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Dim objStream : Set objStream = CreateObject("ADODB.Stream")
Dim objElement : Set objElement = CreateObject("MSXML2.DOMDocument").CreateElement("tmp")
objWshShl.CurrentDirectory = objFso.GetParentFolderName(WScript.ScriptFullName)

' Check settings file first
If Settings.Load Then
    ' Check if everything legit
    Settings.CheckFiles
    If Err.Number <> 0 Then
        ' Maybe game has been restored?
        ' Check Patched bool first
        Settings.Patched = False
        ' Check if everything legit
        Settings.CheckFiles
        If Err.Number <> 0 Then
            ' Update Settings
            Settings.GetLocation
            Settings.GetPatched
            ' Check if everything legit
            Settings.CheckFiles
            If Err.Number <> 0 Then
                ' Raise Error, quit
                patcherError Err.Description
            End If
        End If
        ' Save settings
        Settings.Save
    End If
Else
    ' Update Settings
    Settings.GetLocation
    Settings.GetPatched
    ' Validate Files
    Settings.CheckFiles
    If Err.Number <> 0 Then
        ' Raise Error, quit
        patcherError Err.Description
    End If
    ' Save new settings
    Settings.Save
End If
If Settings.Patched Then
    strPatchMessage = "un"
    arrSaveEndings = Array(".core", ".bak")
Else
    strPatchMessage = ""
    arrSaveEndings = Array(".bak", ".core")
End If
strSaveLocation = objWshShl.ExpandEnvironmentStrings("%APPDATA%") & "\Cuphead\"
' Got locations

' Check for xdelta3.exe
If (NOT objFso.FileExists("data\xdelta3.exe")) Then
    patcherError "Could not locate xdelta3"
End If

' Last check before patching
intOKCancel = MsgBox("Click OK to " & strPatchMessage & "patch" & vbCrLf & vbCrLf & "(" & Settings.CupheadLocation & ")", vbOKCancel, "CupCore Patcher")
if intOKCancel = 2 Then
    WScript.Quit()
End If

' Patching gets done here
for each file in arrPatches
    CurrentFile = Settings.CupheadLocation & file(0)
    ' If flag is set we are unpatching
    If Settings.Patched Then
        objFso.DeleteFile CurrentFile
        objFso.MoveFile CurrentFile & ".bak", CurrentFile
    Else
        ' Check for external manipulation
        If objFso.FileExists(CurrentFile & ".bak") Then
            objFso.DeleteFile CurrentFile & ".bak"
        End If
        objFso.MoveFile CurrentFile, CurrentFile & ".bak"
        objWshShl.Run "data\xdelta3 -d -s """ & CurrentFile & ".bak"" ""data\" & file(0) & ".xdelta"" """ & CurrentFile & """", 0, True
    End If
Next
' Backup save files
If objFso.FolderExists(strSaveLocation) Then
    For each file in arrSaveFiles
        strSaveFile = strSaveLocation & file
        If (Settings.Patched) Then
            If objFso.FileExists(strSaveFile) Then
                objFso.DeleteFile strSaveFile
            End If
            If objFso.FileExists(strSaveFile & ".bak") Then
                objFso.MoveFile strSaveFile & ".bak", strSaveFile
            End If
        Else
            If objFso.FileExists(strSaveFile) Then
                objFso.CopyFile strSaveFile, strSaveFile & ".bak"
            End If
        End If
    Next
Else
    MsgBox "Saves could not be located, backups were not created", 32, "CupCore Patcher"
End If
'''''''''
' Old code
' If objFso.FolderExists(strSaveLocation) Then
'     For each file in arrSaveFiles
'         strSaveFile = strSaveLocation & file
'         ' Rename .sav to backups
'         If objFso.FileExists(strSaveFile) Then
'             objFso.MoveFile strSaveFile, strSaveFile & arrSaveEndings(0)
'         End If
'         ' Rename backups to .sav
'         If objFso.FileExists(strSaveFile & arrSaveEndings(1)) Then
'             objFso.MoveFile strSaveFile & arrSaveEndings(1), strSaveFile
'         End If
'     Next
' Else
'     MsgBox "Saves could not be located, backups were not created", 32, "CupCore Patcher"
' End If



' Change Patched variable and save
Settings.Patched = not Settings.Patched
Settings.Save
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

Class XmlSettings
    Private strSettingsFile, strCupheadPath, blnPatched
    Public Default Function Init( strSettingFileArg )
        strSettingsFile = strSettingFileArg
        Set Init = Me
    End Function

    Public Property Let CupheadLocation(byVal value)
        strCupheadPath = value
    End Property

    Public Property Get CupheadLocation
        CupheadLocation = strCupheadPath
    End Property

    Public Property Let Patched(byVal value)
        blnPatched = value
    End Property

    Public Property Get Patched
        Patched = blnPatched
    End Property

    ' Loading helper function
    Private Function GetSetting(byRef objXML, byVal strTag)
        Dim objNode
        For each objNode in objXML.getElementsByTagName(strTag)
            GetSetting = objNode.Text
            Exit Function
        Next
        GetSetting = Null
    End Function

    ' Load the Settings file into class members
    Public Function Load()
        Load = false
        Dim objXML : Set objXML = CreateObject("Microsoft.XMLDOM")
        objXML.async = false
        If Not objXML.load( strSettingsFile ) Then
            Exit Function
        End If
        strCupheadPath = GetSetting(objXML, CUPHEAD_PATH)
        blnPatched = GetSetting(objXML, PATCHED)
        If IsNull(strCupheadPath) Or IsNull(blnPatched) Or _
            IsEmpty(strCupheadPath) Or IsEmpty(blnPatched) Then
            Exit Function
        End If
        ' Postprocess variables
        blnPatched = CBool( blnPatched )
        Load = true
    End Function

    ' Writes class member settings into settings file
    Public Function Save()
        Save = false
        Dim objSettings, objPath, objBool
        Dim objXML : Set objXML = CreateObject("Microsoft.XMLDOM")
        Set objSettings = objXML.createElement("Settings")
        Set objPath = objXML.createElement(CUPHEAD_PATH)
        Set objBool = objXML.createElement(PATCHED)
        objPath.Text = strCupheadPath
        objBool.Text = blnPatched
        objXML.appendChild objSettings
        objSettings.appendChild objPath
        objSettings.appendChild objBool
        Save = objXML.Save( strSettingsFile )
    End Function

    Public Sub GetLocation
        ' Find Cuphead location.
        ' If cannot find, let select manually
        On Error Resume Next
        strRegValue = objWshShl.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 268910\InstallLocation")
        If Not objFso.FolderExists(strRegValue & "\Cuphead_Data\") Then
            strRegValue = ""
        End If
        If len(strRegValue) = 0 or Err.Number <> 0 Then
            Set objFolder = objShl.BrowseForFolder(0,"Cuphead not found, please select location manually.",0,17)

            If objFolder is Nothing Then
                Wscript.Quit()
            Else
                If Not objFso.FileExists(objFolder.Self.Path & "\Cuphead.exe") Then
                    If Not objFso.FileExists(objFolder.Self.Path & "\..\Cuphead.exe") Then
        		        patcherError "Cuphead executable not found!"
                    Else
                        strCupheadPath = objFolder.Self.Path & "\"
                    End If
                Else
                    strCupheadPath = objFolder.Self.Path & "\Cuphead_Data\"
                End If
            End If
        Else
            strCupheadPath = strRegValue & "\Cuphead_Data\"
        End If
    End Sub

    Public Sub GetPatched
        ' Check Assembly-CSharp.dll as significant file
        If ( objFso.FileExists(strCupheadPath & "Managed\Assembly-CSharp.dll" & ".bak") ) Then
            blnPatched = True
        ElseIf ( objFso.FileExists(strCupheadPath & "Managed\Assembly-CSharp.dll") ) Then
            ' Check for 1.1
            If verifyMd5("e39a8a234edb59c07087a829de4fac34", strCupheadPath & "Managed\Assembly-CSharp.dll") Then
                patcherError "Cuphead v1.1 detected! Please install the LEGACY version."
			ElseIf verifyMd5("bdebd14be8a36c516c37d7930697d185", strCupheadPath & "Managed\Assembly-CSharp.dll") Then
                patcherError "Cuphead v1.2 detected! Please install the LEGACY version."
            End If
            blnPatched = False
        Else
            ' Significant file is missing
            patcherError "Could not locate ""Assembly-CSharp.dll""" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
        End If
    End Sub

    Sub CheckFiles
        ' ************ ERRORS ************
        '  1 - xdelta not found
        '  2 - bak not found, reinstall
        '  3 - md5 failed, reinstall
        '  4 - file not found, reinstall
        ' ********************************
        On Error Resume Next
        Err.Clear
        Dim CurrentFile, file
        For each file in arrPatches
            ' Check delta files
            If NOT objFso.FileExists("data\" & file(0) & ".xdelta") Then
                Err.Raise 1,, "Could not locate ""data\" & file(0) & ".xdelta"""
            End If
            ' Check Cuphead files
            CurrentFile = Settings.CupheadLocation & file(0)
            If Settings.Patched Then
                ' For unpatching we need the backup
                If NOT objFso.FileExists(CurrentFile & ".bak") Then
                    Err.Raise 2,, "Could not locate """ & file(0) & ".bak""" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
                End If
            Else
                ' Verifying md5, since xdelta will throw an error if file not matching
                If NOT verifyMd5( file(1), CurrentFile ) Then
                    Err.Raise 3,, "Could not verify """ & file(0) & """" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
                End If
            End If
            ' Lethal file not found, cannot patch
            If NOT objFso.FileExists(CurrentFile) Then
                Err.Raise 4,, "Could not locate """ & file(0) & """" & vbCrLf & vbCrLf & "Please reinstall Cuphead"
            End If
        Next
    End Sub
End Class
