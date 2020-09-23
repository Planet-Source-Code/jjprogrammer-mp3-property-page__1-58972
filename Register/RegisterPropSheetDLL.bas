Attribute VB_Name = "RegisterPropSheetDLL"
'Each <Projectname>.PropertySheetHandler has its own entry (classID) in the registry,
'and it can change with each compile depending on the compatibility setting.
'This program sets the app extension's Property Sheet Handler to the correct classID
'and also creates .reg file
Sub Main()
SetContextMenuHandlerID ".mp3", "mp3 file", "MP3PropertySheet"
End Sub
Sub SetContextMenuHandlerID(AppExtension As String, AppExtensionName As String, ProjectName As String)
 Dim keyvalue As String

 keyvalue = GetString(HKEY_CLASSES_ROOT, AppExtension, "") 'check for default type
 If Not keyvalue > "" Then
    zRegistryCreateKey HKEY_CLASSES_ROOT, AppExtension, , AppExtensionName 'create if not present
 End If
 keyvalue = GetString(HKEY_CLASSES_ROOT, ProjectName & ".clsPropSheet\Clsid", "") 'get the classID for .clsPropSheet
 'set the app extension's Context Menu Handler to the correct classID
 If keyvalue > "" Then zRegistryCreateKey HKEY_CLASSES_ROOT, AppExtensionName & "\shellex\PropertySheetHandlers\" & ProjectName, , keyvalue
 MakeRegfile AppExtension, AppExtensionName, ProjectName, keyvalue
End Sub
Sub MakeRegfile(AppExtension, AppExtensionName As String, ProjectName As String, keyvalue As String)
    Dim appstring As String
    If Right$(App.Path, 1) <> "\" Then appstring = App.Path & "\" Else appstring = App.Path
    WritePrivateProfileString "HKEY_CLASSES_ROOT\" & AppExtension, "@", AppExtensionName, appstring & ProjectName & ".reg"
    WritePrivateProfileString "HKEY_CLASSES_ROOT\" & AppExtensionName & "\shellex\PropertySheetHandlers\" & ProjectName, "@", keyvalue, appstring & ProjectName & ".reg"
End Sub

