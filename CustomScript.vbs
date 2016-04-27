Const CSIDL_COMMON_PROGRAMS = &H17
Const CSIDL_PROGRAMS = &H2
Const CSIDL_STARTMENU = &HB
Const CSIDL_PROFILE = &H28&

Dim dteWait
dteWait = DateAdd("s", 60, Now())
Do Until (Now() > dteWait)
Loop

dim linkFolder
Dim objShell, objFSO
Dim objCurrentUserStartFolder
Dim strCurrentUserStartFolderPath
Dim objAllUsersProgramsFolder
Dim strAllUsersProgramsPath

Dim colVerbs
Dim objVerb
dim objWSHShell
dim objFolder
dim objFolderItem

Set objWSHShell = CreateObject("Wscript.Shell")
Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.NameSpace(CSIDL_PROFILE)
Set linkFolder = objShell.NameSpace("C:\Users\" & objFolder & "\Links")
Set objFolderItem = objFolder.Self



'Delete Windows Personal Settings
dim tempfile
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("G:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\G Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("H:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\H Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("I:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\I Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("J:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\J Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("K:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\K Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If
set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("L:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\L Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("M:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\M Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("N:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\N Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("O:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\O Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("P:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\P Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("Q:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\Q Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("R:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\R Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("S:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\S Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("T:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\T Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("U:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\U Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("V:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\V Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("W:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\W Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

set tempfile=CreateObject("Scripting.FileSystemObject")
If tempfile.FolderExists("X:\") Then
tempfile.CopyFile "C:\Masters\DriveShortcuts\X Drive.lnk", "C:\Users\" & objFolder & "\Links\"
End If

dim filedel
	Set filedel = CreateObject("Scripting.FileSystemObject") 
If filedel.FileExists("C:\Users\" & objFolder & "\Links\Desktop.lnk") Then 
 	filedel.DeleteFile("C:\Users\" & objFolder & "\Links\Desktop.lnk")
End If 
If filedel.FileExists("C:\Users\" & objFolder & "\Links\Downloads.lnk") Then 
 	filedel.DeleteFile("C:\Users\" & objFolder & "\Links\Downloads.lnk")
End If 


Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCurrentUserStartFolder = objShell.NameSpace (CSIDL_STARTMENU)
strCurrentUserStartFolderPath = objCurrentUserStartFolder.Self.Path
Set objAllUsersProgramsFolder = objShell.NameSpace(CSIDL_COMMON_PROGRAMS)
strAllUsersProgramsPath = objAllUsersProgramsFolder.Self.Path

' - Remove From Taskbar -

'Windows Media Player
If objFSO.FileExists(strAllUsersProgramsPath & "\Windows Media Player.lnk") Then
    Set objFolder = objShell.Namespace(strAllUsersProgramsPath)
    Set objFolderItem = objFolder.ParseName("Windows Media Player.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Unpin from Taskbar" Then objVerb.DoIt
    Next
End If

'Explorer Default to Libraries
If objFSO.FileExists(strCurrentUserStartFolderPath & "\Programs\Accessories\Windows Explorer.lnk") Then
	Set objFolder = objShell.Namespace(strCurrentUserStartFolderPath & "\Programs\Accessories")
	Set objFolderItem = objFolder.ParseName("Windows Explorer.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Unpin from Taskbar" Then objVerb.DoIt
	Next
End If

' - Pin to Taskbar -

'Windows Explorer
If objFSO.FileExists("C:\Masters\DriveShortcuts\WindowsExComp.lnk") Then
	Set objFolder = objShell.Namespace("C:\Masters\DriveShortcuts")
	Set objFolderItem = objFolder.ParseName("WindowsExComp.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If

'Internet Explorer
If objFSO.FileExists(strCurrentUserStartFolderPath & "\Programs\Internet Explorer.lnk") Then
    Set objFolder = objShell.Namespace(strCurrentUserStartFolderPath & "\Programs")
    Set objFolderItem = objFolder.ParseName("Internet Explorer.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
    Next
End If

'OWA
If objFSO.FileExists("C:\Masters\DriveShortcuts\Outlook E-Mail.lnk") Then
    Set objFolder = objShell.Namespace("C:\Users\Public\Desktop")
    Set objFolderItem = objFolder.ParseName("Outlook E-Mail.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
    Next
End If

'Microsoft Word 2016
If objFSO.FileExists(strAllUsersProgramsPath & "\Word 2016.lnk") Then
	Set objFolder = objShell.Namespace(strAllUsersProgramsPath)
	Set objFolderItem = objFolder.ParseName("Word 2016.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If

'Microsoft Excel 2016
If objFSO.FileExists(strAllUsersProgramsPath & "\Excel 2016.lnk") Then
	Set objFolder = objShell.Namespace(strAllUsersProgramsPath)
	Set objFolderItem = objFolder.ParseName("Excel 2016.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If

'Microsoft PowerPoint 2016
If objFSO.FileExists(strAllUsersProgramsPath & "\PowerPoint 2016.lnk") Then
	Set objFolder = objShell.Namespace(strAllUsersProgramsPath)
	Set objFolderItem = objFolder.ParseName("PowerPoint 2016.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If


'Chrome
If objFSO.FileExists(strAllUsersProgramsPath & "\Google Chrome\Google Chrome.lnk") Then
	Set objFolder = objShell.Namespace(strAllUsersProgramsPath & "\Google Chrome")
	Set objFolderItem = objFolder.ParseName("Google Chrome.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If

'LogOff Button
If objFSO.FileExists("C:\windows\system32\logoff.exe") Then
	Set objFolder = objShell.Namespace("C:\Masters\DriveShortcuts")
	Set objFolderItem = objFolder.ParseName("Logoff.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If

'Service Desk
If objFSO.FileExists("C:\Masters\DriveShortcuts\ServiceDesk.lnk") Then
	Set objFolder = objShell.Namespace("C:\Masters\DriveShortcuts")
	Set objFolderItem = objFolder.ParseName("ServiceDesk.lnk")
	Set colVerbs = objFolderItem.Verbs
	For Each objVerb in colVerbs
		If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then objVerb.DoIt
	Next
End If


dim shelly
dim appDataLocation

set shelly = CreateObject("WScript.Shell")
appDataLocation = shelly.ExpandEnvironmentStrings("%APPDATA%")
Set filedel = CreateObject("Scripting.FileSystemObject") 


'Delete Script Link

If filedel.FileExists(appDataLocation & "\Microsoft\Windows\Start Menu\Programs\Startup\CustomScript.lnk") Then 
 	filedel.DeleteFile appDataLocation & "\Microsoft\Windows\Start Menu\Programs\Startup\CustomScript.lnk"
End If 