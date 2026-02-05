Option Explicit

Dim fso, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' ========================
' Base directory (where the VBS lives)
' ========================

Dim baseDir
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)

' ========================
' App settings
' ========================

Dim appName
appName = "Yield Dashboard"

Dim iconSource
iconSource = baseDir & "\favicon.ico"

Dim htmlTarget
htmlTarget = baseDir & "\index.html"   ' HTML in same folder

' ========================
' Validate files
' ========================

If Not fso.FileExists(iconSource) Then
    MsgBox "Icon file (app.ico) not found.", vbCritical
    WScript.Quit
End If

If Not fso.FileExists(htmlTarget) Then
    MsgBox "HTML file not found.", vbCritical
    WScript.Quit
End If

' ========================
' Install icon
' ========================

Dim iconDir, iconDest
iconDir = shell.ExpandEnvironmentStrings("%ProgramData%") & "\Icons"

If Not fso.FolderExists(iconDir) Then
    fso.CreateFolder iconDir
End If

iconDest = iconDir & "\" & appName & ".ico"
fso.CopyFile iconSource, iconDest, True

' ========================
' Create desktop shortcut
' ========================

Dim desktopPath
desktopPath = shell.SpecialFolders("Desktop")

Dim shortcutPath
shortcutPath = desktopPath & "\" & appName & ".lnk"

Dim lnk
Set lnk = shell.CreateShortcut(shortcutPath)

lnk.TargetPath = htmlTarget
lnk.WorkingDirectory = baseDir
lnk.IconLocation = iconDest
lnk.WindowStyle = 1
lnk.Save

MsgBox appName & " installed successfully!", vbInformation
