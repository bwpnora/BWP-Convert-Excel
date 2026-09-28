' ==============================================================================
' BWPConvertTTNVN - Non-elevated Excel Add-in Installer
' Copyright (c) 2026 IT Leon. All rights reserved.
' ==============================================================================
Option Explicit

Dim fso, wshShell, objWMIService, colProcesses
Dim scriptDir, xlamSource, destDir, destPath, shaFile
Dim expectedHash, actualHash, isQuiet, arg

Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' Check for silent / quiet command line argument
isQuiet = False
For Each arg In WScript.Arguments
    If LCase(arg) = "/quiet" Or LCase(arg) = "/silent" Or LCase(arg) = "-quiet" Or LCase(arg) = "--quiet" Then
        isQuiet = True
    End If
Next

' ------------------------------------------------------------------------------
' 1. Running Excel Detection
' ------------------------------------------------------------------------------
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
If Err.Number = 0 Then
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    If colProcesses.Count > 0 Then
        If Not isQuiet Then
            MsgBox "Phat hien Microsoft Excel dang chay." & vbCrLf & vbCrLf & _
                   "Vui long luu cong viec va dong Excel truoc khi cai dat BWPConvertTTNVN.", _
                   vbExclamation, "BWPConvertTTNVN Setup"
        End If
        WScript.Quit 1
    End If
End If
On Error GoTo 0

' ------------------------------------------------------------------------------
' 2. Locate Source Add-in File
' ------------------------------------------------------------------------------
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlamSource = fso.BuildPath(scriptDir, "BWPConvertTTNVN.xlam")

' If running from installer/ subdirectory within repo, fallback to dist/
If Not fso.FileExists(xlamSource) Then
    Dim parentDir, distSource
    parentDir = fso.GetParentFolderName(scriptDir)
    distSource = fso.BuildPath(fso.BuildPath(parentDir, "dist"), "BWPConvertTTNVN.xlam")
    If fso.FileExists(distSource) Then
        xlamSource = distSource
        scriptDir = fso.GetParentFolderName(distSource)
    End If
End If

If Not fso.FileExists(xlamSource) Then
    If Not isQuiet Then
        MsgBox "Khong tim thay tap tin cai dat BWPConvertTTNVN.xlam tai:" & vbCrLf & xlamSource, _
               vbCritical, "BWPConvertTTNVN Setup"
    End If
    WScript.Quit 2
End If

' ------------------------------------------------------------------------------
' 3. Cryptographic SHA-256 Checksum Verification
' ------------------------------------------------------------------------------
shaFile = fso.BuildPath(scriptDir, "BWPConvertTTNVN-v1.0.0.sha256")
If Not fso.FileExists(shaFile) Then
    shaFile = fso.BuildPath(scriptDir, "BWPConvertTTNVN.sha256")
End If
If Not fso.FileExists(shaFile) Then
    ' Check for any .sha256 file in the directory
    Dim folderObj, fileObj
    Set folderObj = fso.GetFolder(scriptDir)
    For Each fileObj In folderObj.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "sha256" Then
            shaFile = fileObj.Path
            Exit For
        End If
    Next
End If

If fso.FileExists(shaFile) Then
    expectedHash = ExtractSha256FromFile(shaFile)
    If Len(expectedHash) = 64 Then
        actualHash = ComputeSha256(xlamSource)
        If LCase(expectedHash) <> LCase(actualHash) Then
            If Not isQuiet Then
                MsgBox "Loi xac thuc tinh toan ven (SHA-256 Checksum Mismatch)!" & vbCrLf & vbCrLf & _
                       "Tap tin cai dat co the bi hong hoac bi can thiep." & vbCrLf & _
                       "Ky vong: " & expectedHash & vbCrLf & _
                       "Thuc te:  " & actualHash, _
                       vbCritical, "BWPConvertTTNVN Setup"
            End If
            WScript.Quit 3
        End If
    End If
End If

' ------------------------------------------------------------------------------
' 4. Mark of the Web (MOTW) Removal from Source
' ------------------------------------------------------------------------------
RemoveMotw xlamSource

' ------------------------------------------------------------------------------
' 5. Copy Add-in to User's Excel AddIns Directory
' ------------------------------------------------------------------------------
Dim appData
appData = wshShell.ExpandEnvironmentStrings("%APPDATA%")
destDir = fso.BuildPath(appData, "Microsoft\AddIns")
If Not fso.FolderExists(destDir) Then
    CreateFolderRecursive destDir
End If

destPath = fso.BuildPath(destDir, "BWPConvertTTNVN.xlam")

On Error Resume Next
fso.CopyFile xlamSource, destPath, True
If Err.Number <> 0 Then
    Dim copyErr
    copyErr = Err.Description
    On Error GoTo 0
    If Not isQuiet Then
        MsgBox "Khong the sao chep tap tin add-in vao:" & vbCrLf & destPath & vbCrLf & vbCrLf & _
               "Loi: " & copyErr, vbCritical, "BWPConvertTTNVN Setup"
    End If
    WScript.Quit 4
End If
On Error GoTo 0

' Remove MOTW from destination file as well
RemoveMotw destPath

' ------------------------------------------------------------------------------
' 6. Idempotent Excel Registration
' ------------------------------------------------------------------------------
Dim xlApp, targetAddin, i, currentAddin
Dim regSucceeded
regSucceeded = False

On Error Resume Next
Set xlApp = CreateObject("Excel.Application")
If Err.Number = 0 And Not (xlApp Is Nothing) Then
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set targetAddin = Nothing
    For i = 1 To xlApp.AddIns.Count
        Set currentAddin = xlApp.AddIns(i)
        If UCase(currentAddin.Name) = "BWPCONVERTTTNVN.XLAM" Or UCase(currentAddin.Title) = "BWPCONVERTTTNVN" Then
            Set targetAddin = currentAddin
            Exit For
        End If
    Next

    If targetAddin Is Nothing Then
        Set targetAddin = xlApp.AddIns.Add(destPath, True)
    End If

    If Not (targetAddin Is Nothing) Then
        targetAddin.Installed = False
        targetAddin.Installed = True
        regSucceeded = True
    End If

    xlApp.Quit
    Set xlApp = Nothing
End If
On Error GoTo 0

If Not regSucceeded Then
    If Not isQuiet Then
        MsgBox "Canh bao: Khong the tu dong kich hoat add-in qua Excel COM." & vbCrLf & _
               "File da duoc copy vao: " & destPath & vbCrLf & _
               "Vui long mo Excel -> File -> Options -> Add-ins de kich hoat thu cong.", _
               vbExclamation, "BWPConvertTTNVN Setup"
    End If
    WScript.Quit 5
End If

' ------------------------------------------------------------------------------
' 7. User Confirmation
' ------------------------------------------------------------------------------
If Not isQuiet Then
    MsgBox "Cai dat BWPConvertTTNVN thanh cong!" & vbCrLf & vbCrLf & _
           "Vui long mo Excel de su dung tien ich tren the BWPConvertTTNVN.", _
           vbInformation, "BWPConvertTTNVN Setup"
End If

WScript.Quit 0


' ==============================================================================
' Helper Functions
' ==============================================================================

Function ExtractSha256FromFile(filePath)
    On Error Resume Next
    Dim ts, content, re, matches
    Set ts = fso.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        ExtractSha256FromFile = ""
        Exit Function
    End If
    content = ts.ReadAll()
    ts.Close
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[0-9a-fA-F]{64}"
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(content)
    If matches.Count > 0 Then
        ExtractSha256FromFile = LCase(matches(0).Value)
    Else
        ExtractSha256FromFile = ""
    End If
    On Error GoTo 0
End Function

Function ComputeSha256(targetFile)
    On Error Resume Next
    Dim tempFile, psCmd, runRes, ts, hashStr
    tempFile = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
    
    psCmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command " & _
            """(Get-FileHash -LiteralPath '" & Replace(targetFile, "'", "''") & "' -Algorithm SHA256).Hash.ToLower() | Out-File -FilePath '" & Replace(tempFile, "'", "''") & "' -Encoding ascii"""
    
    runRes = wshShell.Run(psCmd, 0, True)
    If runRes = 0 And fso.FileExists(tempFile) Then
        Set ts = fso.OpenTextFile(tempFile, 1, False)
        hashStr = Trim(ts.ReadLine())
        ts.Close
        fso.DeleteFile tempFile, True
        ComputeSha256 = LCase(hashStr)
    Else
        ComputeSha256 = ""
    End If
    On Error GoTo 0
End Function

Sub RemoveMotw(filePath)
    On Error Resume Next
    Dim psCmd
    psCmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command " & _
            """Unblock-File -LiteralPath '" & Replace(filePath, "'", "''") & "' -ErrorAction SilentlyContinue"""
    wshShell.Run psCmd, 0, True
    On Error GoTo 0
End Sub

Sub CreateFolderRecursive(folderPath)
    Dim parent
    If fso.FolderExists(folderPath) Then Exit Sub
    parent = fso.GetParentFolderName(folderPath)
    If parent <> "" And Not fso.FolderExists(parent) Then
        CreateFolderRecursive parent
    End If
    fso.CreateFolder folderPath
End Sub
