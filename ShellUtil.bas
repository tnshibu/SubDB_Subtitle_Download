Attribute VB_Name = "ShellUtil"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function getSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const SYNCHRONIZE       As Long = &H100000

Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const LANG_NEUTRAL As Long = &H0&
Private Declare Function FormatMessageA _
   Lib "kernel32" _
   ( _
       ByVal dwFlags As Long, _
       ByRef lpSource As Any, _
       ByVal dwMessageId As Long, _
       ByVal dwLanguageId As Long, _
       ByVal lpBuffer As String, _
       ByVal nSize As Long, _
       ByRef Arguments As Long _
   ) As Long
''''''''''''''''''''''''
Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SE_ERR_ACCESSDENIED = 5            '  access denied
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8                     '  out of memory
Public Const SE_ERR_SHARE = 26

Public Const STYLE_NORMAL = 11


Global Const NORMAL_PRIORITY_CLASS = &H20&
Global Const INFINITE = -1&
'Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

''''''''''''''''''''''''
Public Sub ExecuteAndNoWait(cmdline As String)
    'Call Shell_Command(cmdline, vbNormalFocus, "", 2000000000)
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim X As Long
    NameStart.cb = Len(NameStart)
    NameStart.dwFlags = 1         'STARTF_USESHOWWINDOW
    NameStart.wShowWindow = 0     'SW_HIDE
    X = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
    'X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    'X = CloseHandle(NameOfProc.hProcess)
End Sub
Public Sub ExecuteAndWait(cmdline As String)
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim X As Long
    NameStart.cb = Len(NameStart)
    NameStart.dwFlags = 1         'STARTF_USESHOWWINDOW
    NameStart.wShowWindow = 0     'SW_HIDE
    X = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
    X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    X = CloseHandle(NameOfProc.hProcess)
End Sub

Public Function FormatMessage(ByRef lErrorNumber As Long) As String
    'this function is only for debugging purpose
    FormatMessage = Space$(255)
    Call FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrorNumber, LANG_NEUTRAL, FormatMessage, 255, ByVal 0&)
    FormatMessage = Left$(FormatMessage, InStr(FormatMessage, vbNewLine) - 1)
End Function
Public Function ShellAndWaitForTermination_X( _
        sShell As String, _
        Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, _
        Optional ByRef sError As String, _
        Optional ByVal lTimeOut As Long = 2000000000 _
    ) As Boolean
Dim hProcess As Long
Dim lR As Long
Dim lTimeStart As Long
Dim bSuccess As Boolean
    
On Error GoTo ShellAndWaitForTerminationError
    
    ' This is v2 which is somewhat more reliable:
    'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ShellExecute(0&, vbNullString, sShell, vbNullString, vbNullString, vbHide))
    Dim procID As Long
    procID = ShellExecute(0&, vbNullString, sShell, vbNullString, vbNullString, vbHide)
    hProcess = OpenProcess(SYNCHRONIZE, 0, procID)
    If (hProcess = 0) Then
        'MsgBox FormatMessage(Err.LastDllError)
        sError = "This program could not determine whether the process started." & _
             "Please watch the program and check it completes."
        ' Only fail if there is an error - this can happen
        ' when the program completes too quickly.
    Else
        bSuccess = True
        lTimeStart = timeGetTime()
        Do
            ' Get the status of the process
            GetExitCodeProcess hProcess, lR
            ' Sleep during wait to ensure the other process gets
            ' processor slice:
            DoEvents: Sleep 100
            If (timeGetTime() - lTimeStart > lTimeOut) Then
                ' Too long!
                sError = "The process has timed out."
                lR = 0
                bSuccess = False
            End If
        Loop While lR = STILL_ACTIVE
    End If
    ShellAndWaitForTermination = bSuccess
        
    Exit Function

ShellAndWaitForTerminationError:
    sError = Err.Description
    Exit Function
End Function

Public Function Shell_Command( _
        sShell As String, _
        Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, _
        Optional ByRef sError As String, _
        Optional ByVal lTimeOut As Long = 2000000000 _
    ) As Boolean
Dim hProcess As Long
Dim lR As Long
Dim lTimeStart As Long
Dim bSuccess As Boolean
    
On Error GoTo Shell_CommandError
    
    ' This is v2 which is somewhat more reliable:
    'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ShellExecute(0&, vbNullString, sShell, vbNullString, vbNullString, vbHide))
    Dim procID As Long
    procID = ShellExecute(0&, vbNullString, sShell, vbNullString, vbNullString, vbHide)
    hProcess = OpenProcess(SYNCHRONIZE, 0, procID)
    If (hProcess = 0) Then
        'MsgBox FormatMessage(Err.LastDllError)
        sError = "This program could not determine whether the process started." & _
             "Please watch the program and check it completes."
        ' Only fail if there is an error - this can happen
        ' when the program completes too quickly.
    End If
    Shell_Command = bSuccess
        
    Exit Function

Shell_CommandError:
    sError = Err.Description
    Exit Function
End Function

Public Function launchFile(fileName As String)
    ShellExecute hwnd, "open", fileName, vbNullString, vbNullString, SW_SHOWNORMAL
End Function

