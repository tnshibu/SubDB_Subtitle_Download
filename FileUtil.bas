Attribute VB_Name = "FileUtil"
Option Explicit
'
Private Type FILETIME
    dwLowDate As Long
    dwHighDate As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMillisecs As Integer
End Type

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public fileUtilCancelFlag As Boolean
'usage Call SetFileDate("c:\autoexec.bat", "01/01/00 12:30 pm")

Public Sub SetFileDate(sFileName As String, sDate As String)

    Dim hFile As Long
    Dim lResult As Long
    Dim udtSysTime As SYSTEMTIME
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    
    With udtSysTime
        .wYear = Year(sDate)
        .wMonth = Month(sDate)
        .wDay = Day(sDate)
        .wDayOfWeek = Weekday(sDate) - 1
        .wHour = Hour(sDate)
        .wMinute = Minute(sDate)
        .wSecond = Second(sDate)
    End With
    
    lResult = SystemTimeToFileTime(udtSysTime, udtLocalTime)
    lResult = LocalFileTimeToFileTime(udtLocalTime, udtFileTime)
    
    hFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    lResult = SetFileTime(hFile, ByVal 0&, ByVal 0&, udtFileTime)

    Call CloseHandle(hFile)

End Sub



Public Function getFilesRecursive(ByVal FullPath As String) As String()

Dim oFs As New FileSystemObject
Dim sAns() As String
Dim oFolder As Folder
Dim oFolder2 As Folder
Dim oFile2 As Folder
Dim oFile As file
Dim lElement As Long
Dim i As Integer
'''''''''''''''''''''''
If (InStr(FullPath, "target") > 0) Then
    getFilesRecursive = sAns
    Exit Function
End If
DoEvents
If (fileUtilCancelFlag = True) Then GoTo STOP_INDEXING
On Error GoTo STOP_INDEXING

'''''''''''''''''''''''
FullPath = Trim(FullPath)
'get last char
Dim last_char As String
last_char = Right(FullPath, 1)
If (last_char = "\") Then
    FullPath = Left(FullPath, Len(FullPath) - 1) ' then remove the last back slash
End If


ReDim sAns(0) As String
If oFs.FolderExists(FullPath) Then
    Set oFolder = oFs.GetFolder(FullPath)
    For Each oFile In oFolder.Files
        lElement = IIf(sAns(0) = "", 0, lElement + 1)
        ReDim Preserve sAns(lElement) As String
        sAns(lElement) = FullPath & "\" & oFile.Name
    Next
    For Each oFile2 In oFolder.SubFolders
        If (oFile2.Name = ".svn") Then
            GoTo end_of_for
        End If
        lElement = lElement + 1
        ReDim Preserve sAns(lElement) As String
        sAns(lElement) = FullPath & "\" & oFile2.Name 'first add the subfolder name itself
        
        Dim filesInSubDir() As String
        filesInSubDir = getFilesRecursive(FullPath & "\" & oFile2.Name)
        If (fileUtilCancelFlag = True) Then GoTo STOP_INDEXING
        
        ReDim Preserve sAns(lElement + UBound(filesInSubDir) + 1) As String
        For i = LBound(filesInSubDir) To UBound(filesInSubDir)
            lElement = lElement + 1
            sAns(lElement) = filesInSubDir(i)
        Next i
end_of_for:
    Next
End If
STOP_INDEXING:

ErrHandler:
Set oFs = Nothing
Set oFolder = Nothing
Set oFile = Nothing

getFilesRecursive = sAns
End Function

'
'Private Function getAllFileNamesInCurrentDirectory(dir As String) As String()
'    Dim arr() As String
'    ReDim arr(0)
'    Dim i As Integer
'    Dim fso As New FileSystemObject
'    Dim fld As Folder
'    Dim fil As File
'    Set fld = fso.GetFolder(dir)
'    i = 0
'    For Each fil In fld.Files
'        ReDim Preserve arr(0 To i)
'        'arr(i) = fil.Name
'        arr = addToArray(arr, fil.Name)
'        i = i + 1
'        'Debug.Print fil.Name
'    Next
'    Set fil = Nothing
'    Set fld = Nothing
'    Set fso = Nothing
'    getAllFileNamesInDirectory = arr
'End Function
Function isDir(dirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    isDir = GetAttr(dirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
Public Function getFileSize(fileName As String) As Long
    Dim fso As New FileSystemObject
    Dim f As file
    'Get a reference to the File object.
    If fso.fileExists(fileName) Then
        Set f = fso.GetFile(fileName)
        getFileSize = f.Size
    Else
        MsgBox "File not found"
    End If
End Function
Public Sub readFileContentsToArray(ByVal fileName As String, ByRef TheArray As Variant)
'PURPOSE:    Puts all lines of file into a string array
'PARAMETERS: FileName = FullPath of File
'            TheArray = StringArray to which contents
'                       Of File will be added.
'Example
'  Dim sArray() as String
'  FileToArray "C:\MyTextFile.txt", sArray
'  For lCtr = 0 to Ubound(sArray)
'  Debug.Print sArray(lCtr)
'  Next

'NOTES:
'  --  Requires a reference to Microsoft Scripting Runtime
'      Library
'  --  You can write this method in a number of different ways
'      For instance, you can take advantage of VB 6's ability to
'      return an array.
' --   You can also read all the contents of the file and use the
'      Split function with vbCrlf as the delimiter, but I
'      wanted to illustrate use of the ReadLine
'      and AtEndOfStream methods.
'**********************************************************

  Dim oFSO As New FileSystemObject
  Dim oFSTR As Scripting.TextStream
  Dim ret As Long
  Dim lCtr As Long

  If Dir(fileName) = "" Then Exit Sub

'Check if string array was passed
'If you want to permit other type of arrays (e.g.,
'variant) remove or modify this line
If VarType(TheArray) <> vbArray + vbString Then Exit Sub
  
  On Error GoTo ErrorHandler
     Set oFSTR = oFSO.OpenTextFile(fileName)
      
     Do While Not oFSTR.AtEndOfStream
            ReDim Preserve TheArray(lCtr) As String
            TheArray(lCtr) = oFSTR.ReadLine
            lCtr = lCtr + 1
            DoEvents 'optional but with large file
                     'program will appear to hang
                     'without it
    Loop
     oFSTR.Close
     
ErrorHandler:
     Set oFSTR = Nothing
End Sub

Public Function fileExists(file As String) As Boolean
    fileExists = False
    If Dir(file) <> "" Then
      fileExists = True
    Else
      fileExists = False
    End If
End Function
Function getFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        getFilenameFromPath = getFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Function getParentFolderName(strFullFilePath As String) As String
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  getParentFolderName = fso.getParentFolderName(strFullFilePath)
End Function
Function fileRename(sSourceFile As String, sDestinationFile As String)
   Dim oFSO As FileSystemObject
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   oFSO.MoveFile sSourceFile, sDestinationFile
   Set oFSO = Nothing
End Function
Function fileMoveToFolder(sSourceFile As String, sDestinationFolder As String)
   If (sSourceFile = "") Then
        Exit Function
   End If
   If (sDestinationFolder = "") Then
        Exit Function
   End If
   Dim oFSO As FileSystemObject
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   oFSO.MoveFile sSourceFile, sDestinationFolder & "\" & getFilenameFromPath(sSourceFile)
   Set oFSO = Nothing
End Function
Function getFileExtention(sSourceFile As String) As String
    Dim arr() As String ' string array
    arr = Split(sSourceFile, ".")
    getFileExtention = arr(UBound(arr))
End Function
