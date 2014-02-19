VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SubDB Subtitles Download/Upload"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   11730
   ClientWidth     =   10455
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   10455
   Begin VB.CommandButton btn_ViewLog 
      Caption         =   "&View Log"
      Height          =   390
      Left            =   5703
      TabIndex        =   7
      Top             =   4395
      Width           =   1725
   End
   Begin VB.CommandButton btn_About 
      Caption         =   "&About"
      Height          =   390
      Left            =   7530
      TabIndex        =   6
      Top             =   4395
      Width           =   1725
   End
   Begin VB.CommandButton btn_Download 
      Caption         =   "&Download"
      Height          =   390
      Left            =   225
      TabIndex        =   5
      Top             =   4395
      Width           =   1725
   End
   Begin VB.CommandButton btn_Options 
      Caption         =   "&Options"
      Height          =   390
      Left            =   3877
      TabIndex        =   4
      Top             =   4395
      Width           =   1725
   End
   Begin VB.CommandButton btn_Upload 
      Caption         =   "&Upload"
      Height          =   390
      Left            =   2051
      TabIndex        =   2
      Top             =   4395
      Width           =   1725
   End
   Begin VB.ListBox List1 
      Height          =   3570
      ItemData        =   "Form1.frx":058A
      Left            =   75
      List            =   "Form1.frx":058C
      TabIndex        =   1
      ToolTipText     =   "Drag and drop files here..."
      Top             =   645
      Width           =   10230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Drag files here"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   405
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SubDB Subtitle Download/Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2010
      TabIndex        =   0
      Top             =   90
      Width           =   3930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim userNme$, timerInterval$, userAgentString, messageAddedFileURL$, proxyHost$, proxyPort$, proxyPortInt, proxyHost2$, proxyPort2$, proxyPortInt2, notificationBlinker$, notificationSound$
Dim lastRefreshedFileContents_Previous$
Dim messageAddedFileContents_Previous$
Dim lastRefreshedList_Previous As Variant
Dim blinkerToggle As Boolean
Dim notifyRefreshed As Boolean
Dim notifyMessageAdded As Boolean
Dim currentBlinkerIcon As StdPicture
Dim monitoredFolder As String
Dim completedFolder As String
Dim messageMonitorInterval As String
Dim messageMonitorIntervalInt As String
Dim soundFile As String
Dim prevMessageCount As Integer
Dim currMessageCount As Integer
Dim TEMP_BATCH_FILE As String
Dim OPTION_FILE As String
Dim LOG_FILE_NAME As String

Private Sub btn_Hide_Click()
    Me.Hide
End Sub

Private Sub btn_About_Click()
    frm_About.Show vbModal
End Sub

Private Sub btn_Download_Click()
    downloadFilesFromScreen
End Sub

Private Sub btn_Options_Click()
    frm_Options.Show vbModal
End Sub

Private Sub btn_Upload_Click()
    'uploadFilesFromScreen
End Sub
Private Sub downloadFilesFromScreen()
    writeToLog ("download from screen initiated")
    Dim intX As Integer
    For intX = 0 To List1.ListCount - 1
        Dim listItem1 As Variant
        listItem1 = List1.List(intX)
        If (listItem1 = "") Then
            GoTo next_in_loop
        End If
        If (isDir(CStr(listItem1))) Then
            writeToLog (listItem1 & " is a folder, all files inside it will be tried")
            Call downloadSubsForFolder(CStr(listItem1))
        Else
            Call downloadOneFile(CStr(listItem1))
        End If
next_in_loop:
    Next
        
    List1.Clear
End Sub
Private Sub downloadSubsForFolder(folderName As String)
    Dim files() As String
    files = getFilesRecursive(folderName)
    Dim intX As Integer
    For intX = 0 To UBound(files)
        Dim listItem1 As String
        listItem1 = files(intX)
        If listItem1 <> "" Then
            Call downloadOneFile(CStr(listItem1))
            'Sleep 1000
        End If
    Next
End Sub
Private Sub downloadOneFile(filePath As String)
        
        'writeToLog ("trying : " & filePath)
        If (isDir(filePath) = True) Then
            writeToLog (filePath & " is a directory.  skipping")
            Exit Sub
        End If
        Dim extn As String
        extn = getFileExtention(filePath)
        If ( _
                (extn = "avi") _
             Or (extn = "mp4") _
             Or (extn = "mkv") _
             Or (extn = "vob") _
             Or (extn = "mpeg") _
             Or (extn = "mpg") _
        ) Then
            GoTo continue_flow
        Else
            writeToLog (filePath & " does not have a valid extention.  skipping")
            Exit Sub
        End If
continue_flow:
        If (FileLen(filePath) < 1000000) Then
            writeToLog (filePath & " file size is less than 1 MB.  skipping")
            Exit Sub
        End If
        Dim messageFound As Boolean
        Dim fileFound As Boolean
        fileFound = True
        If (messageFound Or fileFound) Then
            
            ''---------------------------------------------------------------------
            Dim f As Long
            Dim X(65535) As Byte  ' creates an array with index from 0 to 65535(65536 bytes)
            
            '  extract 64k bytes from begining of file
            'read begining of file
            f = FreeFile()
            Open filePath For Binary As #f
            Get #f, , X 'read it into an array
            Close #f
            
            Dim w As Long
            Dim tempPath  As String
            tempPath = "temp.bin"
            w = FreeFile()
            Open tempPath For Binary As #w 'open this file in write mode
            Put #w, , X 'write the begining part to temp file
            
            '  extract 64k bytes from end of file
            'read end part of file
            f = FreeFile()
            Open filePath For Binary As #f
            Dim fileSize As Long
            fileSize = FileLen(filePath)
            Dim filePointer As Long
            filePointer = fileSize - 65535
            Seek #f, filePointer
            Get #f, , X 'read it into an array
            Close #f
            
            Put #w, , X 'write the end part to temp file
            Close #w
            
            ''---------------------------------------------------------------------
            'generate the MD5 into a temp file 'temp.md5'
            writeToLog (filePath & ". computng MD5 - start")
            Open App.Path & "\" & TEMP_BATCH_FILE For Output As #2
            Print #2, "md5 " & tempPath & " > temp.md5" 'output to temp.md5
            Close #2
            ExecuteAndWait (App.Path & "\" & TEMP_BATCH_FILE)
            writeToLog (filePath & ". computng MD5 - end")
            ''---------------------------------------------------------------------
            ''---------------------------------------------------------------------
            ''---------------------------------------------------------------------
            
            
            
            'extract md5 value from temp file
            Dim md5FileNumber  As Integer
            md5FileNumber = FreeFile
            Open "temp.md5" For Input As md5FileNumber
            
            Dim md5OneLine As String
            Dim md5Value As String
            Do Until EOF(md5FileNumber)
                Line Input #md5FileNumber, md5OneLine
                md5OneLine = Trim(md5OneLine)
                md5Value = Left(md5OneLine, 32)
            Loop
                        writeToLog (filePath & ". md5Value = " & md5Value)



            Dim currentCommandString1 As String
            '
            '
            Dim proxyServerURL As String
            proxyServerURL = load_One_Value_From_INI_File("PROXY_SERVER")
            currentCommandString1 = "curl "
            If (proxyServerURL <> "") Then
                currentCommandString1 = currentCommandString1 & " --proxy " & proxyServerURL
            End If
            '
            Dim userAgent As String
            userAgent = load_One_Value_From_INI_File("USER_AGENT")
            currentCommandString1 = currentCommandString1 & " --user-agent """ & userAgent & """"
            '
            currentCommandString1 = currentCommandString1 & " --dump-header header.txt"
            currentCommandString1 = currentCommandString1 & " -o temp.srt"
            '
            Dim serverURL As String
            serverURL = load_One_Value_From_INI_File("SERVER_URL")
            serverURL = serverURL & "?action=download&hash=" & md5Value
            currentCommandString1 = currentCommandString1 & " """ & serverURL & """"
            '
            
            currentCommandString1 = Replace(currentCommandString1, "%", "%%") ' batch file needs double percentage symbol
            writeToLog (filePath & ". CommandString = " & currentCommandString1)
            Open App.Path & "\" & TEMP_BATCH_FILE For Output As #2
            Print #2, currentCommandString1
            Close #2
            ''---------------------------------------------------------------------
            ExecuteAndWait (App.Path & "\" & TEMP_BATCH_FILE)
            writeToLog (filePath & ". curl command execution complete ")
            ''---------------------------------------------------------------------
            ''---------------------------------------------------------------------
            ''---------------------------------------------------------------------
            ' now copy the downloaded SRT to the movie folder
            If (fileExists(App.Path & "\temp.srt") = False) Then
                ' if no file was downloaded, skip it...
                writeToLog (filePath & ". srt file was not downloaded. Skipping")
                Exit Sub
            End If
            Dim parentPath As String
            parentPath = getParentFolderName(filePath)
            Dim fileName As String
            Dim fileNameNoExtn As String
            fileName = getFilenameFromPath(filePath)
            fileNameNoExtn = Left(fileName, InStrRev(fileName, ".") - 1)
            Dim srtFileName As String
            srtFileName = fileNameNoExtn + ".srt"
            Call fileRename("temp.srt", srtFileName)
            writeToLog (filePath & ". srt file was renamed.")
            Call fileMoveToFolder(srtFileName, parentPath)
            writeToLog (filePath & ". srt file was moved to movie folder.")
        End If
End Sub
Private Sub btn_ViewLog_Click()
    launchFile (App.Path & "\" & LOG_FILE_NAME)
End Sub

Private Sub Form_Load()
    Form1.Left = Screen.Width \ 2 - Width \ 2
    Form1.Top = Screen.Height \ 2 - Height \ 2
    
    'a listbox control in the form and set its DragMode property to Automatic
    List1.DragMode = vbManual
    List1.OLEDragMode = vbAutomatic
    List1.OLEDropMode = vbAutomatic
    
    TEMP_BATCH_FILE = "temp_batch_file.bat"
    OPTION_FILE = "Message_Poster_Client.ini"
    LOG_FILE_NAME = "subdb.log.txt"
        

'********************************************************************************
'***    Read settings from options file     *************************************
'********************************************************************************
    Dim optionsFile$
    optionsFile$ = App.Path & "\" & "Message_Poster_Client.ini"
    
    userNme$ = load_One_Value_From_INI_File("USERNAME")
    timerInterval$ = load_One_Value_From_INI_File("MESSAGE_CHECK_TIMER_INTERVAL")
    userAgentString = load_One_Value_From_INI_File("USER_AGENT_STRING")
    messageAddedFileURL$ = load_One_Value_From_INI_File("MESSAGEADDEDFILE")
    notificationBlinker$ = load_One_Value_From_INI_File("NOTIFICATIONBLINKER")
    notificationSound$ = load_One_Value_From_INI_File("NOTIFICATIONSOUND")
    monitoredFolder = load_One_Value_From_INI_File("MONITORED_FOLDER")
    messageMonitorInterval = load_One_Value_From_INI_File("MESSAGE_MONITOR_INTERVAL")
    soundFile = load_One_Value_From_INI_File("SOUND_FILE")
    completedFolder = load_One_Value_From_INI_File("COMPLETED_FOLDER")

    Dim timerIntervalInt As Integer
    timerIntervalInt = getTimerIntervalInt(timerInterval$)
    messageMonitorIntervalInt = getTimerIntervalInt(messageMonitorInterval)
'********************************************************************************
'***    End of Read settings from options file     ******************************
'********************************************************************************


End Sub

Private Sub Form_Resize()
    Debug.Print Me.Height
    On Error GoTo EXIT_SUB
    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub
    If (Me.Height < 4155) Then
        Me.Height = 4155
        Exit Sub
    End If

    '************************************************************************
    Label1.Left = (Form1.Width / 2) - (Label1.Width / 2)
    
    btn_Download.Top = Me.Height - 900
    btn_Upload.Top = Me.Height - 900
    btn_Options.Top = Me.Height - 900
    btn_ViewLog.Top = Me.Height - 900
    btn_About.Top = Me.Height - 900

    '************************************************************************
    '************************************************************************
    List1.Width = Me.Width - 250
    List1.Height = Me.Height - 1600
    '************************************************************************
    '************************************************************************
EXIT_SUB:
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim ItemCount As Integer
    Dim i As Integer
    On Error Resume Next
    ItemCount = Data.files.Count
    
    For i = 1 To ItemCount
    List1.AddItem Data.files(i)
    Next
    
    If Err Then Err.Clear
End Sub
Private Sub aplicationExit()
    Unload Me
End Sub

Private Function getTimerIntervalInt(timerInterval$)
    Dim defaultTimer, returnTimer As Integer
    defaultTimer = 15
    On Error GoTo timerIntervalError
    If (timerInterval$ = "") Then
        returnTimer = defaultTimer
    Else
        returnTimer = val(timerInterval$)
    End If
    GoTo noTimerIntervalError
timerIntervalError:
    returnTimer = defaultTimer
noTimerIntervalError:
    getTimerIntervalInt = returnTimer
    End Function

'********************************************************************************
'***    End   *******************************************************************
'********************************************************************************
Public Function isDirectoryExists(dirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    isDirectoryExists = GetAttr(dirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

