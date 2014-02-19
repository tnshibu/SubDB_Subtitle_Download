VERSION 5.00
Begin VB.Form frm_Options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SubDB Downloader - Options"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Sleep 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1350
      Width           =   6255
   End
   Begin VB.TextBox txt_Proxy 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.TextBox txt_User_Agent 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   975
      Width           =   6255
   End
   Begin VB.TextBox txt_ServerURL 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   225
      Width           =   6255
   End
   Begin VB.CommandButton btn_Close 
      Caption         =   "C&ancel"
      Height          =   420
      Left            =   4860
      TabIndex        =   5
      Top             =   1845
      Width           =   1455
   End
   Begin VB.CommandButton btn_Save 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   420
      Left            =   2550
      TabIndex        =   4
      Top             =   1845
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pause between requests :"
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   1350
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Proxy :"
      Height          =   195
      Left            =   1515
      TabIndex        =   8
      Top             =   622
      Width           =   480
   End
   Begin VB.Label lblUserAgent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Agent String :"
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   990
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server URL :"
      Height          =   195
      Left            =   1065
      TabIndex        =   6
      Top             =   240
      Width           =   930
   End
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Close_Click()
    Unload Me
End Sub
Private Sub btn_Save_Click()
    Call save_Values_To_INI_File
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    'a listbox control in the form and set its DragMode property to Automatic

    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Call load_Values_From_INI_File
End Sub
Private Sub Form_Resize()
    Call ResizeControls
End Sub
Private Sub ResizeControls()
    If WindowState = vbMinimized Then Exit Sub

    btn_Save.Top = Me.Height - 1000
    btn_Close.Top = Me.Height - 1000
End Sub
Private Sub load_Values_From_INI_File()
    Dim val As String
    val = load_One_Value_From_INI_File("SERVER_URL")
    txt_ServerURL.Text = val
    val = load_One_Value_From_INI_File("PROXY_SERVER")
    txt_Proxy.Text = val
    val = load_One_Value_From_INI_File("USER_AGENT")
    txt_User_Agent.Text = val
    val = load_One_Value_From_INI_File("SLEEP_INTERVAL")
    txt_Sleep.Text = val
End Sub
Private Sub save_Values_To_INI_File()
    Dim val As String
    val = txt_ServerURL.Text
    Call save_One_Value_To_INI_File("SERVER_URL", val)
    val = txt_Proxy.Text
    Call save_One_Value_To_INI_File("PROXY_SERVER", val)
    val = txt_User_Agent.Text
    Call save_One_Value_To_INI_File("USER_AGENT", val)
    val = txt_Sleep.Text
    Call save_One_Value_To_INI_File("SLEEP_INTERVAL", val)
End Sub
Private Sub txt_ServerURL_GotFocus()
    txt_ServerURL.SelStart = 0
    txt_ServerURL.SelLength = Len(txt_ServerURL.Text)
End Sub
Private Sub txt_Proxy_GotFocus()
    txt_Proxy.SelStart = 0
    txt_Proxy.SelLength = Len(txt_Proxy.Text)
End Sub
Private Sub txt_User_Agent_Change()
    txt_User_Agent.SelStart = 0
    txt_User_Agent.SelLength = Len(txt_User_Agent.Text)
End Sub
Private Sub txt_Sleep_GotFocus()
    txt_Sleep.SelStart = 0
    txt_Sleep.SelLength = Len(txt_Sleep.Text)
End Sub

