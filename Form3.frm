VERSION 5.00
Begin VB.Form frm_About 
   Caption         =   "About"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Caption         =   "Developer : Shibu TN"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1095
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "ver 1.0 ($BUILD_TIME_STAMP$)"
      Height          =   255
      Left            =   345
      TabIndex        =   1
      Top             =   810
      Width           =   4710
   End
   Begin VB.Label Label1 
      Caption         =   "Vypeen Soft - SubDB subtitle download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6480
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
