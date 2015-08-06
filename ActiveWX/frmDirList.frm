VERSION 5.00
Begin VB.Form frmDirList 
   Caption         =   "Select Directory"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   1005
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   390
      Width           =   1005
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   3330
      Width           =   3195
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   210
      TabIndex        =   0
      Top             =   390
      Width           =   3165
   End
   Begin VB.Label Label2 
      Caption         =   "Drives:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   3090
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Directories:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2505
   End
End
Attribute VB_Name = "frmDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sDirPath As String
Private bIsCancel As Boolean

Private Sub cmdClose_Click()
  bIsCancel = True
  Unload Me
End Sub

Private Sub cmdSelect_Click()
  Unload Me
End Sub

Private Sub Dir1_Change()
  sDirPath = Dir1.Path
End Sub

Private Sub Drive1_Change()
 Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  bIsCancel = False
End Sub
Public Function Show_Dialog() As String
   Me.Show vbModal
   If Not bIsCancel Then
        Show_Dialog = sDirPath
   End If
End Function
