VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "comct232.ocx"
Begin VB.Form frmStatus 
   Caption         =   "Download Status"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
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
      Left            =   1710
      TabIndex        =   2
      Top             =   1350
      Width           =   1245
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   979
      _Version        =   327681
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   267
      FullHeight      =   37
   End
   Begin VB.Label labInetStatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   930
      Width           =   4065
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStop_Click()
    mdiMain.ActiveForm.sbStopInet
End Sub

Private Sub Form_Load()
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  Animation1.Open App.Path & "\" & "filecopy.avi"
End Sub
