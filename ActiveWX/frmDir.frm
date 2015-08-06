VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmDir 
   Caption         =   "Directories"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Simulator Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   180
      TabIndex        =   12
      Top             =   3120
      Width           =   8865
      Begin VB.ComboBox cboFSVersion 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "FS Version:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   420
         Width           =   1755
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8730
      Top             =   3570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6870
      TabIndex        =   11
      Top             =   4110
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7950
      TabIndex        =   10
      Top             =   4110
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "APLC Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   150
      TabIndex        =   5
      Top             =   1620
      Width           =   8865
      Begin VB.TextBox txtAPLC 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         TabIndex        =   7
         Top             =   390
         Width           =   5265
      End
      Begin VB.CommandButton cmdAPLC 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7560
         TabIndex        =   6
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Path for APLC program:"
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
         Left            =   270
         TabIndex        =   9
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Enter the path and filename for the  APLC program. This is REQUIRED to create Active Weather 98 Adventures.  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   8
         Top             =   780
         Width           =   5265
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Simulator Adventure Directory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   8835
      Begin VB.CommandButton cmdPathADV 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7620
         TabIndex        =   3
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox txtADVPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2130
         TabIndex        =   2
         Top             =   390
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   $"frmDir.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   4
         Top             =   780
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Path FS Adv Directory:"
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
         Left            =   270
         TabIndex        =   1
         Top             =   450
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAPLC_Click()
  ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo errhandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Executable Programs" & _
    "(*.exe)|*.exe"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtAPLC.Text = CommonDialog1.filename
    Exit Sub

errhandler:
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdPathADV_Click()
  Dim sADVPath As String
  sADVPath = frmDirList.Show_Dialog
  If sADVPath <> "" Then
    txtADVPath.Text = sADVPath
  End If
End Sub

Private Sub cmdSave_Click()
  SaveSetting "ActiveWx98", "Sim", "AdvPath", txtADVPath.Text
  gsAdvPath = txtADVPath.Text
  
  SaveSetting "ActiveWx98", "Sim", "APLCPath", txtAPLC.Text
  gsAPLCPath = txtAPLC.Text
  
  If cboFSVersion.ListIndex = 0 Then
     SaveSetting "ActiveWx98", "Sim", "Version", "98"
     gsSimVersion = "98"
  Else
     SaveSetting "ActiveWx98", "Sim", "Version", "95"
     gsSimVersion = "95"
  End If
  Unload Me
End Sub

Private Sub Command2_Click()
'  Dim sPath As String
'  sPath = Chr(34) & txtPathADV.Text & "\aplc32.bat" & Chr(34)
'  MsgBox sPath
'  Call Shell(sPath, vbMaximizedFocus)
'   sPath = txtPathADV.Text & "\test.bat"
End Sub

Private Sub Form_Load()
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  txtADVPath.Text = gsAdvPath
  txtAPLC.Text = gsAPLCPath
  cboFSVersion.AddItem "Flight Simulator 98"
  cboFSVersion.AddItem "Flight Simulator for Windows 95"
  If gsSimVersion = "98" Then
     cboFSVersion.ListIndex = 0
  Else
     cboFSVersion.ListIndex = 1
  End If
End Sub
