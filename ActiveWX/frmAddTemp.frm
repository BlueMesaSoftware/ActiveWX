VERSION 5.00
Begin VB.Form frmAddTemp 
   Caption         =   "Add Temperature"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTempAdd 
      Caption         =   "&Add Temp."
      Default         =   -1  'True
      Height          =   315
      Left            =   660
      TabIndex        =   5
      Top             =   1140
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Frame fraTemp 
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3255
      Begin VB.TextBox txtTemp 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1380
         TabIndex        =   1
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label57 
         Caption         =   "&Temperature:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label60 
         Caption         =   "degrees F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2010
         TabIndex        =   2
         Top             =   390
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmAddTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sStationID As String
Private treWeather As TreeView
Public Sub sbLoad_Form(treWx As TreeView)
 
 Dim iCount As Integer
 
 sStationID = Mid(treWx.SelectedItem.Key, 2, 3)
 iCount = 1
 Do Until iCount > treWx.Nodes.Count
     If treWx.Nodes(iCount).Key = "C" & sStationID & "T1" Then
        MsgBox "You already have a temperture specified."
     End If
     iCount = iCount + 1
 Loop
 Set treWeather = treWx
 Me.Show vbModal
End Sub



Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdTempAdd_Click()
  Dim nodX As Node
  Dim sParentKey As String
  Dim sNodeKey As String
  Dim sTemp As String
  
  sNodeKey = "C" & sStationID & "T" & "1"
  sParentKey = "S" & sStationID
  sTemp = "Temperature: " & txtTemp.Text & " F"
  Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
             sTemp)
  nodX.Tag = CInt(txtTemp.Text)
  nodX.Image = "thermo"
  nodX.EnsureVisible
  Unload Me
End Sub

Private Sub Form_Load()
   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub
