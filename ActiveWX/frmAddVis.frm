VERSION 5.00
Begin VB.Form frmAddVis 
   Caption         =   "Add Visibility"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTemp 
      Caption         =   "Visibility"
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
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   3855
      Begin VB.TextBox txtVis 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label60 
         Caption         =   "Miles (fraction in decimal)"
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
         Left            =   1710
         TabIndex        =   5
         Top             =   390
         Width           =   1965
      End
      Begin VB.Label Label57 
         Caption         =   "&Visibility:"
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
         TabIndex        =   0
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   1140
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddVis 
      Caption         =   "&Add Visibility"
      Default         =   -1  'True
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   1140
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sStationID As String
Private treWeather As TreeView

Private Sub cmdAddVis_Click()
  Dim nodX As Node
  Dim sParentKey As String
  Dim sNodeKey As String
  Dim sVis As String
  
  sNodeKey = "C" & sStationID & "V"
  sParentKey = "S" & sStationID
  sVis = "Visibility: " & txtVis.Text & " Miles"
  Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
             sVis)
  nodX.Tag = CInt(txtVis.Text)
  nodX.Image = "vis"
  nodX.EnsureVisible
  Unload Me
End Sub

Public Sub sbLoad_Form(treWx As TreeView)
 
 Dim iCount As Integer
 
 sStationID = Mid(treWx.SelectedItem.Key, 2, 3)
 iCount = 1
 Do Until iCount > treWx.Nodes.Count
     If treWx.Nodes(iCount).Key = "C" & sStationID & "V" Then
        MsgBox "You already have visibility specified for this station."
     End If
     iCount = iCount + 1
 Loop
 Set treWeather = treWx
 Me.Show vbModal
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub
