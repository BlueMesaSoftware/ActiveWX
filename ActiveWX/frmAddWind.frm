VERSION 5.00
Begin VB.Form frmAddWind 
   Caption         =   "Add Wind Layer"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddWind 
      Caption         =   "&Add Wind"
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
      Left            =   1740
      TabIndex        =   21
      Top             =   3330
      Width           =   1155
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
      Left            =   2970
      TabIndex        =   20
      Top             =   3330
      Width           =   1155
   End
   Begin VB.Frame fraWind 
      Caption         =   "Winds"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      TabIndex        =   14
      Top             =   90
      Width           =   4065
      Begin VB.TextBox txtWindLayer 
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
         Left            =   1500
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox txtWindBase 
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
         Left            =   1500
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txtWindTop 
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
         Left            =   1500
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox txtWindTurb 
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   13
         Top             =   2520
         Width           =   525
      End
      Begin VB.TextBox txtWindDir 
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1410
         Width           =   585
      End
      Begin VB.TextBox txtWindSpeed 
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1770
         Width           =   585
      End
      Begin VB.ComboBox cboWindType 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "&Layer:"
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
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label44 
         Caption         =   "feet AGL"
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
         Left            =   2250
         TabIndex        =   19
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label46 
         Caption         =   "&Base:"
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
         Left            =   300
         TabIndex        =   2
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label47 
         Caption         =   "&Top:"
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
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label49 
         Caption         =   "T&urbulence:"
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
         Left            =   300
         TabIndex        =   12
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label50 
         Caption         =   "1-360"
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
         Left            =   2250
         TabIndex        =   18
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label51 
         Caption         =   "0-none   255-severe"
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
         Left            =   2280
         TabIndex        =   17
         Top             =   2580
         Width           =   1485
      End
      Begin VB.Label Label52 
         Caption         =   "&Direction:"
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
         Left            =   300
         TabIndex        =   6
         Top             =   1470
         Width           =   945
      End
      Begin VB.Label Label53 
         Caption         =   "feet AGL"
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
         Left            =   2250
         TabIndex        =   16
         Top             =   750
         Width           =   705
      End
      Begin VB.Label Label43 
         Caption         =   "Knots"
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
         Left            =   2250
         TabIndex        =   15
         Top             =   1830
         Width           =   465
      End
      Begin VB.Label Label45 
         Caption         =   "&Speed:"
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
         Left            =   300
         TabIndex        =   8
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label Label48 
         Caption         =   "T&ype:"
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
         Left            =   300
         TabIndex        =   10
         Top             =   2220
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmAddWind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddWind_Click()
  Dim asWindData(8) As Variant
  Dim sParentKey As String
  Dim sNodeKey As String
  Dim sWind As String
  Dim nodX As Node
  Dim iWindCount As Integer
  Dim vStation As Variant
  
   '==add additional layers of wind
     '0-wind direction
     '1-wind speed
     '2-wind type
     '3-turbulence
     '4-wind base feet AGL
     '5-wind top feet  AGL
     '6-SURFACE/Layer
  asWindData(0) = txtWindDir.Text
  asWindData(1) = txtWindSpeed.Text
  asWindData(2) = cboWindType.List(cboWindType.ListIndex)
  asWindData(3) = txtWindTurb.Text
  asWindData(4) = txtWindBase.Text
  asWindData(5) = txtWindTop.Text
  
  sWind = "Wind Layer: " & txtWindLayer.Text & " " & asWindData(0) & _
             " at " & asWindData(1) & " knots " & _
             "(" & asWindData(4) & " - " & asWindData(5) & " feet AGL)"
   
  If Left(mdiMain.ActiveForm.treWeather.SelectedItem.Key, 1) = "S" Then
         '0- ICAO
        vStation = mdiMain.ActiveForm.treWeather.SelectedItem.Tag
        vStation(15) = CInt(vStation(15)) + 1
        mdiMain.ActiveForm.treWeather.SelectedItem.Tag = vStation
        sParentKey = mdiMain.ActiveForm.treWeather.SelectedItem.Key
        sNodeKey = "C" & vStation(0) & "C" & _
                   vStation(15)
        Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
  Else
        vStation = mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Tag
        vStation(15) = CInt(vStation(15)) + 1
        sParentKey = mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Key
        sNodeKey = "C" & vStation(0) & "C" & _
                   vStation(15)
        mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Tag = vStation
        Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
  End If
  
  
  nodX.Image = "wind"
  nodX.Tag = asWindData
  nodX.Sorted = True
  nodX.EnsureVisible
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  
  cboWindType.AddItem "steady"
  cboWindType.AddItem "gusty"
  cboWindType.ListIndex = 0
  txtWindTurb.Text = "0"
End Sub
