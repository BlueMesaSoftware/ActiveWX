VERSION 5.00
Begin VB.Form frmAddCloud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Cloud Layer"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   3090
      TabIndex        =   17
      Top             =   3810
      Width           =   1335
   End
   Begin VB.CommandButton cmdCloudAdd 
      Caption         =   "&Add Layer"
      Default         =   -1  'True
      Height          =   315
      Left            =   1710
      TabIndex        =   16
      Top             =   3810
      Width           =   1335
   End
   Begin VB.Frame fraClouds 
      Caption         =   "Clouds"
      Height          =   3555
      Left            =   90
      TabIndex        =   18
      Top             =   120
      Width           =   4305
      Begin VB.TextBox txtCloudLayer 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   585
      End
      Begin VB.TextBox txtCloudDev 
         Height          =   315
         Left            =   1500
         TabIndex        =   13
         Top             =   2520
         Width           =   525
      End
      Begin VB.TextBox txtCloudTurb 
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   2160
         Width           =   525
      End
      Begin VB.TextBox txtCloudTop 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   585
      End
      Begin VB.TextBox txtCloudBase 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   585
      End
      Begin VB.ComboBox cboCloudType 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   2355
      End
      Begin VB.ComboBox cboCloudIce 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2880
         Width           =   2355
      End
      Begin VB.ComboBox cboCloudCov 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "&Cloud Layer:"
         Height          =   195
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "feet AGL"
         Height          =   195
         Left            =   2250
         TabIndex        =   22
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label20 
         Caption         =   "&Deviation:"
         Height          =   195
         Left            =   330
         TabIndex        =   12
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "0-none   255-severe"
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   2220
         Width           =   1725
      End
      Begin VB.Label Label22 
         Caption         =   "Miles"
         Height          =   195
         Left            =   2160
         TabIndex        =   20
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label Label23 
         Caption         =   "T&urbulence:"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "T&ype:"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1530
         Width           =   1245
      End
      Begin VB.Label Label25 
         Caption         =   "&Top:"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label26 
         Caption         =   "&Base:"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label19 
         Caption         =   "&Icing:"
         Height          =   195
         Left            =   330
         TabIndex        =   14
         Top             =   2940
         Width           =   945
      End
      Begin VB.Label Label28 
         Caption         =   "feet AGL"
         Height          =   195
         Left            =   2250
         TabIndex        =   19
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label27 
         Caption         =   "&Coverage:"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   1860
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmAddCloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCloudAdd_Click()
  If mdiMain.ActiveForm.treWeather.SelectedItem Is Nothing Then Exit Sub
  
  Dim asCloudData(8) As Variant
  Dim sParentKey As String
  Dim sNodeKey As String
  Dim sCloud As String
  Dim nodX As Node
  Dim iCloudCount As Integer
  Dim vStation As Variant
  

  
  asCloudData(1) = txtCloudBase.Text
  asCloudData(2) = txtCloudTop.Text
  asCloudData(3) = cboCloudType.List(cboCloudType.ListIndex)
  asCloudData(4) = cboCloudCov.List(cboCloudCov.ListIndex)
  asCloudData(5) = txtCloudTurb.Text
  asCloudData(6) = txtCloudDev.Text
  asCloudData(7) = CStr(cboCloudIce.ListIndex)
  asCloudData(8) = txtCloudLayer.Text
  sCloud = "Cloud: " & txtCloudLayer.Text & " (" & _
           Format(txtCloudBase.Text, "##,##0") & " - " & _
           Format(txtCloudTop.Text, "##,##0") & " feet AGL) " & _
           cboCloudCov.List(cboCloudCov.ListIndex) & " " & _
           cboCloudType.List(cboCloudType.ListIndex)
 
  If Left(mdiMain.ActiveForm.treWeather.SelectedItem.Key, 1) = "S" Then
         '0- ICAO
        vStation = mdiMain.ActiveForm.treWeather.SelectedItem.Tag
        vStation(16) = CInt(vStation(16)) + 1
        mdiMain.ActiveForm.treWeather.SelectedItem.Tag = vStation
        sParentKey = mdiMain.ActiveForm.treWeather.SelectedItem.Key
        sNodeKey = "C" & vStation(0) & "C" & _
                   vStation(16)
        Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sCloud)
  Else
        vStation = mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Tag
        vStation(16) = CInt(vStation(16)) + 1
        sParentKey = mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Key
        sNodeKey = "C" & vStation(0) & "C" & _
                   vStation(16)
        mdiMain.ActiveForm.treWeather.SelectedItem.Parent.Tag = vStation
        Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sCloud)
  End If
  
  
  nodX.Image = "cloud"
  nodX.Tag = asCloudData
  nodX.Sorted = True
  nodX.EnsureVisible
End Sub

Private Sub Form_Load()
   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 2 - Me.Height / 2
   cboCloudType.AddItem "Cirrus"
   cboCloudType.AddItem "CirroStratus"
   cboCloudType.AddItem "CirroCumulus"
   cboCloudType.AddItem "AltoStratus"
   cboCloudType.AddItem "AltoCumulus"
   cboCloudType.AddItem "StratoCumulus"
   cboCloudType.AddItem "NimboStratus"
   cboCloudType.AddItem "Stratus"
   cboCloudType.AddItem "Cumulus"
   cboCloudType.AddItem "Cumulonimbus"
   cboCloudType.AddItem "Userdefined"
   cboCloudType.ListIndex = 10
   
   cboCloudCov.AddItem "Clear"
   cboCloudCov.AddItem "Scattered1"
   cboCloudCov.AddItem "Scattered2"
   cboCloudCov.AddItem "Scattered3"
   cboCloudCov.AddItem "Scattered4"
   cboCloudCov.AddItem "Broken5"
   cboCloudCov.AddItem "Broken6"
   cboCloudCov.AddItem "Broken7"
   cboCloudCov.AddItem "Overcast"
   cboCloudCov.ListIndex = 1
      
   cboCloudIce.AddItem "None"
   cboCloudIce.AddItem "Cloud Icing"
   cboCloudIce.ListIndex = 0
   txtCloudTurb.Text = "0"
   txtCloudDev.Text = "0"
End Sub
