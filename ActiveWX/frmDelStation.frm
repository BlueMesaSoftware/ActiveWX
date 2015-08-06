VERSION 5.00
Begin VB.Form frmDelStation 
   Caption         =   "Delete Station From Database"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2790
      TabIndex        =   13
      Top             =   4320
      Width           =   1065
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
      Left            =   3930
      TabIndex        =   12
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Frame fraStations 
      Caption         =   "Select Station"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4905
      Begin VB.TextBox txtICAO 
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
         MaxLength       =   4
         TabIndex        =   7
         Top             =   3240
         Width           =   525
      End
      Begin VB.ListBox lstICAO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   885
      End
      Begin VB.ListBox lstStations 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1140
         TabIndex        =   5
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox txtStationName 
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
         TabIndex        =   4
         Top             =   3570
         Width           =   2865
      End
      Begin VB.ComboBox cboRegions 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   3195
      End
      Begin VB.OptionButton optSortStaIATA 
         Caption         =   "IATA"
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
         TabIndex        =   2
         Top             =   750
         Width           =   795
      End
      Begin VB.OptionButton optSortStaName 
         Caption         =   "Station Name"
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
         Left            =   1140
         TabIndex        =   1
         Top             =   750
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.Label Label38 
         Caption         =   "State/&Region:"
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
         Left            =   270
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "&Search station/IATA list for:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2970
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "&IATA:"
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
         Left            =   420
         TabIndex        =   9
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label Label54 
         Caption         =   "S&tation:"
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
         Left            =   390
         TabIndex        =   8
         Top             =   3660
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmDelStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bIsLoad As Boolean
Private Sub cboRegions_Click()
  If bIsLoad Then Exit Sub
  sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_STATION_NAME
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdDelete_Click()
  If MsgBox("Do you want to delete " & lstICAO.List(lstICAO.ListIndex) & _
     " from the database?", vbYesNo) = 7 Then Exit Sub
  Dim sSQL As String
  sSQL = "DELETE from Airports where IATA='" & lstICAO.List(lstICAO.ListIndex) & "'"
  gDB.Execute sSQL
  If optSortStaIATA.Value = True Then
    sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_IATA
  Else
    sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_STATION_NAME
  End If
  
End Sub

Private Sub Form_Load()
  bIsLoad = True
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  Call fnLoad_Regions
  bIsLoad = False
End Sub
Private Function fnLoad_Regions() As Integer
  
  Dim rs As Recordset
  Dim sSQL As String
    
  sSQL = "select distinct State from Airports order by State"
  Set rs = gDB.OpenRecordset(sSQL, dbOpenSnapshot)
  Do Until rs.EOF
     cboRegions.AddItem rs!State
     rs.MoveNext
  Loop
  rs.Close
End Function
Private Sub sbLoad_Stations(sStateName As String, iSort As Integer)
  Dim rs As Recordset
  Dim sSQL As String
  Dim iCount As Integer
  Dim sTemp As String
  Dim sOrderBy As String
  
  If iSort = SORT_IATA Then
     sOrderBy = " order by IATA"
  ElseIf iSort = SORT_STATION_NAME Then
     sOrderBy = " order by FacilityName"
  End If
 
  
  lstICAO.Clear
  lstStations.Clear
  
  sSQL = "select * from Airports where State='" & sStateName & "'" & sOrderBy
  Set rs = gDB.OpenRecordset(sSQL, dbOpenSnapshot)
  Do Until rs.EOF
     lstICAO.AddItem Trim(rs!IATA)
     If Not IsNull(rs!ICAO_Prefix) Then
        lstICAO.ItemData(lstICAO.NewIndex) = Asc(Trim(rs!ICAO_Prefix))
     End If
     If Not IsNull(rs!facilityname) Then
        lstStations.AddItem Trim(rs!facilityname)
     Else
        lstStations.AddItem " "
     End If
     rs.MoveNext
     iCount = iCount + 1
  Loop
  rs.Close
End Sub

Private Sub lstICAO_Click()
    lstStations.ListIndex = lstICAO.ListIndex
End Sub

Private Sub lstStations_Click()
    lstICAO.ListIndex = lstStations.ListIndex
End Sub

Private Sub txtICAO_Change()
  Dim strCheckParm As String
  Dim iCount As Integer
   
  If Len(txtICAO.Text) = 0 Then Exit Sub
  If lstICAO.ListCount = 0 Then Exit Sub
  
  strCheckParm = UCase(Left(txtICAO.Text, Len(txtICAO.Text)))
  iCount = lstICAO.ListCount
  Do Until iCount < 0
    If UCase(Left(lstICAO.List(iCount), Len(txtICAO.Text))) = strCheckParm Then
       lstICAO.ListIndex = iCount
       Exit Sub
    End If
    iCount = iCount - 1
  Loop
End Sub

Private Sub txtStationName_Change()
  Dim strCheckParm As String
  Dim iCount As Integer
   
  If Len(txtStationName.Text) = 0 Then Exit Sub
  If lstStations.ListCount = 0 Then Exit Sub
  
  strCheckParm = UCase(Left(txtStationName.Text, Len(txtStationName.Text)))
  iCount = lstStations.ListCount
  Do Until iCount < 0
    If UCase(Left(lstStations.List(iCount), Len(txtStationName.Text))) = strCheckParm Then
       lstStations.ListIndex = iCount
       Exit Sub
    End If
    iCount = iCount - 1
  Loop
End Sub
