VERSION 5.00
Begin VB.Form frmAddStation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Station to Database"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   3300
      TabIndex        =   17
      Top             =   3210
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4650
      TabIndex        =   18
      Top             =   3210
      Width           =   1275
   End
   Begin VB.Frame fraStation 
      Caption         =   "Station Data"
      Height          =   3045
      Left            =   120
      TabIndex        =   19
      Top             =   90
      Width           =   5865
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   3675
      End
      Begin VB.ComboBox cboRegions 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1440
         Width           =   3675
      End
      Begin VB.TextBox txtICAOPrefix 
         Height          =   315
         Left            =   3690
         MaxLength       =   1
         TabIndex        =   3
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtIATA 
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtStationLat 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   1800
         Width           =   1485
      End
      Begin VB.TextBox txtStationLong 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   2160
         Width           =   1485
      End
      Begin VB.TextBox txtStationElev 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtStationName 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   3675
      End
      Begin VB.Label Label10 
         Caption         =   "City:"
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label Label9 
         Caption         =   "K for USA lower 48"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "State/Region:"
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "hhh-mm-ss.sss(E/W)"
         Height          =   255
         Left            =   3510
         TabIndex        =   22
         Top             =   2190
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "hh-mm-ss.sss(N/S)"
         Height          =   255
         Left            =   3510
         TabIndex        =   21
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "ICAO Prefix:"
         Height          =   255
         Left            =   2670
         TabIndex        =   2
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "IATA 3 letter:"
         Height          =   255
         Left            =   270
         TabIndex        =   0
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Latitiude:"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   1860
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Longitude:"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   2220
         Width           =   1185
      End
      Begin VB.Label Label7 
         Caption         =   "Elevation:"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   2580
         Width           =   1185
      End
      Begin VB.Label Label39 
         Caption         =   "Station Name:"
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label Label40 
         Caption         =   "Feet"
         Height          =   255
         Left            =   2970
         TabIndex        =   20
         Top             =   2580
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmAddStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub sbLoad_Regions()
  Dim rs As Recordset
  Dim sSQL As String
    
  sSQL = "select distinct State from Airports order by State"
  Set rs = gDB.OpenRecordset(sSQL, dbOpenSnapshot)
  Do Until rs.EOF
     cboRegions.AddItem rs!State
     rs.MoveNext
  Loop
  rs.Close
End Sub

Private Sub cmdAdd_Click()
  Dim sSQL As String
  If txtStationName.Text = "" Or txtCity = "" Or cboRegions.Text = "" Or _
     txtStationLat.Text = "" Or txtStationLong.Text = "" Or txtStationElev.Text = "" _
     Or txtIATA.Text = "" Or txtICAOPrefix.Text = "" Then
     MsgBox "You did not fill in all the required fields.", vbExclamation
     Exit Sub
  End If
'  On Error GoTo err_exit
  
  sSQL = "INSERT into Airports(IATA, TypeFacility, FacilityName, City, State, " & _
         "Latitude, Longitude, ICAO_Prefix, Elevation) Values('" & _
         txtIATA.Text & "', 'AIRPORT', '" & txtStationName.Text & _
         "', '" & txtCity & "', '" & _
         cboRegions.Text & "', '" & txtStationLat.Text & "', '" & _
         txtStationLong.Text & "', '" & txtICAOPrefix.Text & _
         "', " & txtStationElev & ")"
  
  gDB.Execute sSQL
  mdiMain.ActiveForm.cboRegions.ListIndex = cboRegions.ListIndex
  MsgBox "Station has been added to the database.", vbInformation
  Unload Me
  Exit Sub
err_exit:
  MsgBox "There was an error while attempting to add a new station.", vbExclamation
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  sbLoad_Regions
End Sub

Private Sub txtIATA_KeyPress(KeyAscii As Integer)
 Dim sChar As String
 sChar = Chr(KeyAscii)
 KeyAscii = Asc(UCase(sChar))
End Sub

Private Sub txtICAOPrefix_KeyPress(KeyAscii As Integer)
 Dim sChar As String
 sChar = Chr(KeyAscii)
 KeyAscii = Asc(UCase(sChar))
End Sub
