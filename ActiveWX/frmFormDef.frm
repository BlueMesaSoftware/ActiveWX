VERSION 5.00
Begin VB.Form frmFormDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Form CGI"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraWeb 
      Caption         =   "CGI Form Lines"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   90
      TabIndex        =   6
      Top             =   840
      Width           =   4965
      Begin VB.ComboBox cboSub 
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
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3150
         Width           =   2295
      End
      Begin VB.TextBox txtURL 
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
         Height          =   285
         Left            =   300
         TabIndex        =   10
         Top             =   1020
         Width           =   4425
      End
      Begin VB.TextBox txtCGI2 
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
         Height          =   285
         Left            =   300
         TabIndex        =   9
         Top             =   2700
         Width           =   4425
      End
      Begin VB.TextBox txtCGI1 
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
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   1740
         Width           =   4245
      End
      Begin VB.TextBox txtFormName 
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
         Height          =   285
         Left            =   1410
         TabIndex        =   7
         Top             =   420
         Width           =   3285
      End
      Begin VB.Label Label5 
         Caption         =   "&Type Submission:"
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
         Left            =   300
         TabIndex        =   17
         Top             =   3210
         Width           =   1335
      End
      Begin VB.Label labURL 
         Caption         =   "&URL:"
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
         Left            =   330
         TabIndex        =   16
         Top             =   750
         Width           =   375
      End
      Begin VB.Label labName2 
         Caption         =   "&End CGI Line (Completion of CGI line):"
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
         Left            =   300
         TabIndex        =   15
         Top             =   2460
         Width           =   4425
      End
      Begin VB.Label labName1 
         Caption         =   "&Begin CGI Line (Starting at the ? in URL)"
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
         Left            =   330
         TabIndex        =   14
         Top             =   1470
         Width           =   4125
      End
      Begin VB.Label Label1 
         Caption         =   "&Form Name:"
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
         Left            =   330
         TabIndex        =   13
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   12
         Top             =   1770
         Width           =   105
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "+ Wx Station Textbox +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   11
         Top             =   2130
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pre-Programmed Sites"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   4995
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
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
         Left            =   3600
         TabIndex        =   5
         Top             =   270
         Width           =   1035
      End
      Begin VB.ComboBox cboSites 
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
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "Site:"
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
         TabIndex        =   4
         Top             =   330
         Width           =   825
      End
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
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
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
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   1245
   End
End
Attribute VB_Name = "frmFormDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtName2_Change()

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
   SaveSetting "ActiveWx98", "Web", "FormName", txtFormName.Text
   SaveSetting "ActiveWx98", "Web", "URL", txtURL.Text
   SaveSetting "ActiveWx98", "Web", "CGI1", txtCGI1.Text
   SaveSetting "ActiveWx98", "Web", "CGI2", txtCGI2.Text
   SaveSetting "Activewx98", "Web", "CGISub", cboSub.ListIndex
   Call mdiMain.ActiveForm.fnReadRegSettings
   Unload Me
End Sub

Private Sub cmdSelect_Click()
    Select Case cboSites.ListIndex
    Case 0
       txtFormName = "NWS Current METAR"
       txtURL.Text = "http://tgsv5.nws.noaa.gov/cgi-bin/mgetmetar.pl"
       txtCGI1.Text = "cccc="
       cboSub.ListIndex = 0
    Case 1
       txtFormName = "Texas A&M Weather Interface"
       txtURL.Text = "http://www.met.tamu.edu/cgi-bin/post-weather"
       txtCGI1.Text = "station="
       txtCGI2.Text = "&time=0&subcom1=3&groups=1"
       cboSub.ListIndex = 1
 End Select
End Sub



Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    cboSites.AddItem "NWS Current METAR"
    cboSites.AddItem "Texas A&M Weather Interface"
    cboSites.ListIndex = 0
    
    cboSub.AddItem "URL Line"
    cboSub.AddItem "POST method"
    cboSub.ListIndex = 0
    
    
    txtFormName.Text = GetSetting("ActiveWx98", "Web", "FormName", "")
    txtURL.Text = GetSetting("ActiveWx98", "Web", "URL", "")
    txtCGI1.Text = GetSetting("ActiveWx98", "Web", "CGI1", "")
    txtCGI2.Text = GetSetting("ActiveWx98", "Web", "CGI2", "")
    cboSub.ListIndex = CInt(GetSetting("ActiveWx98", "Web", "CGISub", "0"))
End Sub
