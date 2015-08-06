VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmWeather 
   Caption         =   "New Weather Adventure"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   10950
   Visible         =   0   'False
   Begin VB.Frame fraStation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   5790
      TabIndex        =   46
      Top             =   540
      Visible         =   0   'False
      Width           =   4965
      Begin VB.TextBox txtStationInRange 
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
         Left            =   1890
         TabIndex        =   136
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txtStationWxWidth 
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
         Left            =   1890
         TabIndex        =   134
         Top             =   2460
         Width           =   555
      End
      Begin VB.TextBox txtStationN 
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
         Left            =   1890
         TabIndex        =   73
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtStationWxTran 
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
         Left            =   1890
         TabIndex        =   58
         Top             =   2850
         Width           =   555
      End
      Begin VB.TextBox txtStationWxHeight 
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
         Left            =   1890
         TabIndex        =   57
         Top             =   2070
         Width           =   555
      End
      Begin VB.TextBox txtStationElev 
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
         Left            =   1890
         TabIndex        =   51
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtStationLong 
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
         Left            =   1890
         TabIndex        =   50
         Top             =   1320
         Width           =   2025
      End
      Begin VB.TextBox txtStationLat 
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
         Left            =   1890
         TabIndex        =   49
         Top             =   960
         Width           =   2025
      End
      Begin VB.TextBox txtStationICAO 
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
         Left            =   1890
         MaxLength       =   4
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Wx In Range:"
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
         Left            =   210
         TabIndex        =   138
         Top             =   3300
         Width           =   1545
      End
      Begin VB.Label Label12 
         Caption         =   "Miles"
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
         Left            =   2610
         TabIndex        =   137
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Miles Wide (West to East)"
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
         Left            =   2610
         TabIndex        =   135
         Top             =   2520
         Width           =   2145
      End
      Begin VB.Label Label8 
         Caption         =   "Wx Longitude:"
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
         Left            =   210
         TabIndex        =   133
         Top             =   2520
         Width           =   1545
      End
      Begin VB.Label Label40 
         Caption         =   "Feet"
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
         Left            =   2970
         TabIndex        =   75
         Top             =   1710
         Width           =   645
      End
      Begin VB.Label Label39 
         Caption         =   "Station Name:"
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
         TabIndex        =   71
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label Label31 
         Caption         =   "Decimal"
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
         Left            =   4020
         TabIndex        =   62
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label30 
         Caption         =   "Decimal"
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
         Left            =   4020
         TabIndex        =   61
         Top             =   990
         Width           =   645
      End
      Begin VB.Label Label16 
         Caption         =   "Miles"
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
         Left            =   2610
         TabIndex        =   60
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label Label15 
         Caption         =   "Miles Wide (North to South)"
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
         Left            =   2610
         TabIndex        =   59
         Top             =   2100
         Width           =   2085
      End
      Begin VB.Label Label35 
         Caption         =   "Wx Transition:"
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
         Left            =   210
         TabIndex        =   56
         Top             =   2910
         Width           =   1545
      End
      Begin VB.Label Label34 
         Caption         =   "Wx Latitude:"
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
         Left            =   210
         TabIndex        =   55
         Top             =   2130
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Elevation:"
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
         TabIndex        =   54
         Top             =   1740
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Longitude:"
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
         TabIndex        =   53
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Latitiude:"
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
         TabIndex        =   52
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "IATA 3 letter:"
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
         TabIndex        =   47
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.Frame fraFiles 
      Caption         =   "From Local Weather Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   5790
      TabIndex        =   39
      Top             =   510
      Visible         =   0   'False
      Width           =   4995
      Begin VB.OptionButton optProcWxPlus 
         Caption         =   "Process Wx/Add Stations"
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
         Left            =   2670
         TabIndex        =   23
         Top             =   810
         Width           =   2235
      End
      Begin VB.OptionButton optProcWx 
         Caption         =   "Process Wx"
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
         Left            =   1380
         TabIndex        =   22
         Top             =   810
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Process Wx"
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   21
         Top             =   750
         Width           =   1185
      End
      Begin VB.TextBox txtMetarPath 
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
         Left            =   1410
         TabIndex        =   20
         Top             =   390
         Width           =   3345
      End
      Begin VB.CommandButton cmdMetarFile 
         Caption         =   "Metar File..."
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
         Left            =   150
         TabIndex        =   19
         Top             =   390
         Width           =   1185
      End
   End
   Begin VB.Frame fraWeb 
      Caption         =   "From World Wide Web"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   5790
      TabIndex        =   35
      Top             =   2010
      Visible         =   0   'False
      Width           =   4965
      Begin VB.TextBox txtForm 
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
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   600
         Width           =   2475
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   210
         ScaleHeight     =   345
         ScaleWidth      =   4515
         TabIndex        =   120
         Top             =   2520
         Width           =   4515
         Begin VB.OptionButton optProcURLPlus 
            Caption         =   "Process Wx/Add Stations"
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
            Left            =   2130
            TabIndex        =   132
            Top             =   60
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optProcURL 
            Caption         =   "Process Wx"
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
            Left            =   450
            TabIndex        =   121
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdProcWebWx 
         Caption         =   "&Process Wx"
         Enabled         =   0   'False
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
         Left            =   3270
         TabIndex        =   119
         Top             =   2130
         Width           =   1455
      End
      Begin VB.ComboBox cboURL 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   330
         TabIndex        =   118
         Top             =   1620
         Width           =   4365
      End
      Begin VB.OptionButton optURL 
         Caption         =   "URL"
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
         Left            =   150
         TabIndex        =   117
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&New/Edit..."
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
         Left            =   3810
         TabIndex        =   116
         Top             =   600
         Width           =   1005
      End
      Begin VB.OptionButton optCGI 
         Caption         =   "CGI/Form"
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
         Left            =   150
         TabIndex        =   24
         Top             =   330
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.TextBox txtCGIStation 
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
         Left            =   1260
         TabIndex        =   26
         Top             =   960
         Width           =   1365
      End
      Begin VB.CommandButton cmdViewHTML 
         Caption         =   "&View HTML"
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   28
         Top             =   2130
         Width           =   1515
      End
      Begin VB.CommandButton cmdGetURL 
         Caption         =   "&Get URL"
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
         TabIndex        =   27
         Top             =   2130
         Width           =   1365
      End
      Begin VB.Label labForms 
         Caption         =   "Web Form:"
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
         TabIndex        =   115
         Top             =   630
         Width           =   945
      End
      Begin VB.Label labCGIStation 
         Caption         =   "Wx Station:"
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
         TabIndex        =   25
         Top             =   990
         Width           =   975
      End
   End
   Begin VB.Frame fraStations 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   5790
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdAddGlobal 
         Caption         =   "Add &Global Wx to Adventure"
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
         Left            =   2490
         TabIndex        =   14
         Top             =   4050
         Width           =   2235
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
         TabIndex        =   10
         Top             =   750
         Value           =   -1  'True
         Width           =   1275
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
         TabIndex        =   9
         Top             =   750
         Width           =   795
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
         TabIndex        =   8
         Top             =   300
         Width           =   3195
      End
      Begin VB.CommandButton cmdAddWeather 
         Caption         =   "&Add Station to Adventure"
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
         Left            =   270
         TabIndex        =   13
         Top             =   4050
         Width           =   2145
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
         TabIndex        =   18
         Top             =   3570
         Width           =   2865
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
         TabIndex        =   12
         Top             =   1080
         Width           =   3555
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
         TabIndex        =   11
         Top             =   1080
         Width           =   885
      End
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
         TabIndex        =   16
         Top             =   3240
         Width           =   525
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
         TabIndex        =   17
         Top             =   3660
         Width           =   675
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
         TabIndex        =   15
         Top             =   3300
         Width           =   495
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
         TabIndex        =   77
         Top             =   2970
         Width           =   2535
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
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraClouds 
      Caption         =   "Cloud Layer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   5790
      TabIndex        =   67
      Top             =   540
      Visible         =   0   'False
      Width           =   4995
      Begin VB.TextBox txtCloudLayer 
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
         TabIndex        =   64
         Top             =   330
         Width           =   585
      End
      Begin VB.ComboBox cboCloudCov 
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
         TabIndex        =   76
         Top             =   1770
         Width           =   2355
      End
      Begin VB.ComboBox cboCloudIce 
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
         TabIndex        =   84
         Top             =   2850
         Width           =   2355
      End
      Begin VB.ComboBox cboCloudType 
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
         TabIndex        =   72
         Top             =   1410
         Width           =   2355
      End
      Begin VB.TextBox txtCloudBase 
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
         TabIndex        =   66
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txtCloudTop 
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
         TabIndex        =   69
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox txtCloudTurb 
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
         TabIndex        =   80
         Top             =   2130
         Width           =   525
      End
      Begin VB.TextBox txtCloudDev 
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
         TabIndex        =   82
         Top             =   2490
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "&Cloud Layer:"
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
         TabIndex        =   63
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "&Coverage:"
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
         TabIndex        =   74
         Top             =   1830
         Width           =   1245
      End
      Begin VB.Label Label28 
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
         TabIndex        =   45
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label19 
         Caption         =   "&Icing:"
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
         Left            =   330
         TabIndex        =   83
         Top             =   2910
         Width           =   945
      End
      Begin VB.Label Label26 
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
         TabIndex        =   65
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label25 
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
         TabIndex        =   68
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label24 
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
         TabIndex        =   70
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label Label23 
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
         TabIndex        =   78
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Miles"
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
         TabIndex        =   44
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label21 
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
         Left            =   2250
         TabIndex        =   43
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label Label20 
         Caption         =   "&Deviation:"
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
         Left            =   330
         TabIndex        =   81
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label Label18 
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
         TabIndex        =   42
         Top             =   720
         Width           =   705
      End
   End
   Begin VB.Frame fraWind 
      Caption         =   "Wind Layer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   5790
      TabIndex        =   79
      Top             =   540
      Visible         =   0   'False
      Width           =   4995
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
         Left            =   1470
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   390
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   2190
         Width           =   1575
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
         Left            =   1470
         TabIndex        =   100
         Top             =   1830
         Width           =   585
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
         Left            =   1470
         TabIndex        =   97
         Top             =   1470
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
         Left            =   1470
         TabIndex        =   104
         Top             =   2580
         Width           =   525
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
         Left            =   1470
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   1110
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
         Left            =   1470
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   750
         Width           =   585
      End
      Begin VB.Label Label55 
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
         Left            =   270
         TabIndex        =   88
         Top             =   450
         Width           =   885
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
         Left            =   270
         TabIndex        =   101
         Top             =   2250
         Width           =   1245
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
         Left            =   270
         TabIndex        =   99
         Top             =   1890
         Width           =   945
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
         Left            =   2220
         TabIndex        =   93
         Top             =   1890
         Width           =   465
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
         Left            =   2220
         TabIndex        =   90
         Top             =   810
         Width           =   705
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
         Left            =   270
         TabIndex        =   96
         Top             =   1530
         Width           =   945
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
         Left            =   2250
         TabIndex        =   87
         Top             =   2640
         Width           =   1725
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
         Left            =   2220
         TabIndex        =   86
         Top             =   1530
         Width           =   465
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
         Left            =   270
         TabIndex        =   103
         Top             =   2640
         Width           =   1215
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
         Left            =   270
         TabIndex        =   94
         Top             =   1170
         Width           =   885
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
         Left            =   270
         TabIndex        =   91
         Top             =   810
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
         Left            =   2220
         TabIndex        =   85
         Top             =   1170
         Width           =   705
      End
   End
   Begin VB.Frame fraTemp 
      Caption         =   "Altimeter,Temperature and Visibility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   5790
      TabIndex        =   106
      Top             =   540
      Visible         =   0   'False
      Width           =   4995
      Begin VB.TextBox txtVis 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         TabIndex        =   111
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox txtTemp 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         TabIndex        =   109
         Top             =   810
         Width           =   525
      End
      Begin VB.TextBox txtAltimeter 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         MaxLength       =   5
         TabIndex        =   107
         Top             =   420
         Width           =   525
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
         Left            =   2190
         TabIndex        =   114
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label59 
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
         Left            =   2190
         TabIndex        =   113
         Top             =   1260
         Width           =   1995
      End
      Begin VB.Label Label58 
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
         Left            =   240
         TabIndex        =   112
         Top             =   1260
         Width           =   1215
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
         Left            =   240
         TabIndex        =   110
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label56 
         Caption         =   "&Altimeter:"
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
         TabIndex        =   108
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdGeneral 
      Height          =   315
      Left            =   6030
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   555
   End
   Begin VB.CommandButton cmdStations 
      Height          =   315
      Left            =   6930
      Picture         =   "Form1.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5160
      Width           =   555
   End
   Begin VB.CommandButton cmdWx 
      Height          =   315
      Left            =   7860
      Picture         =   "Form1.frx":03B4
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5160
      Width           =   555
   End
   Begin VB.Frame fraGeneral 
      Height          =   3285
      Left            =   5790
      TabIndex        =   40
      Top             =   480
      Width           =   4995
      Begin VB.CommandButton cmdSit 
         Caption         =   "&Situation Name"
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
         TabIndex        =   123
         Top             =   1890
         Width           =   1365
      End
      Begin VB.TextBox txtAdvSituation 
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
         Left            =   1620
         TabIndex        =   6
         Top             =   1890
         Width           =   3105
      End
      Begin VB.TextBox txtAdvFilename 
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
         Left            =   1620
         TabIndex        =   5
         Top             =   1530
         Width           =   1875
      End
      Begin VB.TextBox txtAdvDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1620
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   750
         Width           =   3105
      End
      Begin VB.TextBox txtAdvTitle 
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
         Left            =   1620
         TabIndex        =   1
         Top             =   360
         Width           =   3105
      End
      Begin VB.Label Label17 
         Caption         =   "*.adv"
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
         Left            =   3600
         TabIndex        =   41
         Top             =   1590
         Width           =   585
      End
      Begin VB.Label Label10 
         Caption         =   "&Filename:"
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
         TabIndex        =   4
         Top             =   1590
         Width           =   885
      End
      Begin VB.Label Label9 
         Caption         =   "&Description:"
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
         TabIndex        =   2
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "&Title:"
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
         TabIndex        =   0
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weather Adventure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   120
      TabIndex        =   98
      Top             =   90
      Width           =   5565
      Begin ComctlLib.TreeView treWeather 
         Height          =   5415
         Left            =   120
         TabIndex        =   105
         Top             =   270
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9551
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   318
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5190
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5760
      Picture         =   "Form1.frx":04B6
      ScaleHeight     =   300
      ScaleWidth      =   5160
      TabIndex        =   36
      Top             =   180
      Width           =   5160
      Begin VB.Label labTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Adventure"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label labTitleDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "General adventure settings"
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
         TabIndex        =   37
         Top             =   60
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdMakeADV 
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
      Left            =   8790
      Picture         =   "Form1.frx":59F8
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5160
      Width           =   555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   9630
      TabIndex        =   34
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      Caption         =   "Compile Adventure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8640
      TabIndex        =   131
      Top             =   5520
      Width           =   945
   End
   Begin VB.Label Label64 
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   8580
      TabIndex        =   130
      Top             =   5220
      Width           =   195
   End
   Begin VB.Label Label63 
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   7650
      TabIndex        =   129
      Top             =   5220
      Width           =   195
   End
   Begin VB.Label Label29 
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   6720
      TabIndex        =   128
      Top             =   5220
      Width           =   195
   End
   Begin VB.Label Label14 
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   5820
      TabIndex        =   127
      Top             =   5220
      Width           =   195
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "General Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5850
      TabIndex        =   126
      Top             =   5520
      Width           =   945
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      Caption         =   "Add Wx Stations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6780
      TabIndex        =   125
      Top             =   5520
      Width           =   945
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      Caption         =   "Process METAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7710
      TabIndex        =   124
      Top             =   5520
      Width           =   945
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   5310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5AFA
            Key             =   "stationic"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":601C
            Key             =   "propic"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":653E
            Key             =   "detailic"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6A60
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6F82
            Key             =   "cloud"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":74A4
            Key             =   "thermo"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":79C6
            Key             =   "alt"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7EE8
            Key             =   "wind"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":840A
            Key             =   "vis"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iWindCount As Integer
Private iCloudLayer As Integer
Private iTempCount As Integer
Public strHTML As String
Private sFilePath As String
Private sACWPath As String
Private bIsLoad As Boolean
Private sCGIURL As String
Private sCGI1 As String
Private sCGI2 As String
Private iCGISub As Integer
Public giCurrentTip As Integer


Private Type ADV_WX_TYPE
   station_ID As String
   miles_to As String
   inRange As String
   weather As String
   weather_char As String
   wind_count As Integer
   winds(9) As String
   visibility As String
   cloud_count As Integer
   clouds(9) As String
   temp_count As Integer
   temperature(9) As String
   baro_pressure As String
End Type

Private Type SAVE_TREE_TYPE
   sParentKey As String * 80
   sParentText As String * 80
   sNodeKey As String * 80
   sNodeText As String * 80
   NodeTag As Variant
End Type
'office 97 toolbar globals
Private iOldX As Integer, iOldY As Integer
Private oldIndex As Byte

Public Sub sbClear_Buttons()
   MoveMouse 100
End Sub
Public Sub sbLoadTenURLs()
   Dim iCount As Integer
   
   cboURL.Clear
   iCount = 1
   Do Until iCount > 11
     If Trim(GetSetting("ActiveWx98", "Web", "URL" & iCount, "")) <> "" Then
         cboURL.AddItem Trim(GetSetting("ActiveWx98", "Web", "URL" & iCount, ""))
         If iCount = 1 Then
             cboURL.Text = Trim(GetSetting("ActiveWx98", "Web", "URL" & iCount, ""))
         End If
     End If
     iCount = iCount + 1
   Loop
End Sub
Private Sub sbSetToolTips()
    cmdGeneral.ToolTipText = "Add or change adventure settings utilized by Flight Simulator."
    cmdStations.ToolTipText = "Add local or global weather stations to the adventure." & _
                              ""
    cmdWx.ToolTipText = "Process METAR data locally or from the Internet."
End Sub
Private Sub sbCompileADV()
        On Error GoTo err_exit
        If txtAdvFilename.Text <> "" Then
              sAdvFileName = "\" & txtAdvFilename
        Else
              sAdvFileName = "\aw98.adv "
        End If
        
        sShell = Chr(34) & gsAPLCPath & Chr(34) & _
                         " -" & gsSimVersion & " " & _
                          Chr(34) & gsAdvPath & "\aw98.txt" & Chr(34) & " " & Chr(34) & _
                          gsAdvPath & sAdvFileName
        
        
        'Clipboard.Clear
        'Clipboard.SetText sShell
        Call Shell(sShell)
        
        Screen.MousePointer = vbDefault
        MsgBox "Adventure has been generated", vbExclamation
'        Kill gsAdvPath & "\aw98.txt"
        Exit Sub
err_exit:
  If Err.Number = 5 Then
     MsgBox "You did not set your APLC filname and path correctly.", vbExclamation
     Exit Sub
  End If
  If Err.Number = 53 Then
     MsgBox "You need to specify the path and filename of the Aplc program" & sCR & _
     "you are utilizing. Set this from menu item File Settings.", vbExclamation
     Exit Sub
  End If
  MsgBox "An error occurred which prevented your adventure from compiling. " & _
  "The error was: " & Err.Description & ":" & Err.Number
End Sub
Private Sub sbChangeVis()
   If treWeather.SelectedItem Is Nothing Then Exit Sub
   
   treWeather.SelectedItem.Tag = txtVis.Text
   treWeather.SelectedItem.Text = "Visibility: " & txtVis.Text
End Sub
Private Sub sbChangeTemp()
   If treWeather.SelectedItem Is Nothing Then Exit Sub
   
   treWeather.SelectedItem.Tag = CInt(txtTemp.Text)
   treWeather.SelectedItem.Text = "Temperature: " & txtTemp.Text & " F"
End Sub
Private Sub sbChangeAltimeter()
   If treWeather.SelectedItem Is Nothing Then Exit Sub
   
   treWeather.SelectedItem.Tag = txtAltimeter.Text
   treWeather.SelectedItem.Text = "Altimeter: " & txtAltimeter.Text
End Sub
Private Sub sbChangeCloud()
      '1 raw base feet (AGL)
      '2 raw top feet  (AGL)
      '3 type clouds
      '4 coverage
      '5 turbulence
      '6 devation
      '7 icing
  If treWeather.SelectedItem Is Nothing Then Exit Sub
  
  Dim vCloudData As Variant
  vCloudData = treWeather.SelectedItem.Tag
  vCloudData(1) = txtCloudBase.Text
  vCloudData(2) = txtCloudTop.Text
  vCloudData(3) = cboCloudType.List(cboCloudType.ListIndex)
  vCloudData(4) = cboCloudCov.List(cboCloudCov.ListIndex)
  vCloudData(5) = CInt(txtCloudTurb.Text)
  vCloudData(6) = CInt(txtCloudDev.Text)
  vCloudData(7) = CStr(cboCloudIce.ListIndex)
  vCloudData(8) = txtCloudLayer.Text
  treWeather.SelectedItem.Tag = vCloudData
  
  treWeather.SelectedItem.Text = "Cloud: " & vCloudData(8) & " (" & _
           Format(txtCloudBase.Text, "##,##0") & "-" & _
           Format(txtCloudTop.Text, "##,##0") & " feet AGL) " & _
           cboCloudCov.List(cboCloudCov.ListIndex) & " " & _
           cboCloudType.List(cboCloudType.ListIndex)
  treWeather.SelectedItem.Parent.Sorted = True
  
End Sub
Private Sub sbChangeStation()
     'sStation
     '0- ICAO
     '1- Station Latitude (Dec)
     '2- Station Longitude (Dec)
     '3- Station Elevation (meters)
     '4- Area Begin Lat
     '5- Area Begin Long
     '6- Area End Lat
     '7- Area End Long
     '8- Weather Area Width
     '9- Weather Area Transition
     '10- Weather Course
     '11- Weather Velocity
     '12- Station Name
     '13- Weather time/date data
  If treWeather.SelectedItem Is Nothing Then Exit Sub
  
  Dim vStationData As Variant
  vStationData = treWeather.SelectedItem.Tag
  vStationData(0) = txtStationICAO.Text
  vStationData(1) = txtStationLat.Text
  vStationData(2) = txtStationLong.Text
  vStationData(3) = txtStationElev.Text
  vStationData(4) = txtStationWxHeight.Text
  vStationData(5) = txtStationWxWidth.Text
  vStationData(6) = txtStationInRange.Text
  vStationData(12) = txtStationN.Text
  treWeather.SelectedItem.Tag = vStationData
  treWeather.SelectedItem.Text = vStationData(0) & "(" & vStationData(12) & ")" & _
          vStationData(13)
End Sub
Private Sub sbChangeWind()
  
  If treWeather.SelectedItem Is Nothing Then Exit Sub
  
  Dim vWindData As Variant
  '0-wind direction
  '1-wind speed
  '2-wind type
  '3-turbulence
  '4-wind base feet
  '5-wind top feet
  '6-surface/layer
  
   vWindData = treWeather.SelectedItem.Tag
   vWindData(4) = txtWindBase.Text
   vWindData(5) = txtWindTop.Text
   vWindData(1) = txtWindSpeed.Text
   vWindData(0) = txtWindDir.Text
   vWindData(3) = cboWindType.List(cboWindType.ListIndex)
   vWindData(6) = txtWindLayer.Text
   
   treWeather.SelectedItem.Tag = vWindData
'  If vWindData(6) = "L" Then
    treWeather.SelectedItem.Text = "Wind Layer: " & _
             vWindData(6) & " " & _
             vWindData(0) & _
             " at " & vWindData(1) & " knots " & _
             "(" & vWindData(4) & " - " & vWindData(5) & " feet AGL)"
'  Else
'     treWeather.SelectedItem.Text = "Wind Surface: " & vWindData(0) & _
'             " at " & vWindData(1) & " knots " & _
'             "(" & vWindData(4) & " - " & vWindData(5) & " feet AGL)"
'  End If
End Sub
  
Private Sub DownMouse(ByVal bIndex As Byte)

' If bIndex <> 100 Then
'
' Dim iX As Integer, iY As Integer
'
'    iX = Picture1(bIndex).Left - 4
'    iY = Picture1(bIndex).Top - 5
'
'    Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B
'    Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014
'    Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8 + 1, iY + 17 + 8), &H80000014
'
'
'End If

End Sub
Private Sub MoveMouse(ByVal bIndex As Byte)

'  If bIndex <> oldIndex Then  ' If it is already drawn, then don't do it again!
'
'  Dim iX As Integer, iY As Integer
'
'    If oldIndex <> 100 Then ' Index 100 = No button selected!
'        Picture2.Line (iOldX, iOldY)-(iOldX + 17 + 8, iOldY + 17 + 8), &H8000000A, B
'        ' Remove the 3D-effect of the old button.
'    End If
'
'    If bIndex <> 100 Then
'
'        iX = Picture1(bIndex).Left - 4
'        iY = Picture1(bIndex).Top - 5
'
'
'        Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014, B
'        Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B
'        Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B
'
'        iOldX = iX: iOldY = iY
'    End If
'
'    oldIndex = bIndex
'
'End If
End Sub
Private Sub UpMouse(ByVal bIndex As Byte)

' If bIndex <> 100 Then
'
' Dim iX As Integer, iY As Integer
'
'    iX = Picture1(bIndex).Left - 4
'    iY = Picture1(bIndex).Top - 5
'
'    Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014, B
'    Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010
'    Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8 + 1, iY + 17 + 8), &H80000010
'
'
'
'End If

End Sub
Private Sub sbRefresh_Selected_Screen(treNode As Node)
 
 If Left(treNode.Key, 1) = "S" Then
         labTitle.Caption = "Weather Area Data"
         labTitleDesc.Caption = "Edit adventure wx area data"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = True
         fraWind.Visible = False
         fraTemp.Visible = False
         sbSetStationControls treNode.Tag
  End If
  
  Select Case Mid(treNode.Key, 5, 1)
     Case "C" 'clouds
         Dim sCloudCov As String
         Dim sCloudType As String
         labTitle.Caption = "Edit Cloud Layer"
         labTitleDesc.Caption = "Edit selected cloud layer"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = True
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = False
          '1 raw base feet (AGL)
          '2 raw top feet  (AGL)
          '3 type clouds
          '4 coverage
          '5 turbulence
          '6 devation
          '7 icing
          '8-layer
          txtCloudBase.Text = Format(treNode.Tag(1), "###0")
          txtCloudTop.Text = treNode.Tag(2)
          sCloudType = treNode.Tag(3)
          cboCloudType.ListIndex = fnFindCloudType(cboCloudType, sCloudType)
          sCloudCov = treNode.Tag(4)
          cboCloudCov.ListIndex = fnFindCloudCov(cboCloudCov, sCloudCov)
          txtCloudTurb.Text = treNode.Tag(5)
          txtCloudDev.Text = treNode.Tag(6)
          cboCloudIce.ListIndex = treNode.Tag(7)
          txtCloudLayer.Text = treNode.Tag(8)
     Case "W" 'winds
         Dim sWindType As String
         labTitle.Caption = "Edit Wind Layer"
         labTitleDesc.Caption = "Edit selected wind layer"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = True
         fraTemp.Visible = False
        '0-wind direction
        '1-wind speed
        '2-wind type
        '3-turbulence
        '4-wind base feet
         '5-wind top feet
          txtWindLayer.Text = treNode.Tag(6)
          txtWindDir.Text = treNode.Tag(0)
          txtWindSpeed.Text = treNode.Tag(1)
          sWindType = treNode.Tag(2)
          cboWindType.ListIndex = fnFindWindType(cboWindType, sWindType)
          txtWindTurb.Text = treNode.Tag(3)
          txtWindBase.Text = Format(treNode.Tag(4), "###0")
          txtWindTop.Text = treNode.Tag(5)
     Case "A" 'altimeter
         labTitle.Caption = "Edit Altimeter"
         labTitleDesc.Caption = "Edit selected altimeter setting"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
          
         txtAltimeter.Text = Node.Tag
         txtAltimeter.BackColor = ACTIVE_COLOR
         txtAltimeter.Enabled = True
         txtTemp.Text = ""
         txtTemp.BackColor = INACTIVE_COLOR
         txtTemp.Enabled = False
         txtVis.Text = ""
         txtVis.BackColor = INACTIVE_COLOR
         txtVis.Enabled = False
      Case "T" 'Temp
         labTitle.Caption = "Edit Temperature"
         labTitleDesc.Caption = "Edit selected temperature"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
         txtAltimeter.Text = ""
         txtAltimeter.BackColor = INACTIVE_COLOR
         txtAltimeter.Enabled = False
         txtTemp.Text = treNode.Tag
         txtTemp.BackColor = ACTIVE_COLOR
         txtTemp.Enabled = True
         txtVis.Text = ""
         txtVis.BackColor = INACTIVE_COLOR
         txtVis.Enabled = False
  
   Case "V" 'visibility
         labTitle.Caption = "Edit Visibility"
         labTitleDesc.Caption = "Edit selected visibility"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
          
         txtAltimeter.BackColor = INACTIVE_COLOR
         txtAltimeter.Enabled = True
         txtTemp.Text = ""
         txtTemp.BackColor = INACTIVE_COLOR
         txtTemp.Enabled = False
         txtVis.Text = treNode.Tag
         txtVis.BackColor = ACTIVE_COLOR
         txtVis.Enabled = True
  End Select
End Sub
Private Function fnSetFilePath(iTypeDialog As Integer) As Boolean
   CommonDialog1.CancelError = True
   
   On Error GoTo errhandler
   
   CommonDialog1.Flags = cdlOFNHideReadOnly
  
   CommonDialog1.Filter = "All Files (*.*)|*.*|Active Wx Files" & _
   "(*.acw)|*.acw"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    If iTypeDialog = TYPE_DLG_OPEN Then
'        CommonDialog1.ShowOpen
    Else
        CommonDialog1.ShowSave
    End If
    
    ' Display name of selected file
    sACWPath = CommonDialog1.filename
    fnSetFilePath = True
  Exit Function
   
errhandler:
   sACWPath = ""
   fnSetFilePath = False
End Function
Public Sub sbLoad_Form(iTypeLoad As Integer, sFileName As String)
    If iTypeLoad = TYPE_LOAD_OLD Then
        sFilePath = sFileName
        If UCase(Right(sFileName, 3)) = "ACW" Then
            sbReadACWFile
        Else
            sbReadWxPlus 1
        End If
    End If
   giCurrentTip = TIP_GENERAL_ADV
'   If gbShowTips = True Then
'
'   End If
End Sub
Private Sub sbReadACWFile()
   Dim tTreeData As SAVE_TREE_TYPE
   Dim iCount As Integer
   Dim nodX As Node
   
   On Error GoTo err_exit
   
   Screen.MousePointer = vbHourglass
   Open sFilePath For Random Access Read As #1 Len = 640
   Get #1, 1, tTreeData
     
   If Trim(tTreeData.sNodeText) <> "Active Weather 98 Ver 1" Then
       Close #1
       Exit Sub
   End If
  
   txtAdvTitle.Text = tTreeData.NodeTag(0)
   txtAdvDesc.Text = tTreeData.NodeTag(1)
   txtAdvFilename.Text = tTreeData.NodeTag(2)
   txtAdvSituation.Text = tTreeData.NodeTag(3)
   
   treWeather.Nodes.Clear
   
   iCount = 2
   Get #1, iCount, tTreeData
   Do Until EOF(1)
         If Trim(tTreeData.sParentKey) = "none" Then
             Set nodX = treWeather.Nodes.Add(, , Trim(tTreeData.sNodeKey), _
                        Trim(tTreeData.sNodeText))
             nodX.Tag = tTreeData.NodeTag
             nodX.Image = "stationic"
         End If
         iCount = iCount + 1
         Get #1, iCount, tTreeData
   Loop
   iCount = 2
   Close #1
   Open sFilePath For Random Access Read As #1 Len = 640
   
  
   Do Until EOF(1)
        Get #1, iCount, tTreeData
        If Trim(tTreeData.sParentKey) <> "none" Then
             Set nodX = treWeather.Nodes.Add(Trim(tTreeData.sParentKey), tvwChild, _
                        Trim(tTreeData.sNodeKey), Trim(tTreeData.sNodeText))
             nodX.Tag = tTreeData.NodeTag
             nodX.Parent.Sorted = True
             Select Case Mid(Trim(tTreeData.sNodeKey), 5, 1)
                Case "A"
                   nodX.Image = "alt"
                Case "C"
                   nodX.Image = "cloud"
                Case "W"
                   nodX.Image = "wind"
                Case "T"
                   nodX.Image = "thermo"
                Case "M"
                   nodX.Image = "clock"
                Case "V"
                   nodX.Image = "vis"
             End Select
        End If
         
         iCount = iCount + 1
   Loop
   Close #1
   Me.Caption = sFilePath
   sACWPath = sFilePath
   mdiMain.mnuFileSave.Enabled = True
   Screen.MousePointer = vbDefault
   
   mdiMain.sbSaveFiveFiles sACWPath
   mdiMain.sbReadFiveFiles
   
   Exit Sub
metar_file_type:
   Close #1
   sbReadWxPlus 1
   
err_exit:
   Close #1
   Screen.MousePointer = vbDefault
End Sub
Public Sub sbSaveWx(bIsSaveAs As Boolean)
   If bIsSaveAs Then
       If Not fnSetFilePath(TYPE_DLG_SAVE) Then Exit Sub
       sbSaveTree
   Else
        If sACWPath <> "" Then
         sbSaveTree
        Else
            If Not fnSetFilePath(TYPE_DLG_SAVE) Then Exit Sub
            sbSaveTree
        End If
   End If
   Me.Caption = sACWPath
End Sub
Private Sub sbSaveTree()
   Dim iCount As Integer
   Dim tTreeData As SAVE_TREE_TYPE
   Dim asACWData(5) As String
   Dim iPOSCount As Integer
   
   On Error Resume Next
   Kill sACWPath
   Open sACWPath For Random Access Write As #1 Len = 640
   
   'write the version information for the file
   tTreeData.sNodeText = "Active Weather 98 Ver 1"
   asACWData(0) = txtAdvTitle.Text
   asACWData(1) = txtAdvDesc.Text
   asACWData(2) = txtAdvFilename.Text
   asACWData(3) = txtAdvSituation.Text
   tTreeData.NodeTag = asACWData
   Put #1, 1, tTreeData
   
   If treWeather.Nodes.Count = 0 Then
      Close #1
      Exit Sub
   End If
   
   iCount = 1
   iPOSCount = 2
   Do Until iCount > treWeather.Nodes.Count
       tTreeData.sNodeKey = treWeather.Nodes(iCount).Key
       tTreeData.NodeTag = treWeather.Nodes(iCount).Tag
       tTreeData.sNodeText = treWeather.Nodes(iCount).Text
       If Left(treWeather.Nodes(iCount).Key, 1) = "S" Then
          tTreeData.sParentKey = "none"
       Else
          tTreeData.sParentKey = treWeather.Nodes(iCount).Parent.Key
          tTreeData.sParentText = treWeather.Nodes(iCount).Parent.Text
       End If
       
       Put #1, iPOSCount, tTreeData
       iCount = iCount + 1
       iPOSCount = iPOSCount + 1
   Loop
   mdiMain.SBar1.SimpleText = "File saved"
   Close #1
   
   mdiMain.sbSaveFiveFiles sACWPath
   mdiMain.sbReadFiveFiles
End Sub
Private Sub sbSetStationControls(vTagData As Variant)
      'sStation
     '0- ICAO
     '1- Station Latitude (Dec)
     '2- Station Longitude (Dec)
     '3- Station Elevation (meters)
     '4- Area Begin Lat
     '5- Area Begin Long
     '6- Area End Lat
     '7- Area End Long
     '8- Weather Area Width
     '9- Weather Area Transition
     '10- Weather Course
     '11- Weather Velocity
     '12 -Station Name
     txtStationICAO.Text = vTagData(0)
     txtStationLat.Text = vTagData(1)
     txtStationLong.Text = vTagData(2)
     txtStationElev.Text = vTagData(3)
     txtStationWxHeight.Text = vTagData(4)
     txtStationWxWidth.Text = vTagData(5)
     txtStationInRange.Text = vTagData(6)
     txtStationWxTran.Text = vTagData(9)
     txtStationN.Text = Trim(vTagData(12))
     
End Sub
Private Sub sbInetStopped()
   cmdGetURL.Enabled = True
   cmdProcWebWx.Enabled = True
   cmdViewHTML.Enabled = True
   cmdStop.Enabled = False
   Unload frmStatus
End Sub
Public Sub sbStopInet()
   Inet1.Cancel
'   Do Until Not Inet1.StillExecuting
     frmStatus.labInetStatus.Caption = "Closing connection..."
'     DoEvents
'   Loop
   cmdGetURL.Enabled = True
   cmdProcWebWx.Enabled = True
   cmdViewHTML.Enabled = True
   Unload frmStatus
 
End Sub
Private Sub sbGetURL()

    Dim sURL As String
    Dim sCGI As String
    
'    On Error GoTo err_handler
    strHTML = ""
    If Not optCGI.Value Then
        If cboURL.Text = "" Then GoTo err_input
        Inet1.Execute cboURL.Text
        sbSaveTenURLs cboURL.Text
    Else
        If sCGIURL = "" Then GoTo err_input
'        sFormData = Trim(sCGI1)
        sURL = sCGIURL
        sCGI = Trim(sCGI1) & Trim(txtCGIStation.Text) & Trim(sCGI2)
        If iCGISub = 1 Then
            Inet1.Execute sURL, "POST", sCGI
        Else
            sURL = sURL & "?" & sCGI
            Inet1.Execute sURL
        End If
    End If
       
    cmdGetURL.Enabled = False
    frmStatus.Show vbModal
    
    
    Exit Sub
err_handler:
   MsgBox Err.Description, vbExclamation
   sbStopInet
   Exit Sub
err_input:
   MsgBox "You did not enter a URL to process.", vbExclamation
End Sub
Private Sub sbSaveRegSettings()
  SaveSetting "ActiveWx98", "Stations", "Regions", cboRegions.ListIndex
  SaveSetting "ActiveWx98", "Web", "CGIStation", txtCGIStation.Text
  SaveSetting "ActiveWx98", "Stations", "optCGI", optCGI.Value
  SaveSetting "ActiveWx98", "Stations", "optURL", optURL.Value
End Sub
Private Sub sbSetDefaultCGI()
        
 
End Sub
Public Function fnReadRegSettings() As Integer
 
  Dim sSetting As String
   
  txtForm.Text = GetSetting("ActiveWx98", "Web", "FormName", "")
  sCGIURL = GetSetting("ActiveWx98", "Web", "URL", "")
  sCGI1 = GetSetting("ActiveWx98", "Web", "CGI1", "")
  sCGI2 = GetSetting("ActiveWx98", "Web", "CGI2", "")
  txtCGIStation.Text = GetSetting("ActiveWx98", "Web", "CGIStation", "")
  cboRegions.ListIndex = CInt(GetSetting("ActiveWx98", "Stations", "Regions", 0))
  optCGI.Value = CBool(GetSetting("ActiveWx98", "Stations", "optCGI", 0))
  optURL.Value = CBool(GetSetting("ActiveWx98", "Stations", "optURL", 0))
  iCGISub = CInt(GetSetting("ActiveWx98", "Web", "CGISub", "0"))
End Function

Private Sub sbSaveTenURLs(sNewURL As String)
   Dim iCount As Integer
   Dim iDupIndex As Integer
   Dim sFileNum As String
   
   iDupIndex = 4
   iCount = 1
   'look for duplicates
   Do Until iCount > 11
     If sNewURL = Trim(GetSetting("ActiveWx98", "Web", "URL" & iCount, "")) Then
       iDupIndex = iCount
     End If
     iCount = iCount + 1
   Loop
   
   iCount = iDupIndex - 1
   Do Until iCount = 0
       
       SaveSetting "ActiveWx98", "Web", "URL" & iCount + 1, _
                   GetSetting("ActiveWx98", "Web", "URL" & iCount, "")
       iCount = iCount - 1
   Loop
   
   SaveSetting "ActiveWx98", "Web", "URL1", sNewURL
   sbLoadTenURLs
   
End Sub

Private Sub cboCloudCov_LostFocus()
   sbChangeCloud
End Sub

Private Sub cboCloudIce_LostFocus()
  sbChangeCloud
End Sub

Private Sub cboCloudType_LostFocus()
   sbChangeCloud
End Sub

Private Sub cboRegions_Click()
  If bIsLoad Then Exit Sub
  sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_STATION_NAME
End Sub

Private Sub cboWindType_LostFocus()
  sbChangeWind
End Sub

Private Sub cmdAddGlobal_Click()
    Dim nodX As Node
     Dim sStationName As String
     Dim sFullStationName As String
     Dim sStationID As String
     Dim sStation(16) As String
     Dim sDescID As String
     Dim asCloudData(8) As String
     Dim asWindData(6) As String
     Dim asWindLayer1(6) As String
     Dim asWindLayer2(6) As String
     Dim sVis As String
     Dim iWindCount As Integer
     
     On Error GoTo error_dup
     
     'sStation
     '0- ICAO
     '1- Station Latitude (Dec)
     '2- Station Longitude (Dec)
     '3- Station Elevation (feet)
     '4- Weather Area Height
     '5- Weather Area Width
     '6- Inrange (turn on weather area)
     '7-
     '8-
     '9- Weather Area Transition
     '10- Weather Course
     '11- Weather Velocity
     '12 -Station Name
     '13 -weather date time data
     '14 -ICAO Prefix
     '15 -Wind Count
     '16 -Cloud Count
        
     sStation(0) = "GLB"
     sStationName = "GlobalWX"
     sStation(1) = "1"
     sStation(2) = "1"
     sStation(3) = "10"
     sStation(4) = "500"
     sStation(5) = "500"
     sStation(6) = "400"
     
     sStation(8) = "300"
     sStation(9) = "150"
     sStation(10) = "0"
     sStation(11) = "0"
     
     sStationID = "SGLB"
     sFullStationName = "GLB (GlobalWX)"
     sStation(12) = sStationName
     sStation(14) = "G"
     sStation(15) = "3"
     sStation(16) = "1"
     
     Set nodX = treWeather.Nodes.Add(, , sStationID, sFullStationName)
     nodX.Image = "stationic"
     nodX.Tag = sStation
     nodX.Sorted = True
'     nodX.EnsureVisible
     
      '=======Wind===================================
       '0-wind direction
     '1-wind speed
     '2-wind type
     '3-turbulence
     '4-wind base feet
     '5-wind top feet
  
     '==add additional layers of wind
     '0-wind direction
     '1-wind speed
     '2-wind type
     '3-turbulence
     '4-wind base feet AGL
     '5-wind top feet  AGL
     '6-wind level
     asWindData(0) = "275"
     asWindData(1) = 4
     asWindData(2) = "steady"
     asWindData(3) = 0
     asWindData(4) = 0
     asWindData(5) = 1700
     asWindData(6) = "1"
     
     asWindLayer1(0) = "275"
     asWindLayer1(1) = asWindData(1) + 11
     asWindLayer1(2) = "steady"
     asWindLayer1(3) = 0
     asWindLayer1(4) = asWindData(5)
     asWindLayer1(5) = asWindData(5) + 6800
     asWindLayer1(6) = "2"
     
     asWindLayer2(0) = asWindLayer1(0)
     asWindLayer2(1) = asWindLayer1(1) + 11
     asWindLayer2(2) = "steady"
     asWindLayer2(3) = 0
     asWindLayer2(4) = asWindLayer1(5)
     asWindLayer2(5) = asWindLayer1(5) + 14800
     asWindLayer2(6) = "3"
     
     '====Add surface layer to the tree
     sNodeKey = "CGLBW1"
     sParentKey = "SGLB"
     sWind = "Wind Layer: " & asWindData(6) & " " & asWindData(0) & _
             " at " & asWindData(1) & " knots " & _
             "(" & asWindData(4) & " - " & asWindData(5) & " feet AGL)"
     Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindData
     nodX.Image = "wind"
          
     '====Add wind layer1 to the tree=======================================
     sNodeKey = "CGLBW2"
     
     sWind = "Wind Layer: " & asWindLayer1(6) & " " & asWindLayer1(0) & _
             " at " & asWindLayer1(1) & " knots " & _
             "(" & asWindLayer1(4) & " - " & asWindLayer1(5) & " feet AGL)"
     Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindLayer1
     nodX.Image = "wind"
       
     '====Add wind layer2 to the tree=======================================
     sNodeKey = "CGLBW3"
    
     sWind = "Wind Layer: " & asWindLayer2(6) & " " & asWindLayer2(0) & _
             " at " & asWindLayer2(1) & " knots " & _
             "(" & asWindLayer2(4) & " - " & asWindLayer2(5) & " feet AGL)"
     Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindLayer2
     nodX.Image = "wind"
             
     '====Altmeter=====================
     sNodeKey = "CGLBA"
                 
     sParentKey = "SGLB"
     sAltimeter = "Altimeter: 29.92"
     
     Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sAltimeter)
     nodX.Tag = "29.92"
     nodX.Image = "alt"
     
     '1 raw base feet (AGL)
     '2 raw top feet  (AGL)
     '3 type clouds
     '4 coverage
     '5 turbulence
     '6 devation
     '7 icing
   
 
   asCloudData(1) = "12190"
   asCloudData(2) = 12190 + 1100
   asCloudData(3) = "Userdefined"
   asCloudData(4) = "Scattered3"
   
   sCloud = "Cloud: 1 (" & _
           Format(asCloudData(1), "##,##0") & " - " & _
           Format(asCloudData(2), "##,##0") & " feet AGL) " & asCloudData(4) & _
           " " & asCloudData(3)
   
   asCloudData(5) = "0"
   asCloudData(6) = "0"
   asCloudData(7) = "0"
   asCloudData(8) = 1
   sNodeKey = "CGLBC" & 1
   sParentKey = "SGLB"
   Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sCloud)
   nodX.Image = "cloud"
   nodX.Tag = asCloudData
   
   
   '==========Temperature=======================
   sNodeKey = "CGLBT1"
   sParentKey = "SGLB"
   sTemp = "Tempature: 65 F"
   iTempC = "65"
   Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sTemp)
   nodX.Tag = iTempC
   nodX.Image = "thermo"

   '===========Visibility=========================
   sNodeKey = "CGLBV"
   sParentKey = "SGLB"
   sVisibility = "Visibility: 15 miles"
   sVis = "15"
   Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sVisibility)
   nodX.Image = "vis"
   nodX.Tag = sVis
   nodX.EnsureVisible
   Exit Sub
error_dup:
  MsgBox "You have already added a global weather area. You must delete the current one to " & _
  "add a new one.", vbInformation
End Sub



Private Sub cmdAddWeather_Click()
  If lstStations.ListIndex = -1 Then Exit Sub
  
  Call fnAddStation(lstICAO.List(lstICAO.ListIndex), _
                    Chr(lstICAO.ItemData(lstICAO.ListIndex)), 0)
End Sub

Private Sub cmdAdvSituation_Click()
   Dim sSituation As String
   
  ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo errhandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Situation Files" & _
    "(*.stn)|*.stn"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
   
    txtAdvSituation.Text = CommonDialog1.filename
    Exit Sub

errhandler:

End Sub

Private Sub cmdAddWind_Click()
  If treWeather.SelectedItem Is Nothing Then Exit Sub
  
  Dim aWindData(6) As String
  Dim sParentKey As String
  Dim sNodeKey As String
  Dim sCloud As String
  Dim nodX As Node
  Dim iCloudCount As Integer
  Dim vStation As Variant
  
  '0-wind direction
  '1-wind speed
  '2-wind type
  '3-turbulence
  '4-wind base feet
  '5-wind top feet
  '6-surface/layer
  
 
   aWindData(4) = txtWindBase.Text
   aWindData(5) = txtWindTop.Text
   aWindData(1) = txtWindSpeed.Text
   aWindData(0) = txtWindDir.Text
   aWindData(3) = cboWindType.List(cboWindType.ListIndex)
 
  sWind = "Wind Layer: " & aWindData(0) & _
             " at " & aWindData(1) & " knots " & _
             "(" & aWindData(4) & " - " & aWindData(5) & " feet AGL)"
  
 
  If Left(treWeather.SelectedItem.Key, 1) = "S" Then
        '0- ICAO
        vStation = treWeather.SelectedItem.Tag
        vStation(15) = CInt(vStation(15)) + 1
        sParentKey = treWeather.SelectedItem.Key
        sNodeKey = "C" & treWeather.SelectedItem.Tag(0) & "W" & _
                   vStation(15)
        Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
  Else
        vStation = treWeather.SelectedItem.Parent.Tag
        vStation(15) = CInt(vStation(15)) + 1
        sParentKey = treWeather.SelectedItem.Parent.Key
        sNodeKey = "C" & treWeather.SelectedItem.Parent.Tag(0) & "W" & _
                   vStation(15)
        treWeather.SelectedItem.Parent.Tag = vStation
        Set nodX = treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
  End If
  
  
  nodX.Image = "wind"
  nodX.Tag = aWindData
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub



Private Sub cmdCloudClear_Click()
      
  txtCloudBase.Text = ""
  txtCloudTop.Text = ""
  cboCloudType.ListIndex = 0
  cboCloudCov.ListIndex = 0
  txtCloudTurb.Text = ""
  txtCloudDev.Text = ""
  cboCloudIce.ListIndex = 0
End Sub



Private Sub cmdDelStation_Click()

End Sub

Private Sub cmdEdit_Click()
  frmFormDef.Show vbModal
End Sub

Private Sub cmdGeneral_Click()
         labTitle.Caption = "Adventure"
         labTitleDesc.Caption = "General adventure settings"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = True
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = False
         giCurrentTip = TIP_GENERAL_ADV
End Sub

Private Sub cmdGetURL_Click()
    sbGetURL
End Sub

Private Sub cmdImport_Click()
  If optProcWx.Value Then
        Call fnReadWeather
  Else
        sbReadWxPlus 0
  End If
End Sub

Private Function fnReadWxTree(sStation As String, nNodes As Nodes) As ADV_WX_TYPE
    Dim iCount As Integer
    Dim sChildKey As String
    Dim lCloudBaseMt As Long
    Dim lCloudTopMt As Long
    Dim lWindBaseFt As Long
    Dim lWindTopFt As Long
    Dim sngDegWidth As Single
    Dim sngBeginLong As Single
    Dim sngBeginLat As Single
    Dim sngEndLong As Single
    Dim sngEndLat As Single
    
    iWindCount = 0
    iCloudLayer = 0
    iTempCount = 0
    
    sChildKey = "C" & sStation
      'sStation (parent)
     '0- ICAO
     '1- Station Latitude (Dec)
     '2- Station Longitude (Dec)
     '3- Station Elevation (feet)
     '4- Area Begin Lat
     '5- Area Begin Long
     '6- Area End Lat
     '7- Area End Long
     '8- Weather Area Width
     '9- Weather Area Transition
     '10- Weather Course
     '11- Weather Velocity
     
   
    iCount = 1
    Do Until iCount > nNodes.Count
       If Left(nNodes(iCount).Key, 4) = sChildKey Then
           Select Case Mid(nNodes(iCount).Key, 5, 1)
              Case "A" 'Altimeter
                fnReadWxTree.baro_pressure = "  BARO_PRESSURE " & nNodes(iCount).Tag & _
                                             ",0"
              Case "C" 'CLOUDS
                '1 raw base feet (AGL)
                '2 raw top feet  (AGL)
                '3 type clouds
                '4 coverage
                '5 turbulence
                '6 devation
                '7 icing
                iCloudLayer = iCloudLayer + 1
                If iCloudLayer <= 9 Then
                    lCloudBaseMt = _
                      fnFTOM(nNodes(iCount).Tag(1)) + fnFTOM(nNodes(iCount).Parent.Tag(3))
                    lCloudTopMt = _
                      fnFTOM(nNodes(iCount).Tag(2)) + fnFTOM(nNodes(iCount).Parent.Tag(3))
                    
                    fnReadWxTree.clouds(iCloudLayer) = "  CLOUDS " & iCloudLayer & _
                      ", " & lCloudBaseMt & ", " & lCloudTopMt & ", " & _
                      nNodes(iCount).Tag(3) & ", " & nNodes(iCount).Tag(4) & _
                      ", " & nNodes(iCount).Tag(5) & ", " & _
                      nNodes(iCount).Tag(6) & ", " & nNodes(iCount).Tag(7)
                    
                    fnReadWxTree.cloud_count = iCloudLayer
                                        
                End If
              Case "T" 'TEMPERATURES
                iTempCount = iTempCount + 1
                If iTempCount = 1 Then
                   fnReadWxTree.temperature(iTempCount) = "  TEMPERATURE " & iTempCount & _
                         ", FTOM(" & nNodes(iCount).Parent.Tag(3) & "), " & _
                         fnFTOC(nNodes(iCount).Tag) & _
                         ", 0"
                   fnReadWxTree.temp_count = iTempCount
                End If
              Case "V" 'VISIBILITIES
                fnReadWxTree.visibility = "  VISIBILITY " & nNodes(iCount).Tag
              Case "W" 'WINDS
                    iWindCount = iWindCount + 1
                    If iWindCount = 1 Then
                         'set intial station data
                         'sStation
                        '0- ICAO
                        '1- Station Latitude (Dec)
                        '2- Station Longitude (Dec)
                        '3- Station Elevation (feet)
                        '4- Weather Area Height North to South
                        '5- Weather Area Width  West to East
                        '6- Inrange (turn on weather area)
                        '7-
                        '8-
                        '9- Weather Area Transition
                        '10- Weather Course
                        '11- Weather Velocity
                        '12 -Station Name
                        '13 -weather date time data
                        '14 -ICAO Prefix
                        '15 -Wind Count
                        '16 -Cloud Count
                         
                         fnReadWxTree.station_ID = nNodes(iCount).Parent.Tag(0)
                         fnReadWxTree.miles_to = "MILESTOWX = GROUND_DISTANCE(" & _
                                         nNodes(iCount).Parent.Tag(1) & ", " & _
                                         nNodes(iCount).Parent.Tag(2) & ")"
                        
                         fnReadWxTree.inRange = "IF MILESTOWX < " & nNodes(iCount).Parent.Tag(6) & " THEN"
                        
                        'convert Miles to degrees
                        sngDegWidth = CSng(nNodes(iCount).Parent.Tag(4)) / 40
                        'calculate the beginning and end of the wx area
                        
                        'begining longitude
                        sngBeginLong = CSng(nNodes(iCount).Parent.Tag(2)) + sngDegWidth / 2
                        'beginning latitiude
                        sngBeginLat = CSng(nNodes(iCount).Parent.Tag(1))
                        'ending longitude
                        sngEndLong = CSng(nNodes(iCount).Parent.Tag(2)) - sngDegWidth / 2
                        'ending latitiude
                        sngEndLat = sngBeginLat
     
                    
                         
                         fnReadWxTree.weather = "  WEATHER " & Chr(34) & _
                                              nNodes(iCount).Parent.Tag(0) & Chr(34) & ", " & _
                                              sngBeginLat & ", " & _
                                              sngBeginLong & ", " & _
                                              sngEndLat & ", " & _
                                              sngEndLong
                        
                        '8- Weather Area Width
                        '9- Weather Area Transition
                        '10- Weather Course
                        '11- Weather Velocity
                         fnReadWxTree.weather_char = "  WEATHER_CHAR " & _
                                    nNodes(iCount).Parent.Tag(4) & ", " & _
                                    nNodes(iCount).Parent.Tag(9) & ", " & _
                                    nNodes(iCount).Parent.Tag(10) & ", " & _
                                    nNodes(iCount).Parent.Tag(11)
                         'Wind
                         '0-wind direction
                         '1-wind speed
                         '2-wind type
                         '3-turbulence
                         '4-wind base feet AGL
                         '5-wind top feet  AGL
                         'parent
                         '3- Station Elevation (feet)
                         lWindBaseFt = CLng(nNodes(iCount).Tag(4)) + _
                                       CLng(nNodes(iCount).Parent.Tag(3))
                         lWindTopFt = CLng(nNodes(iCount).Tag(5)) + _
                                       CLng(nNodes(iCount).Parent.Tag(3))
                         fnReadWxTree.winds(iWindCount) = _
                             "  WINDS " & iWindCount & ", FTOM( " & lWindBaseFt & ")" & _
                             ", FTOM( " & lWindTopFt & "), " & _
                             nNodes(iCount).Tag(2) & ", " & nNodes(iCount).Tag(0) & _
                             ", " & nNodes(iCount).Tag(1) & ", " & nNodes(iCount).Tag(3)
                         fnReadWxTree.wind_count = iWindCount
                    ElseIf iWindCount > 1 And iWindCount < 10 Then
                         lWindBaseFt = CLng(nNodes(iCount).Tag(4)) + _
                                       CLng(nNodes(iCount).Parent.Tag(3))
                         lWindTopFt = CLng(nNodes(iCount).Tag(5)) + _
                                       CLng(nNodes(iCount).Parent.Tag(3))
                         fnReadWxTree.winds(iWindCount) = _
                             "  WINDS " & iWindCount & ", FTOM( " & lWindBaseFt & ")" & _
                             ", FTOM( " & lWindTopFt & "), " & _
                             nNodes(iCount).Tag(2) & ", " & nNodes(iCount).Tag(0) & _
                             ", " & nNodes(iCount).Tag(1) & ", " & nNodes(iCount).Tag(3)
                         fnReadWxTree.wind_count = iWindCount
                    End If
             Case " "
           End Select
       End If
       iCount = iCount + 1
    Loop
    
    
End Function
Private Sub cmdMakeADV_Click()
  Dim sFSPath As String
  Dim iCount As Integer
  Dim iSubCount As Integer
  Dim iWxCount As Integer
  Dim tAdvWeather As ADV_WX_TYPE
  Dim sCR As String
  Dim sShell As String
  Dim sAdvFileName As String
  Dim bIsGlobal As Boolean
  
  Screen.MousePointer = vbHourglass
  
  On Error GoTo error_exit
  sCR = Chr(13) + Chr(10)

  Open gsAdvPath & "\aw98.txt" For Output As #1    ' Open file for output.
  
  Print #1, "Title " & Chr(34) & txtAdvTitle.Text & Chr(34)
  Print #1, "Description " & Chr(34) & txtAdvDesc.Text & Chr(34)
  Print #1, ";THIS CODE IS COPYRIGHT 1997 BLUE MESA SOFTWARE"
  'Print #1, "DEBUG_WINDOW ON"
  Print #1, "ADV_KEYS ADD, 557  ;CTRL-X"
  Print #1, "DECLARE WXSECTOR"
  Print #1, "DECLARE INRANGEMILES"
  Print #1, "DECLARE SELECTEDWX"
  Print #1, "DECLARE MILESTOWX"
  Print #1, "DECLARE LAST_PLANE_LAT"
  Print #1, "DECLARE LAST_PLANE_LON"
  Print #1, ""
  Print #1, "SELECTEDWX = 0"
  Print #1, "WXSECTOR = 0"
  Print #1, "INRANGEMILES = 35"
  Print #1, "Scroll " & Chr(34) & "Welcome to Active Weather 98 Generated: " & _
                    Format(Date, "long date") & " at "; Format(Time, "hh:mm:ss AMPM") & _
                    "....Press CTRL-X to exit this adventure. " & Chr(34)
  
  If txtAdvSituation.Text <> "" Then
      Print #1, "LOAD_SITUATION " & Chr(34) & txtAdvSituation.Text & Chr(34)
  End If
  
  Print #1, ""
  Print #1, "While 1"
  Print #1, "ACTIVEWX:"
  Print #1, "  ONKEY 557 GOSUB ENDADV"
  Print #1, "  GoSub DISTANCE"
  Print #1, "ENDWHILE"
  Print #1, " "
  Print #1, "DISTANCE:"
  Print #1, "WXSECTOR=-1"
  iCount = 1
  Do Until iCount > treWeather.Nodes.Count
     If Left(treWeather.Nodes(iCount).Key, 1) = "S" Then
         iWxCount = iWxCount + 1

         tAdvWeather = fnReadWxTree(Mid(treWeather.Nodes(iCount).Key, 2), treWeather.Nodes)

         If tAdvWeather.cloud_count > 2 Then
                 If MsgBox("Flight Simulator only allows up to 2 cloud layers active at" & sCR & _
                       "any given time. Currently your defined weather has " & sCR & _
                        tAdvWeather.cloud_count & _
                       " cloud levels. If you continue and create the adventure, " & sCR & _
                       "FS will ignore one of the cloud layers.  If cancel " & sCR & _
                       "you may delete a cloud layer. Do you want to continue " & sCR & _
                       "and create this adventure?", vbOKCancel) = vbCancel Then
                           Close #1
                           Exit Sub
                  End If
          End If
          If tAdvWeather.station_ID <> "GLB" Then
                Print #1, ";" & tAdvWeather.station_ID
                Print #1, tAdvWeather.miles_to
                Print #1, tAdvWeather.inRange
                Print #1, " WXSECTOR=" & iWxCount
                Print #1, " IF SELECTEDWX!=WXSECTOR THEN"
                Print #1, tAdvWeather.weather
                Print #1, tAdvWeather.weather_char
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.wind_count
                   Print #1, tAdvWeather.winds(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.cloud_count
                   Print #1, tAdvWeather.clouds(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                Print #1, tAdvWeather.visibility
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.temp_count
                   Print #1, tAdvWeather.temperature(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                Print #1, tAdvWeather.baro_pressure
                Print #1, "  SELECTEDWX=" & iWxCount
                Print #1, " ENDIF"
                Print #1, "ENDIF"
                Print #1, " "
           End If
        End If
     iCount = iCount + 1
  Loop

  
  
 
  
  iCount = 1
  Do Until iCount > treWeather.Nodes.Count
     If Left(treWeather.Nodes(iCount).Key, 1) = "S" Then
         iWxCount = iWxCount + 1
         tAdvWeather = fnReadWxTree(Mid(treWeather.Nodes(iCount).Key, 2), treWeather.Nodes)
         If tAdvWeather.station_ID = "GLB" Then
                bIsGlobal = True
                Print #1, ";OUT OF RANGE OF ALL WX STATIONS"
                Print #1, "If WXSECTOR = -1 THEN"
                Print #1, "  If SELECTEDWX!=WXSECTOR THEN"
                Print #1, "     GOSUB GLOBAL_WX"
                Print #1, "  EndIf"
                Print #1, "  MILESTOWX = GROUND_DISTANCE(LAST_PLANE_LAT, LAST_PLANE_LON)"
                Print #1, "  If MILESTOWX > INRANGEMILES Then"
                Print #1, "      GoSub GLOBAL_WX"
                Print #1, "  ENDIF"
                Print #1, "ENDIF"
                Print #1, "WAIT 60"
                Print #1, "RETURN"
                Print #1, ""
                Print #1, ""
                Print #1, "GLOBAL_WX:"
                Print #1, "  LAST_PLANE_LAT = PLANE_LAT"
                Print #1, "  LAST_PLANE_LON = PLANE_LON"
                Print #1, "  ;" & tAdvWeather.station_ID
                Print #1, "  WEATHER " & Chr(34) & _
                          "GLB" & Chr(34) & _
                          ", PLANE_LAT, PLANE_LON + 50, PLANE_LAT, PLANE_LON - 50"
                Print #1, "  WEATHER_CHAR 300, 200, 0, 0"
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.wind_count
                   Print #1, tAdvWeather.winds(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.cloud_count
                   Print #1, tAdvWeather.clouds(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                Print #1, tAdvWeather.visibility
                iSubCount = 1
                Do Until iSubCount > tAdvWeather.temp_count
                   Print #1, tAdvWeather.temperature(iSubCount)
                   iSubCount = iSubCount + 1
                Loop
                Print #1, tAdvWeather.baro_pressure
                Print #1, "  SELECTEDWX = -1"
                Print #1, "  RETURN"
                Print #1, ""
           End If
        End If
     iCount = iCount + 1
  Loop
  
  If Not bIsGlobal Then
       Print #1, "WAIT 60"
       Print #1, "RETURN"
  End If
  
  Print #1, ""
  Print #1, "ENDADV:"
  Print #1, " END"
  Close #1
  
  sbCompileADV
  Screen.MousePointer = vbDefault
  Exit Sub
error_exit:
  Screen.MousePointer = vbDefault
  
  MsgBox "The make of the ADV failed with: " & sCR & Err.Description & " " & Err.Number
  Close #1
End Sub

Private Sub cmdMetarFile_Click()
 ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo errhandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtMetarPath.Text = CommonDialog1.filename
    Exit Sub

errhandler:

End Sub








Private Sub sbReadWxPlus(iTypeRun As Integer)
  Dim sInputLine As String
  Dim iPOS As Integer
  Dim sMetarLines() As String
  Dim sSelStations() As String
  Dim iCount As Integer
  Dim iSubCount As Integer
  Dim rs As Recordset
  Dim sSQL As String
  Dim lRCode As Long
  
  ReDim sMetarLines(1)
  ReDim sSelStations(1)
  sCR = Chr(13) + Chr(10)

  sICAO = txtICAO.Text & " "
 
  '==============================================================================
  iCount = 0
  Dim sWxFile As String
  
  If iTypeRun = 0 Then
    Open txtMetarPath.Text For Input As #1 ' Open file.
  Else  'opening
    'path of file being opened
    Open sFilePath For Input As #1 ' Open file.
    mdiMain.sbSaveFiveFiles sFilePath
    mdiMain.sbReadFiveFiles
  End If
  'Add the weather stations to the adventure
  Do While Not EOF(1)
       Line Input #1, sInputLine
       If Mid(sInputLine, 5, 1) = " " And IsNumeric(Mid(sInputLine, 6, 1)) Then
           lRCode = fnAddStation(Mid(sInputLine, 2, 3), Left(sInputLine, 1), 1)
           If lRCode = 1 Or lRCode = 35602 Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = sInputLine 'minus the new and -1 diff
                iCount = iCount + 1
           End If
        ElseIf InStr(1, sInputLine, "Z ") > 12 Then
             lRCode = fnAddStation(Mid(sInputLine, InStr(1, sInputLine, "Z ") - 10, 3), _
                      Mid(sInputLine, InStr(1, sInputLine, "Z ") - 11, 1), 1)
             If lRCode = 1 Or lRCode = 35602 Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = _
                           Mid(sLine, InStr(1, sInputLine, "Z ") - 11)
                iStationCount = iStationCount + 1
                mdiMain.SBar1.SimpleText = iStationCount
                DoEvents
             End If
       End If
  iSubCount = iSubCount + 1
  Loop
  
  mdiMain.SBar1.SimpleText = iCount & " station(s) found"
  Close #1
  
  sbReadMetarLines sMetarLines
  

  
End Sub
Public Sub sbReadWxWebPlus()
  Dim sLine As String
  Dim sMetarLines() As String
  Dim iCount As Long
  Dim iStationCount As Integer
  Dim iLastBeginPOS As Long
  Dim lRCode As Long
  
  ReDim sMetarLines(1)
'  treWeather.Nodes.Clear
 
  '==============================================================================
  iCount = 1
  iLastBeginPOS = 1
  Do Until iCount > Len(strHTML) Or iStationCount > MAX_STATION_COUNT
     'if chr10 then break the line
     If Asc(Mid(strHTML, iCount, 1)) = 10 Then
          sLine = Mid(strHTML, iLastBeginPOS, iCount - iLastBeginPOS)
        
           '==========add stations to adventure============================
           If Mid(sLine, 5, 1) = " " And IsNumeric(Mid(sLine, 6, 1)) Then
            
             lRCode = fnAddStation(Mid(sLine, 2, 3), Left(sLine, 1), 1)
             If lRCode = 1 Or lRCode = 35602 Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = sLine 'minus the new and -1 diff
                iStationCount = iStationCount + 1
                mdiMain.SBar1.SimpleText = iStationCount
                DoEvents
             End If
           
           ElseIf InStr(1, sLine, "Z ") > 12 Then
             lRCode = fnAddStation(Mid(sLine, InStr(1, sLine, "Z ") - 10, 3), _
                      Mid(sLine, InStr(1, sLine, "Z ") - 11, 1), 1)
             If lRCode = 1 Or lRCode = 35602 Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = _
                           Mid(sLine, InStr(1, sLine, "Z ") - 11)
                iStationCount = iStationCount + 1
                mdiMain.SBar1.SimpleText = iStationCount
                DoEvents
             End If
           End If
           
           
           iLastBeginPOS = iCount + 1
     End If
     iCount = iCount + 1
  Loop
   
  mdiMain.SBar1.SimpleText = iStationCount & " station(s) found"
 
  sbReadMetarLines sMetarLines
  

  
End Sub

Private Function fnReadWeather() As Integer
  Dim sInputLine As String
  Dim iPOS As Integer
  Dim sMetarLines() As String
  Dim sSelStations() As String
  Dim iCount As Integer
  Dim iSubCount As Integer
  
  
  ReDim sMetarLines(1)
  ReDim sSelStations(1)
  sCR = Chr(13) + Chr(10)
'  txtALC.Text = ""
  sICAO = txtICAO.Text & " "
  
  '===Read the selected stations from the tree====================================
  iCount = 1
  iSubCount = 1
  Do Until iCount > treWeather.Nodes.Count
     If Left(treWeather.Nodes(iCount).Key, 1) = "S" Then
            sSelStations(iSubCount) = Mid(treWeather.Nodes(iCount).Key, 2)
            ReDim Preserve sSelStations(UBound(sSelStations) + 1)
            iSubCount = iSubCount + 1
     End If
     iCount = iCount + 1
  Loop
   
  '==============================================================================
  iCount = 0
  
  Open txtMetarPath.Text For Input As #1 ' Open file.


  Do While Not EOF(1)
        Line Input #1, sInputLine
        iSubCount = 1
        Do Until iSubCount > UBound(sSelStations) - 1
            iPOS = InStr(1, sInputLine, sSelStations(iSubCount))
            If iPOS > 0 And IsNumeric(Mid(sInputLine, iPOS + 5, 1)) Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = sInputLine 'minus the new and -1 diff
            End If
            iSubCount = iSubCount + 1
        Loop
  Loop
  Close #1
  
 
 sbReadMetarLines sMetarLines
  
 
  

End Function
Private Sub sbReadMetarLines(sMetarLines() As String)
  Dim iCharPOS As Integer
  Dim sTempStr As String
  Dim bVisSet As Boolean
  Dim sICAO As String
  Dim sCR As String
  Dim sOutput As String
  Dim sPrevTempStr As String
  
  'process the lines
  iCharPOS = 1
  iCount = 1
  Do Until iCount > UBound(sMetarLines) - 1
     'write the station ALC line

'     Debug.Print sMetarLines(iCount)
     Do Until iCharPOS > Len(sMetarLines(iCount))
      
         If Mid(sMetarLines(iCount), iCharPOS, 1) = Chr(32) Or _
            iCharPOS = Len(sMetarLines(iCount)) Then
            
               'correct for end of line by adding last byte
               If iCharPOS = Len(sMetarLines(iCount)) Then
                   sTempStr = sTempStr & Right(sMetarLines(iCount), 1)
               End If
               If iCharPOS = 5 Then
                   Call fnStationID(sTempStr)
               End If
               'time section
               If Right(sTempStr, 1) = "Z" Then  'time section
                   Call fnTime(Mid(sTempStr, 2))
               End If
               'Wind section
               If Right(sTempStr, 2) = "KT" Then  'wind section
                  Call fnWind(Mid(sTempStr, 2))
               End If
                'visibility section
               If Right(sTempStr, 2) = "SM" Then
                  If InStr(1, sTempStr, "/") > 0 Then
                      If Right(sPrevTempStr, 2) = "KT" Then
                         Call fnVisibility(Mid(sTempStr, 2))
                      Else
                         Call fnVisibility(Mid(sPrevTempStr, 2) & " " & Mid(sTempStr, 2))
                      End If
                  Else
                     Call fnVisibility(Mid(sTempStr, 2))
                  End If
               End If
               
               'cloud section for standard types
               Select Case Mid(sTempStr, 2, 3)
                 Case "OVC"
                   Call fnClouds(Mid(sTempStr, 2))
                 Case "FEW"
                   Call fnClouds(Mid(sTempStr, 2))
                 Case "SCT"
                   Call fnClouds(Mid(sTempStr, 2))
                 Case "BKN"
                   Call fnClouds(Mid(sTempStr, 2))
                 Case "CLR"
                  Call fnClouds(Mid(sTempStr, 2))
               End Select
               'cloud section for vertical vis
               If Mid(sTempStr, 2, 2) = "VV" Then
                  Call fnClouds(Mid(sTempStr, 2))
               End If
               
               'Altimeter section
               If Mid(sTempStr, 2, 1) = "A" And Len(sTempStr) = 6 Then
                   sOutput = sOutput & fnAltimeter(Mid(sTempStr, 2)) & sCR
               End If
               
               'tempurature section
               If InStr(1, sTempStr, "/") > 0 And Right(sTempStr, 2) <> "SM" _
                 And Len(sTempStr) <= 8 Then
                  sOutput = sOutput & fnTemp(Mid(sTempStr, 2)) & sCR
               End If
                             
               'Remark to end
               If Mid(sTempStr, 2, 3) = "RMK" Then
                 iCharPOS = Len(sMetarLines(iCount)) + 1
               End If
               sPrevTempStr = sTempStr
               sTempStr = ""
          End If
          sTempStr = sTempStr & Mid(sMetarLines(iCount), iCharPOS, 1)
          iCharPOS = iCharPOS + 1
     Loop
     
     giCloudCount = 0
     giWindCount = 0
     giTempCount = 0
     sTempStr = ""
     bVisSet = False
     iCount = iCount + 1
     iCharPOS = 1
  Loop
 
End Sub
Private Sub sbReadWxWeb()
  Dim sInputLine As String
  Dim iPOS As Integer
  Dim sMetarLines() As String
  Dim sSelStations() As String
  Dim iCount As Integer
  Dim iSubCount As Integer
  
  
  ReDim sMetarLines(1)
  ReDim sSelStations(1)
  sCR = Chr(13) + Chr(10)
'  txtALC.Text = ""
  sICAO = txtICAO.Text & " "
  
  '===Read the selected stations from the tree====================================
  iCount = 1
  iSubCount = 1
  Do Until iCount > treWeather.Nodes.Count
     If Left(treWeather.Nodes(iCount).Key, 1) = "S" Then
            sSelStations(iSubCount) = Mid(treWeather.Nodes(iCount).Key, 2)
            ReDim Preserve sSelStations(UBound(sSelStations) + 1)
            iSubCount = iSubCount + 1
     End If
     iCount = iCount + 1
  Loop
   
  '==============================================================================
    
 
  iCount = 1
  iLastBeginPOS = 1
  Do Until iCount > Len(strHTML)
     'if chr10 then break the line
     If Asc(Mid(strHTML, iCount, 1)) = 10 Then
          sLine = Mid(strHTML, iLastBeginPOS, iCount - iLastBeginPOS)

        
         iSubCount = 1
         Do Until iSubCount > UBound(sSelStations) - 1
            iPOS = InStr(1, sLine, sSelStations(iSubCount))
            If iPOS > 0 And IsNumeric(Mid(sLine, iPOS + 5, 1)) Then
                ReDim Preserve sMetarLines(UBound(sMetarLines) + 1)
                sMetarLines(UBound(sMetarLines) - 1) = sLine 'minus the new and -1 diff
            End If
            iSubCount = iSubCount + 1
         Loop
  
     iLastBeginPOS = iCount + 1
     End If
     iCount = iCount + 1
  Loop
 
 sbReadMetarLines sMetarLines
   
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
Private Function fnAddStation(sIATA As String, sICAOPrefix As String, _
                              iTypeRun As Integer) As Long
     Dim nodX As Node
     Dim sStationName As String
     Dim sFullStationName As String
     Dim sStationID As String
     Dim sStation(16) As String
     Dim sDescID As String
     Dim sngDegWidth As Single
     Dim sngLong As Single
     Dim sngLat As Single
          
     If iTypeRun = 0 Then
            On Error GoTo err_dup
     End If
     If iTypeRun = 1 Then
           On Error GoTo err_exit
     End If
        
     'sStation
     '0- ICAO
     '1- Station Latitude (Dec)
     '2- Station Longitude (Dec)
     '3- Station Elevation (feet)
     '4- Weather Area Height
     '5- Weather Area Width
     '6- Inrange (turn on weather area)
     '7-
     '8-
     '9- Weather Area Transition
     '10- Weather Course
     '11- Weather Velocity
     '12 -Station Name
     '13 -weather date time data
     '14 -ICAO Prefix
     '15 -Wind Count
     '16 -Cloud Count
     
     sStation(0) = sIATA
     
     sSQL = "select * from Airports where IATA='" & sIATA & "' AND ICAO_Prefix='" & _
             sICAOPrefix & "'"
  
     Set rs = gDB.OpenRecordset(sSQL, dbOpenSnapshot)
          
     If rs.EOF Then GoTo err_exit
     'convert station lat to dec
     sngLat = fnConvertLatLon(rs!Latitude, TYPE_LATITUDE)
     sStation(1) = sngLat
     'convert station long to dec
     sngLong = fnConvertLatLon(rs!Longitude, TYPE_LONGITUDE)
     sStation(2) = sngLong
     
     sStation(3) = rs!Elevation
     sStation(4) = rs!Height
     sStation(5) = rs!Width
     sStation(6) = rs!inRange
     
     sStationName = Trim(rs!facilityname)
     rs.Close
         
     sStation(9) = "30"
     sStation(10) = "0"
     sStation(11) = "0"
     
     sStationID = "S" & sIATA
     sFullStationName = sIATA & "(" & sStationName & ")"
     sStation(12) = Trim(sStationName)
     sStation(15) = "0"
     sStation(16) = "0"
     
     Set nodX = treWeather.Nodes.Add(, , sStationID, sFullStationName)
     nodX.Image = "stationic"
     nodX.Tag = sStation
     nodX.Sorted = True
     
'     nodX.EnsureVisible
     fnAddStation = 1
     Exit Function
err_dup:
   MsgBox "You have already selected this station.", vbExclamation
   Exit Function
err_exit:
   If Err.Number = 35602 Then
        fnAddStation = Err.Number
   Else
        fnAddStation = -1
   End If
End Function

Private Sub cmdProcURL_Click()
  
End Sub

Private Sub cmdNext_Click()
'  select gi
End Sub

Private Sub cmdProcWebWx_Click()
 If optProcURL.Value = True Then
   sbReadWxWeb
 Else
   sbReadWxWebPlus
 End If
End Sub

Private Sub cmdStop_Click()
  sbStopInet
End Sub

Private Sub cmdSit_Click()
   Dim iFileNameBegin As Integer
   Dim iFileNameLen As Integer
 
 ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo errhandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Situation Files (*.stn)|*.stn"
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    iFileNameBegin = Len(CommonDialog1.filename) - _
                     (InStr(1, fnReverseString(CommonDialog1.filename), "\") - 2)
    iFileNameLen = (Len(CommonDialog1.filename) - 3) - iFileNameBegin
    txtAdvSituation.Text = Mid(CommonDialog1.filename, iFileNameBegin, iFileNameLen)
    Exit Sub

errhandler:
End Sub

Private Sub cmdStations_Click()
         labTitle.Caption = "Add Stations"
         labTitleDesc.Caption = "Add weather stations to adventure"
         fraStations.Visible = True
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = False
         giCurrentTip = TIP_STATIONS
        
End Sub

Private Sub cmdViewHTML_Click()
  frmWeb.sbLoad_Form strHTML
End Sub

Private Function fnFindCloudType(oCloud As Object, sCloudType As String) As Integer
  Dim iCount As Integer
  Do Until iCount > oCloud.ListCount - 1
     If sCloudType = oCloud.List(iCount) Then
        fnFindCloudType = iCount
        Exit Function
     End If
     iCount = iCount + 1
  Loop
End Function
Private Function fnFindWindType(oWind As Object, sWindType As String) As Integer
  Dim iCount As Integer
  Do Until iCount > oWind.ListCount - 1
     If sWindType = oWind.List(iCount) Then
        fnFindWindType = iCount
        Exit Function
     End If
     iCount = iCount + 1
  Loop
End Function
Private Function fnFindCloudCov(oCloud As Object, sCloudType As String) As Integer
  Dim iCount As Integer
  Do Until iCount > oCloud.ListCount - 1
     If sCloudType = oCloud.List(iCount) Then
        fnFindCloudCov = iCount
        Exit Function
     End If
     iCount = iCount + 1
  Loop
End Function





Private Sub Command4_Click()

End Sub

Private Sub cmdWx_Click()
         labTitle.Caption = "Metar Weather Data"
         labTitleDesc.Caption = "Process/Update adventure wx data"
         fraStations.Visible = False
         fraWeb.Visible = True
         fraFiles.Visible = True
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = False
         giCurrentTip = TIP_WX_DATA
       
End Sub

Private Sub Command1_Click()
  Open "c:\wx\cos.txt" For Input As #1 ' Open file.
'  Open "c:\temp\temp.txt" For Output As #2    ' Open file for output.
  Dim sInputLine As String
  Do While Not EOF(1)
        Input #1, sInputLine
        strHTML = strHTML & sInputLine
'        Debug.Print sInputLine
   Loop
 Close #1
 cmdProcWebWx.Enabled = True
 cmdViewHTML.Enabled = True
 
' iCount = 1
' Do Until iCount > Len(sInputLine)
'     If Asc(Mid(sInputLine, iCount, 1)) = 10 Then
'          MsgBox iCount
'     End If
'     iCount = iCount + 1
' Loop
'
' 'Close #2
End Sub

Private Sub Form_Activate()
   If sACWPath <> "" Then
        mdiMain.mnuFileSave.Enabled = True
   End If
   mdiMain.mnuFileSaveAs.Enabled = True
   mdiMain.mnuWxCloud.Enabled = True
   mdiMain.mnuWxTemp.Enabled = True
   mdiMain.mnuWxVis.Enabled = True
   mdiMain.mnuWxWind.Enabled = True
   mdiMain.mnuEditPaste.Enabled = True
   mdiMain.sbWx_Buttons BUTTON_DISPLAY_ENABLED
End Sub

Private Sub Form_Deactivate()
   mdiMain.mnuFileSave.Enabled = False
   mdiMain.mnuFileSaveAs.Enabled = False
   mdiMain.SBar1.SimpleText = ""
   mdiMain.mnuFileSaveAs.Enabled = False
   mdiMain.mnuWxCloud.Enabled = False
   mdiMain.mnuWxTemp.Enabled = False
   mdiMain.mnuWxVis.Enabled = False
   mdiMain.mnuWxWind.Enabled = False
   mdiMain.mnuEditPaste.Enabled = False
   mdiMain.sbWx_Buttons BUTTON_DISPLAY_DISABLED
End Sub

Private Sub Form_Load()
   With Me
    .Top = 1
    .Left = 1
    .Height = 6360
    .Width = 11040
   End With
      
   bIsLoad = True
   oldIndex = 100
   treWeather.ImageList = ImageList1
   sACWPath = ""
   
   bManualURL = False
   
   sbSetToolTips
   
   mdiMain.mnuFileSaveAs.Enabled = True
      
   Call fnLoad_Regions
   fnReadRegSettings
   sbLoadTenURLs
   sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_IATA
 
   
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
   cboCloudCov.ListIndex = 0
   
   cboCloudIce.AddItem "None"
   cboCloudIce.AddItem "Cloud Icing"
   cboCloudIce.ListIndex = 0
   
   cboWindType.AddItem "steady"
   cboWindType.AddItem "gusty"
   
   txtCloudDev.Text = "0"
   txtCloudTurb.Text = "0"
   bIsLoad = False
  
End Sub

Private Sub lstIATA_Click()
  lstStations.ListIndex = lstIATA.ListIndex
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
'  Call MoveMouse(100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   sbSaveRegSettings
   
   mdiMain.mnuFileSave.Enabled = False
   mdiMain.mnuFileSaveAs.Enabled = False
   mdiMain.mnuFileSave.Enabled = False
   mdiMain.mnuFileSaveAs.Enabled = False
   mdiMain.SBar1.SimpleText = ""
   mdiMain.mnuFileSaveAs.Enabled = False
   mdiMain.mnuWxCloud.Enabled = False
   mdiMain.mnuWxTemp.Enabled = False
   mdiMain.mnuWxVis.Enabled = False
   mdiMain.mnuWxWind.Enabled = False
   mdiMain.mnuEditPaste.Enabled = False
   mdiMain.sbWx_Buttons BUTTON_DISPLAY_DISABLED
End Sub

Private Sub fraClouds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
End Sub

Private Sub fraFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
End Sub

Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
End Sub

Private Sub fraWeather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
'  MoveMouse 100
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    ' Retrieve server response using the GetChunk
    ' method when State = 12. This example assumes the
    ' data is text.
    Debug.Print State
    On Error GoTo error_exit
    Select Case State
        Case icHostResolvingHost
             frmStatus.labInetStatus.Caption = "Looking up host..."
        Case icHostResolved
             frmStatus.labInetStatus.Caption = "Host found..."
        Case icConnecting
             frmStatus.labInetStatus.Caption = "Connecting to host..."
        Case icConnected
             frmStatus.labInetStatus.Caption = "Connected to host..."
        Case icRequesting
             frmStatus.labInetStatus.Caption = "Sending request to host..."
        Case icRequestSent
             frmStatus.labInetStatus.Caption = "Request sent..."
        Case icReceivingResponse
             frmStatus.labInetStatus.Caption = "Connected to Host. Waiting for response..."
        Case icError
             sbStopInet
             MsgBox "Error connecting to URL.", vbExclamation
        Case icDisconnected
             
        Case icResponseCompleted ' 12
            Dim strData As Variant
            Dim bDone As Boolean: bDone = False
           ' Get first chunk.
            strData = Inet1.GetChunk(2048, icString)
            DoEvents
            Do While Not bDone
                strHTML = strHTML & strData
                ' Get next chunk.
                strData = Inet1.GetChunk(2048, icString)
                DoEvents
                If Len(strData) = 0 Then
                    bDone = True
                End If
            Loop
            On Error Resume Next
            Kill "c:\temp\temp.txt"
            Open "c:\temp\temp.txt" For Output As #1    ' Open file for output.
            Print #1, strHTML
            Close #1
            
            sbStopInet
        End Select
Exit Sub
error_exit:
   sbStopInet
End Sub


Private Sub lstICAO_Click()
  lstStations.ListIndex = lstICAO.ListIndex
End Sub

Private Sub lstICAO_DblClick()
   cmdAddWeather_Click
End Sub

Private Sub lstStations_Click()
  lstICAO.ListIndex = lstStations.ListIndex
End Sub

Private Sub optCGI_Click()
     cboURL.Enabled = False
     cboURL.BackColor = INACTIVE_COLOR
     txtForm.BackColor = ACTIVE_COLOR
     txtForm.Enabled = True
     labForms.Enabled = True
     txtCGIStation.BackColor = ACTIVE_COLOR
     txtCGIStation.Enabled = True
     labCGIStation.Enabled = True
     cmdEdit.Enabled = True
End Sub

Private Sub optSortStaIATA_Click()
  If optSortStaIATA.Value = True Then
     sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_IATA
  End If
End Sub

Private Sub optSortStaName_Click()
   If optSortStaName.Value = True Then
     sbLoad_Stations cboRegions.List(cboRegions.ListIndex), SORT_STATION_NAME
  End If
End Sub

Private Sub optURL_Click()
     cboURL.Enabled = True
     cboURL.BackColor = ACTIVE_COLOR
     txtForm.BackColor = INACTIVE_COLOR
     txtForm.Enabled = False
     labForms.Enabled = False
     txtCGIStation.BackColor = INACTIVE_COLOR
     txtCGIStation.Enabled = False
     labCGIStation.Enabled = False
     cmdEdit.Enabled = False

End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call DownMouse(Index)
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 0 Then
    Call MoveMouse(Index)
  End If
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call UpMouse(Index)
   Select Case Index
     
      Case 1 'file and web
        
      Case 2 'General adv options
        
      Case 3 'clouds
'         labTitle.Caption = "Cloud Layers"
'         labTitleDesc.Caption = "Set or add cloud layers"
'         fraStations.Visible = False
'         fraWeb.Visible = False
'         fraFiles.Visible = False
'         fraGeneral.Visible = False
'         fraClouds.Visible = True
'         fraStation.Visible = False
'         fraWind.Visible = False
     Case 4 'winds
'         labTitle.Caption = "Wind Layers"
'         labTitleDesc.Caption = "Set or add wind layers"
'         fraStations.Visible = False
'         fraWeb.Visible = False
'         fraFiles.Visible = False
'         fraGeneral.Visible = False
'         fraClouds.Visible = False
'         fraStation.Visible = False
'         fraWind.Visible = True
      Case 7 'station
'         labTitle.Caption = "Station"
'         labTitleDesc.Caption = "View or change station data"
'         fraStations.Visible = False
'         fraWeb.Visible = False
'         fraFiles.Visible = False
'         fraGeneral.Visible = False
'         fraClouds.Visible = False
'         fraStation.Visible = True
'         fraWind.Visible = False
   End Select
End Sub

Private Sub treWeather_KeyDown(KeyCode As Integer, Shift As Integer)
  'Processes the delete key being pressed
  If treWeather.SelectedItem Is Nothing Then Exit Sub
  If KeyCode = 46 Then
      treWeather.Nodes.Remove (treWeather.SelectedItem.Index)
      sbRefresh_Selected_Screen treWeather.SelectedItem
  End If
End Sub

Private Sub treWeather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mdiMain.sbClear_Buttons
'  MoveMouse 100
End Sub

Private Sub treWeather_NodeClick(ByVal Node As ComctlLib.Node)
 
'  On Error GoTo error_exit
  
  If Left(Node.Key, 1) = "S" Then
         labTitle.Caption = "Weather Area Data"
         labTitleDesc.Caption = "Edit adventure wx area data"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = True
         fraWind.Visible = False
         fraTemp.Visible = False
         sbSetStationControls Node.Tag
  End If
  
  Select Case Mid(Node.Key, 5, 1)
     Case "C" 'clouds
         Dim sCloudCov As String
         Dim sCloudType As String
         labTitle.Caption = "Edit Cloud Layer"
         labTitleDesc.Caption = "Edit selected cloud layer"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = True
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = False
          '1 raw base feet (AGL)
          '2 raw top feet  (AGL)
          '3 type clouds
          '4 coverage
          '5 turbulence
          '6 devation
          '7 icing
          '8-layer
          txtCloudBase.Text = Format(Node.Tag(1), "###0")
          txtCloudTop.Text = Node.Tag(2)
          sCloudType = Node.Tag(3)
          cboCloudType.ListIndex = fnFindCloudType(cboCloudType, sCloudType)
          sCloudCov = Node.Tag(4)
          cboCloudCov.ListIndex = fnFindCloudCov(cboCloudCov, sCloudCov)
          txtCloudTurb.Text = Node.Tag(5)
          txtCloudDev.Text = Node.Tag(6)
          cboCloudIce.ListIndex = Node.Tag(7)
          txtCloudLayer.Text = Node.Tag(8)
     Case "W" 'winds
         Dim sWindType As String
         labTitle.Caption = "Edit Wind Layer"
         labTitleDesc.Caption = "Edit selected wind layer"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = True
         fraTemp.Visible = False
        '0-wind direction
        '1-wind speed
        '2-wind type
        '3-turbulence
        '4-wind base feet
         '5-wind top feet
          txtWindLayer.Text = Node.Tag(6)
          txtWindDir.Text = Node.Tag(0)
          txtWindSpeed.Text = Node.Tag(1)
          sWindType = Node.Tag(2)
          cboWindType.ListIndex = fnFindWindType(cboWindType, sWindType)
          txtWindTurb.Text = Node.Tag(3)
          txtWindBase.Text = Format(Node.Tag(4), "###0")
          txtWindTop.Text = Node.Tag(5)
     Case "A" 'altimeter
         labTitle.Caption = "Edit Altimeter"
         labTitleDesc.Caption = "Edit selected altimeter setting"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
          
         txtAltimeter.Text = Node.Tag
         txtAltimeter.BackColor = ACTIVE_COLOR
         txtAltimeter.Enabled = True
         txtTemp.Text = ""
         txtTemp.BackColor = INACTIVE_COLOR
         txtTemp.Enabled = False
         txtVis.Text = ""
         txtVis.BackColor = INACTIVE_COLOR
         txtVis.Enabled = False
      Case "T" 'Temp
         labTitle.Caption = "Edit Temperature"
         labTitleDesc.Caption = "Edit selected temperature"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
         txtAltimeter.Text = ""
         txtAltimeter.BackColor = INACTIVE_COLOR
         txtAltimeter.Enabled = False
         txtTemp.Text = Node.Tag
         txtTemp.BackColor = ACTIVE_COLOR
         txtTemp.Enabled = True
         txtVis.Text = ""
         txtVis.BackColor = INACTIVE_COLOR
         txtVis.Enabled = False
  
   Case "V" 'visibility
         labTitle.Caption = "Edit Visibility"
         labTitleDesc.Caption = "Edit selected visibility"
         fraStations.Visible = False
         fraWeb.Visible = False
         fraFiles.Visible = False
         fraGeneral.Visible = False
         fraClouds.Visible = False
         fraStation.Visible = False
         fraWind.Visible = False
         fraTemp.Visible = True
          
         txtAltimeter.BackColor = INACTIVE_COLOR
         txtAltimeter.Enabled = True
         txtTemp.Text = ""
         txtTemp.BackColor = INACTIVE_COLOR
         txtTemp.Enabled = False
         txtVis.Text = Node.Tag
         txtVis.BackColor = ACTIVE_COLOR
         txtVis.Enabled = True
  End Select
  Exit Sub
error_exit:
  MsgBox Err.Description
End Sub

Private Sub txtAltimeter_LostFocus()
 sbChangeAltimeter
End Sub

Private Sub txtCloudBase_LostFocus()
 sbChangeCloud
End Sub

Private Sub txtCloudDev_LostFocus()
  sbChangeCloud
End Sub

Private Sub txtCloudLayer_LostFocus()
   sbChangeCloud
End Sub

Private Sub txtCloudTop_LostFocus()
   sbChangeCloud
End Sub

Private Sub txtCloudTurb_LostFocus()
  sbChangeCloud
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

Private Sub txtMetarPath_Change()
  If Len(txtMetarPath) > 4 Then
     cmdImport.Enabled = True
  End If
End Sub

Private Sub txtStationElev_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationICAO_LostFocus()
 sbChangeStation
End Sub

Private Sub txtStationInRange_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationLat_LostFocus()
   sbChangeStation
End Sub

Private Sub txtStationLong_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationN_LostFocus()
  sbChangeStation
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

Private Sub txtStationWxBLat_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationWxBLong_LostFocus()
   sbChangeStation
End Sub

Private Sub txtStationWxELat_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationWxELong_LostFocus()
   sbChangeStation
End Sub

Private Sub txtStationWxHeight_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationWxTran_LostFocus()
   sbChangeStation
End Sub

Private Sub txtStationWxVol_LostFocus()
   sbChangeStation
End Sub

Private Sub txtStationWxWid_LostFocus()
  sbChangeStation
End Sub

Private Sub txtStationWxWidth_LostFocus()
  sbChangeStation
End Sub

Private Sub txtTemp_LostFocus()
 sbChangeTemp
End Sub

Private Sub txtVis_LostFocus()
 sbChangeVis
End Sub

Private Sub txtWindBase_LostFocus()
 sbChangeWind
End Sub

Private Sub txtWindDir_LostFocus()
  sbChangeWind
End Sub

Private Sub txtWindSpeed_LostFocus()
  sbChangeWind
End Sub

Private Sub txtWindTop_LostFocus()
  sbChangeWind
End Sub

Private Sub txtWindTurb_LostFocus()
  sbChangeWind
End Sub
