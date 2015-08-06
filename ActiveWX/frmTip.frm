VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmTip 
   Caption         =   "Fast Tips"
   ClientHeight    =   4320
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5700
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4065
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   4005
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   120
      Width           =   4155
      Begin RichTextLib.RichTextBox rtxTip 
         Height          =   3405
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   6006
         _Version        =   327681
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmTip.frx":030A
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   2
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4380
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
  
End Sub





Private Sub cmdOK_Click()
    Unload Me
End Sub
Public Sub Load_Tip(iCurrentTip As Integer)
   On Error GoTo err_exit
   Select Case iCurrentTip
        Case TIP_GENERAL_ADV
           lblTitle.Caption = "Adventure Settings"
           rtxTip.LoadFile App.Path & "\adv1.rtf", rtfRTF
        Case TIP_STATIONS
           lblTitle.Caption = "Weather Stations"
           rtxTip.LoadFile App.Path & "\wxsta.rtf", rtfRTF
        Case TIP_WX_DATA
           lblTitle.Caption = "Process Weather Data"
           rtxTip.LoadFile App.Path & "\wxdata.rtf", rtfRTF
   End Select
   Me.Show vbModal
   Exit Sub
err_exit:
   Unload Me
End Sub
Public Sub DisplayCurrentTip()
 
End Sub

Private Sub Form_Load()
 Me.Left = 10
 Me.Top = 10
End Sub

