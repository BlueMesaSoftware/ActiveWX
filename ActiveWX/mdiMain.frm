VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Active Weather 98"
   ClientHeight    =   6345
   ClientLeft      =   1380
   ClientTop       =   1860
   ClientWidth     =   10245
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin ComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6090
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5790
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   683
      TabIndex        =   0
      Top             =   0
      Width           =   10245
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   4740
         Picture         =   "mdiMain.frx":0442
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   11
         ToolTipText     =   "Fast Tips"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   1200
         Picture         =   "mdiMain.frx":0544
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   10
         ToolTipText     =   "Save Wx"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   4140
         Picture         =   "mdiMain.frx":0646
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   9
         ToolTipText     =   "Add Visibility"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   3780
         Picture         =   "mdiMain.frx":0988
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   8
         ToolTipText     =   "Add Temperature"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   3390
         Picture         =   "mdiMain.frx":0CCA
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   7
         ToolTipText     =   "Add Wind Layer"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3000
         Picture         =   "mdiMain.frx":100C
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   6
         ToolTipText     =   "Add Cloud Layer"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   1890
         Picture         =   "mdiMain.frx":110E
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   5
         ToolTipText     =   "Add Station to DB"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   2400
         Picture         =   "mdiMain.frx":1450
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   4
         ToolTipText     =   "Delete Station from DB"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   390
         Picture         =   "mdiMain.frx":1792
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   2
         ToolTipText     =   "New Wx"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   780
         Picture         =   "mdiMain.frx":1894
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   1
         ToolTipText     =   "Open Wx"
         Top             =   120
         Width           =   285
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000014&
         X1              =   1
         X2              =   1
         Y1              =   1
         Y2              =   29
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000010&
         X1              =   303
         X2              =   303
         Y1              =   2
         Y2              =   30
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   188
         X2              =   188
         Y1              =   0
         Y2              =   28
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   189
         X2              =   189
         Y1              =   1
         Y2              =   29
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   792
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   792
         X2              =   0
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   794
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   304
         X2              =   304
         Y1              =   2
         Y2              =   30
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   348
         X2              =   348
         Y1              =   2
         Y2              =   30
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   10
         X2              =   10
         Y1              =   4
         Y2              =   30
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   13
         X2              =   13
         Y1              =   4
         Y2              =   30
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   14
         X2              =   14
         Y1              =   4
         Y2              =   30
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000010&
         X1              =   17
         X2              =   17
         Y1              =   4
         Y2              =   30
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000014&
         X1              =   114
         X2              =   114
         Y1              =   0
         Y2              =   28
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000010&
         X1              =   113
         X2              =   113
         Y1              =   -1
         Y2              =   27
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1996
            Key             =   "cloude"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1AA8
            Key             =   "cloudd"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1BBA
            Key             =   "winde"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":210C
            Key             =   "windd"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":265E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2BB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Wx..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Wx..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Wx"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Wx As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePaths 
         Caption         =   "Setup..."
      End
      Begin VB.Menu mnuFileNames 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuWx 
      Caption         =   "W&eather"
      Begin VB.Menu mnuWxCloud 
         Caption         =   "Add Cloud Layer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWxWind 
         Caption         =   "Add Wind Layer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWxTemp 
         Caption         =   "Add Temperature..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWxVis 
         Caption         =   "Add Visibility..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "&Database"
      Begin VB.Menu mnuDBAdd 
         Caption         =   "Add Station to DB..."
      End
      Begin VB.Menu mnuDBDelete 
         Caption         =   "Delete Station from DB..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTips 
         Caption         =   "&Fast Tips..."
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Active Weather 98..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'office 97 toolbar globals
Private iOldX As Integer, iOldY As Integer
Private oldIndex As Byte
Public Sub sbClear_Buttons()
    Call MoveMouse(100)
End Sub
Public Sub sbLoadFiveFiles()
   Dim iCount As Integer
   
   iCount = 1
   Do Until iCount > 6
     If Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, "")) <> "" Then
       If iCount = 1 Then
          mnuFileNames(0).Visible = True
       End If
       Load mnuFileNames(iCount)
       mnuFileNames(iCount).Caption = "&" & iCount & " " & _
                            Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, ""))
       mnuFileNames(iCount).Tag = _
                            Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, ""))
       mnuFileNames(iCount).Visible = True
       giFileCount = giFileCount + 1
     End If
     iCount = iCount + 1
   Loop
End Sub
Public Sub sbReadFiveFiles()
   Dim iCount As Integer
   
   mnuFileNames(0).Visible = True
   iCount = 1
   Do Until iCount > 6
     If Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, "")) <> "" Then
       If iCount > giFileCount Then
            Load mnuFileNames(iCount)
            giFileCount = giFileCount + 1
       End If
       mnuFileNames(iCount).Caption = "&" & iCount & " " & _
                            Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, ""))
       mnuFileNames(iCount).Tag = _
                            Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, ""))
       mnuFileNames(iCount).Visible = True
     End If
     iCount = iCount + 1
   Loop
End Sub

Public Sub sbSaveFiveFiles(sNewFile As String)
 
   Dim iCount As Integer
   Dim iDupIndex As Integer
   Dim sFileNum As String
   
   iDupIndex = 4
   iCount = 1
   'look for duplicates
   Do Until iCount > 6
     If sNewFile = Trim(GetSetting("ActiveWx98", "Files", "File" & iCount, "")) Then
       iDupIndex = iCount
     End If
     iCount = iCount + 1
   Loop
   
   iCount = iDupIndex - 1
   Do Until iCount = 0
       
       SaveSetting "ActiveWx98", "Files", "File" & iCount + 1, _
                   GetSetting("ActiveWx98", "Files", "File" & iCount, "")
       iCount = iCount - 1
   Loop
   
   SaveSetting "ActiveWx98", "Files", "File1", sNewFile
   
   
   
End Sub
Private Sub MDIForm_Load()
 oldIndex = 100
 sbLoadFiveFiles
End Sub
Public Sub sbWx_Buttons(iTypeDisplay As Integer)
  If iTypeDisplay = BUTTON_DISPLAY_ENABLED Then
        Picture1(3).Enabled = True
        Picture1(3).Picture = ImageList1.ListImages(1).Picture
        Picture1(4).Enabled = True
        Picture1(4).Picture = ImageList1.ListImages(3).Picture
        Picture1(5).Enabled = True
        Picture1(5).Picture = ImageList1.ListImages(5).Picture
        Picture1(6).Enabled = True
  Else
        Picture1(3).Enabled = False
        Picture1(3).Picture = ImageList1.ListImages(2).Picture
        Picture1(4).Enabled = False
        Picture1(4).Picture = ImageList1.ListImages(4).Picture
        Picture1(5).Enabled = False
        Picture1(5).Picture = ImageList1.ListImages(6).Picture
        Picture1(6).Enabled = False
  End If
End Sub
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call MoveMouse(100)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   sbSaveRegData
End Sub

Private Sub mnuDBAdd_Click()
   frmAddStation.Show
End Sub

Private Sub mnuDBDelete_Click()
   frmDelStation.Show vbModal
End Sub

Private Sub mnuEditPaste_Click()
  mdiMain.ActiveForm.strHTML = Clipboard.GetText
  mdiMain.ActiveForm.sbReadWxWebPlus
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
End Sub

Private Sub mnuFileNames_Click(Index As Integer)
   Dim frmWtr As New frmWeather
   frmWtr.Refresh
   frmWtr.sbLoad_Form TYPE_LOAD_OLD, mnuFileNames(Index).Tag
End Sub

Private Sub mnuFileNew_Click()
    Dim frmWtr As New frmWeather
    frmWtr.Refresh
    frmWtr.sbLoad_Form TYPE_LOAD_NEW, ""
End Sub

Private Sub mnuFileOpen_Click()
   Dim frmWtr As New frmWeather
   If Not fnSetFilePath Then Exit Sub
   frmWtr.Refresh
   frmWtr.sbLoad_Form TYPE_LOAD_OLD, CommonDialog1.filename
    
End Sub
Private Function fnSetFilePath() As Boolean
   CommonDialog1.CancelError = True
   
   On Error GoTo errhandler
   
   CommonDialog1.Flags = cdlOFNHideReadOnly
   
   CommonDialog1.Filter = "All Files (*.*)|*.*|Active Wx Files" & _
   "(*.acw)|*.acw|METAR Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    If iTypeDialog = TYPE_DLG_OPEN Then
        CommonDialog1.ShowOpen
    Else
        CommonDialog1.ShowOpen
    End If
    
    ' Display name of selected file
    sFilePath = CommonDialog1.filename
    fnSetFilePath = True
    Exit Function
errhandler:
     fnSetFilePath = False
End Function
   
Private Sub mnuFilePaths_Click()
    frmDir.Show vbModal
End Sub

Private Sub mnuFileSave_Click()
   On Error GoTo err_exit
   mdiMain.ActiveForm.sbSaveWx False
err_exit:
   
End Sub

Private Sub mnuFileSaveAs_Click()
   mdiMain.ActiveForm.sbSaveWx True
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuWeatherCreate_Click()
   
End Sub
Private Sub MoveMouse(ByVal bIndex As Byte)

  If bIndex <> oldIndex Then  ' If it is already drawn, then don't do it again!

  Dim iX As Integer, iY As Integer

    If oldIndex <> 100 Then ' Index 100 = No button selected!
        mdiMain.Picture2.Line (iOldX, iOldY)-(iOldX + 17 + 8, iOldY + 17 + 8), &H8000000A, B
        ' Remove the 3D-effect of the old button.
    End If

    If bIndex <> 100 Then

        iX = mdiMain.Picture1(bIndex).Left - 4
        iY = mdiMain.Picture1(bIndex).Top - 5


        mdiMain.Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014, B
        mdiMain.Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B
        mdiMain.Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B

        iOldX = iX: iOldY = iY
    End If

    oldIndex = bIndex

End If
End Sub
Private Sub DownMouse(ByVal bIndex As Byte)

 If bIndex <> 100 Then

 Dim iX As Integer, iY As Integer

    iX = mdiMain.Picture1(bIndex).Left - 4
    iY = mdiMain.Picture1(bIndex).Top - 5

    mdiMain.Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010, B
    mdiMain.Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014
    mdiMain.Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8 + 1, iY + 17 + 8), &H80000014


End If

End Sub
Private Sub UpMouse(ByVal bIndex As Byte)

 If bIndex <> 100 Then

 Dim iX As Integer, iY As Integer

    iX = mdiMain.Picture1(bIndex).Left - 4
    iY = mdiMain.Picture1(bIndex).Top - 5

    mdiMain.Picture2.Line (iX, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000014, B
    mdiMain.Picture2.Line (iX + 17 + 8, iY)-(iX + 17 + 8, iY + 17 + 8), &H80000010
    mdiMain.Picture2.Line (iX, iY + 17 + 8)-(iX + 17 + 8 + 1, iY + 17 + 8), &H80000010



End If

End Sub

Private Sub mnuHelpTips_Click()
    On Error GoTo err_exit
    frmTip.Load_Tip mdiMain.ActiveForm.giCurrentTip
err_exit:
End Sub

Private Sub mnuWindowCascade_Click()
 mdiMain.Arrange vbCascade
End Sub

Private Sub mnuWxCloud_Click()
  
  If mdiMain.ActiveForm.treWeather.SelectedItem Is Nothing Then
     MsgBox "Select a station to add the cloud layer to.", vbInformation
     Exit Sub
  End If
  frmAddCloud.Show vbModal
End Sub

Private Sub mnuWxTemp_Click()
  If mdiMain.ActiveForm.treWeather.SelectedItem Is Nothing Then
     MsgBox "Select a station to add the temperature to.", vbInformation
     Exit Sub
  End If
  frmAddTemp.sbLoad_Form mdiMain.ActiveForm.treWeather
End Sub

Private Sub mnuWxVis_Click()
  If mdiMain.ActiveForm.treWeather.SelectedItem Is Nothing Then
     MsgBox "Select a station to add visibility to.", vbInformation
     Exit Sub
  End If
  frmAddVis.sbLoad_Form mdiMain.ActiveForm.treWeather
End Sub

Private Sub mnuWxWind_Click()
  If mdiMain.ActiveForm.treWeather.SelectedItem Is Nothing Then
     MsgBox "Select a station to add the wind layer to.", vbInformation
     Exit Sub
  End If
  frmAddWind.Show vbModal
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
  sbClear_Buttons
  Select Case Index
     Case 0 'new
        mnuFileNew_Click
     Case 1 'open
        mnuFileOpen_Click
     Case 2 'save
        mnuFileSave_Click
     Case 3 'add cloud
        mnuWxCloud_Click
     Case 4 'wind
        mnuWxWind_Click
     Case 5 'temperature
        mnuWxTemp_Click
     Case 6 'Visibility
        mnuWxVis_Click
     Case 7 'add station
        mnuDBAdd_Click
     Case 9 'fast tips
        mnuHelpTips_Click
  End Select
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call MoveMouse(100)
End Sub
