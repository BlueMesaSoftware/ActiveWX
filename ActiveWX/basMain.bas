Attribute VB_Name = "basMain"
Global gsActiveWxPath As String
Global gsAdvPath As String
Global gsAPLCPath As String
Global gsInstallDir As String
Global gsSimVersion As String
Global giFileCount As Integer
Global gbShowTips As Boolean



Public Function fnReverseString(sInputString As String) As String
   Dim iCount As Integer
   Dim sOutStr As String
   
   iCount = Len(sInputString)
   Do Until iCount = 0
      sOutStr = sOutStr & Mid(sInputString, iCount, 1)
      iCount = iCount - 1
   Loop
   fnReverseString = sOutStr
End Function


Public Sub sbLoadRegData()
   gsAdvPath = GetSetting("ActiveWx98", "Sim", "AdvPath", "none")
   If gsAdvPath = "none" Then
       gsAdvPath = ""
       frmDir.Show vbModal
   End If
   gsAPLCPath = GetSetting("ActiveWx98", "Sim", "APLCPath", "none")
   If gsAPLCPath = "none" Then
       gsAPLCPath = ""
       frmDir.Show vbModal
   End If
   If GetSetting("ActiveWx98", "Settings", "ShowTips", "1") = 1 Then
     gbShowTips = True
   End If
   gsSimVersion = GetSetting("ActiveWx98", "Sim", "Version", "98")
   mdiMain.Top = CInt(GetSetting("ActiveWx98", "Settings", "Top", "400"))
   mdiMain.Left = CInt(GetSetting("ActiveWx98", "Settings", "Left", "300"))
   mdiMain.Width = CInt(GetSetting("ActiveWx98", "Settings", "Width", "11500"))
   mdiMain.Height = CInt(GetSetting("ActiveWx98", "Settings", "Height", "8000"))
   mdiMain.WindowState = CInt(GetSetting("ActiveWx98", "Settings", "Windowstate", "0"))
End Sub
Public Sub sbSaveRegData()
   SaveSetting "ActiveWx98", "Settings", "Windowstate", mdiMain.WindowState
   If mdiMain.WindowState = 0 Then
        SaveSetting "ActiveWx98", "Settings", "Top", mdiMain.Top
        SaveSetting "ActiveWx98", "Settings", "Left", mdiMain.Left
        SaveSetting "ActiveWx98", "Settings", "Width", mdiMain.Width
        SaveSetting "ActiveWx98", "Settings", "Height", mdiMain.Height
  End If
End Sub
Public Sub sbDeleteChildren(nParent As Node, nNodes As Nodes)
   Dim sChildKey As String
   Dim iCount As Integer
   Dim iSubCount As Integer
   Dim iNodeCount As Integer
   Dim sKeys() As String
   
   ReDim sKeys(1)
   
   sChildKey = Left(nParent.Child.Key, 4)
   iCount = 1
   iNodeCount = nNodes.Count
   iSubCount = 1
   Do Until iCount > iNodeCount
       If Left(nNodes(iCount).Key, 4) = sChildKey Then
          ReDim Preserve sKeys(UBound(sKeys) + 1)
          sKeys(iSubCount) = nNodes(iCount).Key
          iSubCount = iSubCount + 1
       End If
       iCount = iCount + 1
   Loop
   
   iCount = 1
   Do Until iCount > UBound(sKeys) - 1
       nNodes.Remove (sKeys(iCount))
       iCount = iCount + 1
   Loop
End Sub
Public Sub Main()
 
 
' If Date > CDate("01/12/1998") Or Date < CDate("11/24/1997") Then
'    MsgBox "This beta version of Active Weather 98 has expired."
'    End
' End If
  
 Set gDB = Workspaces(0).OpenDatabase(App.Path & "\wx1.mdb")
 
 Load mdiMain
 sbLoadRegData
 mdiMain.Show
 
 Exit Sub
error_exit:
   MsgBox "Can not connect to database"
   End
End Sub
