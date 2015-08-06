VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Station_ID As String
Public Metar As String
Public StationCount As Integer
Public Wind As String
Public WindCount As Integer
Public Visibility As String
Public Cloud As String
Public CloudCount As Integer


Private piStationIndex As Integer
Private piWindIndex As Integer

Private Type STATION_TYPE
   Station_ID As String
   Metar As String
   WindList(10) As String
   CloudList(10) As String
   Visibility As String
End Type

Private Stations() As STATION_TYPE
Property Let WindIndex(iIndex As Integer)
   Wind = Stations(piStationIndex).WindList(iIndex)
   piWindIndex = iIndex
End Property
Property Let StationIndex(iIndex As Integer)
'    If piStationIndex < StationCount - 1 Then
        Station_ID = Stations(iIndex).Station_ID
        piStationIndex = iIndex
'    Else
'        MsgBox "Error"
'    End If
End Property
Public Sub SetStationID(sMetarStationID As String)
      StationCount = StationCount + 1
      ReDim Preserve Stations(UBound(Stations) + 1)
      Stations(StationCount - 1).Station_ID = sMetarStationID
End Sub

Public Sub SetWind(sMetarWind As String)
     Dim St As Integer
     Dim sWind As String
     
     WindCount = WindCount + 1
     sWind = "WINDS " & WindCount & ", 0, FTOM( 1015 ), "
           
     'gusting or steady wind
     If InStr(1, sMetarWind, "G") > 0 Then
          sWind = sWind & "gusty, "
     Else
          sWind = sWind & "steady, "
     End If
     
     sWind = sWind & Left(sMetarWind, 3) & ", " & _
              Left(Right(sMetarWind, 4), 2) & ", 00"

     Stations(StationCount - 1).WindList(WindCount - 1) = sWind
'                      240, 11, 11.427183"
End Sub
Public Sub SetVisibility(sMetarVis As String)
  Dim iSlashPOS As Integer
  Dim sDecimalVis As String
  Dim sWholeVis As String
  Dim sTempVis As String
  Dim iCount As Integer
  
  iSlashPOS = InStr(1, sMetarVis, "/")
  ' VISIBILITY visibility
  If iSlashPOS > 0 Then
     Select Case Mid(sMetarVis, iSlashPOS - 1, 3)
         Case "1/4"
              sDecimalVis = ".25"
         Case "1/2"
              sDecimalVis = ".50"
         Case "3/4"
              sDecimalVis = ".75"
     End Select
  End If
  iCount = 1
  Do Until iCount > Len(sMetarVis)
    If Mid(sMetarVis, iCount, 1) = Chr(32) Or _
       Mid(sMetarVis, iCount, 1) = "S" Then
       sWholeVis = sTempVis
       iCount = 99
    End If
    sTempVis = sTempVis & Mid(sMetarVis, iCount, 1)
    iCount = iCount + 1
  Loop
'  fnVisibility = "VISIBILITY " & sWholeVis & sDecimalVis
 End Sub
Public Sub setClouds(sMetarCloud As String)
' CLOUDS 1, FTOM(vCld1Bot), FTOM(1000), userdefined, broken5, 3.919924, 300, 0
  Dim iCloudBaseMeter As Integer
  Dim iCloudBaseFeet As Integer
  Dim rs As Recordset
  Dim sSQL As String
   
  sSQL = "select * from Stations where IATA='COS'"
  
  Set rs = gDB.OpenRecordset(sSQL, dbOpenSnapshot)
  Do Until rs.EOF
      iCloudBaseFeet = CInt(Mid(sMetarCloud, 4, 3) & "00")
      iCloudBaseMeter = fnFTOM(iCloudBaseFeet) + rs!Elevation
      rs.MoveNext
  Loop
  rs.Close
  
  
  
'  fnClouds = "CLOUDS 1, " & iCloudBaseMeter & ", FTOM(1000), userdefined, "
End Sub
Private Function fnFTOM(iFeet) As Integer
    fnFTOM = iFeet / 3.2808
    
End Function


Private Sub Class_Initialize()
  ReDim Stations(1)
 
End Sub
