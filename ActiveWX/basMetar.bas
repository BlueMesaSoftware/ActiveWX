Attribute VB_Name = "basMetar"
Global gDB As Database
Global giCloudCount As Integer
Global giTempCount As Integer
Global gsStationID As String
Global giFieldElevMeters As Integer
Global gsTopBuffer As String
Global gsMidBuffer As String
Global gsBottomBuffer As String
Global gsMilesToSection As String
Global arrTest(4) As String

Public Function fnConvertLatLon(sLatLong As String, iType As Integer) As Single
   Dim lLatLong1 As Long
   Dim lLatLong2 As Long
   Dim lLatLong3 As Long
   
   Select Case iType
     '34-59-53.03N
     Case TYPE_LATITUDE
         lLatLong1 = CLng(Left(sLatLong, 2))
         lLatLong2 = CLng(Mid(sLatLong, 4, 2))
         lLatLong2 = lLatLong2 * 100 / 60
         lLatLong3 = CLng(Mid(sLatLong, 7, 2))
         lLatLong3 = lLatLong3 * 100 / 60
         fnConvertLatLon = CSng(lLatLong1 & "." & lLatLong2 & lLatLong3)
         If Right(sLatLong, 1) = "S" Then
            fnConvertLatLon = 0 - CSng(lLatLong1 & "." & lLatLong2 & lLatLong3)
         End If
     Case TYPE_LONGITUDE
         '089-48-30.366W
         lLatLong1 = CLng(Left(sLatLong, 3))
         lLatLong2 = CLng(Mid(sLatLong, 5, 2))
         lLatLong2 = lLatLong2 * 100 / 60
         lLatLong3 = CLng(Mid(sLatLong, 8, 2))
         lLatLong3 = lLatLong3 * 100 / 60
         fnConvertLatLon = CSng(lLatLong1 & "." & lLatLong2 & lLatLong3)
        
   End Select
End Function
Public Sub sbReadFiles()
   Dim sInputLine As String
   Open "d:\wx\top.txt" For Input As #1
    Do While Not EOF(1) ' Loop until end of file.
        sInputLine = Input(1, #1)
        gsTopBuffer = gsTopBuffer & sInputLine
    Loop
    Close #1
    Open "d:\wx\bottom.txt" For Input As #1
    Do While Not EOF(1) ' Loop until end of file.
        Input #1, sInputLine
        gsBottomBuffer = gsBottomBuffer & sInputLine & Chr(13) + Chr(10)
    Loop
    MsgBox gsBottomBuffer
    Close #1
End Sub
Public Function fnStationID(sMetarStationID As String) As String
     Dim nodX As Node
     Dim sKey As String
     Dim rs As Recordset
     Dim sSQL As String
     Dim sngLat As Single
     Dim iLat2 As Integer
     Dim sngLong As Single
     Dim iLong2 As Integer
     Dim sngBeginLat As Single
     Dim sngEndLat As Single
     Dim sngBeginLong As Single
     Dim sngEndLong As Single
     Dim sStation(11) As String
     
     On Error GoTo err_exit
     
     sKey = "S" & Mid(sMetarStationID, 2, 3)
     Set nodX = mdiMain.ActiveForm.treWeather.Nodes(sKey)
     If nodX.Children > 0 Then
        sbDeleteChildren nodX, mdiMain.ActiveForm.treWeather.Nodes
     End If
     gsStationID = Mid(sMetarStationID, 2, 3)
     Exit Function
     
err_exit:
End Function
Public Function fnAltimeter(sMetarAltimeter As String) As String
   Dim nodX As Node
   Dim sParentKey As String
   Dim sNodeKey As String
   Dim sAltimeter As String
   
   On Error GoTo err_exit
   
   sNodeKey = "C" & gsStationID & "A"
                 
   sParentKey = "S" & gsStationID
   sAltimeter = "Altimeter: " & Mid(sMetarAltimeter, 2, 2) & "." & _
                 Mid(sMetarAltimeter, 4, 2)
                 
   Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sAltimeter)
   nodX.Tag = Mid(sMetarAltimeter, 2, 2) & "." & _
                 Mid(sMetarAltimeter, 4, 2)
   nodX.Image = "alt"
   Exit Function
err_exit:
   
 
End Function
Public Function fnTime(sMetarTime As String) As String
  '092217Z
  On Error GoTo error_exit
  Dim nodX As Node
  Dim cnodX As Node
  Dim sParentKey As String
  Dim sKey As String
  Dim sFullStationName As String
  Dim sNodeText As String
  Dim vTagData As Variant
  sKey = "C" & gsStationID & "M"
  sParentKey = "S" & gsStationID

  sNodeText = sFullStationName & "Wx Day: " & Left(sMetarTime, 2) & " Time: " & _
              Format(Mid(sMetarTime, 3, 4), "##:##") & "Z"
  
  Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add _
              (sParentKey, tvwChild, sKey, sNodeText)
  nodX.Image = "clock"
  
   
error_exit:

End Function
Public Function fnWind(sMetarWind As String) As String
     Dim St As Integer
     Dim nodX As Node
     Dim sNodeKey As String
     Dim sWind As String
     Dim sParentKey As String
     Dim asWindData(6) As String
     Dim asWindLayer1(6) As String
     Dim asWindLayer2(6) As String
     Dim vStation As Variant
     
     '0-wind direction
     '1-wind speed
     '2-wind type
     '3-turbulence
     '4-wind base feet
     '5-wind top feet
     '6-Wind Level
     
     On Error GoTo err_exit
     
     giWindCount = giWindCount + 1
'     fnWind = "WINDS " & giWindCount & ", 0, FTOM( 1015 ), "
           
     'gusting or steady wind
     If InStr(1, sMetarWind, "G") > 0 Then
          fnWind = fnWind & "gusty, "
          asWindData(2) = "gusty"
     Else
          fnWind = fnWind & "steady, "
          asWindData(2) = "steady"
     End If
     
     'turbulence section
     If Left(Right(sMetarWind, 4), 2) <= 10 Then
       asWindData(3) = 0
     ElseIf Left(Right(sMetarWind, 4), 2) > 10 And _
         Left(Right(sMetarWind, 4), 2) <= 30 Then
       asWindData(3) = 30
     ElseIf Left(Right(sMetarWind, 4), 2) > 30 Then
       asWindData(3) = 120
     End If
     If asWindData(2) = "gusty" Then
        asWindData(3) = CInt(asWindData(3)) + 40
     End If
     
     If Left(sMetarWind, 3) = "VRB" Then
        asWindData(0) = 360
     Else
        asWindData(0) = Left(sMetarWind, 3)
     End If
     asWindData(1) = Left(Right(sMetarWind, 4), 2)
     asWindData(4) = "0"
     asWindData(5) = "1700"
     asWindData(6) = "1"
     
     '==add additional layers of wind
     '0-wind direction
     '1-wind speed
     '2-wind type
     '3-turbulence
     '4-wind base feet AGL
     '5-wind top feet  AGL
     '6-Wind Level
     
     asWindLayer1(0) = asWindData(0)
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
     asWindLayer2(5) = asWindLayer1(5) + 21800
     asWindLayer2(6) = "3"
     
     '====Add surface layer to the tree
     sNodeKey = "C" & gsStationID & "W" & giWindCount
     sParentKey = "S" & gsStationID
     sWind = "Wind Layer: " & giWindCount & " " & Left(sMetarWind, 3) & _
             " at " & asWindData(1) & " knots " & _
             "(" & asWindData(4) & " - " & asWindData(5) & " feet AGL)"
     Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindData
     nodX.Image = "wind"
     
     giWindCount = giWindCount + 1
     '====Add wind layer1 to the tree=======================================
     sNodeKey = "C" & gsStationID & "W" & giWindCount
     sParentKey = "S" & gsStationID
     sWind = "Wind Layer: " & giWindCount & " " & asWindLayer1(0) & _
             " at " & asWindLayer1(1) & " knots " & _
             "(" & asWindLayer1(4) & " - " & asWindLayer1(5) & " feet AGL)"
     Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindLayer1
     nodX.Image = "wind"
     
     giWindCount = giWindCount + 1
     '====Add wind layer2 to the tree=======================================
     sNodeKey = "C" & gsStationID & "W" & giWindCount
     sParentKey = "S" & gsStationID
     sWind = "Wind Layer: " & giWindCount & " " & asWindLayer2(0) & _
             " at " & asWindLayer2(1) & " knots " & _
             "(" & asWindLayer2(4) & " - " & asWindLayer2(5) & " feet AGL)"
     Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sWind)
     nodX.Tag = asWindLayer2
     nodX.Image = "wind"
     
     '=================set tag data for wind count on parent
     Set nodX = mdiMain.ActiveForm.treWeather.Nodes(sParentKey)
     vStation = nodX.Tag
     vStation(15) = "3"
     nodX.Tag = vStation
     
     Exit Function
err_exit:

End Function
Public Function fnTemp(sMetarTemp As String) As String
   Dim nodX As Node
   Dim sParentKey As String
   Dim sNodeKey As String
   Dim sTemp As String
   Dim iTempF As Integer
   
   On Error GoTo error_exit
   
   sNodeKey = "C" & gsStationID & "T" & "1"
                 
   sParentKey = "S" & gsStationID
   If Left(sMetarTemp, 1) = "M" Then
         If InStr(1, sMetarTemp, "/") = 3 Then
            sTemp = "Temperature: " & fnCTOF(CInt("-" & Mid(sMetarTemp, 2, 1))) & " F"
            iTempF = fnCTOF(CInt("-" & Mid(sMetarTemp, 2, 1)))
         Else
            sTemp = "Temperature: " & fnCTOF(CInt("-" & Mid(sMetarTemp, 2, 2))) & " F"
            iTempF = fnCTOF(CInt("-" & Mid(sMetarTemp, 2, 2)))
        End If
   Else
         sTemp = "Temperature: " & fnCTOF(Left(sMetarTemp, 2)) & " F"
         iTempF = fnCTOF(Left(sMetarTemp, 2))
   End If
   
   Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sTemp)
   nodX.Tag = iTempF
   nodX.Image = "thermo"
   
   
   giTempCount = giTempCount + 1
   Exit Function
error_exit:
  
End Function
Public Function fnVisibility(sMetarVis As String) As String
  Dim iSlashPOS As Integer
  Dim sDecimalVis As String
  Dim sWholeVis As String
  Dim sTempVis As String
  Dim iCount As Integer
  Dim sVisibility As String
  Dim sNodeKey As String
  Dim sParentKey As String
  Dim sVis As String
    
  On Error GoTo err_dup
  
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
    'exit out of loop
    If Mid(sMetarVis, iCount, 1) = Chr(32) Or _
       Mid(sMetarVis, iCount, 1) = "S" Then
        sWholeVis = sTempVis
        If Left(sWholeVis, 1) = "P" Then
          sWholeVis = Right(sWholeVis, Len(sWholeVis) - 1)
        End If
        iCount = 99
    End If
    sTempVis = sTempVis & Mid(sMetarVis, iCount, 1)
    iCount = iCount + 1
  Loop
        
  If Mid(sMetarVis, 2, 1) = "/" Or Mid(sMetarVis, 3, 1) = "/" Then
        sVis = "0" & sDecimalVis
  Else
        sVis = sWholeVis & sDecimalVis
  End If
  sNodeKey = "C" & gsStationID & "V"
  sParentKey = "S" & gsStationID
  sVisibility = "Visibility: " & sVis & " miles"
  Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sVisibility)
  nodX.Image = "vis"
  nodX.Tag = sVis
'  nodX.EnsureVisible
  
  
'  fnVisibility = "VISIBILITY " & sWholeVis & sDecimalVis
  Exit Function

err_dup:
  
  
 End Function
Public Function fnCTOF(iTempCelcius As Integer) As Integer
   fnCTOF = iTempCelcius * 1.8 + 32
End Function
Public Function fnFTOC(iTempF As Integer) As Integer
   fnFTOC = CInt((iTempF - 32) / 1.8)
End Function
Public Function fnClouds(sMetarCloud As String) As String

  Dim iCloudBaseMeter As Integer
  Dim iCloudBaseFeet As Integer
  Dim iCloudTopMeter As Integer
  Dim iCloudTopFeet As Integer
  Dim sCoverage As String
  Dim sCloud As String
  Dim asCloudData(8) As String
  Dim iCloudThickness As Integer
  Dim vStation As Variant
  
  On Error GoTo err_exit
  
  '1 raw base feet (AGL)
  '2 raw top feet  (AGL)
  '3 type clouds
  '4 coverage
  '5 turbulence
  '6 devation
  '7 icing
  '8-level
  
  Select Case Left(sMetarCloud, 3)
      Case "CLR"
         sCoverage = "Clear"
         iCloudThickness = 300
      Case "FEW"
         sCoverage = "Scattered2"
         iCloudThickness = 300
      Case "SCT"
         sCoverage = "Scattered4"
         iCloudThickness = 300
      Case "BKN"
         sCoverage = "Broken7"
         iCloudThickness = 800
      Case "OVC"
         sCoverage = "Overcast"
         iCloudThickness = 1500
  End Select

  
  giCloudCount = giCloudCount + 1
  
  '1 raw base feet (AGL)
  '2 raw top feet  (AGL)
  '3 type clouds
  '4 coverage
  '5 turbulence
  '6 devation
  '7 icing
  '8 -level
  
  If sCoverage = "Clear" Then
        asCloudData(1) = "3300"
        asCloudData(2) = 3300 + iCloudThickness
        asCloudData(3) = "Userdefined"
        asCloudData(4) = "Scattered1"
        asCloudData(8) = giCloudCount
        
        sCloud = "Cloud Level: 1 " & _
                 " (" & asCloudData(1) & " - " & asCloudData(2) & " feet AGL) " & _
                 asCloudData(4) & " " & asCloudData(3)
  
  ElseIf Left(sMetarCloud, 2) = "VV" Then
        asCloudData(1) = "1100"
        asCloudData(2) = 1100 + 1000
        asCloudData(3) = "Userdefined"
        asCloudData(4) = "Overcast"
        asCloudData(8) = giCloudCount
        
        sCloud = "Cloud Level: " & giCloudCount & _
                 " (1,100 - 2,100 feet AGL) Overcast " & asCloudData(3)
  Else 'All other cloud types
        asCloudData(1) = Mid(sMetarCloud, 4, 3) & "00"
        asCloudData(2) = CInt(Mid(sMetarCloud, 4, 3) & "00") + iCloudThickness
        asCloudData(3) = "Userdefined"
        asCloudData(4) = sCoverage
        asCloudData(8) = giCloudCount
        iCloudBaseFeet = CInt(Mid(sMetarCloud, 4, 3) & "00")
        iCloudTopFeet = iCloudBaseFeet + iCloudThickness
        iCloudTopMeter = iCloudBaseMeter + fnFTOM(iCloudThickness)
        
        sCloud = "Cloud: " & giCloudCount & " (" & _
           Format(iCloudBaseFeet, "##,##0") & " - " & _
           Format(iCloudTopFeet, "##,##0") & " feet AGL) " & sCoverage & _
           " " & asCloudData(3)

  
  End If
  
  
  asCloudData(5) = "0"
  asCloudData(6) = "0"
  asCloudData(7) = "0"
  
  sNodeKey = "C" & gsStationID & "C" & giCloudCount
  sParentKey = "S" & gsStationID
 
  
  Set nodX = mdiMain.ActiveForm.treWeather.Nodes.Add(sParentKey, tvwChild, sNodeKey, _
                sCloud)
  nodX.Image = "cloud"
  nodX.Tag = asCloudData
  
  '=================set tag data for wind count on parent
  Set nodX = mdiMain.ActiveForm.treWeather.Nodes(sParentKey)
  vStation = nodX.Tag
  vStation(16) = giCloudCount
  nodX.Tag = vStation
  Exit Function

err_exit:
  
End Function
Public Function fnFTOM(iFeet) As Integer
  On Error GoTo err_exit
  fnFTOM = iFeet / 3.2808
  Exit Function
err_exit:
 fnFTOM = 0
End Function

