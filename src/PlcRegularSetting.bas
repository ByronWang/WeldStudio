Attribute VB_Name = "PlcRegularSetting"

Public Function DefalutStagesParameters() As RegularSettingType

Dim DefalutParam As RegularSettingType

'0   1   Parameter set index
'1   2   High volt timer in seconds
'2   3   Low volt timer in seconds
'3   4   Current in upset timer in seconds
'4   5   Upset timer in seconds
'5   6   High volt in volts
'6   7   Low volt in volts
'7   8   Boost volt in volts
'8   9   Current setpoint in amps for stage I
'9   10  Current setpoint in amps for stage II-1
'10  11  Current setpoint in amps for stage II-2
'11  12  Upset in millimeter
'12  13  Flash speed in mm/s
'13  14  Boost speed in mm/s
'14  15  Pre-flash distance in millimeter
'

    DefalutParam.Value(0) = 0
    DefalutParam.Value(1) = 0
    DefalutParam.Value(2) = 0
    DefalutParam.Value(3) = 0
    DefalutParam.Value(4) = 0
    DefalutParam.Value(5) = 0
    DefalutParam.Value(6) = 0
    DefalutParam.Value(7) = 0
    DefalutParam.Value(8) = 0
    DefalutParam.Value(9) = 0
    DefalutParam.Value(10) = 0
    DefalutParam.Value(11) = 0
    DefalutParam.Value(12) = 0
    DefalutParam.Value(13) = 0
    DefalutParam.Value(14) = 0
        
DefalutStagesParameters = DefalutParam
End Function

Public Function LoadAll(filename As String) As RegularFileItemType()
    Dim pFileHeader As FileHeaderType
    Dim pFileItem As RegularFileItemType
    Dim pFileItemList() As RegularFileItemType
    
    Dim i As Integer
    Dim pos As Integer
    pos = 0
    
    Open filename For Binary As #1
    Get 1, 1, pFileHeader
    
    ReDim pFileItemList(pFileHeader.Count)
    
    pos = pos + LenB(pFileHeader)
    
    For i = 0 To pFileHeader.Count - 1
        Get 1, pos + 1, pFileItem
        pos = pos + LenB(pFileItem)
        pFileItemList(i) = pFileItem
    Next i
Close 1

LoadAll = pFileItemList
End Function



Public Function LoadConfig(filename As String, configName As String) As RegularSettingType
    Dim pFileItemList() As RegularFileItemType
    pFileItemList = LoadAll(filename)
    
    Dim i As Integer
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        If Trim(pFileItemList(i).name) = Trim(configName) Then
            LoadConfig = pFileItemList(i).RegularSetting
            Exit Function
        End If
    Next i
    
    LoadConfig = DefalutStagesParameters
End Function


Public Function SaveConfig(filename As String, configName As String, RegularSetting As RegularSettingType)
    Dim pFileItemList() As RegularFileItemType
    pFileItemList = LoadAll(filename)


    Dim haved As Boolean
    Dim i As Integer
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        If Trim(pFileItemList(i).name) = Trim(configName) Then
            haved = True
            Exit For
        End If
    Next i
    
    

    Dim pFileHeader As FileHeaderType
    Dim pFileItem As RegularFileItemType
    Dim pos As Integer
    
    pFileHeader.Count = UBound(pFileItemList)
    
    pos = 0
    If haved Then
        pos = pos + LenB(pFileHeader)
        pos = pos + (i) * LenB(pFileItem)
    Else
        pos = pos + LenB(pFileHeader)
        pos = pos + pFileHeader.Count * LenB(pFileItem)
        pFileHeader.Count = pFileHeader.Count + 1
    End If
    
    pFileItem.name = configName
    pFileItem.RegularSetting = RegularSetting
    
    Open filename For Binary As #1
        Put 1, 1, pFileHeader
        Put 1, pos + 1, pFileItem
    Close 1
End Function