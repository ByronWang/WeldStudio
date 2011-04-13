Attribute VB_Name = "PlcRegularSetting"
Public fileName As String


Type RegularSettingType
    Value(15 - 1) As Single
End Type

Type RegularFileItemType
    name As String * 20
    regularSetting As RegularSettingType
End Type

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

    DefalutParam.Value(0) = 47
    DefalutParam.Value(1) = 110
    DefalutParam.Value(2) = 0.9
    DefalutParam.Value(3) = 2.5
    
    DefalutParam.Value(4) = 100
    DefalutParam.Value(5) = 75
    DefalutParam.Value(6) = 100
    
    DefalutParam.Value(7) = 200
    DefalutParam.Value(8) = 250
    DefalutParam.Value(9) = 290
    DefalutParam.Value(10) = 12.1
    
    DefalutParam.Value(13) = 3
    DefalutParam.Value(11) = 0.19
    DefalutParam.Value(12) = 1.3
    'DefalutParam.Value(14) = 0
        
DefalutStagesParameters = DefalutParam
End Function

Public Function LoadAll() As RegularFileItemType()
    Dim fileName As String
    fileName = App.path & "\" & SETTING_PATH & "RegularSetting.config"
    
    
    Dim pFileHeader As FileHeaderType
    Dim pFileItem As RegularFileItemType
    Dim pFileItemList() As RegularFileItemType
    
    Dim i As Integer
    Dim pos As Integer
    pos = 0
    
    Open fileName For Binary As #1
    Get 1, 1, pFileHeader
    
    ReDim pFileItemList(pFileHeader.count)
    
    pos = pos + LenB(pFileHeader)
    
    For i = 0 To pFileHeader.count - 1
        Get 1, pos + 1, pFileItem
        pos = pos + LenB(pFileItem)
        pFileItemList(i) = pFileItem
    Next i
Close 1

LoadAll = pFileItemList
End Function



Public Function LoadConfig(configName As String) As RegularSettingType
    Dim fileName As String
    fileName = App.path & "\" & SETTING_PATH & "RegularSetting.config"
    
    Dim pFileItemList() As RegularFileItemType
    pFileItemList = LoadAll()
    
    Dim i As Integer
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        If Trim(pFileItemList(i).name) = Trim(configName) Then
            LoadConfig = pFileItemList(i).regularSetting
            Exit Function
        End If
    Next i
    
    LoadConfig = DefalutStagesParameters
End Function


Public Function SaveConfig(configName As String, regularSetting As RegularSettingType)
    Dim fileName As String
    fileName = App.path & "\" & SETTING_PATH & "RegularSetting.config"
    
    Dim pFileItemList() As RegularFileItemType
    pFileItemList = LoadAll()


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
    
    pFileHeader.count = UBound(pFileItemList)
    
    pos = 0
    If haved Then
        pos = pos + LenB(pFileHeader)
        pos = pos + (i) * LenB(pFileItem)
    Else
        pos = pos + LenB(pFileHeader)
        pos = pos + pFileHeader.count * LenB(pFileItem)
        pFileHeader.count = pFileHeader.count + 1
    End If
    
    pFileItem.name = configName
    pFileItem.regularSetting = regularSetting
    
    Open fileName For Binary As #1
        Put 1, 1, pFileHeader
        Put 1, pos + 1, pFileItem
    Close 1
End Function

Public Function DeleteConfig(ByVal i As Integer) As Boolean
    Dim fileName As String
    fileName = App.path & "\" & SETTING_PATH & "RegularSetting.config"
    
    Dim pFileItemList() As RegularFileItemType
    pFileItemList = LoadAll()

    Dim pFileHeader As FileHeaderType
    Dim pFileItem As RegularFileItemType
    Dim pos As Integer
    
    pFileHeader.count = UBound(pFileItemList)
    pFileHeader.count = pFileHeader.count - 1
        
    pos = 0
    pos = pos + LenB(pFileHeader)
    pos = pos + (i) * LenB(pFileItem)
        
    Open fileName For Binary As #1
        Put 1, 1, pFileHeader
            
        For i = i To pFileHeader.count - 1
            pFileItem.name = pFileItemList(i + 1).name
            pFileItem.regularSetting = pFileItemList(i + 1).regularSetting
            Put 1, pos + 1, pFileItem
            pos = pos + LenB(pFileItem)
        Next i
    Close 1
End Function


Public Function AssertEqualRegularData(ByRef regularSetting As RegularSettingType, ByRef dest As RegularSettingType) As Boolean
 
    Dim j As Integer
    
    DoEvents
    
    For j = 1 To 14
        If (regularSetting.Value(j - 1) <> dest.Value(j - 1)) Then
            GoTo NotEqual
        End If
    Next
            
    AssertEqualRegularData = True

Exit Function
NotEqual:
    AssertEqualRegularData = False
End Function


