Attribute VB_Name = "PlcPulseSetting"
Option Explicit

Type FileHeaderType
    Count As Integer
End Type

Type StageParametersType
    Value(8 - 1) As Single
'    Distance As Single
'    Voltage As Single
'    Time As Single
'    CurrentSetpoint1 As Single
'    CurrentSetpoint2 As Single
'    CurrentSetpoint3 As Single
'    ForwardSpeed As Single
'    ReverseSpeed As Single
End Type

Type GeneralParametersType
    Value(4 - 1) As Single
    'Parameter set index     'TODO  meaning unknown
'    CurrentInUpsetTimerInSeconds As Single
'    UpsetInMillimeter As Single
'    HoldingTimerForTensionUseInSeconds As Single
'    ForgingForceInTonnes As Single
End Type

Type PulseSettingType
    Stages(7 - 1) As StageParametersType
    General As GeneralParametersType
End Type

Type PulseFileItemType
    name As String * 20
    PulseSetting As PulseSettingType
End Type


Type RegularSettingType
    Value(15 - 1) As Single
End Type

Type RegularFileItemType
    name As String * 20
    RegularSetting As RegularSettingType
End Type

Public Function DefalutStagesParameters() As PulseSettingType

Dim DefalutParam As PulseSettingType

    DefalutParam.General.Value(0) = 0
    DefalutParam.General.Value(1) = 9
    DefalutParam.General.Value(2) = 20
    DefalutParam.General.Value(3) = 0
        
    Dim i As Integer
    Dim stage As StageParametersType
    
    For i = 0 To 6
        DefalutParam.Stages(i).Value(0) = 0.1
        DefalutParam.Stages(i).Value(1) = 0
        DefalutParam.Stages(i).Value(2) = 250
        DefalutParam.Stages(i).Value(3) = 150
        DefalutParam.Stages(i).Value(4) = 150
        DefalutParam.Stages(i).Value(5) = 150
        DefalutParam.Stages(i).Value(6) = 0.1
        DefalutParam.Stages(i).Value(7) = 0.1
'
'
'        DefalutParam.Stages(i).Distance = 0.1
'        DefalutParam.Stages(i).Time = 0
'        DefalutParam.Stages(i).Voltage = 250
'        DefalutParam.Stages(i).CurrentSetpoint1 = 150
'        DefalutParam.Stages(i).CurrentSetpoint2 = 150
'        DefalutParam.Stages(i).CurrentSetpoint3 = 150
'        DefalutParam.Stages(i).ForwardSpeed = 0.1
'        DefalutParam.Stages(i).ReverseSpeed = 0.1
    Next i
DefalutStagesParameters = DefalutParam
End Function

Public Function LoadAll(filename As String) As PulseFileItemType()
    Dim pFileHeader As FileHeaderType
    Dim pFileItem As PulseFileItemType
    Dim pFileItemList() As PulseFileItemType
    
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



Public Function LoadConfig(filename As String, configName As String) As PulseSettingType
    Dim pFileItemList() As PulseFileItemType
    pFileItemList = LoadAll(filename)
    
    Dim i As Integer
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        If Trim(pFileItemList(i).name) = Trim(configName) Then
            LoadConfig = pFileItemList(i).PulseSetting
            Exit Function
        End If
    Next i
    
    LoadConfig = DefalutStagesParameters
End Function


Public Function SaveConfig(filename As String, configName As String, PulseSetting As PulseSettingType)
    Dim pFileItemList() As PulseFileItemType
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
    Dim pFileItem As PulseFileItemType
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
    pFileItem.PulseSetting = PulseSetting
    
    Open filename For Binary As #1
        Put 1, 1, pFileHeader
        Put 1, pos + 1, pFileItem
    Close 1
End Function