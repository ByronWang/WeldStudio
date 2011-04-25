Attribute VB_Name = "PlcPulseSetting"
Option Explicit

Private Type FileHeaderType
    count As Integer
End Type

'    Distance As Single
'    Voltage As Single
'    Time As Single
'    CurrentSetpoint1 As Single
'    CurrentSetpoint2 As Single
'    CurrentSetpoint3 As Single
'    ForwardSpeed As Single
'    ReverseSpeed As Single
Private Type StageParametersType
    Value(8 - 1) As Single
End Type

'    Parameter set index     'TODO  meaning unknown
'    CurrentInUpsetTimerInSeconds As Single
'    UpsetInMillimeter As Single
'    HoldingTimerForTensionUseInSeconds As Single
'    ForgingForceInTonnes As Single
Private Type GeneralParametersType
    Value(4 - 1) As Single
End Type

Type PulseSettingType
    Stages(7 - 1) As StageParametersType
    General As GeneralParametersType
End Type

Type PulseFileItemType
    name As String * 20
    pulseSetting As PulseSettingType
End Type

Public Function DefalutStagesParameters() As PulseSettingType
    Dim DefalutParam As PulseSettingType
    
    DefalutParam.General.Value(1) = 12
    DefalutParam.General.Value(0) = 0.5
    DefalutParam.General.Value(3) = 55
    DefalutParam.General.Value(2) = 0
    
    Dim i As Integer
    Dim stage As StageParametersType
    
    i = -1
    
    i = i + 1 ' Preflash
    DefalutParam.Stages(i).Value(0) = 4.5
    DefalutParam.Stages(i).Value(2) = 30
    DefalutParam.Stages(i).Value(1) = 98
    DefalutParam.Stages(i).Value(3) = 300
    DefalutParam.Stages(i).Value(4) = 350
    DefalutParam.Stages(i).Value(5) = 400
    DefalutParam.Stages(i).Value(6) = 1.4
    DefalutParam.Stages(i).Value(7) = 0.7
    
    i = i + 1 'Flash-I
    DefalutParam.Stages(i).Value(0) = 6
    DefalutParam.Stages(i).Value(2) = 40
    DefalutParam.Stages(i).Value(1) = 85
    DefalutParam.Stages(i).Value(3) = 350
    DefalutParam.Stages(i).Value(4) = 500
    DefalutParam.Stages(i).Value(5) = 600
    DefalutParam.Stages(i).Value(6) = 2.75
    DefalutParam.Stages(i).Value(7) = 0.8
    
    i = i + 1 'Flash-II
    DefalutParam.Stages(i).Value(0) = 7
    DefalutParam.Stages(i).Value(2) = 30
    DefalutParam.Stages(i).Value(1) = 77
    DefalutParam.Stages(i).Value(3) = 200
    DefalutParam.Stages(i).Value(4) = 300
    DefalutParam.Stages(i).Value(5) = 400
    DefalutParam.Stages(i).Value(6) = 2.75
    DefalutParam.Stages(i).Value(7) = 0.7
    
    i = i + 1 'Flash-III
    DefalutParam.Stages(i).Value(0) = 6
    DefalutParam.Stages(i).Value(2) = 15
    DefalutParam.Stages(i).Value(1) = 89
    DefalutParam.Stages(i).Value(3) = 190
    DefalutParam.Stages(i).Value(4) = 350
    DefalutParam.Stages(i).Value(5) = 450
    DefalutParam.Stages(i).Value(6) = 0.4
    DefalutParam.Stages(i).Value(7) = 0.2
    
    i = i + 1 'Flash-IV
    DefalutParam.Stages(i).Value(0) = 6
    DefalutParam.Stages(i).Value(2) = 3
    DefalutParam.Stages(i).Value(1) = 100
    DefalutParam.Stages(i).Value(3) = 180
    DefalutParam.Stages(i).Value(4) = 500
    DefalutParam.Stages(i).Value(5) = 600
    DefalutParam.Stages(i).Value(6) = 0.8
    DefalutParam.Stages(i).Value(7) = 0.2
    
    i = i + 1 'Boost-I
    DefalutParam.Stages(i).Value(0) = 2.8
    DefalutParam.Stages(i).Value(2) = 3
    DefalutParam.Stages(i).Value(1) = 100
    DefalutParam.Stages(i).Value(3) = 200
    DefalutParam.Stages(i).Value(4) = 500
    DefalutParam.Stages(i).Value(5) = 600
    DefalutParam.Stages(i).Value(6) = 1.2
    DefalutParam.Stages(i).Value(7) = 0.2
    
    i = i + 1 'Boost-II
    DefalutParam.Stages(i).Value(0) = 10
    DefalutParam.Stages(i).Value(2) = 3
    DefalutParam.Stages(i).Value(1) = 100
    DefalutParam.Stages(i).Value(3) = 225
    DefalutParam.Stages(i).Value(4) = 500
    DefalutParam.Stages(i).Value(5) = 600
    DefalutParam.Stages(i).Value(6) = 1.6
    DefalutParam.Stages(i).Value(7) = 0.1
    
DefalutStagesParameters = DefalutParam
End Function

Public Function LoadAll() As PulseFileItemType()
    Dim FileName As String
    FileName = App.path & "\" & SETTING_PATH & "PulseSetting.cfg"
    
    out.log "<<<<<<<<<<<<<<<<<<<<<     LoadAll  PulseFileItemType <<<<<<<<<<<<<<<<<<<"
    
    Dim pFileHeader As FileHeaderType
    Dim pFileItem As PulseFileItemType
    Dim pFileItemList() As PulseFileItemType
    
    Dim i As Integer
    Dim pos As Integer
    pos = 0
    
    Open FileName For Binary As #1
    Get 1, 1, pFileHeader
    
    out.log " " & 1 & " > pFileHeader.count = " & pFileHeader.count
    
    ReDim pFileItemList(pFileHeader.count)
    
    pos = pos + LenB(pFileHeader)
    
    For i = 0 To pFileHeader.count - 1
        Get 1, pos + 1, pFileItem
        pos = pos + LenB(pFileItem)
        pFileItemList(i) = pFileItem
        
        out.log " " & pos & " > pFileItem.name = " & pFileItem.name
        out.logSingleArray "pFileItem.pulseSetting.General.Value", pFileItem.pulseSetting.General.Value
        out.logSingleArray "pFileItem.pulseSetting.Stages(1).Value", pFileItem.pulseSetting.Stages(1).Value
    Next i
    Close 1
    
    out.log ">>>>>>>>>>>>>>>>>>     Finish LoadAll PulseFileItemType  >>>>>>>>>>>>>>>>>>>>"

LoadAll = pFileItemList
End Function

Public Function LoadConfig(configName As String) As PulseSettingType
    Dim pFileItemList() As PulseFileItemType
    pFileItemList = LoadAll()
    
    Dim i As Integer
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        If Trim(pFileItemList(i).name) = Trim(configName) Then
            LoadConfig = pFileItemList(i).pulseSetting
            Exit Function
        End If
    Next i
    
    LoadConfig = DefalutStagesParameters
End Function

Public Function SaveConfig(configName As String, pulseSetting As PulseSettingType)
    Dim FileName As String
    FileName = App.path & "\" & SETTING_PATH & "PulseSetting.cfg"
    
    Dim pFileItemList() As PulseFileItemType
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
    Dim pFileItem As PulseFileItemType
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
    pFileItem.pulseSetting = pulseSetting
    
    Open FileName For Binary As #1
        Put 1, 1, pFileHeader
        Put 1, pos + 1, pFileItem
    Close 1
End Function

Public Function DeleteConfig(ByVal i As Integer) As Boolean
    Dim FileName As String
    FileName = App.path & "\" & SETTING_PATH & "PulseSetting.cfg"
    
    Dim pFileItemList() As PulseFileItemType
    pFileItemList = LoadAll()

    Dim pFileHeader As FileHeaderType
    Dim pFileItem As PulseFileItemType
    Dim pos As Integer
    
    pFileHeader.count = UBound(pFileItemList)
    pFileHeader.count = pFileHeader.count - 1
        
    pos = 0
    pos = pos + LenB(pFileHeader)
    pos = pos + (i) * LenB(pFileItem)
        
    Open FileName For Binary As #1
        Put 1, 1, pFileHeader
            
        For i = i To pFileHeader.count - 1
            pFileItem.name = pFileItemList(i + 1).name
            pFileItem.pulseSetting = pFileItemList(i + 1).pulseSetting
            Put 1, pos + 1, pFileItem
            pos = pos + LenB(pFileItem)
        Next i
    Close 1
End Function

Public Function AssertEqualPulseData(ByRef pulseSetting As PulseSettingType, ByRef dest As PulseSettingType) As Boolean
    Dim i As Integer
    Dim j As Integer
    
    out.log "<<<<<<<<<<<<<<<<<<<<<     AssertEqualPulseData   <<<<<<<<<<<<<<<<<<<"
    out.log "start compare pulseSetting.Stages(j).Value(i) <> dest.Stages(j).Value(i)"
    
    Dim hasNotEqual As Boolean
    hasNotEqual = False
    
    For i = 0 To 7
        For j = 0 To 6
            If Not out.eq(pulseSetting.Stages(j).Value(i), dest.Stages(j).Value(i)) Then
                out.log "<>  i=" & i & " j=" & j & "  " & pulseSetting.Stages(j).Value(i) & "  >-<  " & dest.Stages(j).Value(i)
                hasNotEqual = True
            Else
                out.log "==  i=" & i & " j=" & j & "  " & pulseSetting.Stages(j).Value(i) & "  >-<  " & dest.Stages(j).Value(i)
            End If
        Next
    Next
    
    out.log "pulseSetting.General.Value(j - 1) <> dest.General.Value(j - 1) "
    
    For j = 1 To 4
        If Not out.eq(pulseSetting.General.Value(j - 1), dest.General.Value(j - 1)) Then
            out.log "<> j=" & j & "  " & pulseSetting.General.Value(j - 1) & "  >-<  " & dest.General.Value(j - 1)
            hasNotEqual = True
        Else
            out.log "== j=" & j & "  " & pulseSetting.General.Value(j - 1) & "  >-<  " & dest.General.Value(j - 1)
        End If
    Next
    
    AssertEqualPulseData = Not hasNotEqual
    out.log ">>>>>>>>>>>>>>>>>>       return " & AssertEqualPulseData & "           >>>>>>>>>>>>>>>>>>>>"
End Function
