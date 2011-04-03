Attribute VB_Name = "PlcAnalysiser"
Option Explicit

Const INIT_STAGE  As Integer = 0
Const PREFLASH_STAGE As Integer = 1
Const FLASH_STAGE As Integer = 2
Const BOOST_STAGE As Integer = 3
Const UPSET_STAGE As Integer = 4
Const FORGE_STAGE As Integer = 5
Const SHEAR_STAGE As Integer = 6


Const PRE_FALSH As Integer = 1
Const FLASH_I As Integer = 2
Const FLASH_II As Integer = 3
Const FLASH_III As Integer = 4
Const FLASH_IV As Integer = 5
Const BOOST_I As Integer = 6
Const BOOST_II As Integer = 7
Const UPSET As Integer = 8
Const FORGE As Integer = 9
Const SHEAR_I As Integer = 10
Const SHEAR_II As Integer = 11
Const SHEAR_III As Integer = 12


Private analysisDefine As WeldAnalysisDefineType


Private buf(30000) As WeldData

'1       PRE-FALSH
'2       FLASH-i
'3       FLASH-II
'4       FLASH-III
'5       FLASH-IV
'6       BOOST-I
'7       BOOST-II
'8       UPSET
'9       FORGE
'10      SHEAR-I
'11      SHEAR-II
'12      SHEAR-III


Private Function GetAnalysisDefine()

analysisDefine.FlashEnable = GetSetting(App.EXEName, "AnalysisDefine", "FlashEnable", 1)
analysisDefine.BoostEnable = GetSetting(App.EXEName, "AnalysisDefine", "BoostEnable", 1)
analysisDefine.UpsetEnable = GetSetting(App.EXEName, "AnalysisDefine", "UpsetEnable", 1)
analysisDefine.ForgeEnable = GetSetting(App.EXEName, "AnalysisDefine", "ForgeEnable", 1)
analysisDefine.SlippageEnable = GetSetting(App.EXEName, "AnalysisDefine", "SlippageEnable", 1)
analysisDefine.CurrentInterruptEnable = GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptEnable", 1)
analysisDefine.ShortCircuitEnable = GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitEnable", 1)
analysisDefine.TotalRailUsageEnable = GetSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageEnable", 1)

analysisDefine.FlashMin = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashMin", 0.14))
analysisDefine.FlashMax = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashMax", 0.25))
analysisDefine.BoostMin = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostMin", 0.75))
analysisDefine.BoostMax = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostMax", 1.2))
analysisDefine.UpsetMin = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetMin", 14#))
analysisDefine.UpsetMax = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetMax", 20#))
analysisDefine.ForgeMin = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ForgeMin", 30))
analysisDefine.ForgeMax = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ForgeMax", 60))
analysisDefine.SlippageUpsetTime = CSng(GetSetting(App.EXEName, "AnalysisDefine", "SlippageUpsetTime", 0.75))
analysisDefine.SlippageUpset = CSng(GetSetting(App.EXEName, "AnalysisDefine", "SlippageUpset", 22#))
analysisDefine.CurrentInterruptCurrent = CSng(GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptCurrent", 100))
analysisDefine.CurrentInterruptTime = CSng(GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptTime", 2#))
analysisDefine.ShortCircuitCurrent = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitCurrent", 550))
analysisDefine.ShortCircuitTime = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitTime", 0.8))
analysisDefine.TotalRailUsageTotalRail = CSng(GetSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageTotalRail", 30))
analysisDefine.InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 430))
analysisDefine.BoostSpeedTimeRange = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostSpeedTimeRange", 2))
analysisDefine.UpsetCurrentMinimum = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetCurrentMinimum", 0))
analysisDefine.UpsetDiameter_Pistonside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Pistonside)", 0))
analysisDefine.UpsetDiameter_Rodside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Rodside)", 0))


End Function

Private Function anaPreFlash(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
Dim i As Integer

Dim sumVol As Double
Dim sumCurrent As Double

    For i = startPos To stopPos
        sumVol = sumVol + buf(i).Volt
        sumCurrent = sumCurrent + buf(i).Amp
    Next
        
    r.PreFlashVoltage = CInt(sumVol / (stopPos - startPos + 1))
    r.PreFlashCurrent = CInt(sumCurrent / (stopPos - startPos + 1))
    
    r.PreFlashRailUsed = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.PreFlashDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
    anaPreFlash = r
End Function


Private Function anaFlash(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
Dim i As Integer

Dim sumVol As Double
Dim sumCurrent As Double

    For i = startPos To stopPos
        sumVol = sumVol + buf(i).Volt
        sumCurrent = sumCurrent + buf(i).Amp
    Next
        
    r.FlashVoltage = CInt(sumVol / (stopPos - startPos + 1))
    r.FlashCurrent = CInt(sumCurrent / (stopPos - startPos + 1))
    
    r.FlashRailUsed = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.FlashDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
           
    ' OLD LOGIC
    '    For i = stopPos To startPos Step -1
    '        If buf(stopPos).Time - buf(i).Time >= analysisDefine.FlashSpeedTimeRange Then
    '            Exit For
    '        End If
    '    Next
    '
    '    r.FlashSpeed = (buf(stopPos).Dist - buf(i + 1).Dist) / (buf(stopPos).Time - buf(i + 1).Time)
    ' OLD LOGIC
    
    '   NEW  Logic
    
    ' Change at 2011-2-24 by wangshilian
    Dim bStartOK As Boolean
    Dim bStart As Long
    Dim bStop As Long
    
    
    bStartOK = False
    For i = startPos To stopPos
        If buf(i).PlcStage = FLASH_I Then
            bStartOK = True
            bStart = i
            Exit For
        End If
    Next i
        
    If bStartOK Then
        For i = i To stopPos
            If buf(i).PlcStage <> FLASH_I Then
                Exit For
            End If
        Next i
        For i = i To stopPos
            If buf(i).PlcStage <> FLASH_II Then
                Exit For
            End If
        Next i
        
        For i = i To stopPos
            If buf(i).PlcStage <> FLASH_III Then
                Exit For
            End If
        Next i
        
        bStop = i
        r.FlashSpeed = (buf(bStop - 1).Dist - buf(bStart).Dist) / (buf(bStop - 1).Time - buf(bStart).Time)
    Else
        r.FlashSpeed = 0
    End If
    
'   End of new logic


    
    If analysisDefine.FlashEnable Then
    
    If analysisDefine.FlashMin > r.FlashSpeed Or r.FlashSpeed > analysisDefine.FlashMax Then
        r.Succeed = NO
        r.FlashSpeedSucceed = NO
    Else
        r.FlashSpeedSucceed = OK
    End If
        
    End If
    
        
    anaFlash = r
End Function




Private Function anaBoost(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
    
Dim i As Integer
Dim sumVol As Double
Dim sumCurrent As Double

    For i = startPos To stopPos
        sumVol = sumVol + buf(i).Volt
        sumCurrent = sumCurrent + buf(i).Amp
    Next
        
    r.BoostVoltage = CInt(sumVol / (stopPos - startPos + 1))
    r.BoostCurrent = CInt(sumCurrent / (stopPos - startPos + 1))
    
    r.BoostRailUsed = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.BoostDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
    
'    'TODO   OLD LOGIC
'    For i = stopPos To startPos Step -1
'        If buf(stopPos).Time - buf(i).Time >= 1 Then
'            Exit For
'        End If
'    Next
'
'    r.BoostSpeed = (buf(stopPos).Dist - buf(i + 1).Dist) / (buf(stopPos).Time - buf(i + 1).Time)

'   NEW  Logic
'   Change at 2011-2-24 by wangshilian
    Dim bStartOK As Boolean
    Dim bStart As Long
    Dim bStop As Long
    
    
    bStartOK = False
    For i = startPos To stopPos
        If buf(i).PlcStage = BOOST_II Then
            bStartOK = True
            bStart = i
            Exit For
        End If
    Next i
        
    If bStartOK Then
        For i = i To stopPos
            If buf(i).PlcStage <> BOOST_II Then
                Exit For
            End If
        Next i
        
        bStop = i
        r.BoostSpeed = (buf(bStop - 1).Dist - buf(bStart).Dist) / (buf(bStop - 1).Time - buf(bStart).Time)
    Else
        r.BoostSpeed = 0
    End If
    
'   End of new logic
                       
    If analysisDefine.BoostEnable Then
    
    If analysisDefine.BoostMin > r.BoostSpeed Or r.BoostSpeed > analysisDefine.BoostMax Then
        r.Succeed = NO
        r.BoostSpeedSucceed = NO
    Else
        r.BoostSpeedSucceed = OK
    End If
    
    End If
    
    
    Dim bIn As Boolean
    Dim sTime As Single
    
    If analysisDefine.CurrentInterruptEnable Then
        r.HasCurrentInterruptinBoost = False
        
        For i = startPos To stopPos
            If buf(i).Amp <= analysisDefine.CurrentInterruptCurrent Then

                If bIn = True Then
                    If buf(i).Time - sTime >= analysisDefine.CurrentInterruptTime Then
                        r.Succeed = NO
                        r.HasCurrentInterruptinBoost = NO
                        Exit For
                    End If
                Else
                    bIn = True
                    sTime = buf(i).Time
                End If
            Else
                bIn = False
                sTime = buf(i).Time
            End If
        Next
    
    End If
    
    
    
    bIn = False
    sTime = 0
    
    If analysisDefine.ShortCircuitEnable Then
        r.HasShortCircuitinBoost = False
        
        For i = startPos To stopPos
            If buf(i).Amp > analysisDefine.ShortCircuitCurrent Then
                        
                If bIn = True Then
                    If buf(i).Time - sTime >= analysisDefine.ShortCircuitTime Then
                        r.HasShortCircuitinBoost = NO
                        r.Succeed = NO
                        Exit For
                    End If
                Else
                    bIn = True
                    sTime = buf(i).Time
                End If
            Else
                bIn = False
                sTime = buf(i).Time
            End If
        Next
    End If
    
    
    anaBoost = r
End Function



Private Function anaUpset(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
    
Dim i As Integer
Dim maxCurrent As Long
Dim bIn As Boolean
Dim sTime As Single

    For i = startPos To stopPos
        If buf(i).Amp > maxCurrent Then
            maxCurrent = buf(i).Amp
        End If
    Next
        
    r.UpsetMaxCurrent = maxCurrent
    r.UpsetRailUsage = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.UpsetDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
    
    If analysisDefine.SlippageEnable Then
        If r.UpsetDuration < analysisDefine.SlippageUpsetTime Or r.UpsetRailUsage > analysisDefine.SlippageUpset Then
            r.Succeed = NO
            r.HasSlippage = NO
        Else
            r.HasSlippage = OK
        End If
    End If
    
    bIn = False
    For i = startPos To stopPos
        If buf(i).Amp > analysisDefine.UpsetCurrentMinimum Then
            If bIn = False Then
                bIn = True
                sTime = buf(i).Time
            End If
        ElseIf bIn = True Then
            r.UpsetCurrentOnTime = buf(i - 1).Time - sTime
            Exit For
        End If
    Next
        
    If analysisDefine.UpsetEnable Then
    If analysisDefine.UpsetMin > r.UpsetRailUsage Or r.UpsetRailUsage > analysisDefine.UpsetMax Then
        r.Succeed = NO
        r.UpsetRailUsageSucceed = NO
    Else
        r.UpsetRailUsageSucceed = OK
    End If
    End If
        
    anaUpset = r
End Function



Private Function anaForge(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
Dim i As Integer


Dim force As Single
Dim sumForce As Double

' OLD LOGIC
'    For i = stopPos To startPos Step -1
'        force = PlcAnalysiser.toForce(buf(i).PsiUpset, buf(i).PsiOpen)
'        sumForce = sumForce + force
'    Next
'
'    r.ForgeAverageForce = CInt(sumForce / (stopPos - startPos + 1))

    If stopPos - startPos > 3 Then
        i = startPos
        sumForce = PlcAnalysiser.toForce(buf(i).PsiUpset, buf(i).PsiOpen, analysisDefine)
        i = i + 1
        sumForce = sumForce + PlcAnalysiser.toForce(buf(i).PsiUpset, buf(i).PsiOpen, analysisDefine)
        i = i + 1
        sumForce = sumForce + PlcAnalysiser.toForce(buf(i).PsiUpset, buf(i).PsiOpen, analysisDefine)
        r.ForgeAverageForce = CInt(sumForce / 3)
    Else
        r.ForgeAverageForce = 0
    End If
    
    If analysisDefine.ForgeEnable Then
        If analysisDefine.ForgeMin > r.ForgeAverageForce Or r.ForgeAverageForce > analysisDefine.ForgeMax Then
            r.Succeed = NO
            r.ForgeForceSucceed = NO
        Else
            r.ForgeForceSucceed = OK
        End If
    End If
    r.ForgeDuration = buf(stopPos).Time - buf(startPos - 1).Time
            
    anaForge = r
End Function

Private Function anaAll(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
Dim i As Integer

          
            
    r.TotalRailUsage = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.TotalDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
    If analysisDefine.TotalRailUsageEnable Then
    
    If r.TotalRailUsage < analysisDefine.TotalRailUsageTotalRail Then
        r.Succeed = NO
        r.TotalRailUsageSucceed = NO
    Else
        r.TotalRailUsageSucceed = True
    End If
    
    End If
            
    Dim sumVol As Double
    Dim sumCurrent As Double

    For i = startPos To stopPos
        sumVol = sumVol + buf(i).Volt
        sumCurrent = sumCurrent + buf(i).Amp
    Next
                
    'TODO
    r.OverallImpedance = (sumVol / sumCurrent) * (1000000 / 3600) / 2.4
    
    
    'TODO Holding Time
        
    anaAll = r
End Function




Public Function Analysis(buf() As WeldData, count As Integer) As WeldAnalysisResultType

Call GetAnalysisDefine

Dim stage As Integer
Dim pos As Integer
Dim startPos As Integer
Dim lastPos As Integer

Dim r As WeldAnalysisResultType
r.Succeed = PlcDeclare.OK


stage = INIT_STAGE
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next

startPos = pos

stage = PREFLASH_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If


'Nav to  preflash end
lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
r = PlcAnalysiser.anaPreFlash(buf, lastPos, pos - 1, r)
 
'=====================================================
stage = FLASH_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If

lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
r = PlcAnalysiser.anaFlash(buf, lastPos, pos - 1, r)
 
'=====================================================
stage = BOOST_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If

lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
r = PlcAnalysiser.anaBoost(buf, lastPos, pos - 1, r)
 
'=====================================================
stage = UPSET_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If
  
lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
r = PlcAnalysiser.anaUpset(buf, lastPos, pos - 1, r)
 
'=====================================================
stage = FORGE_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If
  
lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
r = PlcAnalysiser.anaForge(buf, lastPos, pos - 1, r)
r = PlcAnalysiser.anaAll(buf, startPos, pos - 1, r)
 
'=====================================================
stage = SHEAR_STAGE
If buf(pos).WeldStage <> stage Then
    GoTo OVER
End If
  
lastPos = pos
For pos = pos To count
    If buf(pos).WeldStage <> stage Then
        Exit For
    End If
Next
 


Analysis = r
Exit Function
OVER:
    r.Succeed = INTERRUPT
    Analysis = r
End Function


Public Function toForce(Pupset As Long, Popen As Long, anaDefine As WeldAnalysisDefineType) As Double

    Dim Dpiston As Double
    Dim Drod As Double
    'Call GetAnalysisDefine
    Dpiston = anaDefine.UpsetDiameter_Pistonside
    Drod = anaDefine.UpsetDiameter_Rodside
    
'    toForce = 3.1415926 * ( _
'        0.0703 * Pupset * ((Dpiston / 2) * (Dpiston / 2) - (Drod / 2) * (Drod / 2)) - _
'         0.0703 * Popen * (Dpiston / 2) * (Dpiston / 2) _
'        ) _
'        / 100 / 1000
        
        toForce = 2 * 3.1415926 * ( _
        0.0703 * (Abs(Pupset) - Abs(Popen)) * ((Dpiston / 2) * (Dpiston / 2) - ((Drod / 2) * (Drod / 2))) _
        ) _
        / 100 / 1000

        
'
'        toForce = 3.1415926 * ( _
'            0.0703 * ( _
'                Abs(Pupset) * (Dpiston / 2) * (Dpiston / 2) - _
'                Abs(Popen) * ((Dpiston / 2) * (Dpiston / 2) - (Drod / 2) * (Drod / 2)) _
'            ) _
'        ) _
'        / 100 / 1000
        
End Function
        

