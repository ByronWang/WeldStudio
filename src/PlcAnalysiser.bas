Attribute VB_Name = "PlcAnalysiser"
Option Explicit

Const INIT_STAGE  As Integer = 0
Const PREFLASH_STAGE As Integer = 1
Const FLASH_STAGE As Integer = 2
Const BOOST_STAGE As Integer = 3
Const UPSET_STAGE As Integer = 4
Const FORGE_STAGE As Integer = 5
Const SHEAR_STAGE As Integer = 6


Public analysisDefine As WeldAnalysisDefineType


Public buf(30000) As WeldData


Public Function GetAnalysisDefine()

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
analysisDefine.FlashSpeedTimeRange = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashSpeedTimeRange", 10))
analysisDefine.BoostSpeedTimeRange = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostSpeedTimeRange", 2))
analysisDefine.UpsetCurrentMinimum = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetCurrentMinimum", 0))
analysisDefine.UpsetDiameter_Pistonside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Pistonside)", 0))
analysisDefine.UpsetDiameter_Rodside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Rodside)", 0))


End Function

Public Function anaPreFlash(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
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


Public Function anaFlash(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
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
    
           
           
    For i = stopPos To startPos Step -1
        If buf(stopPos).Time - buf(i).Time >= analysisDefine.FlashSpeedTimeRange Then
            Exit For
        End If
    Next
    
    r.FlashSpeed = (buf(stopPos).Dist - buf(i + 1).Dist) / (buf(stopPos).Time - buf(i + 1).Time)
    
    If analysisDefine.FlashEnable Then
    
    If analysisDefine.FlashMin > r.FlashSpeed Or r.FlashSpeed > analysisDefine.FlashMax Then
        r.succeed = False
        r.FlashSpeedSucceed = False
    Else
        r.FlashSpeedSucceed = True
    End If
        
    End If
    
        
    anaFlash = r
End Function




Public Function anaBoost(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
    
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
    
    
    'TODO
    For i = stopPos To startPos Step -1
        If buf(stopPos).Time - buf(i).Time >= 1 Then
            Exit For
        End If
    Next
    
    r.BoostSpeed = (buf(stopPos).Dist - buf(i + 1).Dist) / (buf(stopPos).Time - buf(i + 1).Time)
           
                       
    If analysisDefine.BoostEnable Then
    
    If analysisDefine.BoostMin > r.BoostSpeed Or r.BoostSpeed > analysisDefine.BoostMax Then
        r.succeed = False
        r.BoostSpeedSucceed = False
    Else
        r.BoostSpeedSucceed = True
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
                        r.succeed = False
                        r.HasCurrentInterruptinBoost = True
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
                        r.HasShortCircuitinBoost = True
                        r.succeed = False
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



Public Function anaUpset(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
    
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
            r.succeed = False
            r.HasSlippage = True
        Else
            r.HasSlippage = False
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
        r.succeed = False
        r.UpsetRailUsageSucceed = False
    Else
        r.UpsetRailUsageSucceed = True
    End If
    End If
        
    anaUpset = r
End Function



Public Function anaForge(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
   
Dim i As Integer


Dim force As Single
Dim sumForce As Double

    For i = stopPos To startPos Step -1
        force = (buf(i).PsiUpset - buf(i).PsiOpen) / 25.4
        sumForce = sumForce + force
    Next
            
    r.ForgeAverageForce = CInt(sumForce / (stopPos - startPos + 1))
    
    If analysisDefine.ForgeEnable Then
        If analysisDefine.ForgeMin > r.ForgeAverageForce Or r.ForgeAverageForce > analysisDefine.ForgeMax Then
            r.succeed = False
            r.ForgeForceSucceed = False
        Else
            r.ForgeForceSucceed = True
        End If
    End If
    r.ForgeDuration = buf(stopPos).Time - buf(startPos - 1).Time
            
    anaForge = r
End Function

Public Function anaAll(buf() As WeldData, startPos As Integer, stopPos As Integer, r As WeldAnalysisResultType) As WeldAnalysisResultType
Dim i As Integer

          
            
    r.TotalRailUsage = buf(stopPos + 1).Dist - buf(startPos).Dist
    r.TotalDuration = buf(stopPos + 1).Time - buf(startPos).Time
    
    If analysisDefine.TotalRailUsageEnable Then
    
    If r.TotalRailUsage < analysisDefine.TotalRailUsageTotalRail Then
        r.succeed = False
        r.TotalRailUsageSucceed = False
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
                
    r.OverallImpedance = (sumVol / sumCurrent) * (1000000 / 3600) / 3
    
    
    'TODO Holding Time
        
    anaAll = r
End Function




Public Function ANALYSIS(buf() As WeldData, count As Integer) As WeldAnalysisResultType

Call GetAnalysisDefine

Dim stage As Integer
Dim pos As Integer
Dim startPos As Integer
Dim lastPos As Integer

Dim r As WeldAnalysisResultType
r.succeed = True

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
 


ANALYSIS = r
Exit Function
OVER:
    ANALYSIS = r
End Function