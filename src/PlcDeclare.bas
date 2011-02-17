Attribute VB_Name = "PlcDeclare"
Option Explicit

Public Const INIT_STAGE  As Integer = 0
Public Const PREFLASH_STAGE As Integer = 1
Public Const FLASH_STAGE As Integer = 2
Public Const BOOST_STAGE As Integer = 3
Public Const UPSET_STAGE As Integer = 4
Public Const FORGE_STAGE As Integer = 5
Public Const SHEAR_STAGE As Integer = 6

Type WeldData
     Dist As Single
     Time As Single
     Amp As Long
     Volt As Long
     PsiUpset As Long
     PsiOpen As Long
     PlcStage As Long
     WeldStage As Long
End Type

Type WeldMonitor
    data As WeldData
    WeldCycle As Long
    BoschValve As Long
End Type

Type WeldHeader
     Init(5) As Byte  'INIT..
     INIT_I As Long
        
     PRE_FLASH(9) As Byte 'PRE-FLASH.
    
     FLASH(5) As Byte 'FLASH.
     FLASH_I As Long
    
     BOOST(5) As Byte 'BOOST.
     BOOST_I As Long
    
     UPSET(5) As Byte 'UPSET.
     UPSETI As Long
    
     FORGE(5) As Byte 'FORGE.
     FORGE_I As Long
    
     SHEAR(5) As Byte 'SHEAR.
     SHEAR_I As Long
    
     HOLDING(7) As Byte 'HOLDING.
     HOLDING_I As Integer
    
End Type

Type Record
    header As WeldHeader
    data As WeldData
End Type


Type FileHeader
     'H1(&HC14 - 1) As Byte
     H1H(&H90A - 1) As Byte
     ParamName As String * &H19
     H1L(&H2EA - 1 + 7) As Byte
     filename As String * &H6
     Date As String * &HC
     H3(&HF4 - 1) As Byte
     'H4(&H8 - 1) As Byte
     Time As String * 8
     H5(&HFA - 1) As Byte
     RecordCount As Long
     H6(&H630 - 1) As Byte
End Type



Type WeldAnalysisDefineType
    FlashEnable As Boolean
    X1 As Integer
    FlashMin As Single
    FlashMax As Single
    BoostEnable As Boolean
    X2 As Integer
    BoostMin As Single
    BoostMax As Single
    UpsetEnable As Boolean
    X3 As Integer
    UpsetMin As Single
    UpsetMax As Single
    ForgeEnable As Boolean
    X4 As Integer
    ForgeMin As Long
    ForgeMax As Long
    SlippageEnable As Boolean
    X5 As Integer
    SlippageUpsetTime As Single
    SlippageUpset As Single
    CurrentInterruptEnable As Boolean
    X6 As Integer
    CurrentInterruptCurrent As Long
    CurrentInterruptTime As Single
    ShortCircuitEnable As Boolean
    X7 As Integer
    ShortCircuitCurrent As Long
    ShortCircuitTime As Single
    TotalRailUsageEnable As Boolean
    X8 As Integer
    TotalRailUsageTotalRail As Long
    FlashSpeedTimeRange As Long
    BoostSpeedTimeRange As Long
    UpsetCurrentMinimum As Long
    UpsetDiameter_Pistonside As Single
    UpsetDiameter_Rodside As Single
End Type

Type InnerAna
    X1 As Single
    X2 As Single
    X3 As Single
    X4 As Single
    X5 As Single
    X6 As Single
    X7 As Single
    X8 As Single
    X9 As Single
    X10 As Single
End Type


Type WeldAnalysisResultType
    succeed  As Integer
    X110 As Integer
    FlashSpeedSucceed  As Boolean
    X111 As Integer
    BoostSpeedSucceed  As Boolean
    X112 As Integer
    UpsetRailUsageSucceed  As Boolean
    X113 As Integer
    ForgeForceSucceed  As Boolean
    X114 As Integer
    HasCurrentInterruptinBoost  As Boolean
    X115 As Integer
    HasSlippage     As Boolean 'TODO
    X116 As Integer
    HasShortCircuitinBoost  As Boolean
    X117 As Integer
    TotalRailUsageSucceed    As Boolean
    X118 As Integer
    X129 As Single
    PreFlashVoltage     As Long
    FlashVoltage     As Long
    BoostVoltage     As Long
    X13 As Single
    X14 As Single
    X15 As Single
    X16 As Single
    X17 As Single
    PreFlashCurrent     As Long
    FlashCurrent     As Long
    BoostCurrent     As Long
    X18 As Single
    X19 As Single
    X20 As Single
    X21 As Single
    UpsetMaxCurrent     As Long
    HoldingTime As Long
    ForgeAverageForce     As Long
    X23 As Single
    PreFlashRailUsed     As Single
    FlashRailUsed     As Single
    BoostRailUsed     As Single
    UpsetRailUsage     As Single
    X24 As Single
    X25 As Single
    X26 As Single
    X27 As Single
    PreFlashDuration     As Single
    FlashDuration     As Single
    BoostDuration     As Single
    UpsetDuration     As Single
    ForgeDuration     As Single
    X29 As Single
    X30 As Single
    X31 As Single
    X32 As Single
    TotalRailUsage     As Single
    TotalDuration     As Single
    FlashSpeed     As Single
    BoostSpeed     As Single
    UpsetCurrentOnTime     As Single
    OverallImpedance As Single
    V33 As Single
    V34 As Single

End Type


Type FileR
    header As FileHeader
    data() As Record
    analysisDefine As WeldAnalysisDefineType
    analysisResult As WeldAnalysisResultType
End Type



Global Const DTL_VERSION_ID = -1
Global Const DTL_SUCCESS = 0
Global Const DTL_PENDING = 1
Global Const DTL_E_FAIL = 24
