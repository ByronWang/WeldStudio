Attribute VB_Name = "PlcDeclare"
Option Explicit

Public Const GeneralMode As String = "G"
Public Const JinanMode As String = "J"
Public Const EngMode As String = "E"


Public Const ReadOnly As Boolean = True
Public Const WeldNumberMode As String = EngMode   'GeneralMode EngMode

Public Const LOAD_ALL_PARAMETER  As Integer = 0
Public Const LOAD_PULSE_SETTING As Integer = 1
Public Const LOAD_REGULAR_SETTING As Integer = 2
Public Const LOAD_CALIBRATION_SETTING As Integer = 3

Public Const REGULAR_MODE As Integer = 1
Public Const PULSE_MODE As Integer = 2

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


Type FileHeader1
     X00 As String * 1
     X01 As Byte
     X02 As String * 1
     X03 As Byte
     X04 As String * 4
     X08 As Byte
     X09 As Byte
     CompanyName As String * &H20 ' X10

     
     X20 As String * &H8D6
     
     
     X900 As String * &HA
     unitName As String * &H1A
     operator As String * &H2
     X922 As String * &HA
     
     X930 As String * &HD0
     
     Xa00 As String * &HA
     Location As String * &H10
     Xa14 As String * &H6
     
     Xa20 As String * &H1F0

     
     Xc10 As String * &H4

End Type

Type FileHeader2
     CompactedWeldNumber As String * &H5
     WeldNumberMode As String * &H1
     Date As String * &HB
     ParamSettingMode As String * &H1
     Xc25 As String * &HA
     
     XC30 As String * &HE0
     
     
     Xd10 As String * &HA
     Time As String * &H8
     Xd22 As String * &HE
     
     H5(&HEC - 1) As Byte
     RecordCount As Integer
     XE1E As String * 4
     XE20 As String * &H5F0

     X1410 As String * &HA
     ParamSettingName As String * &H7
     X1423 As String * &HD

     H6(&H20 - 1) As Byte
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
    'FlashSpeedTimeRange As Long
    InitialVoltage As Long
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
    Succeed  As Integer
    X110 As Integer
    FlashSpeedSucceed  As Integer
    X111 As Integer
    BoostSpeedSucceed  As Integer
    X112 As Integer
    UpsetRailUsageSucceed  As Integer
    X113 As Integer
    ForgeForceSucceed  As Integer
    X114 As Integer
    HasCurrentInterruptinBoost  As Integer
    X115 As Integer
    HasSlippage     As Integer 'TODO
    X116 As Integer
    HasShortCircuitinBoost  As Integer
    X117 As Integer
    TotalRailUsageSucceed    As Integer
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
    header1 As FileHeader1
    header2 As FileHeader2
    data() As Record
    analysisDefine As WeldAnalysisDefineType
    analysisResult As WeldAnalysisResultType
End Type



Global Const DTL_VERSION_ID = -1
Global Const DTL_SUCCESS = 0
Global Const DTL_PENDING = 1
Global Const DTL_E_FAIL = 24


Global Const OK As Integer = 1
Global Const NO As Integer = 2
Global Const INTERRUPT As Integer = 3
Global Const NotUsed As Integer = 4


Public Function Floor(i As Long, modBy As Long) As Long
    Floor = CInt(i / modBy)
    
    If Floor * modBy > i Then
        Floor = Floor - 1
    End If
End Function

