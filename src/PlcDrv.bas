Attribute VB_Name = "PLCDrv"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'Function Prototypes
Declare Function DTL_INIT Lib "DTL32.DLL" (ByVal def&) As Long
Declare Sub DTL_UNINIT Lib "DTL32.DLL" (ByVal param As Long)

Declare Function DTL_C_DEFINE Lib "DTL32.DLL" (NameId&, ByVal def As String) As Long
Declare Function DTL_UNDEF Lib "DTL32.DLL" (ByVal NameId&) As Long

Declare Function DTL_DRIVER_OPEN Lib "DTL32.DLL" (ByVal nDriverId&, ByVal szDriverName As String, ByVal timeout&) As Long
Declare Function DTL_DRIVER_CLOSE Lib "DTL32.DLL" (ByVal nDriverId&, ByVal timeout&) As Long

Declare Function DTL_READ_W Lib "DTL32.DLL" (ByVal NameId&, Variable As Any, Iostat&, ByVal timeout&) As Long
Declare Function DTL_WRITE_W Lib "DTL32.DLL" (ByVal NameId&, Variable As Any, Iostat&, ByVal timeout&) As Long

Declare Function DTL_PUT_SLC500_FLT Lib "DTL32.DLL" (ByVal invalue As Single, ByRef buf() As Byte) As Long


Declare Sub DTL_ERROR_S Lib "DTL32.DLL" (ByVal Status&, ByVal errstr$, ByVal StrSize%)



Global buffer(50) As Integer
Global bufSingle(50) As Single

Public Const DATA_PATH As String = "Data"
Public Const SETTING_PATH As String = "Set"

Public Const WeldStages_Res_Start As Integer = 2000
Public Const PlcStages_Res_Start As Integer = 3000
Public WeldStages(7) As String
Public PlcStages(12) As String

Public ProcessSetting As Integer

Public beActive As Boolean
'dim
Dim UtlServer As IServer

Dim handle, handle_PC_Data&, handle_PC_Monitor&, response&



Public g_Error_String As String * 80


Public RUN_PHASE As String

Private Status As Long
Public IO_STATUS As Long

Public IsSimulate As Integer
Public SimulatePath As String

Public Calibrate_Distance As Boolean


Public Function OpenPLCConnection() As Integer
    IsSimulate = GetSetting(App.EXEName, "Simulate", "IsSimulate", 0)
    SimulatePath = GetSetting(App.EXEName, "Simulate", "SimulateFilename", App.path & "\T0039.WLD")
    Calibrate_Distance = "1" = Left(GetSetting(App.EXEName, "Calibration", "value", "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"), 1)

    beActive = True

    If IsSimulate = 1 Then
        Set UtlServer = New CPlcSimulate
    Else
        Set UtlServer = New CPlcDrv
    End If
    Dim i As Integer
    
    For i = 0 To 6
        WeldStages(i) = LoadResString(WeldStages_Res_Start + i + 1) 'init
    Next
    
    For i = 1 To 12
        PlcStages(i) = LoadResString(PlcStages_Res_Start + i)  'init
    Next


   'Create definition table
    Status = UtlServer.Init(10)
    If (Status <> DTL_SUCCESS) Then
        Call ClosePLCConection
        RUN_PHASE = "DTL_INIT"
        GoTo ERROR_HANDLE
    End If
    
    Status = UtlServer.OpenDriver(1, "AB_232_DF1-1", 2000)
    If (Status <> DTL_SUCCESS) Then
        Call ClosePLCConection
        RUN_PHASE = "DTL_DRIVER_OPEN"
        GoTo ERROR_HANDLE
    End If
    
    OpenPLCConnection = 0
Exit Function
ERROR_HANDLE:


    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    OpenPLCConnection = Status
    Exit Function
End Function

Public Function ClosePLCConection()
    Call UtlServer.CloseDriver(1, 1000)
    Call UtlServer.Uninit(0)
End Function

Public Function ReadCurrentProcessSetting(ByRef ps As Integer) As Long
    Dim buffer_signal(1) As Integer
    Dim def_signal As String
    def_signal = "N21:3,1,WORD,READ,AB:LOCAL,1,SLC500,1"
    
    Dim handle_signal As Long
    
    DoEvents
    
    Status = UtlServer.Define(handle_signal, def_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.ReadInt(handle_signal, buffer_signal, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    ps = buffer_signal(0)
    
    DoEvents
    
    Status = UtlServer.Undef(handle_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If

    DoEvents
    
Exit Function
ERROR_HANDLE:
    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    ReadCurrentProcessSetting = Status
    Exit Function
End Function

Public Function PreparePcMonitor() As Long
    Status = UtlServer.Define(handle_PC_Monitor, "N22:0,12,WORD,READ,AB:LOCAL,0,SLC500,1")
    If (Status <> DTL_SUCCESS) Then
      Call UtlServer.ErrorStr(Status, g_Error_String, 80)
        RUN_PHASE = "DEF FOR PC MONITOR"
        GoTo ERROR_HANDLE
    End If
    
Exit Function
ERROR_HANDLE:
    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    PreparePcMonitor = Status
    Exit Function
End Function

Public Function ClosePcMonitor() As Long
    Status = UtlServer.Undef(handle_PC_Monitor)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR PC MONITOR"
        GoTo ERROR_HANDLE
    End If

Exit Function
ERROR_HANDLE:
    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    ClosePcMonitor = Status
    Exit Function
End Function

Public Function ReadPcMonitor(ByRef wm As WeldMonitor) As Long
    Status = UtlServer.ReadInt(handle_PC_Monitor, buffer, IO_STATUS, 12345)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "READER PC MONITOR"
        GoTo ERROR_HANDLE
    End If

    If Calibrate_Distance Then
        wm.data.Dist = buffer(0) / 100
    Else
        wm.data.Dist = buffer(0)
    End If
    
    wm.data.Amp = buffer(1)
    wm.data.PsiUpset = buffer(2)
    wm.data.Volt = buffer(3)
    wm.data.PsiOpen = buffer(4)
    
    
    wm.WeldCycle = buffer(8)
    wm.data.WeldStage = buffer(9)
    wm.BoschValve = buffer(10)
    wm.data.PlcStage = buffer(11)
    
    '0   DIST scaled reading in mm * 100
    '1   AMP scaled reading in A
    '2   PSI scaled reading in psi
    '3   VOLT scaled reading in V
    '4   PSI2 scaled reading in psi
    '8   Weld cycle status 0-Idle, 1-Cycle
    '9   Weld stage 0-init, 1-preflash 2-flash 3-boost 4-upset 5-forge 6-shear
    '10  Bosch valve
    '11  PLC stage


Exit Function
ERROR_HANDLE:

    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    ReadPcMonitor = Status
    Exit Function
End Function


Public Function WritePulseData(pulseSetting As PulseSettingType) As Long
    
    '    Distance As Single
    '    Voltage As Single
    '    Time As Single
    '    CurrentSetpoint1 As Single
    '    CurrentSetpoint2 As Single
    '    CurrentSetpoint3 As Single
    '    ForwardSpeed As Single
    '    ReverseSpeed As Single
    
    Dim def(8) As String
    def(0) = "F62:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(1) = "F64:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(2) = "F66:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(3) = "F68:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(4) = "F70:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(5) = "F72:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(6) = "F74:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    def(7) = "F76:1,7,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1"
    
    def(8) = "F78:0,5,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1" 'General
    
    Dim i As Integer
    Dim j As Integer
    
    Dim handle As Long
    
    Dim InitialVoltage As Long
    InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 430))
    
    DoEvents
    
    For i = 0 To 7
        If i = 1 Then
            For j = 0 To 6
                bufSingle(j) = CInt(pulseSetting.Stages(j).Value(i) * InitialVoltage / 100) 'Voltage
            Next j
        Else
            For j = 0 To 6
                bufSingle(j) = pulseSetting.Stages(j).Value(i)
            Next
        End If
        
        
    
        DoEvents
    
        Status = UtlServer.Define(handle, def(i))
        If (Status <> DTL_SUCCESS) Then
            RUN_PHASE = "DEFINE FOR WRITE PULSE SETTING " & i
            GoTo ERROR_HANDLE
        End If
        
        DoEvents
    
        Status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
        If (Status <> DTL_SUCCESS) Then
            RUN_PHASE = "DO WRITE PULSE SETTING " & i
            GoTo ERROR_HANDLE
        End If
        
        DoEvents
    
        Status = UtlServer.Undef(handle)
        If (Status <> DTL_SUCCESS) Then
            RUN_PHASE = "UNDEF FOR WRITE PULSE SETTING " & i
            GoTo ERROR_HANDLE
        End If
    Next
    
    '0   1   Parameter set index
    '1   2   Current in upset timer in seconds
    '2   3   Upset in millimeter
    '3   4   Holding timer for tension use in seconds
    '4   5   Forging force in tonnes
    bufSingle(0) = 6
    For j = 1 To 4
        bufSingle(j) = pulseSetting.General.Value(j - 1)
    Next
    
    DoEvents
    
    Status = UtlServer.Define(handle, def(8))
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    '++
    Dim buffer_signal(1) As Integer
    Dim def_signal As String
    def_signal = "N21:4,1,WORD,MODIFY,AB:LOCAL,1,SLC500,1"
    
    Dim handle_signal As Long
    buffer_signal(0) = 2
    
    DoEvents
    
    Status = UtlServer.Define(handle_signal, def_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteInt(handle_signal, buffer_signal, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    '++

    WritePulseData = 0
    
Exit Function
ERROR_HANDLE:

    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    WritePulseData = Status
    Exit Function
End Function

Public Function WriteRegularData(regularSetting As RegularSettingType) As Long
    
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
    
    Dim def As String
    def = "F60:0,15,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1" 'General
    
    Dim j As Integer
    Dim handle As Long
    
    Dim InitialVoltage As Long
    InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 430))
    
    bufSingle(0) = 6
    For j = 1 To 14
        bufSingle(j) = regularSetting.Value(j - 1)
    Next
    
    j = 4
    bufSingle(j) = CInt(regularSetting.Value(j - 1) * InitialVoltage / 100) 'Voltage
    j = 5
    bufSingle(j) = CInt(regularSetting.Value(j - 1) * InitialVoltage / 100) 'Voltage
    j = 6
    bufSingle(j) = CInt(regularSetting.Value(j - 1) * InitialVoltage / 100) 'Voltage
    
    
    
    DoEvents
    
    Status = UtlServer.Define(handle, def)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR WRITE REGULAR SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO WRITE REGULAR SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR WRITE REGULAR SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    
    
    Dim buffer_signal(1) As Integer
    Dim def_signal As String
    def_signal = "N21:4,1,WORD,MODIFY,AB:LOCAL,1,SLC500,1"
    
    DoEvents
    
    Dim handle_signal As Long
    buffer_signal(0) = 1
    Status = UtlServer.Define(handle_signal, def_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteInt(handle_signal, buffer_signal, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR ACTIVE PULSE SETTING"
        GoTo ERROR_HANDLE
    End If
        
    WriteRegularData = 0

Exit Function
ERROR_HANDLE:

    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    WriteRegularData = Status
    Exit Function
End Function


Public Function WriteCalibrationData(ca() As Single) As Long
    
    '5   1   LVDT calibration rate in mm/DU
    '6   2   LVDT calibration ZeroPoint in mm
    '7   3   LVDT calibration offset in mm
    '8   4   AMP calibration rate in A/DU
    '9   5   AMP calibration ZeroPoint in amps
    '10  6   AMP offset in Amps
    '11  7   Voltage calibration rate in V/DU
    '12  8   Volt calibration ZeroPoint in volts
    '13  9   Volt calibration offset in volts
    '14  10  PSI calibration rate in psi/DU
    '15  11  PSI calibration ZeroPoint in psi
    '16  12  PSI - calibration for offset in psi
    
    
    Dim def As String
    def = "F8:5,12,FLOAT,MODIFY,AB:LOCAL,1,SLC500,1" 'General
    
    Dim j As Integer
    Dim handle As Long
    
    For j = 0 To 11
        bufSingle(j) = ca(j)
    Next
    
    DoEvents
    
    Status = UtlServer.Define(handle, def)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR WRITE PULSE SETTING GENERAL"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    
    '++
    Dim buffer_signal(1) As Integer
    Dim def_signal As String
    def_signal = "N21:5,1,WORD,MODIFY,AB:LOCAL,1,SLC500,1" 'General
    
    Dim handle_signal As Long
    buffer_signal(0) = 1
    
    DoEvents
    
    Status = UtlServer.Define(handle_signal, def_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DEFINE FOR ACTIVE CALIBRATION SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.WriteInt(handle_signal, buffer_signal, IO_STATUS, 1000)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "DO ACTIVE CALIBRATION SETTING"
        GoTo ERROR_HANDLE
    End If
    
    DoEvents
    
    Status = UtlServer.Undef(handle_signal)
    If (Status <> DTL_SUCCESS) Then
        RUN_PHASE = "UNDEF FOR ACTIVE CALIBRATION SETTING"
        GoTo ERROR_HANDLE
    End If
        
    WriteCalibrationData = 0
Exit Function
ERROR_HANDLE:

    beActive = False

    Call UtlServer.ErrorStr(Status, g_Error_String, 80)
    WriteCalibrationData = Status
    Exit Function
End Function
