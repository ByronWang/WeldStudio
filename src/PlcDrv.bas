Attribute VB_Name = "PLCDrv"
Option Explicit


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


Declare Sub DTL_ERROR_S Lib "DTL32.DLL" (ByVal status&, ByVal errstr$, ByVal StrSize%)



Global buffer(50) As Integer
Global bufSingle(50) As Single

Public Const DATA_PATH As String = "Data"
Public Const SETTING_PATH As String = "Set"

'dim
Dim UtlServer As IServer

Dim handle, handle_PC_Data&, handle_PC_Monitor&, status&, response&




Public beActive As Boolean
Public RUN_PHASE As String

Public Const WeldStages_Res_Start As Integer = 2000
Public WeldStages(7) As String
Public Const PlcStages_Res_Start As Integer = 3000
Public PlcStages(12) As String
Public IO_STATUS As Long

Public IsSimulate As Integer
Public SimulatePath As String

Public Sub InitPLCConnection()
IsSimulate = GetSetting(App.EXEName, "Simulate", "IsSimulate", 0)
SimulatePath = GetSetting(App.EXEName, "Simulate", "SimulateFilename", App.path & "\T0039.WLD")

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


Dim errstr As String * 80
   'Create definition table
status = UtlServer.Init(10)
If (status <> DTL_SUCCESS) Then
    RUN_PHASE = "DTL_INIT"
    GoTo ERROR_HANDLE
End If

status = UtlServer.OpenDriver(1, "AB_232_DF1-1", 2000)
If (status <> DTL_SUCCESS) Then
    RUN_PHASE = "DTL_DRIVER_OPEN"
    GoTo ERROR_HANDLE
End If

status = UtlServer.Define(handle_PC_Monitor, "N22:0,12,WORD,READ,AB:LOCAL,0,SLC500,1")
If (status <> DTL_SUCCESS) Then
  Call UtlServer.ErrorStr(status, errstr, 80)
    RUN_PHASE = "DTL_C_DEFINE"
    GoTo ERROR_HANDLE
End If

status = UtlServer.ReadInt(handle_PC_Monitor, buffer, IO_STATUS, 200)
If (status <> DTL_SUCCESS) Then
    RUN_PHASE = "DTL_C_DEFINE"
    GoTo ERROR_HANDLE
End If

    beActive = True
Exit Sub
ERROR_HANDLE:

    beActive = False

    Call UtlServer.ErrorStr(status, errstr, 80)
    Exit Sub
End Sub


Public Function UninitPLCConection()
    Call UtlServer.CloseDriver(1, 1000)
    UtlServer.Uninit (0)
End Function


Public Function readPcMonitor() As WeldMonitor
    status = UtlServer.ReadInt(handle_PC_Monitor, buffer, IO_STATUS, 12345)


    Dim wm As WeldMonitor
    wm.data.Dist = buffer(0) / 100
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


    readPcMonitor = wm
End Function


Public Function WritePulseData(PulseSetting As PulseSettingType)

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

For i = 0 To 7
    For j = 0 To 6
        bufSingle(j) = PulseSetting.Stages(j).Value(i)
    Next

    If UtlServer.Define(handle, def(i)) = 0 Then
        status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
        UtlServer.Undef (handle)
    End If
Next

'0   1   Parameter set index
'1   2   Current in upset timer in seconds
'2   3   Upset in millimeter
'3   4   Holding timer for tension use in seconds
'4   5   Forging force in tonnes
bufSingle(0) = 6
For j = 1 To 4
    bufSingle(j) = PulseSetting.General.Value(j - 1)
Next

If UtlServer.Define(handle, def(8)) = 0 Then
    status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    UtlServer.Undef (handle)
End If

End Function

Public Function WriteRegularData(RegularSetting As RegularSettingType)
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

bufSingle(0) = 6
For j = 1 To 14
    bufSingle(j) = RegularSetting.Value(j - 1)
Next

If UtlServer.Define(handle, def) = 0 Then
    status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    UtlServer.Undef (handle)
End If

End Function




Public Function WriteCalibrationData(ca() As Single)
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

If UtlServer.Define(handle, def) = 0 Then
    status = UtlServer.WriteSingle(handle, bufSingle, IO_STATUS, 1000)
    UtlServer.Undef (handle)
End If

End Function