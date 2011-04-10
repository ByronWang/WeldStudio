VERSION 5.00
Begin VB.Form FrmGraph 
   Appearance      =   0  'Flat
   BackColor       =   &H00400000&
   Caption         =   "FormGraph"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "FrmGraph.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10605
   ScaleWidth      =   15240
   Tag             =   "12000"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picVolt 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      ScaleHeight     =   720
      ScaleWidth      =   8325
      TabIndex        =   18
      Top             =   6535
      Width           =   8355
   End
   Begin VB.PictureBox picDist 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      ScaleHeight     =   720
      ScaleWidth      =   8325
      TabIndex        =   17
      Top             =   5025
      Width           =   8355
   End
   Begin VB.Timer timerDisplay 
      Interval        =   100
      Left            =   3480
      Top             =   4920
   End
   Begin VB.PictureBox picPsi 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      ScaleHeight     =   720
      ScaleWidth      =   8325
      TabIndex        =   1
      Top             =   9240
      Width           =   8355
   End
   Begin VB.PictureBox picAmp 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      ScaleHeight     =   720
      ScaleWidth      =   8325
      TabIndex        =   0
      Top             =   7800
      Width           =   8355
   End
   Begin VB.Timer timerMonitor 
      Interval        =   1
      Left            =   3480
      Top             =   5640
   End
   Begin VB.Label lblProcessSetting 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Regular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   11280
      TabIndex        =   19
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label lblDist 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "67.98"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   10120
      TabIndex        =   16
      Top             =   4680
      Width           =   5000
   End
   Begin VB.Label lblTop 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "A234567890"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   11280
      TabIndex        =   14
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "06:01:01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11640
      TabIndex        =   13
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2011-01-01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblBigCenter 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "A0001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3735
      Left            =   2040
      TabIndex        =   11
      Top             =   720
      Width           =   11295
   End
   Begin VB.Label lblWeldStage 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   13000
      Width           =   1935
   End
   Begin VB.Label lblPlcStage 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   13000
      Width           =   1935
   End
   Begin VB.Label lblVolt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "345"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   10120
      TabIndex        =   8
      Top             =   6120
      Width           =   5000
   End
   Begin VB.Label lblPsi 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   10120
      TabIndex        =   7
      Top             =   8880
      Width           =   5000
   End
   Begin VB.Label lblAmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   10120
      TabIndex        =   6
      Top             =   7440
      Width           =   5000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Force"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   5
      Tag             =   "12040"
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amp"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   4
      Tag             =   "12030"
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Volt"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   3
      Tag             =   "12020"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Tag             =   "12010"
      Top             =   4920
      Width           =   1695
   End
End
Attribute VB_Name = "FrmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fso As FileSystemObject

Dim path As String

Dim timeStart As Long
Dim timePostFromStart As Long
Dim lastStage As Long

Dim dist_scale As Long
Dim volt_scale As Long
Dim amp_scale As Long
Dim psi_scale As Long

Dim buffer(30000) As WeldData
Dim wmRecord As WeldMonitor
Dim wmRecord_Index As Integer

Dim beRecording As Boolean
Dim beSigned As Boolean

Dim Mode_StartRecording As Integer
Dim ModeParam_StartRecoding(5) As Single

Dim ProcessSetting As Integer
Dim weldSerailNumber As Long
'Dim WeldFile As String

Const ANALYSIS_DURATION As Integer = 6000

Dim showMode As ShowModeType

Dim beRequest As Boolean
Dim beUnload As Boolean

Enum ShowModeType
    RECORDING_MODE
    STANDBY_MODE
    ANALYSIS_MODE
End Enum

Dim analysisDefine As WeldAnalysisDefineType
Dim analysisResult As WeldAnalysisResultType

Function SwitchToRecoding(Status As ShowModeType)
    Select Case Status
        Case STANDBY_MODE
            timerMonitor.Enabled = True
            timerDisplay.Tag = ""
            
            lblTop.Caption = "Ready"
            weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
            lblBigCenter.Caption = toWeldNumberShowModel(weldSerailNumber)
            lblBigCenter.ForeColor = &H8000000E
            lblBigCenter.FontSize = 100
            'lblParameter.Caption = GetSetting(App.EXEName, "Parameter", "LastSetting", "DEFAULT")
        Case RECORDING_MODE
            weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
            lblTop.Caption = toWeldNumberShowModel(weldSerailNumber)
            lblBigCenter.FontSize = 140
        Case ANALYSIS_MODE
            timerMonitor.Enabled = False
            timerDisplay.Tag = "ANALYSIS"
    End Select
    showMode = Status
End Function

Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me

    WeldMDIForm.mnuWindow.Enabled = False
    WeldMDIForm.mnuParameters.Enabled = False
    WeldMDIForm.mnuOptions.Enabled = False
    WeldMDIForm.mnuConnect.Enabled = False
    
    amp_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Amp", 500))
    dist_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Dist", 1000))
    volt_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Volt", 100))
    psi_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Force", 120))


    Mode_StartRecording = CInt(GetSetting(App.EXEName, "StartRecording", "StartRecording", 0))
    ModeParam_StartRecoding(1) = CSng(GetSetting(App.EXEName, "StartRecording", "Dist", 2.5))
    ModeParam_StartRecoding(2) = CSng(GetSetting(App.EXEName, "StartRecording", "Amp", 100))
    ModeParam_StartRecoding(3) = CSng(GetSetting(App.EXEName, "StartRecording", "Volt", 450))
    ModeParam_StartRecoding(4) = CSng(GetSetting(App.EXEName, "StartRecording", "Time", 25))

    
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
    analysisDefine.InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 10))
    analysisDefine.BoostSpeedTimeRange = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostSpeedTimeRange", 2))
    analysisDefine.UpsetCurrentMinimum = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetCurrentMinimum", 0))
    analysisDefine.UpsetDiameter_Pistonside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Pistonside)", 0))
    analysisDefine.UpsetDiameter_Rodside = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Rodside)", 0))


    Dim IsSimulate As Integer
    IsSimulate = GetSetting(App.EXEName, "Simulate", "IsSimulate", 0)
    If IsSimulate = 1 Then
        timerMonitor.Interval = 65
    Else
        timerMonitor.Interval = 1
    End If
    

    Set fso = New FileSystemObject
    '
    
    path = App.path & "\Data"
    If Not fso.FolderExists(path) Then
        fso.CreateFolder (path)
    End If
    
    beRecording = False
    timeStart = 0
    timePostFromStart = 0
    wmRecord_Index = 0
    
    ProcessSetting = -1

    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    
    SwitchToRecoding STANDBY_MODE
    
    beSigned = False
    beRequest = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide

    Me.timerDisplay.Enabled = False
    Me.timerMonitor.Enabled = False
    Dim i As Integer
    For i = 0 To 100
        Call Sleep(50)
        If Not beRequest Then
            Exit For
        End If
    Next i
'    While beRequest
'        Call Sleep(300)
'        DoEvents
'    Wend
    
    PLCDrv.ClosePcMonitor
    PLCDrv.ClosePLCConection
    
    WeldMDIForm.mnuWindow.Enabled = True
    WeldMDIForm.mnuParameters.Enabled = True
    WeldMDIForm.mnuOptions.Enabled = True
    WeldMDIForm.mnuConnect.Enabled = True
    If beRequest Then
        beUnload = True
    Else
        beUnload = False
    End If
    Unload Me
ERROR_HANDLE:
End Sub

Private Sub TimerDisplay_Timer()
        
    lblDate.Caption = Format(Date, "YYYY-MM-DD")
    lblTime.Caption = FormatDateTime(Now, vbLongTime)
    
    If Me.timerDisplay.Tag = "ANALYSIS" Then
       If timeGetTime() - timeStart > ANALYSIS_DURATION Then
            SwitchToRecoding (STANDBY_MODE)
       End If
       Exit Sub
    End If
    
    If beRecording Then
        lblBigCenter.Caption = Format(timePostFromStart / 1000, "000")
        'lblBigCenter.Caption = Format(timePostFromStart / 1000, "##0.00")
    End If
    
    If 1 <= wmRecord.data.PlcStage And wmRecord.data.PlcStage <= 12 Then
        lblPlcStage.Caption = PLCDrv.PlcStages(wmRecord.data.PlcStage)
    End If
    
    
    If 1 <= wmRecord.data.WeldStage And wmRecord.data.WeldStage <= 6 Then
        lblWeldStage.Caption = PLCDrv.WeldStages(wmRecord.data.WeldStage)
    End If
    
    
    Dim Dist As Long
    Dim Volt As Long
    Dim Amp As Long
    Dim psi As Long
    
    Dist = wmRecord.data.Dist
    Volt = wmRecord.data.Volt
    Amp = wmRecord.data.Amp
    psi = PlcAnalysiser.toForce(wmRecord.data.PsiUpset, wmRecord.data.PsiOpen, analysisDefine)
    
    'lblTime.Caption = data.Time
    If PLCDrv.Calibrate_Distance Then
        lblDist.FontSize = lblVolt.FontSize
        lblDist.Caption = Format(wmRecord.data.Dist, "##0.0")
    Else
        lblDist.FontSize = lblVolt.FontSize - 3
        lblDist.Caption = Format(wmRecord.data.Dist, "##0")
    End If
    lblVolt.Caption = wmRecord.data.Volt
    lblAmp.Caption = wmRecord.data.Amp
    lblPsi.Caption = psi
    
    If Dist < 0 Then
        Dist = 0
    End If
    
    If Volt < 0 Then
        Volt = 0
    End If
    
    If Amp < 0 Then
        Amp = 0
    End If
    
    If psi < 0 Then
        psi = 0
    End If
    
    
    Dim scale_width As Long
    scale_width = 8000
    Dim scale_height As Long
    scale_height = 4000
    '
    Dim w As Long
    w = Dist * scale_width / dist_scale
    If w >= scale_width Then
        w = scale_width
    End If
    picDist.Width = w
    
    w = Volt * scale_width / volt_scale
    If w >= scale_width Then
        w = scale_width
    End If
    picVolt.Width = w
    
    
    w = Amp * scale_width / amp_scale
    If w >= scale_width Then
        w = scale_width
    End If
    picAmp.Width = w
    
    
    w = psi * scale_width / psi_scale
    If w >= scale_width Then
        w = scale_width
    End If
    picPsi.Width = w
    
    


End Sub



Private Sub TimerMonitor_Timer()
If beRequest Then
    Exit Sub
End If
If beUnload Then
    Unload Me
    Exit Sub
End If


beRequest = True
'9   Weld stage 0-init, 1-preflash 2-flash 3-boost 4-upset 5-forge 6-shear
'11  PLC Stage
'0   DIST scaled reading in mm * 100
'1   AMP scaled reading in A
'3   VOLT scaled reading in V
'2   PSI scaled reading in psi
'4   PSI2 scaled reading in psi
'
'8   Weld cycle status 0-Idle, 1-Cycle
'
'?10 (Force???) Bosch valve

Dim Status As Long


Status = PLCDrv.ReadPcMonitor(wmRecord)
If Status > 0 Then
    GoTo ERROR_HANDLE
End If

If wmRecord.WeldCycle = 1 And 0 <= wmRecord.data.WeldStage And wmRecord.data.WeldStage <= 6 And wmRecord.data.WeldStage >= lastStage Then
    If beRecording Then
        wmRecord.data.Time = timePostFromStart / 1000
        buffer(wmRecord_Index) = wmRecord.data
        wmRecord_Index = wmRecord_Index + 1
    End If
    
    If Not beSigned Then
        timeStart = timeGetTime()
        beSigned = True
        timePostFromStart = 0
    Else
        timePostFromStart = timeGetTime() - timeStart
    End If
        
    If Not beRecording Then ' Start record
        If canStart() Then
            timeStart = timeGetTime()
            timePostFromStart = timeGetTime() - timeStart
            
            beRecording = True
            
            wmRecord_Index = 0
            wmRecord.data.Time = timePostFromStart
            buffer(wmRecord_Index) = wmRecord.data
            SwitchToRecoding RECORDING_MODE
        End If
    End If
    
    lastStage = wmRecord.data.WeldStage
Else
    If beRecording = True Then  ' Finish record current poccess
        timePostFromStart = timeGetTime() - timeStart
        wmRecord.data.Time = timePostFromStart / 1000
        buffer(wmRecord_Index) = wmRecord.data
        wmRecord_Index = wmRecord_Index + 1
     
        timeStart = timeGetTime()
        
        analysisResult = PlcAnalysiser.Analysis(buffer, wmRecord_Index)
        
        If buffer(wmRecord_Index - 2).WeldStage >= 6 Then
            SaveData
        ElseIf GetSetting(App.EXEName, "Weld", "RecordInterrupts", 0) = 1 Then
            SaveData
        End If
     
        If analysisResult.Succeed = PlcDeclare.OK Then
            lblBigCenter.Caption = "OK"
            lblBigCenter.ForeColor = &HFF00&
        ElseIf analysisResult.Succeed = PlcDeclare.NO Then
            lblBigCenter.Caption = "NO"
            lblBigCenter.ForeColor = &HFF&
        Else
            lblBigCenter.Caption = "INT"
            lblBigCenter.ForeColor = &HFF&
        End If
    
        SwitchToRecoding ANALYSIS_MODE
    End If
    beRecording = False
    beSigned = False
    wmRecord_Index = 0
    lastStage = -1
    
    
End If


If ProcessSetting <= 0 Or Not beRecording Then
    Status = PLCDrv.ReadCurrentProcessSetting(ProcessSetting)
    If Status = 0 Then
        If ProcessSetting = 1 Then
            lblProcessSetting.Caption = "Regular"
            lblParameter.Caption = GetSetting(App.EXEName, "Parameter", "LastSetting_Regular", "----")
        ElseIf ProcessSetting = 2 Then
            lblProcessSetting.Caption = "Pulse"
            lblParameter.Caption = GetSetting(App.EXEName, "Parameter", "LastSetting_Pulse", "----")
        Else
            lblProcessSetting.Caption = "Unknown"
            lblParameter.Caption = "----"
        End If
    End If
End If

beRequest = False
Exit Sub
ERROR_HANDLE:
    beRequest = False
    MsgBox "Connection error!¡¡" & vbCrLf & Status
    Unload Me
End Sub

Public Function canStart() As Boolean

    Select Case Mode_StartRecording
        Case 0:
                canStart = True
        Case 1:
            If wmRecord.data.Dist >= ModeParam_StartRecoding(1) Then
                canStart = True
            End If
        Case 2:
            If wmRecord.data.Amp >= ModeParam_StartRecoding(2) Then
                canStart = True
            End If
        Case 3:
            If wmRecord.data.Volt >= ModeParam_StartRecoding(3) Then
                canStart = True
            End If
        Case 4:
            If timePostFromStart / 1000 >= ModeParam_StartRecoding(4) Then
                canStart = True
            End If
    End Select
            
End Function


Function SaveData()
Dim fh1 As FileHeader1
Dim fh2 As FileHeader2
Dim WeldFile As String
    WeldFile = toWeldNumberShowModel(weldSerailNumber)
    
    
    fh1.CompanyName = GetSetting(App.EXEName, "UserData", "CompanyName", "")
    fh1.UnitName = GetSetting(App.EXEName, "UserData", "Unit", "")
    fh1.Location = GetSetting(App.EXEName, "UserData", "Location", "")
        
        
    
    fh2.Date = Date
    fh2.Time = Time
    fh2.filename = WeldFile
    fh2.BaseRed = lblParameter.Caption
        
    If Not fso.FolderExists(path & "\" & Format(Date, "YYYY-MM-DD")) Then
        fso.CreateFolder (path & "\" & Format(Date, "YYYY-MM-DD"))
        fso.CreateTextFile (path & "\" & Format(Date, "YYYY-MM-DD") & "\" & Format(Date, "YYYY-MM-DD") & ".DLY")
    End If
    
    Call PlcWld.SaveData(path & "\" & Format(Date, "YYYY-MM-DD") & "\" & WeldFile & ".WLD", fh1, fh2, buffer, wmRecord_Index, analysisDefine, analysisResult)
    
    Dim dr As DailyReport
    
    dr.Serial = Left(WeldFile, 1)
    dr.Sequence = weldSerailNumber - CInt(weldSerailNumber / 10000) * 10000
    
    
    PlcDailyReport.SaveData path & "\" & Format(Date, "YYYY-MM-DD") & "\" & Format(Date, "YYYY-MM-DD") & ".DLY", dr

    Dim wn As Integer
    wn = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    
    If wn = weldSerailNumber Then
        weldSerailNumber = weldSerailNumber + 1
        Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
       
    End If
    

End Function

