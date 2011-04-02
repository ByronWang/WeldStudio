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
   Begin VB.PictureBox picVolt 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      ScaleHeight     =   720
      ScaleWidth      =   8325
      TabIndex        =   18
      Top             =   1905
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
      Top             =   465
      Width           =   8355
   End
   Begin VB.Timer TimerShow 
      Interval        =   90
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
      Top             =   7560
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
      Top             =   9000
      Width           =   8355
   End
   Begin VB.Timer TimerTest 
      Interval        =   92
      Left            =   3480
      Top             =   5640
   End
   Begin VB.Label lblDist 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "67.98"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1600
      Left            =   10920
      TabIndex        =   16
      Top             =   120
      Width           =   3900
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "A1234"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   5100
      TabIndex        =   15
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label lblParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   4920
      TabIndex        =   14
      Top             =   6360
      Width           =   5535
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11760
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblBigCenter 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "A0001"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2655
      Left            =   3353
      TabIndex        =   11
      Top             =   3600
      Width           =   8535
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
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   15240
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label lblVolt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "345"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1600
      Left            =   10920
      TabIndex        =   8
      Top             =   1560
      Width           =   3900
   End
   Begin VB.Label lblPsi 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1600
      Left            =   10920
      TabIndex        =   7
      Top             =   7200
      Width           =   3900
   End
   Begin VB.Label lblAmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1600
      Left            =   10920
      TabIndex        =   6
      Top             =   8640
      Width           =   3900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Force"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   5
      Tag             =   "12040"
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amp"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   4
      Tag             =   "12030"
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Volt"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   3
      Tag             =   "12020"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Tag             =   "12010"
      Top             =   480
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

Dim lastTimeInMS As Long
Dim timeInMS As Long

Dim dist_scale As Long
Dim volt_scale As Long
Dim amp_scale As Long
Dim psi_scale As Long

Dim rBuf(30000) As WeldData
Dim wm As WeldMonitor
Dim rIndex As Integer

Dim beRecording As Boolean
Dim beSigned As Boolean

Dim StartRecording As Integer
Dim StartRecodingParam(5) As Single

Dim weldSerailNumber As Long
'Dim WeldFile As String

Const ANALYSIS_DURATION As Integer = 6000

Dim showMode As ShowModeType

Enum ShowModeType
    RECORDING_MODE
    STANDBY_MODE
    ANALYSIS_MODE
End Enum

Dim analysisDefine As WeldAnalysisDefineType
Dim analysisResult As WeldAnalysisResultType

Function SwitchToRecoding(status As ShowModeType)
    Select Case status
        Case STANDBY_MODE
            TimerTest.Enabled = True
            TimerShow.Tag = ""
            
            lblTop.Caption = "Ready"
            weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
            lblBigCenter.Caption = toWeldNumberShowModel(weldSerailNumber)
            lblBigCenter.ForeColor = &H8000000E
            lblBigCenter.FontSize = 100
            lblParameter.Caption = GetSetting(App.EXEName, "Parameter", "LastSetting", "DEFAULT")
        Case RECORDING_MODE
            weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
            lblTop.Caption = toWeldNumberShowModel(weldSerailNumber)
            lblBigCenter.FontSize = 120
        Case ANALYSIS_MODE
            TimerTest.Enabled = False
            TimerShow.Tag = "ANALYSIS"
    End Select
    showMode = status
End Function

Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me


    PLCDrv.InitPLCConnection

    amp_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Amp", 500))
    dist_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Dist", 1000))
    volt_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Volt", 100))
    psi_scale = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Press", 50))


    StartRecording = CInt(GetSetting(App.EXEName, "StartRecording", "StartRecording", 0))
    StartRecodingParam(1) = CSng(GetSetting(App.EXEName, "StartRecording", "Dist", 2.5))
    StartRecodingParam(2) = CSng(GetSetting(App.EXEName, "StartRecording", "Amp", 100))
    StartRecodingParam(3) = CSng(GetSetting(App.EXEName, "StartRecording", "Volt", 450))
    StartRecodingParam(4) = CSng(GetSetting(App.EXEName, "StartRecording", "Time", 25))

    
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


    Set fso = New FileSystemObject
    '
    
    path = App.path & "\Data"
    If Not fso.FolderExists(path) Then
        fso.CreateFolder (path)
    End If
    
    beRecording = False
    lastTimeInMS = 0
    timeInMS = 0
    rIndex = 0

    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    
    SwitchToRecoding STANDBY_MODE
    
    beSigned = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PLCDrv.UninitPLCConection
    Unload Me
End Sub

Private Sub TimerShow_Timer()
    
lblDate.Caption = Format(Date, "YYYY-MM-DD")
lblTime.Caption = FormatDateTime(Now, vbLongTime)

If Me.TimerShow.Tag = "ANALYSIS" Then
   If timeGetTime() - lastTimeInMS > ANALYSIS_DURATION Then
        SwitchToRecoding (STANDBY_MODE)
   End If
   Exit Sub
End If

If beRecording Then
    lblBigCenter.Caption = Format(timeInMS / 1000, "000")
    'lblBigCenter.Caption = Format(timeInMS / 1000, "##0.00")
End If

If 1 <= wm.data.PlcStage And wm.data.PlcStage <= 12 Then
    lblPlcStage.Caption = PLCDrv.PlcStages(wm.data.PlcStage)
End If


If 1 <= wm.data.WeldStage And wm.data.WeldStage <= 6 Then
    lblWeldStage.Caption = PLCDrv.WeldStages(wm.data.WeldStage)
End If


Dim Dist As Long
Dim Volt As Long
Dim Amp As Long
Dim psi As Long

Dist = wm.data.Dist
If Dist < 0 Then
    Dist = 0
End If

Volt = wm.data.Volt
If Volt < 0 Then
Volt = 0
End If

Amp = wm.data.Amp
If Amp < 0 Then
    Amp = 0
End If

psi = PlcAnalysiser.toForce(wm.data.PsiUpset, wm.data.PsiOpen, analysisDefine)
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


'lblTime.Caption = data.Time
If PLCDrv.Calibrate_Distance Then
    lblDist.FontSize = lblVolt.FontSize
    lblDist.Caption = Format(wm.data.Dist, "##0.0")
Else
    lblDist.FontSize = lblVolt.FontSize - 3
    lblDist.Caption = Format(wm.data.Dist, "##0")
End If
lblVolt.Caption = wm.data.Volt
lblAmp.Caption = wm.data.Amp
lblPsi.Caption = psi

End Sub

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


Private Sub TimerTest_Timer()


wm = PLCDrv.readPcMonitor


If wm.WeldCycle = 1 And 0 < wm.data.WeldStage And wm.data.WeldStage <= 6 Then
    If Not beSigned Then
        lastTimeInMS = timeGetTime()
        beSigned = True
        timeInMS = 0
    Else
        timeInMS = timeGetTime() - lastTimeInMS
    End If
        
    If Not beRecording Then
        If canStart() Then
            lastTimeInMS = timeGetTime()
            timeInMS = timeGetTime() - lastTimeInMS
            beRecording = True
            rIndex = 0
            wm.data.Time = timeInMS
            rBuf(rIndex) = wm.data
            SwitchToRecoding RECORDING_MODE
        End If
    Else
        wm.data.Time = timeInMS / 1000
        rBuf(rIndex) = wm.data
        rIndex = rIndex + 1
    End If
Else
    If beRecording = True Then
        timeInMS = timeGetTime() - lastTimeInMS
        wm.data.Time = timeInMS / 1000
        rBuf(rIndex) = wm.data
        rIndex = rIndex + 1
     
        lastTimeInMS = timeGetTime()
        
        analysisResult = PlcAnalysiser.Analysis(rBuf, rIndex)
        
        If rBuf(rIndex - 2).WeldStage >= 6 Then
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
            lblBigCenter.ForeColor = &HFFFFFF
        End If
    
        SwitchToRecoding ANALYSIS_MODE
    End If
    beRecording = False
    beSigned = False
    rIndex = 0
End If

End Sub

Public Function canStart() As Boolean

    Select Case StartRecording
        Case 0:
                canStart = True
        Case 1:
            If wm.data.Dist >= StartRecodingParam(1) Then
                canStart = True
            End If
        Case 2:
            If wm.data.Amp >= StartRecodingParam(2) Then
                canStart = True
            End If
        Case 3:
            If wm.data.Volt >= StartRecodingParam(3) Then
                canStart = True
            End If
        Case 4:
            If timeInMS / 1000 >= StartRecodingParam(4) Then
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
    
    Call PlcWld.SaveData(path & "\" & Format(Date, "YYYY-MM-DD") & "\" & WeldFile & ".WLD", fh1, fh2, rBuf, rIndex, analysisDefine, analysisResult)
    
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

