VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading"
   ClientHeight    =   1680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer timerProgress 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1920
      Top             =   120
   End
   Begin VB.Frame frmProgress 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Label lblProgress 
         BackColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const seperator As Integer = 400
Public ParamName As String
Public LoadMode As Integer
Public Status As Long
Dim step As Integer
Dim beRuning As Boolean

Private Sub Form_Load()
    step = (frmProgress.Width - seperator) / 50
    lblProgress.Width = 0
    
    Me.timerProgress.Enabled = True
    DoEvents
    beRuning = False
    
End Sub

Private Sub run()
    
    Select Case LoadMode
        Case LOAD_ALL_PARAMETER:
            Status = LoadAllSetting
        Case LOAD_PULSE_SETTING:
            Status = LoadPulseSetting(ParamName)
        Case LOAD_REGULAR_SETTING:
            Status = LoadRegularSetting(ParamName)
        Case LOAD_CALIBRATION_SETTING:
    End Select


End Sub

Private Sub Finish()
    Me.Hide
    If Status = 0 Then
        If LoadMode = LOAD_ALL_PARAMETER Then
        Else
            MsgBox "Succeed!", vbOKOnly
        End If
    ElseIf Status = 1000 Then
    
    Else
        MsgBox "Connection Error"
    End If
    Unload Me
End Sub

Private Sub timerProgress_Timer()
    If Status <> 0 Then
        lblProgress.BackColor = &HFF
    End If
    If lblProgress.Width < frmProgress.Width - seperator - seperator / 2 Then
        lblProgress.Width = lblProgress.Width + step
    Else
        Call Finish
    End If
    
    If Not beRuning Then
        run
        beRuning = True
    End If
End Sub


Private Function LoadAllSetting()
    Status = PLCDrv.OpenPLCConnection
    If Status <> 0 Then
        LoadAllSetting = Status
        Exit Function
    End If
    
    Status = PLCDrv.PreparePcMonitor
    If Status <> 0 Then
        LoadAllSetting = Status
        Exit Function
    End If
    Dim wm As WeldMonitor
    Status = PLCDrv.ReadPcMonitor(wm)
    If Status <> 0 Then
        LoadAllSetting = Status
        Exit Function
    End If
    
End Function

Private Function LoadRegularSetting(name As String) As Long
    If name = "" Then
        name = GetSetting(App.EXEName, "Parameter", "LastSetting_Regular", "")
    End If
    
    If name = "" Then
        MsgBox "Please config regular setting!"
        LoadRegularSetting = 1000
        Exit Function
    End If
        
    Dim regularSetting As RegularSettingType
    Dim path As String
    path = App.path & "\" & SETTING_PATH & "PulseSetting.config"
    
    regularSetting = PlcRegularSetting.LoadConfig(path, name)
    
    Status = PLCDrv.OpenPLCConnection
    If Status <> 0 Then
        LoadRegularSetting = Status
        Exit Function
    End If
    DoEvents
    Status = PLCDrv.WriteRegularData(regularSetting)
    If Status <> 0 Then
        LoadRegularSetting = Status
        Exit Function
    End If
    DoEvents
    Status = PLCDrv.ClosePLCConection
    
    Call SaveSetting(App.EXEName, "Parameter", "LastSetting_Regular", name)
End Function

Private Function LoadPulseSetting(name As String) As Long
    If name = "" Then
        name = GetSetting(App.EXEName, "Parameter", "LastSetting_Pulse", "")
    End If
    
    If name = "" Then
        MsgBox "Please config pulse setting!"
        LoadPulseSetting = 1000
    End If

    Dim pulseSetting As PulseSettingType
    Dim path As String
    path = App.path & "\" & SETTING_PATH & "PulseSetting.config"
    
    pulseSetting = PlcPulseSetting.LoadConfig(path, name)
    
    Status = PLCDrv.OpenPLCConnection
    If Status <> 0 Then
        LoadPulseSetting = Status
        Exit Function
    End If
    DoEvents
    Status = PLCDrv.WritePulseData(pulseSetting)
    If Status <> 0 Then
        LoadPulseSetting = Status
        Exit Function
    End If
    DoEvents
    Call PLCDrv.ClosePLCConection
    
    Call SaveSetting(App.EXEName, "Parameter", "LastSetting_Pulse", name)
    
End Function

