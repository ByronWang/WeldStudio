VERSION 5.00
Begin VB.Form FrmPulseSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weld parameter for Pulse Process"
   ClientHeight    =   7800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6615
   Icon            =   "FrmPulseSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "13000"
   Begin VB.CommandButton cmdDelete 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   74
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Parameters:"
      Height          =   1935
      Index           =   1
      Left            =   360
      TabIndex        =   50
      Tag             =   "13200"
      Top             =   5160
      Width           =   5895
      Begin VB.TextBox txtValueGeneral 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   66
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   65
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   64
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         Caption         =   "75"
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   63
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   62
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         Caption         =   "20"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   61
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   60
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   59
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         Caption         =   "0.2"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   58
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         Caption         =   "20.0"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   57
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   56
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Forging Force(t):"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   54
         Tag             =   "40"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Tension Holding Time(m): "
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   53
         Tag             =   "30"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset(mm):"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   52
         Tag             =   "20"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current in Upset(s):"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   51
         Tag             =   "10"
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Tag             =   "13020"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.ComboBox cboFileName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Tag             =   "13030"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Tag             =   "13010"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4095
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Tag             =   "13100"
      Top             =   960
      Width           =   5895
      Begin VB.ComboBox cboStage 
         Height          =   300
         ItemData        =   "FrmPulseSetting.frx":000C
         Left            =   1320
         List            =   "FrmPulseSetting.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   7
         Left            =   2640
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   6
         Left            =   2640
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   5
         Left            =   2640
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   4
         Left            =   2640
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   200
         Index           =   3
         Left            =   3375
         TabIndex        =   72
         Top             =   1720
         Width           =   165
      End
      Begin VB.Label lblSign 
         Alignment       =   1  'Right Justify
         Caption         =   "344/430"
         Height          =   225
         Index           =   1
         Left            =   3480
         TabIndex        =   73
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   71
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage:"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Tag             =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "3.0"
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   49
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   48
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.10"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   47
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "3.0"
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   46
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   45
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.10"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   44
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "800"
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   43
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   42
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "50"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   41
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "800"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   40
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   39
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "50"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   38
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "800"
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   37
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   36
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "50"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   35
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "60"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   34
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   32
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "100"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   31
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   30
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "50"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "10.0"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   27
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.1"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   26
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Reverse speed(mm/s):"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Tag             =   "80"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Forward speed(mm/s):"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Tag             =   "70"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 3(A):"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Tag             =   "60"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 2(A):"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Tag             =   "50"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 1(A):"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Tag             =   "40"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Time(s):"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Tag             =   "30"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Voltage(V):"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Tag             =   "20"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Distance(mm):"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Tag             =   "10"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "Max"
         Height          =   255
         Index           =   16
         Left            =   5160
         TabIndex        =   69
         Tag             =   "6"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "Min"
         Height          =   255
         Index           =   15
         Left            =   4320
         TabIndex        =   68
         Tag             =   "2"
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Program name must be 1 to 7 characters long."
      Height          =   255
      Left            =   1800
      TabIndex        =   75
      Top             =   760
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Program Names :"
      Height          =   255
      Left            =   480
      TabIndex        =   70
      Tag             =   "13001"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmPulseSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fso As New FileSystemObject
Dim lastConfigName As String
Dim plcSetting As PulseSettingType
Dim InitialVoltage As Long
Dim lastCboStageIndex As Integer


Const WARNING_COLOR As Long = &H8080FF
Const ERROR_COLOR As Long = &HFF&
Const FINE_COLOR As Long = &HFFFFFF

Const SETTING_FILE_NAME As String = "PulseSetting.cfg"
Const REG_NAME As String = "LastSetting_Pulse"

Dim loading As Boolean

Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cboFileName_Change()
    If lastConfigName <> cboFileName.Text Then
        cmdSave.Enabled = True
        cmdDelete.Enabled = False
        cmdLoad.Enabled = False
    End If
End Sub

Private Sub cboFileName_Click()
    If cboFileName.Text <> "" And cboFileName.Text <> lastConfigName Then
        lastConfigName = cboFileName.Text
        
        Call LoadConfig(lastConfigName)
    End If
End Sub

Private Sub cboStage_Change()
    Call ResetStageValue
End Sub

Private Sub cboStage_Click()

'    If checkInputedDataValidate Then
        Call ResetStageValue
        lastCboStageIndex = cboStage.ListIndex
'    Else
'        cboStage.ListIndex = lastCboStageIndex
'    End If

End Sub

Private Sub cmdDelete_Click()
    If Me.cboFileName.ListCount <= 1 Then
        MsgBox "At lest one setting!"
        Exit Sub
    End If
    
    If cmdSave.Enabled Then
        MsgBox "Please save data first!"
        Exit Sub
    End If
    
    Dim loadedName As String
    loadedName = GetSetting(App.EXEName, "Parameter", REG_NAME, "DEFAULT")
    If loadedName = Me.cboFileName.Text Then
        MsgBox "Can't delete loaded setting ! " & vbCrLf & "Please load another config first!"
        Exit Sub
    End If
    
    Dim index As Integer
    index = Me.cboFileName.ListIndex
    
    Call PlcPulseSetting.DeleteConfig(index)
    
    Form_Load
        
    If index < Me.cboFileName.ListCount Then
        Me.cboFileName.ListIndex = index
    End If
End Sub

Private Sub cmdLoad_Click()
    frmProgress.LoadMode = PlcDeclare.LOAD_PULSE_SETTING
    
    frmProgress.ParamName = cboFileName.Text
    frmProgress.Show vbModal, Me
    If frmProgress.status <> 0 Then
        GoTo ERROR_HANDLE
    End If
    
    cmdLoad.Enabled = False
    
Exit Sub
ERROR_HANDLE:
End Sub

Private Function checkInputedDataValidate() As Boolean
    Dim i As Integer
    
    checkInputedDataValidate = True
    
    For i = 0 To txtValue.count - 1
        If Not IsNumeric(txtValue(i).Text) Then
            checkInputedDataValidate = False
            Exit Function
        End If
    Next i

    For i = 0 To txtValueGeneral.count - 1
        If Not IsNumeric(txtValueGeneral(i).Text) Then
            checkInputedDataValidate = False
            Exit Function
        End If
    Next i

End Function

Private Sub cmdSave_Click()
    If Not checkInputedDataValidate Then
        Exit Sub
    End If
    
    If cboFileName.Text <> "" Then
        Call PlcPulseSetting.SaveConfig(cboFileName.Text, plcSetting)
    End If
    Dim i As Integer
    
    Dim existed As Boolean
    existed = False
    
    For i = 0 To cboFileName.ListCount - 1
        If cboFileName.List(i) = cboFileName.Text Then
            existed = True
            Exit For
        End If
    Next i
    
    If Not existed Then
        cboFileName.AddItem (cboFileName.Text)
    End If
    
    cmdLoad.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = False
    
    lastConfigName = cboFileName.Text
End Sub

Private Function LoadConfig(name As String)
loading = True
    plcSetting = PlcPulseSetting.LoadConfig(name)
    
    Call ResetValue
    
    cmdLoad.Enabled = True
loading = False
End Function

Private Sub ResetValue()
    txtValueGeneral(0).Text = plcSetting.General.Value(0)
    txtValueGeneral(1).Text = plcSetting.General.Value(1)
    txtValueGeneral(2).Text = plcSetting.General.Value(2)
    txtValueGeneral(3).Text = plcSetting.General.Value(3)
    
    cboStage.ListIndex = 0
    Call ResetStageValue
End Sub

Private Sub ResetStageValue()
    Dim i As Integer
    For i = 0 To 7
        txtValue(i).Text = plcSetting.Stages(cboStage.ListIndex).Value(i)
    Next
    i = 1
    lblSign(i).Caption = CInt(plcSetting.Stages(cboStage.ListIndex).Value(i) * InitialVoltage / 100) & "/" & InitialVoltage
End Sub

Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me

    Dim pFileItemList() As PulseFileItemType

    InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 430))
    
    Dim path As String
    path = App.path & "\" & SETTING_PATH & SETTING_FILE_NAME
    If Not fso.FileExists(path) Then
        plcSetting = PlcPulseSetting.DefalutStagesParameters
        PlcPulseSetting.SaveConfig "DEFAULT", plcSetting
    End If

    pFileItemList = PlcPulseSetting.LoadAll()
        
    Dim i As Integer
    For i = 1 To cboFileName.ListCount
        cboFileName.RemoveItem (cboFileName.ListCount - 1)
    Next
    
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        ' out.log Me.name & ".Form_Load > name." & i & " = " & Trim(pFileItemList(i).name)
        cboFileName.AddItem (Trim(pFileItemList(i).name))
    Next
    
    Dim configName As String
    configName = GetSetting(App.EXEName, "Parameter", REG_NAME, "DEFAULT")

    For i = 0 To cboFileName.ListCount - 1
        If cboFileName.List(i) = configName Then
            cboFileName.ListIndex = i
            Exit For
        End If
    Next
    If cboFileName.ListIndex < 0 Then
        cboFileName.ListIndex = 0
    End If
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = True
    cmdLoad.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtValue_Change(index As Integer)
    Dim min As Single
    Dim max As Single
    Dim v As Single
    
    min = CSng(lblMin(index).Caption)
    max = CSng(lblMax(index).Caption)
    
    If IsNumeric(txtValue(index).Text) Then
        v = CSng(txtValue(index).Text)
        
        If index = 1 Then
            Dim i As Integer
            i = index
            lblSign(i).Caption = CStr(CInt(v * InitialVoltage / 100)) & "/" & InitialVoltage
        End If
        
        If min <= v And v <= max Then
            txtValue(index).BackColor = FINE_COLOR
        Else
            txtValue(index).BackColor = WARNING_COLOR
        End If
        
        plcSetting.Stages(cboStage.ListIndex).Value(index) = CDbl(txtValue(index).Text)
        
        If Not loading Then
            cmdSave.Enabled = True
            cmdDelete.Enabled = False
            cmdLoad.Enabled = False
        End If
    Else
        txtValue(index).BackColor = ERROR_COLOR
    End If
End Sub

Private Sub txtValueGeneral_Change(index As Integer)
    Dim min As Single
    Dim max As Single
    Dim v As Single
    
    min = CSng(lblMinGeneral(index).Caption)
    max = CSng(lblMaxGeneral(index).Caption)
    If IsNumeric(txtValueGeneral(index).Text) Then
    
        v = CSng(txtValueGeneral(index).Text)
        If min <= v And v <= max Then
            txtValueGeneral(index).BackColor = FINE_COLOR
        Else
            txtValueGeneral(index).BackColor = WARNING_COLOR
        End If
        
        plcSetting.General.Value(index) = CSng(txtValueGeneral(index).Text)
        
        If Not loading Then
            cmdSave.Enabled = True
            cmdLoad.Enabled = False
            cmdDelete.Enabled = False
        End If
    Else
        txtValueGeneral(index).BackColor = ERROR_COLOR
    End If
End Sub

Private Sub txtValue_GotFocus(index As Integer)
    txtValue(index).SelLength = Len(txtValue(index).Text)
End Sub

Private Sub txtValueGeneral_GotFocus(index As Integer)
    txtValue(index).SelLength = Len(txtValue(index).Text)
End Sub
