VERSION 5.00
Begin VB.Form FrmRegularSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weld parameter for Regular Process"
   ClientHeight    =   7800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6615
   Icon            =   "FrmRegularSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "14000"
   Begin VB.ComboBox cboFileName 
      Height          =   300
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weld Parameters"
      Height          =   6135
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Tag             =   "14100"
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   12
         Left            =   2640
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   5640
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   11
         Left            =   2640
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   13
         Left            =   2640
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   10
         Left            =   2640
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   9
         Left            =   2640
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   8
         Left            =   2640
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   4
         Left            =   2640
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   6
         Left            =   2640
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   7
         Left            =   2640
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   300
         Index           =   2
         Left            =   3360
         TabIndex        =   85
         Top             =   3165
         Width           =   165
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   84
         Top             =   2805
         Width           =   165
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   300
         Index           =   0
         Left            =   3360
         TabIndex        =   83
         Top             =   2445
         Width           =   165
      End
      Begin VB.Label lblSign 
         Alignment       =   1  'Right Justify
         Caption         =   "230/430"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   82
         Top             =   3165
         WhatsThisHelpID =   6
         Width           =   735
      End
      Begin VB.Label lblSign 
         Alignment       =   1  'Right Justify
         Caption         =   "43/430"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   81
         Top             =   2805
         WhatsThisHelpID =   5
         Width           =   735
      End
      Begin VB.Label lblSign 
         Alignment       =   1  'Right Justify
         Caption         =   "344/430"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   79
         Top             =   2445
         WhatsThisHelpID =   4
         Width           =   735
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "1.8"
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   75
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   74
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.7"
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   73
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Preflash Distance (mm):"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   72
         Tag             =   "140"
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "0.30"
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   71
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   70
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.12"
         Height          =   255
         Index           =   11
         Left            =   4320
         TabIndex        =   69
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Boost Speed (mm/s):"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   68
         Tag             =   "130"
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "5.0"
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   67
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   66
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.0"
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   65
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Flash Speed (mm/s):"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   64
         Tag             =   "120"
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "20.0"
         Height          =   255
         Index           =   10
         Left            =   5160
         TabIndex        =   63
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   62
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "9.0"
         Height          =   255
         Index           =   10
         Left            =   4320
         TabIndex        =   61
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset (mm):"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   60
         Tag             =   "110"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "450"
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   59
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   58
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "150"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   57
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint III (A):"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   56
         Tag             =   "100"
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "350"
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   55
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   54
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "150"
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   53
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint II (A):"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   52
         Tag             =   "90"
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage Timer (s):"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Tag             =   "10"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Low Voltage Timer (s):"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   50
         Tag             =   "20"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current in Upset Timer(s):"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Tag             =   "30"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset Timer (s):"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Tag             =   "40"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage (%):"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   47
         Tag             =   "50"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Low Voltage (%):"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   46
         Tag             =   "60"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage Boost(%):"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   45
         Tag             =   "70"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint I (A):"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   44
         Tag             =   "80"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "25"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   43
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   42
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "60"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   41
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "60"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   39
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "200"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   38
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "0.2"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   37
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   36
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "1.0"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   35
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "1.0"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   34
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   33
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "3.00"
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   32
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   31
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   30
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   1  'Right Justify
         Caption         =   "100"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   29
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   28
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   27
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   1  'Right Justify
         Caption         =   "100"
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   26
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   25
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   24
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   1  'Right Justify
         Caption         =   "100"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   23
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Caption         =   "150"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   22
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   21
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         Caption         =   "300"
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   20
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   19
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "Min"
         Height          =   255
         Index           =   15
         Left            =   4320
         TabIndex        =   76
         Tag             =   "2"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "Max"
         Height          =   255
         Index           =   16
         Left            =   5160
         TabIndex        =   77
         Tag             =   "6"
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Tag             =   "14010"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Tag             =   "14030"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Tag             =   "14020"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblSign 
      Alignment       =   1  'Right Justify
      Caption         =   "344/430"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   80
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Program Names :"
      Height          =   255
      Left            =   480
      TabIndex        =   78
      Tag             =   "14001"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRegularSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fso As New FileSystemObject
Dim lastConfigName As String
Dim regularSetting As RegularSettingType
Dim path As String
Dim InitialVoltage As Long


Const WARNING_COLOR As Long = &H8080FF
Const ERROR_COLOR As Long = &HFF&
Const FINE_COLOR As Long = &HFFFFFF

Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cboFileName_Change()
    If lastConfigName <> cboFileName.Text Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub cboFileName_Click()
    If cboFileName.Text <> "" And cboFileName.Text <> lastConfigName Then
        lastConfigName = cboFileName.Text
        Call LoadConfig(lastConfigName)
    End If
End Sub

Private Sub cboStage_Change()

    Dim i As Integer
    For i = 0 To 14 - 1
        txtValue(i).Text = regularSetting.Value(i)
    Next
    
    i = 4
    lblSign(i).Caption = CStr(regularSetting.Value(i) * InitialVoltage / 100) & "/" & InitialVoltage
    i = 5
    lblSign(i).Caption = CStr(regularSetting.Value(i) * InitialVoltage / 100) & "/" & InitialVoltage
    i = 6
    lblSign(i).Caption = CStr(regularSetting.Value(i) * InitialVoltage / 100) & "/" & InitialVoltage
    
End Sub

Private Sub cboStage_Click()
    cboStage_Change
End Sub

Private Sub cmdLoad_Click()
    frmProgress.LoadMode = PlcDeclare.LOAD_REGULAR_SETTING
    frmProgress.ParamName = cboFileName.Text
    frmProgress.Show vbModal, Me
    If frmProgress.Status <> 0 Then
        GoTo ERROR_HANDLE
    End If
    
    cmdLoad.Enabled = False
    
Exit Sub
ERROR_HANDLE:
End Sub

Private Sub cmdSave_Click()
    If Not checkInputedDataValidate Then
        Exit Sub
    End If
    

    If cboFileName.Text <> "" Then
        Call PlcRegularSetting.SaveConfig(path, cboFileName.Text, regularSetting)
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
    
    cmdSave.Enabled = False
    cmdLoad.Enabled = True
End Sub

Private Function LoadConfig(name As String)
    regularSetting = PlcRegularSetting.LoadConfig(path, name)
        
    cboStage_Change
    
End Function


Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me


Dim pFileItemList() As PulseFileItemType

    InitialVoltage = CSng(GetSetting(App.EXEName, "AnalysisDefine", "InitialVoltage", 430))
    
    path = App.path & "\" & SETTING_PATH & "RegularSetting.config"
    
    If Not fso.FileExists(path) Then
        regularSetting = PlcRegularSetting.DefalutStagesParameters
        PlcRegularSetting.SaveConfig path, "DEFAULT", regularSetting
    End If

    pFileItemList = PlcRegularSetting.LoadAll(path)
        
    Dim i As Integer
    For i = 1 To cboFileName.ListCount
        cboFileName.RemoveItem (cboFileName.ListCount - 1)
    Next
        
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        cboFileName.AddItem (Trim(pFileItemList(i).name))
    Next
    
    lastConfigName = GetSetting(App.EXEName, "Parameter", "LastSetting_Regular", "DEFAULT")
    Call LoadConfig(lastConfigName)
    
    For i = 0 To cboFileName.ListCount - 1
        If cboFileName.List(i) = lastConfigName Then
            cboFileName.ListIndex = i
            cboFileName_Click
            Exit For
        End If
    Next
    
    cmdSave.Enabled = False
    cmdLoad.Enabled = False
        
End Sub

Private Sub txtValue_Change(index As Integer)
    Dim min As Single
    Dim max As Single
    Dim v As Single
    
    min = CSng(lblMin(index).Caption)
    max = CSng(lblMax(index).Caption)
    If IsNumeric(txtValue(index).Text) Then
    
        v = CSng(txtValue(index).Text)
        If 4 <= index And index <= 6 Then
            Dim i As Integer
            i = index
            lblSign(i).Caption = CStr(v * InitialVoltage / 100) & "/" & InitialVoltage
        End If
        
        If min <= v And v <= max Then
            txtValue(index).BackColor = FINE_COLOR
        Else
            txtValue(index).BackColor = WARNING_COLOR
        End If
        
        regularSetting.Value(index) = CSng(txtValue(index).Text)
        cmdSave.Enabled = True
    Else
        txtValue(index).BackColor = ERROR_COLOR
    End If
            
End Sub
Private Function checkInputedDataValidate() As Boolean
    Dim min As Single
    Dim max As Single
    Dim v As Single
    
    
Dim i As Integer
    For i = 0 To txtValue.count - 1
        min = CSng(lblMin(i).Caption)
        max = CSng(lblMax(i).Caption)
        If IsNumeric(txtValue(i).Text) Then
            v = CSng(txtValue(i).Text)
'            If Not (min <= v And v <= max) Then
'                checkInputedDataValidate = False
'                Exit Function
'            End If
        Else
            checkInputedDataValidate = False
            Exit Function
        End If
    Next i
    checkInputedDataValidate = True
End Function

Private Sub txtValue_GotFocus(index As Integer)
    txtValue(index).SelLength = Len(txtValue(index).Text)
End Sub
