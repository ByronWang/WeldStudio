VERSION 5.00
Begin VB.Form FrmRegularSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weld parameter for Regular Process"
   ClientHeight    =   7920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   Icon            =   "FrmRegularSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "14000"
   Begin VB.ComboBox cboFileName 
      Height          =   300
      Left            =   2160
      TabIndex        =   77
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weld Parameters"
      Height          =   6375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Tag             =   "14100"
      Top             =   960
      Width           =   5175
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   13
         Left            =   2640
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   5640
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   12
         Left            =   2640
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   11
         Left            =   2640
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   10
         Left            =   2640
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   9
         Left            =   2640
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   8
         Left            =   2640
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   4
         Left            =   2640
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   6
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   7
         Left            =   2640
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   14
         Left            =   4440
         TabIndex        =   74
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   73
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   14
         Left            =   3600
         TabIndex        =   72
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Pre-Flash Distance (mm):"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   71
         Tag             =   "140"
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   13
         Left            =   4440
         TabIndex        =   69
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   68
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   13
         Left            =   3600
         TabIndex        =   67
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Boost Speed (mm/s):"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   66
         Tag             =   "130"
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   64
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   63
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   11
         Left            =   3600
         TabIndex        =   62
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Flash Speed (mm/s):"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   61
         Tag             =   "120"
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   59
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   58
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   10
         Left            =   3600
         TabIndex        =   57
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset (mm):"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   56
         Tag             =   "110"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   54
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   53
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   9
         Left            =   3600
         TabIndex        =   52
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint III (A):"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   51
         Tag             =   "100"
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   49
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   48
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   47
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint II (A):"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   46
         Tag             =   "90"
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage Timer (s):"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Tag             =   "10"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Low Voltage Timer (s):"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Tag             =   "20"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current in Upset Timer(s):"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   42
         Tag             =   "30"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset Timer (s):"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Tag             =   "40"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage (V):"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   40
         Tag             =   "50"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Low Voltage (V):"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   39
         Tag             =   "60"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "High Voltage Boost (V):"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   38
         Tag             =   "70"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint I (A):"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   37
         Tag             =   "80"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.1"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   35
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "10.0"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   32
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "60"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   31
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "250"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   30
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   29
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "460"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   27
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   26
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   25
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   24
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   23
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   22
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   21
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   20
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   19
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   18
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   17
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   16
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   15
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   14
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   13
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Min"
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   75
         Tag             =   "2"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Max"
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   76
         Tag             =   "6"
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Tag             =   "14010"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Tag             =   "14030"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Tag             =   "14020"
      Top             =   7440
      Width           =   1215
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
Dim RegularSetting As RegularSettingType
Dim path As String

Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
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
        txtValue(i).Text = RegularSetting.Value(i)
    Next
    
End Sub

Private Sub cboStage_Click()
    cboStage_Change
End Sub

Private Sub cmdLoad_Click()
    Call PLCDrv.InitPLCConnection
    Call PLCDrv.WriteRegularData(RegularSetting)
    Call PLCDrv.UninitPLCConection
    
    Call SaveSetting(App.EXEName, "Parameter", "LastSetting", "Regular:" & cboFileName.Text)
End Sub

Private Sub cmdSave_Click()
    If cboFileName.Text <> "" Then
        Call PlcRegularSetting.SaveConfig(path, cboFileName.Text, RegularSetting)
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
    
End Sub

Private Function LoadConfig(name As String)
    RegularSetting = PlcRegularSetting.LoadConfig(path, name)
        
    cboStage_Change
    
End Function


Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me

    PLCDrv.InitPLCConnection
    cmdLoad.Enabled = PLCDrv.beActive
    PLCDrv.UninitPLCConection

Dim pFileItemList() As PulseFileItemType

    lastConfigName = ""
    RegularSetting = PlcRegularSetting.DefalutStagesParameters
    
    path = App.path & "\" & SETTING_PATH & "RegularSetting.config"

    pFileItemList = PlcRegularSetting.LoadAll(path)
        
    Dim i As Integer
    For i = 1 To cboFileName.ListCount
        cboFileName.RemoveItem (cboFileName.ListCount - 1)
    Next
    
    For i = LBound(pFileItemList) To UBound(pFileItemList) - 1
        cboFileName.AddItem (Trim(pFileItemList(i).name))
    Next
    
    If UBound(pFileItemList) > 0 Then
        cboFileName.ListIndex = 0
        cboFileName_Click
    Else
        Call LoadConfig("NONE")
    End If
        
End Sub

Private Sub txtValue_Change(index As Integer)
    If IsNumeric(txtValue(index).Text) Then
        RegularSetting.Value(index) = CSng(txtValue(index).Text)
    Else
        txtValue(index).Text = RegularSetting.Value(index)
    End If
End Sub
