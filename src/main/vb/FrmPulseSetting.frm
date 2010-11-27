VERSION 5.00
Begin VB.Form FrmPulseSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weld parameter for Pulse Process"
   ClientHeight    =   7680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   Icon            =   "FrmPulseSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "13000"
   Begin VB.Frame Frame1 
      Caption         =   "General Parameters:"
      Height          =   1935
      Index           =   1
      Left            =   360
      TabIndex        =   45
      Tag             =   "13200"
      Top             =   5160
      Width           =   5175
      Begin VB.TextBox txtValueGeneral 
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtValueGeneral 
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "30"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   65
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   64
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   63
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "75"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   62
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   61
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "20"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   60
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   59
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.00"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   57
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblMaxGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "20.0"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   56
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   55
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMinGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "9"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Forging Force(t):"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   53
         Tag             =   "40"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Tension Holding Time(m): "
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   52
         Tag             =   "30"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Upset(mm):"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   51
         Tag             =   "20"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current in Upset(s):"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   50
         Tag             =   "10"
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Tag             =   "13020"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox cboFileName 
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Tag             =   "13030"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Tag             =   "13010"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4095
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Tag             =   "13100"
      Top             =   960
      Width           =   5175
      Begin VB.ComboBox cboStage 
         Height          =   300
         ItemData        =   "FrmPulseSetting.frx":000C
         Left            =   960
         List            =   "FrmPulseSetting.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   435
         Width           =   2055
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   7
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   6
         Left            =   2640
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   5
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   4
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2400
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
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   71
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage:"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Tag             =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   44
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   43
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   42
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "3.0"
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   41
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   40
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.10"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   39
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   38
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   37
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   36
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   35
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   34
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   33
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "800"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   32
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   31
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "150"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   30
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "460"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   28
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "250"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "60"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   26
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   25
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "10.0"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblSepereter 
         Caption         =   "/"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   22
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "0.1"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Caption         =   "Reverse speed(mm/s):"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   20
         Tag             =   "80"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Forward speed(mm/s):"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Tag             =   "70"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 3(A):"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Tag             =   "60"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 2(A):"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Tag             =   "50"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Current Setpoint 1(A):"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Tag             =   "40"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Time(s):"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Tag             =   "30"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Voltage(V):"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Tag             =   "20"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Distance(mm):"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "10"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Max"
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   69
         Tag             =   "6"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Min"
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   68
         Tag             =   "2"
         Top             =   480
         Width           =   495
      End
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
Dim PulseSetting As PulseSettingType
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
    For i = 0 To 7
        txtValue(i).Text = PulseSetting.Stages(cboStage.ListIndex).Value(i)
    Next
    
End Sub

Private Sub cboStage_Click()
    cboStage_Change
End Sub

Private Sub cmdLoad_Click()

    Call PLCDrv.InitPLCConnection
    Call PLCDrv.WritePulseData(PulseSetting)
    Call PLCDrv.UninitPLCConection
    
    Call SaveSetting(App.EXEName, "Parameter", "LastSetting", "Pulse:" & lastConfigName)
End Sub

Private Sub cmdSave_Click()
    If cboFileName.Text <> "" Then
        Call PlcPulseSetting.SaveConfig(path, cboFileName.Text, PulseSetting)
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
    PulseSetting = PlcPulseSetting.LoadConfig(path, name)
    
    txtValueGeneral(0).Text = PulseSetting.General.Value(0)
    txtValueGeneral(1).Text = PulseSetting.General.Value(1)
    txtValueGeneral(2).Text = PulseSetting.General.Value(2)
    txtValueGeneral(3).Text = PulseSetting.General.Value(3)
    
    cboStage.ListIndex = 0
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
    PulseSetting = PlcPulseSetting.DefalutStagesParameters
    
    path = App.path & "\" & SETTING_PATH & "PulseSetting.config"

    pFileItemList = PlcPulseSetting.LoadAll(path)
        
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

'    Debug.Print Frame1(1).Caption
'
'    Debug.Print txtValue(1).hWnd
End Sub

Private Sub txtValue_Change(index As Integer)
    If IsNumeric(txtValue(index).Text) Then
        PulseSetting.Stages(cboStage.ListIndex).Value(index) = CSng(txtValue(index).Text)
    Else
        txtValue(index).Text = PulseSetting.Stages(cboStage.ListIndex).Value(index)
    End If
End Sub

Private Sub txtValueGeneral_Change(index As Integer)
    If IsNumeric(txtValueGeneral(index).Text) Then
        PulseSetting.General.Value(index) = CSng(txtValueGeneral(index).Text)
    Else
        txtValueGeneral(index).Text = PulseSetting.General.Value(index)
    End If
End Sub
