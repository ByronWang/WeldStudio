VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   7440
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   10920
   Icon            =   "FrmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "16000"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   86
      Tag             =   "16010"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Tag             =   "16020"
      Top             =   6960
      Width           =   1095
   End
   Begin TabDlg.SSTab tabs 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   5
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmOption.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "CommonDialog1"
      Tab(0).Control(3)=   "cboLanguage"
      Tab(0).Control(4)=   "chkOnlineOnStartUp"
      Tab(0).Control(5)=   "lblLanguage"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Simulate"
      TabPicture(1)   =   "FrmOption.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameP(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Sersor Calibration"
      TabPicture(2)   =   "FrmOption.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).Control(1)=   "Frame2(0)"
      Tab(2).Control(2)=   "Frame2(3)"
      Tab(2).Control(3)=   "Frame2(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Sensor Reading Bar"
      TabPicture(3)   =   "FrmOption.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Weld Chart"
      TabPicture(4)   =   "FrmOption.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(1)"
      Tab(4).Control(1)=   "Frame1(0)"
      Tab(4).Control(2)=   "Frame1(2)"
      Tab(4).Control(3)=   "chkFilterData"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Weld Analysis"
      TabPicture(5)   =   "FrmOption.frx":0098
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame1(11)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Weld Recording"
      TabPicture(6)   =   "FrmOption.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label2"
      Tab(6).Control(1)=   "Frame4"
      Tab(6).Control(2)=   "txtWeldNumber"
      Tab(6).Control(3)=   "chkRecordInterrupts"
      Tab(6).Control(4)=   "cmdReset"
      Tab(6).ControlCount=   5
      Begin VB.Frame Frame7 
         Caption         =   "Unit Info"
         Height          =   1335
         Left            =   -74640
         TabIndex        =   187
         Top             =   4800
         Width           =   5775
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   10
            Left            =   1920
            TabIndex        =   191
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   1920
            TabIndex        =   189
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lblComp 
            Caption         =   "Location:"
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   190
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Unit:"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   188
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Company Info"
         Height          =   3855
         Left            =   -74640
         TabIndex        =   168
         Top             =   720
         Width           =   5775
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   1920
            TabIndex        =   186
            Top             =   3240
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   1920
            TabIndex        =   184
            Top             =   2880
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   182
            Top             =   2520
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   180
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   178
            Text            =   "China"
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   176
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   174
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   172
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtComp 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   170
            Text            =   "KIWAY"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lblComp 
            Caption         =   "EMail:"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   185
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Fax:"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   183
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Telephone:"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   181
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Contact Name:"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   179
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Country"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   177
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Zip Code:"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   175
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "City:"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   173
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Address:"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   171
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblComp 
            Caption         =   "Company Name:"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   169
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -67800
         TabIndex        =   167
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkRecordInterrupts 
         Caption         =   "Record Interrupts"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70680
         TabIndex        =   166
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtWeldNumber 
         Height          =   375
         Left            =   -69600
         TabIndex        =   164
         Text            =   "A0001"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   $"FrmOption.frx":00D0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   240
         TabIndex        =   153
         Tag             =   "21900"
         Top             =   4560
         Width           =   10335
         Begin VB.TextBox txtWA 
            Height          =   270
            Index           =   19
            Left            =   6720
            TabIndex        =   162
            Text            =   "82.55"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtWA 
            Height          =   270
            Index           =   18
            Left            =   6720
            TabIndex        =   160
            Text            =   "209.05"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtWA 
            Height          =   270
            Index           =   17
            Left            =   2760
            TabIndex        =   158
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtWA 
            Height          =   270
            Index           =   15
            Left            =   2760
            TabIndex        =   155
            Text            =   "10"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtWA 
            Enabled         =   0   'False
            Height          =   270
            Index           =   16
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   "2"
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblAV 
            Caption         =   "Upset Diameter(Rod side) (mm):"
            Height          =   375
            Index           =   28
            Left            =   3600
            TabIndex        =   163
            Tag             =   "50"
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label lblAV 
            Caption         =   "Upset Diameter(Piston side) (mm):"
            Height          =   375
            Index           =   27
            Left            =   3600
            TabIndex        =   161
            Tag             =   "40"
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label lblAV 
            Caption         =   "Upset Current Minimum (A):"
            Height          =   375
            Index           =   26
            Left            =   240
            TabIndex        =   159
            Tag             =   "30"
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lblAV 
            Caption         =   "Initial Voltage(V):"
            Height          =   375
            Index           =   25
            Left            =   240
            TabIndex        =   157
            Tag             =   "10"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblAV 
            Caption         =   $"FrmOption.frx":00E6
            Enabled         =   0   'False
            Height          =   375
            Index           =   24
            Left            =   4080
            TabIndex        =   156
            Tag             =   "20"
            Top             =   1440
            Width           =   2535
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Criteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         TabIndex        =   105
         Top             =   600
         Width           =   10215
         Begin VB.Frame Frame1 
            Caption         =   "Total Rail Usage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   10
            Left            =   7680
            TabIndex        =   149
            Tag             =   "21800"
            Top             =   2160
            Width           =   2415
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   14
               Left            =   1680
               TabIndex        =   151
               Text            =   "30"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   7
               Left            =   240
               TabIndex        =   150
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblAV 
               Caption         =   "Tolal Rail(mm):"
               Height          =   375
               Index           =   22
               Left            =   240
               TabIndex        =   152
               Tag             =   "10"
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Short-Circuit in Boost "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   9
            Left            =   5160
            TabIndex        =   143
            Tag             =   "21700"
            Top             =   2160
            Width           =   2415
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   6
               Left            =   240
               TabIndex        =   146
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   12
               Left            =   1680
               TabIndex        =   145
               Text            =   "550"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   13
               Left            =   1680
               TabIndex        =   144
               Text            =   "0.80"
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblAV 
               Caption         =   "Current(A):"
               Height          =   375
               Index           =   21
               Left            =   240
               TabIndex        =   148
               Tag             =   "10"
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblAV 
               Caption         =   "Time(sec):"
               Height          =   375
               Index           =   20
               Left            =   240
               TabIndex        =   147
               Tag             =   "20"
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Current Interrupt in Boost"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   8
            Left            =   2640
            TabIndex        =   137
            Tag             =   "21600"
            Top             =   2160
            Width           =   2415
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   11
               Left            =   1680
               TabIndex        =   140
               Text            =   "2.00"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   10
               Left            =   1680
               TabIndex        =   139
               Text            =   "100"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   5
               Left            =   240
               TabIndex        =   138
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblAV 
               Caption         =   "Time(sec):"
               Height          =   375
               Index           =   19
               Left            =   240
               TabIndex        =   142
               Tag             =   "20"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblAV 
               Caption         =   "Current(A):"
               Height          =   375
               Index           =   18
               Left            =   240
               TabIndex        =   141
               Tag             =   "10"
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Slippage Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   7
            Left            =   120
            TabIndex        =   131
            Tag             =   "21500"
            Top             =   2160
            Width           =   2415
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   134
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   8
               Left            =   1680
               TabIndex        =   133
               Text            =   "0.75"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   9
               Left            =   1680
               TabIndex        =   132
               Text            =   "22.00"
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblAV 
               Caption         =   "Upset Time(s):"
               Height          =   375
               Index           =   17
               Left            =   240
               TabIndex        =   136
               Tag             =   "10"
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblAV 
               Caption         =   "Upset(mm):"
               Height          =   375
               Index           =   16
               Left            =   240
               TabIndex        =   135
               Tag             =   "20"
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Forge Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   6
            Left            =   7680
            TabIndex        =   125
            Tag             =   "21400"
            Top             =   360
            Width           =   2415
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   128
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   6
               Left            =   1680
               TabIndex        =   127
               Text            =   "30"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   7
               Left            =   1680
               TabIndex        =   126
               Text            =   "60"
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblAV 
               Caption         =   "Minimum(T):"
               Height          =   375
               Index           =   14
               Left            =   100
               TabIndex        =   130
               Tag             =   "10"
               Top             =   720
               Width           =   1600
            End
            Begin VB.Label lblAV 
               Caption         =   "Maximum(T):"
               Height          =   375
               Index           =   15
               Left            =   100
               TabIndex        =   129
               Tag             =   "20"
               Top             =   1120
               Width           =   1600
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Upset Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   5
            Left            =   5160
            TabIndex        =   119
            Tag             =   "21300"
            Top             =   360
            Width           =   2415
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   5
               Left            =   1680
               TabIndex        =   122
               Text            =   "20.00"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   4
               Left            =   1680
               TabIndex        =   121
               Text            =   "14.00"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   120
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblAV 
               Caption         =   "Maximum(mm):"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   124
               Tag             =   "20"
               Top             =   1100
               Width           =   1605
            End
            Begin VB.Label lblAV 
               Caption         =   "Minimum(mm):"
               Height          =   375
               Index           =   13
               Left            =   100
               TabIndex        =   123
               Tag             =   "10"
               Top             =   720
               Width           =   1600
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Boost Speed Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   4
            Left            =   2640
            TabIndex        =   113
            Tag             =   "21200"
            Top             =   360
            Width           =   2415
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   116
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   115
               Text            =   "0.75"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   3
               Left            =   1680
               TabIndex        =   114
               Text            =   "1.20"
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblAV 
               Caption         =   "Minimum(mm/s):"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   118
               Tag             =   "10"
               Top             =   720
               Width           =   1605
            End
            Begin VB.Label lblAV 
               Caption         =   "Maximum(mm/s):"
               Height          =   375
               Index           =   11
               Left            =   100
               TabIndex        =   117
               Tag             =   "20"
               Top             =   1100
               Width           =   1600
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Flash Speed Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   3
            Left            =   120
            TabIndex        =   107
            Tag             =   "21100"
            Top             =   360
            Width           =   2415
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   110
               Text            =   "0.25"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtWA 
               Height          =   270
               Index           =   0
               Left            =   1680
               TabIndex        =   109
               Text            =   "0.14"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkEnableAnalysis 
               Caption         =   "Enable Analysis"
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   108
               Tag             =   "5"
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblAV 
               Caption         =   "Maximum(mm/s):"
               Height          =   375
               Index           =   9
               Left            =   100
               TabIndex        =   112
               Tag             =   "20"
               Top             =   1100
               Width           =   1600
            End
            Begin VB.Label lblAV 
               Caption         =   "Minimum(mm/s):"
               Height          =   375
               Index           =   10
               Left            =   100
               TabIndex        =   111
               Tag             =   "10"
               Top             =   720
               Width           =   1600
            End
         End
      End
      Begin VB.Frame FrameP 
         Caption         =   "Frame5"
         Height          =   3255
         Index           =   0
         Left            =   -74640
         TabIndex        =   101
         Top             =   600
         Width           =   5415
         Begin VB.CheckBox chkSimulate 
            Caption         =   "Simulate"
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtSimulate 
            Height          =   375
            Left            =   120
            TabIndex        =   103
            Text            =   "Text1"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.CommandButton cmdSimulate 
            Caption         =   "::"
            Height          =   375
            Left            =   3480
            TabIndex        =   102
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Voltage"
         Height          =   1455
         Index           =   2
         Left            =   -74760
         TabIndex        =   87
         Tag             =   "18300"
         Top             =   2040
         Width           =   5055
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   8
            Left            =   1440
            TabIndex        =   92
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   9
            Left            =   1440
            TabIndex        =   91
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   10
            Left            =   3840
            TabIndex        =   90
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   11
            Left            =   3840
            TabIndex        =   89
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkCalibration 
            Caption         =   "Calibrate"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Tag             =   "10"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Zero Point:"
            Height          =   300
            Index           =   8
            Left            =   120
            TabIndex        =   100
            Tag             =   "20"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Maxinum:"
            Height          =   300
            Index           =   9
            Left            =   120
            TabIndex        =   99
            Tag             =   "40"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Range:"
            Height          =   300
            Index           =   10
            Left            =   2520
            TabIndex        =   98
            Tag             =   "30"
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Offset:"
            Height          =   300
            Index           =   11
            Left            =   2520
            TabIndex        =   97
            Tag             =   "50"
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   19
            Left            =   2280
            TabIndex        =   96
            Tag             =   "25"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   18
            Left            =   2280
            TabIndex        =   95
            Tag             =   "45"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "V"
            Height          =   225
            Index           =   17
            Left            =   4680
            TabIndex        =   94
            Tag             =   "35"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "V"
            Height          =   225
            Index           =   16
            Left            =   4680
            TabIndex        =   93
            Tag             =   "55"
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pressure"
         Height          =   1455
         Index           =   3
         Left            =   -69600
         TabIndex        =   72
         Tag             =   "18400"
         Top             =   2040
         Width           =   5055
         Begin VB.CheckBox chkCalibration 
            Caption         =   "Calibrate"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   77
            Tag             =   "10"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   15
            Left            =   3840
            TabIndex        =   76
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   14
            Left            =   3840
            TabIndex        =   75
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   13
            Left            =   1440
            TabIndex        =   74
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   12
            Left            =   1440
            TabIndex        =   73
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "psi"
            Height          =   225
            Index           =   31
            Left            =   4680
            TabIndex        =   85
            Tag             =   "55"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "psi"
            Height          =   225
            Index           =   30
            Left            =   4680
            TabIndex        =   84
            Tag             =   "35"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   29
            Left            =   2280
            TabIndex        =   83
            Tag             =   "45"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   28
            Left            =   2280
            TabIndex        =   82
            Tag             =   "25"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Offset:"
            Height          =   300
            Index           =   15
            Left            =   2520
            TabIndex        =   81
            Tag             =   "50"
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Range:"
            Height          =   300
            Index           =   14
            Left            =   2520
            TabIndex        =   80
            Tag             =   "30"
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Maxinum:"
            Height          =   300
            Index           =   13
            Left            =   120
            TabIndex        =   79
            Tag             =   "40"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Zero Point:"
            Height          =   300
            Index           =   12
            Left            =   120
            TabIndex        =   78
            Tag             =   "20"
            Top             =   600
            Width           =   1215
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -73440
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Caption         =   "Distance(A) and Force(V)"
         Height          =   1575
         Index           =   1
         Left            =   -71160
         TabIndex        =   59
         Tag             =   "20200"
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   3
            Left            =   1680
            TabIndex        =   52
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   4
            Left            =   1680
            TabIndex        =   51
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   5
            Left            =   1680
            TabIndex        =   53
            Text            =   "0"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblAV 
            Caption         =   "Minimum:"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   62
            Tag             =   "10"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblAV 
            Caption         =   "Maxinum:"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   61
            Tag             =   "20"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblAV 
            Caption         =   "Grid Increment :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Tag             =   "30"
            Top             =   1100
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Current(A) and Voltage(V)"
         Height          =   1575
         Index           =   0
         Left            =   -74760
         TabIndex        =   47
         Tag             =   "20100"
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   2
            Left            =   1680
            TabIndex        =   50
            Text            =   "0"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   1
            Left            =   1680
            TabIndex        =   49
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   0
            Left            =   1680
            TabIndex        =   48
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblAV 
            Caption         =   "Grid Increment :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Tag             =   "30"
            Top             =   1100
            Width           =   1455
         End
         Begin VB.Label lblAV 
            Caption         =   "Maxinum:"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Tag             =   "20"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblAV 
            Caption         =   "Minimum:"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Tag             =   "10"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Time (sec)"
         Height          =   1455
         Index           =   2
         Left            =   -74760
         TabIndex        =   44
         Tag             =   "20300"
         Top             =   2280
         Width           =   4575
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   7
            Left            =   1800
            TabIndex        =   57
            Text            =   "0"
            Top             =   705
            Width           =   1215
         End
         Begin VB.TextBox txtWC 
            Height          =   270
            Index           =   6
            Left            =   2760
            TabIndex        =   55
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblAV 
            Caption         =   "Grid Increment :"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Tag             =   "20"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblAV 
            Caption         =   "Minimum Weld Cycle Time(s):"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   45
            Tag             =   "10"
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Limits"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   35
         Tag             =   "19100"
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtSRB 
            Height          =   270
            Index           =   0
            Left            =   1440
            TabIndex        =   39
            Text            =   "1000"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtSRB 
            Height          =   270
            Index           =   1
            Left            =   3600
            TabIndex        =   38
            Text            =   "100"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtSRB 
            Height          =   270
            Index           =   2
            Left            =   1440
            TabIndex        =   37
            Text            =   "500"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtSRB 
            Height          =   270
            Index           =   3
            Left            =   3600
            TabIndex        =   36
            Text            =   "50"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblSRB 
            Alignment       =   1  'Right Justify
            Caption         =   "Current(A):"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Tag             =   "1"
            Top             =   375
            Width           =   1215
         End
         Begin VB.Label lblSRB 
            Alignment       =   1  'Right Justify
            Caption         =   "Distance(mm):"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   42
            Tag             =   "2"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblSRB 
            Alignment       =   1  'Right Justify
            Caption         =   "Voltage(V): "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Tag             =   "3"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblSRB 
            Alignment       =   1  'Right Justify
            Caption         =   "Pressure(psi):"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   40
            Tag             =   "4"
            Top             =   735
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboLanguage 
         Height          =   300
         ItemData        =   "FrmOption.frx":0106
         Left            =   -67440
         List            =   "FrmOption.frx":0113
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkOnlineOnStartUp 
         Caption         =   "OnlineOnStartup"
         Height          =   375
         Left            =   -68400
         TabIndex        =   33
         Tag             =   "16120"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "Distance"
         Height          =   1455
         Index           =   0
         Left            =   -74760
         TabIndex        =   19
         Tag             =   "18100"
         Top             =   480
         Width           =   5055
         Begin VB.CheckBox chkCalibration 
            Caption         =   "Calibrate"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Tag             =   "10"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   3
            Left            =   3840
            TabIndex        =   23
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   2
            Left            =   3840
            TabIndex        =   22
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   1
            Left            =   1440
            TabIndex        =   21
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   0
            Left            =   1440
            TabIndex        =   20
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "mm"
            Height          =   225
            Index           =   7
            Left            =   4680
            TabIndex        =   32
            Tag             =   "55"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "mm"
            Height          =   225
            Index           =   6
            Left            =   4680
            TabIndex        =   31
            Tag             =   "35"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   5
            Left            =   2280
            TabIndex        =   30
            Tag             =   "45"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   4
            Left            =   2280
            TabIndex        =   29
            Tag             =   "25"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Offset:"
            Height          =   300
            Index           =   3
            Left            =   2520
            TabIndex        =   28
            Tag             =   "50"
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Fuly Stroke:"
            Height          =   300
            Index           =   2
            Left            =   2520
            TabIndex        =   27
            Tag             =   "30"
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Fuly Close:"
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Tag             =   "40"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Fuly Open:"
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Tag             =   "20"
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current"
         Height          =   1455
         Index           =   1
         Left            =   -69600
         TabIndex        =   5
         Tag             =   "18200"
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   4
            Left            =   1440
            TabIndex        =   10
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   5
            Left            =   1440
            TabIndex        =   9
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   6
            Left            =   3840
            TabIndex        =   8
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   250
            Index           =   7
            Left            =   3840
            TabIndex        =   7
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkCalibration 
            Caption         =   "Calibrate"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Tag             =   "10"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Zero Point:"
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Tag             =   "20"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Maxinum:"
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Tag             =   "40"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Range:"
            Height          =   300
            Index           =   6
            Left            =   2760
            TabIndex        =   16
            Tag             =   "30"
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lblCalibrate 
            Alignment       =   1  'Right Justify
            Caption         =   "Offset:"
            Height          =   300
            Index           =   7
            Left            =   2520
            TabIndex        =   15
            Tag             =   "50"
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   12
            Left            =   2280
            TabIndex        =   14
            Tag             =   "25"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "DU"
            Height          =   225
            Index           =   13
            Left            =   2280
            TabIndex        =   13
            Tag             =   "45"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "A"
            Height          =   225
            Index           =   14
            Left            =   4680
            TabIndex        =   12
            Tag             =   "35"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "A"
            Height          =   225
            Index           =   15
            Left            =   4680
            TabIndex        =   11
            Tag             =   "55"
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.CheckBox chkFilterData 
         Caption         =   "FilterData"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Tag             =   "20010"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Start Recording"
         Height          =   1935
         Left            =   -74640
         TabIndex        =   2
         Tag             =   "22100"
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtStartRecording 
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   3000
            TabIndex        =   71
            Text            =   "2.50"
            Top             =   4560
            Width           =   855
         End
         Begin VB.TextBox txtStartRecording 
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   3000
            TabIndex        =   70
            Text            =   "2.50"
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox txtStartRecording 
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   600
            TabIndex        =   69
            Text            =   "2.50"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtStartRecording 
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   3000
            TabIndex        =   68
            Text            =   "2.50"
            Top             =   2400
            Width           =   855
         End
         Begin VB.OptionButton optStartRecording 
            Caption         =   "Time (s)"
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   67
            Tag             =   "50"
            Top             =   4200
            Width           =   2055
         End
         Begin VB.OptionButton optStartRecording 
            Caption         =   "Volt (V)"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   66
            Tag             =   "40"
            Top             =   3480
            Width           =   2055
         End
         Begin VB.OptionButton optStartRecording 
            Caption         =   "Current (A)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   65
            Tag             =   "30"
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton optStartRecording 
            Caption         =   "Distance (mm)"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   64
            Tag             =   "20"
            Top             =   2040
            Width           =   2055
         End
         Begin VB.OptionButton optStartRecording 
            Caption         =   "Start of Weld Cycle"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Tag             =   "10"
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Weld Number:"
         Height          =   375
         Left            =   -70800
         TabIndex        =   165
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   5640
         TabIndex        =   106
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblLanguage 
         Caption         =   "Language:"
         Height          =   255
         Left            =   -68520
         TabIndex        =   63
         Tag             =   "16110"
         Top             =   795
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SRB_Data(4) As Integer
Dim Calibration_Data(16) As Single
Dim Calibration_Enable(4) As Single
Dim WeldChart_Data(16) As Single
Dim WeldAnalysis_Data(19) As Single
Dim WeldAnalysisEnable_Data(19) As Single

Dim StartRecording As Integer
Dim StartRecodingParam(5) As Single

Dim LANGUAGE As String
Dim IsSimulate As Integer
Dim SimulateFile As String

Private Sub cboLanguage_Click()
LANGUAGE = cboLanguage.Text
End Sub


Private Sub chkCalibration_Click(Index As Integer)
    Calibration_Enable(Index) = chkCalibration(Index).Value
    If chkCalibration(Index).Value = 1 Then
        Call CalibrationSwitchTo(Index, True)
    Else
        Call CalibrationSwitchTo(Index, False)
    End If
End Sub

Private Function CalibrationSwitchTo(Index As Integer, enable As Boolean)
        txt(Index * 4).Enabled = enable
        lblCalibrate(Index * 4).Enabled = enable
        txt(Index * 4 + 1).Enabled = enable
        lblCalibrate(Index * 4 + 1).Enabled = enable
        txt(Index * 4 + 2).Enabled = enable
        lblCalibrate(Index * 4 + 2).Enabled = enable
        txt(Index * 4 + 3).Enabled = enable
        lblCalibrate(Index * 4 + 3).Enabled = enable
End Function


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Function UpgradeCalibration(cd() As Single, Calibration_Enable() As Single)
    Dim ca(11) As Single
    Dim o As Integer
    Dim p As Integer
    Dim i As Integer
    Dim f As Single
    f = 0
    
    For i = 0 To 3
        o = i * 3
        p = i * 4
        
        If Calibration_Enable(i) = 1 And cd(p + 1) - cd(p + 0) <> 0 Then
                ca(o + 0) = cd(p + 2) / (cd(p + 1) - cd(p + 0))
                ca(o + 1) = cd(p + 0) * ca(o + 0)
                ca(o + 2) = cd(p + 3)
        Else
            ca(o + 0) = 1
            ca(o + 1) = 0
            ca(o + 2) = 0
        End If
     Next
        
    
    Call PLCDrv.InitPLCConnection
    Call PLCDrv.WriteCalibrationData(ca)
    Call PLCDrv.UninitPLCConection
    
End Function

Private Sub cmdOK_Click()

    Call SaveSetting(App.EXEName, "General", "Language", LANGUAGE)
    Call SaveSetting(App.EXEName, "General", "IsSimulate", chkOnlineOnStartUp.Value)
    
    Call SaveSetting(App.EXEName, "Simulate", "IsSimulate", chkSimulate.Value)
    Call SaveSetting(App.EXEName, "Simulate", "SimulateFilename", txtSimulate.Text)
    
    
    
'    Call SaveSetting(App.EXEName, "Calibration", "SendIntoController", True)
    
    
    Dim i As Integer
    
    Dim v As String
    Dim vo As String
    vo = GetSetting(App.EXEName, "Calibration", "value", "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,")
    
    For i = 0 To 19 Step 5
       v = v & Calibration_Enable(i / 5) & ","
       v = v & Calibration_Data(i - i / 5) & ","
       v = v & Calibration_Data(i + 1 - i / 5) & ","
       v = v & Calibration_Data(i + 2 - i / 5) & ","
       v = v & Calibration_Data(i + 3 - i / 5) & ","
    Next
    
    
    If v <> vo Then
        Call SaveSetting(App.EXEName, "Calibration", "value", v)
        UpgradeCalibration Calibration_Data, Calibration_Enable
    End If
    
    


    Call SaveSetting(App.EXEName, "WeldChartSetting", "AVMin", WeldChart_Data(0))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "AVMax", WeldChart_Data(1))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "AVIncr", WeldChart_Data(2))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "DFMin", WeldChart_Data(3))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "DFMax", WeldChart_Data(4))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "DFIncr", WeldChart_Data(5))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "TimeMaxCycleTime", WeldChart_Data(6))
    Call SaveSetting(App.EXEName, "WeldChartSetting", "TimeIncr", WeldChart_Data(7))
    
        
    Call SaveSetting(App.EXEName, "SensorReadingBar", "Amp", SRB_Data(0))
    Call SaveSetting(App.EXEName, "SensorReadingBar", "Dist", SRB_Data(1))
    Call SaveSetting(App.EXEName, "SensorReadingBar", "Volt", SRB_Data(2))
    Call SaveSetting(App.EXEName, "SensorReadingBar", "Press", SRB_Data(3))
    
    
    
    Call SaveSetting(App.EXEName, "StartRecording", "StartRecording", StartRecording)
    Call SaveSetting(App.EXEName, "StartRecording", "Dist", StartRecodingParam(1))
    Call SaveSetting(App.EXEName, "StartRecording", "Amp", StartRecodingParam(2))
    Call SaveSetting(App.EXEName, "StartRecording", "Volt", StartRecodingParam(3))
    Call SaveSetting(App.EXEName, "StartRecording", "Time", StartRecodingParam(4))
    
  
  Call SaveSetting(App.EXEName, "AnalysisDefine", "FlashMin", WeldAnalysis_Data(0))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "FlashMax", WeldAnalysis_Data(1))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "BoostMin", WeldAnalysis_Data(2))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "BoostMax", WeldAnalysis_Data(3))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetMin", WeldAnalysis_Data(4))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetMax", WeldAnalysis_Data(5))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ForgeMin", WeldAnalysis_Data(6))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ForgeMax", WeldAnalysis_Data(7))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "SlippageUpsetTime", WeldAnalysis_Data(8))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "SlippageUpset", WeldAnalysis_Data(9))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptCurrent", WeldAnalysis_Data(10))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptTime", WeldAnalysis_Data(11))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ShortCircuitCurrent", WeldAnalysis_Data(12))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ShortCircuitTime", WeldAnalysis_Data(13))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageTotalRail", WeldAnalysis_Data(14))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "FlashSpeedTimeRange", WeldAnalysis_Data(15))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "BoostSpeedTimeRange", WeldAnalysis_Data(16))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetCurrentMinimum", WeldAnalysis_Data(17))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Pistonside)", WeldAnalysis_Data(18))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Rodside)", WeldAnalysis_Data(19))

    
  Call SaveSetting(App.EXEName, "AnalysisDefine", "FlashEnable", WeldAnalysisEnable_Data(0))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "BoostEnable", WeldAnalysisEnable_Data(1))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "UpsetEnable", WeldAnalysisEnable_Data(2))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ForgeEnable", WeldAnalysisEnable_Data(3))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "SlippageEnable", WeldAnalysisEnable_Data(4))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptEnable", WeldAnalysisEnable_Data(5))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "ShortCircuitEnable", WeldAnalysisEnable_Data(6))
  Call SaveSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageEnable", WeldAnalysisEnable_Data(7))


    
    Me.Hide
    Unload Me
End Sub

Private Sub cmdReset_Click()
   Dim weldSerailNumber As Long
   weldSerailNumber = fromWeldNumberShowModel(txtWeldNumber.Text)
   
   Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    
End Sub

Private Sub cmdSimulate_Click()
     CommonDialog1.filename = txtSimulate.Text
    CommonDialog1.ShowOpen
    txtSimulate.Text = CommonDialog1.filename
    
End Sub

Private Sub Form_Load()
' Resource
PlcRes.LoadResFor Me

LANGUAGE = GetSetting(App.EXEName, "General", "Language", "EN")
cboLanguage.Text = LANGUAGE

chkOnlineOnStartUp.Value = GetSetting(App.EXEName, "Genaral", "IsSimulate", 0)

chkSimulate.Value = GetSetting(App.EXEName, "Simulate", "IsSimulate", 0)
txtSimulate.Text = GetSetting(App.EXEName, "Simulate", "SimulateFilename", App.path & "\T0039.WLD")


    
StartRecording = CInt(GetSetting(App.EXEName, "StartRecording", "StartRecording", 0))
StartRecodingParam(1) = CSng(GetSetting(App.EXEName, "StartRecording", "Dist", 2.5))
StartRecodingParam(2) = CSng(GetSetting(App.EXEName, "StartRecording", "Amp", 100))
StartRecodingParam(3) = CSng(GetSetting(App.EXEName, "StartRecording", "Volt", 450))
StartRecodingParam(4) = CSng(GetSetting(App.EXEName, "StartRecording", "Time", 25))
    
    
    Dim i As Integer
    For i = 0 To 4
        optStartRecording(i).Value = False
    Next
    optStartRecording(StartRecording).Value = True
    
     Dim vo As String
    vo = GetSetting(App.EXEName, "Calibration", "value", "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,")
    Dim cd() As String
    cd = Split(vo, ",")
    
    
    '----------------
 
    Dim k As Integer
    Dim j As Integer
           
    For i = 0 To 3
        k = 4 * i
        j = 5 * i
        
        Calibration_Enable(i) = CSng(cd(j))
        chkCalibration(i).Value = Calibration_Enable(i)
                
        If Calibration_Enable(i) = 1 Then
            Call CalibrationSwitchTo(i, True)
        Else
            Call CalibrationSwitchTo(i, False)
        End If
       
       
        j = j + 1
        Calibration_Data(k) = CSng(cd(j))
        txt(k).Text = Calibration_Data(k)
        
        k = k + 1
        j = j + 1
        Calibration_Data(k) = CSng(cd(j))
        txt(k).Text = Calibration_Data(k)
        
        k = k + 1
        j = j + 1
        Calibration_Data(k) = CSng(cd(j))
        txt(k).Text = Calibration_Data(k)
        
        k = k + 1
        j = j + 1
        Calibration_Data(k) = CSng(cd(j))
        txt(k).Text = Calibration_Data(k)
    Next
    

    
    
txtStartRecording(1).Text = StartRecodingParam(1)
txtStartRecording(2).Text = StartRecodingParam(2)
txtStartRecording(3).Text = StartRecodingParam(3)
txtStartRecording(4).Text = StartRecodingParam(4)
    
    
WeldChart_Data(0) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMin", 0))
WeldChart_Data(1) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMax", 1000))
WeldChart_Data(2) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVIncr", 100))
WeldChart_Data(3) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMin", 0))
WeldChart_Data(4) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMax", 160))
WeldChart_Data(5) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFIncr", 16))
WeldChart_Data(6) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "TimeMaxCycleTime", 200))
WeldChart_Data(7) = CInt(GetSetting(App.EXEName, "WeldChartSetting", "TimeIncr", 10))

SRB_Data(0) = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Amp", 1000))
SRB_Data(1) = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Dist", 100))
SRB_Data(2) = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Volt", 500))
SRB_Data(3) = CInt(GetSetting(App.EXEName, "SensorReadingBar", "Press", 50))


WeldAnalysisEnable_Data(0) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "FlashEnable", 1))
WeldAnalysisEnable_Data(1) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "BoostEnable", 1))
WeldAnalysisEnable_Data(2) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "UpsetEnable", 1))
WeldAnalysisEnable_Data(3) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "ForgeEnable", 1))
WeldAnalysisEnable_Data(4) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "SlippageEnable", 1))
WeldAnalysisEnable_Data(5) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptEnable", 1))
WeldAnalysisEnable_Data(6) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitEnable", 1))
WeldAnalysisEnable_Data(7) = CInt(GetSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageEnable", 1))

WeldAnalysis_Data(0) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashMin", 0.14))
WeldAnalysis_Data(1) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashMax", 0.25))
WeldAnalysis_Data(2) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostMin", 0.75))
WeldAnalysis_Data(3) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostMax", 1.2))
WeldAnalysis_Data(4) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetMin", 14#))
WeldAnalysis_Data(5) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetMax", 20#))
WeldAnalysis_Data(6) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ForgeMin", 30))
WeldAnalysis_Data(7) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ForgeMax", 60))
WeldAnalysis_Data(8) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "SlippageUpsetTime", 0.75))
WeldAnalysis_Data(9) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "SlippageUpset", 22#))
WeldAnalysis_Data(10) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptCurrent", 100))
WeldAnalysis_Data(11) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "CurrentInterruptTime", 2#))
WeldAnalysis_Data(12) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitCurrent", 550))
WeldAnalysis_Data(13) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "ShortCircuitTime", 0.8))
WeldAnalysis_Data(14) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "TotalRailUsageTotalRail", 30))
WeldAnalysis_Data(15) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "FlashSpeedTimeRange", 10))
WeldAnalysis_Data(16) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "BoostSpeedTimeRange", 2))
WeldAnalysis_Data(17) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetCurrentMinimum", 100))
WeldAnalysis_Data(18) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Pistonside)", 209.55))
WeldAnalysis_Data(19) = CSng(GetSetting(App.EXEName, "AnalysisDefine", "UpsetDiameter(Rodside)", 82.55))




    For i = 0 To 3
        txtSRB(i).Text = SRB_Data(i)
    Next
    
    For i = 0 To 7
        txtWC(i).Text = WeldChart_Data(i)
    Next
    
    For i = 0 To 19
        txtWA(i).Text = WeldAnalysis_Data(i)
    Next
    
    
    For i = 0 To 7
        chkEnableAnalysis(i).Value = WeldAnalysisEnable_Data(i)
    Next
    
    Dim weldSerailNumber As Long
    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    txtWeldNumber.Text = PlcCommon.toWeldNumberShowModel(weldSerailNumber)
    
End Sub


Private Sub optStartRecording_Click(Index As Integer)
    StartRecording = Index
    Dim i As Integer
    For i = 1 To 4
        txtStartRecording(i).Enabled = False
    Next
    If Index > 0 Then
        txtStartRecording(Index).Enabled = True
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    If IsNumeric(txt(Index).Text) Then
        Calibration_Data(Index) = CSng(txt(Index).Text)
    Else
        txt(Index).Text = Calibration_Data(Index)
    End If
End Sub

Private Sub txtSRB_Change(Index As Integer)
    If IsNumeric(txtSRB(Index).Text) Then
        SRB_Data(Index) = CInt(txtSRB(Index).Text)
    Else
        txtSRB(Index).Text = SRB_Data(Index)
    End If
End Sub

Private Sub txtStartRecording_Change(Index As Integer)
    If IsNumeric(txtStartRecording(Index).Text) Then
        StartRecodingParam(Index) = CSng(txtStartRecording(Index).Text)
    Else
        txtStartRecording(Index).Text = StartRecodingParam(Index)
    End If
End Sub

Private Sub txtWA_Change(Index As Integer)
    If IsNumeric(txtWA(Index).Text) Then
        WeldAnalysis_Data(Index) = CSng(txtWA(Index).Text)
    Else
        txtWA(Index).Text = WeldAnalysis_Data(Index)
    End If
End Sub

Private Sub txtWC_Change(Index As Integer)
    If IsNumeric(txtWC(Index).Text) Then
        WeldChart_Data(Index) = CInt(txtWC(Index).Text)
    Else
        txtWC(Index).Text = WeldChart_Data(Index)
    End If

End Sub

Private Sub chkEnableAnalysis_Click(Index As Integer)
    WeldAnalysisEnable_Data(Index) = chkEnableAnalysis(Index).Value
End Sub
