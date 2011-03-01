VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmChart 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "FrmChart.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Tag             =   "11000"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   0
      TabIndex        =   1
      Tag             =   "11000"
      Top             =   1080
      Width           =   4335
      Begin VB.CommandButton cmdViewDataDetail 
         Caption         =   "Detail"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   8880
         Width           =   975
      End
      Begin VB.Label lblGroup 
         Caption         =   "Pre-Flash"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Tag             =   "100"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Voltage(V):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Tag             =   "110"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Current(A):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Tag             =   "120"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Rail Used(mm):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   64
         Tag             =   "130"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Duration(s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   63
         Tag             =   "140"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   62
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   61
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   60
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   59
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblGroup 
         Caption         =   "Flash"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Tag             =   "200"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Voltage(V):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   57
         Tag             =   "210"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Current(A):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   56
         Tag             =   "220"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Rail Used(mm): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Tag             =   "230"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Flash Speed(mm/s):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   54
         Tag             =   "240"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   53
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   52
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   51
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   50
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Duration(s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   49
         Tag             =   "250"
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   48
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label lblGroup 
         Caption         =   "Boost"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Tag             =   "300"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Voltage(V):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   46
         Tag             =   "310"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Avg.Current(A):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   45
         Tag             =   "320"
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Rail Used(mm):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   44
         Tag             =   "330"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Max Speed(mm/s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   43
         Tag             =   "340"
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   42
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   41
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   40
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2520
         TabIndex        =   39
         Top             =   4080
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Current Iterrupt(Y/N):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   38
         Tag             =   "350"
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   2520
         TabIndex        =   37
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Short Circuit(Y/N):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   36
         Tag             =   "360"
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2520
         TabIndex        =   35
         Top             =   4560
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Duration(s):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   34
         Tag             =   "370"
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   2520
         TabIndex        =   33
         Top             =   4800
         Width           =   795
      End
      Begin VB.Label lblGroup 
         Caption         =   "Upset"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Tag             =   "400"
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Rail Used(mm):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   31
         Tag             =   "410"
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Slippage(Y/N):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   30
         Tag             =   "420"
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Maximum Current(A): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   29
         Tag             =   "430"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Current ON time(s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   28
         Tag             =   "440"
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   2520
         TabIndex        =   27
         Top             =   5400
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   2520
         TabIndex        =   26
         Top             =   5640
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   2520
         TabIndex        =   25
         Top             =   5880
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   24
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Duration(s):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   23
         Tag             =   "450"
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   22
         Top             =   6360
         Width           =   795
      End
      Begin VB.Label lblGroup 
         Caption         =   "Forge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Tag             =   "500"
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Forge Force(t):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   20
         Tag             =   "510"
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Duration(s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   19
         Tag             =   "520"
         Top             =   7200
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   2520
         TabIndex        =   18
         Top             =   6960
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   17
         Top             =   7200
         Width           =   795
      End
      Begin VB.Label lblGroup 
         Caption         =   "Overall"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Tag             =   "600"
         Top             =   7560
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Impedance(Ohm)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   15
         Tag             =   "610"
         Top             =   7800
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Total Rail used(mm): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   14
         Tag             =   "620"
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Holding Time(m): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   13
         Tag             =   "630"
         Top             =   8280
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Total Duration(s): "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   12
         Tag             =   "640"
         Top             =   8520
         Width           =   2175
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   2520
         TabIndex        =   11
         Top             =   7800
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   2520
         TabIndex        =   10
         Top             =   8040
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   2520
         TabIndex        =   9
         Top             =   8280
         Width           =   795
      End
      Begin VB.Label lblItemData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   2520
         TabIndex        =   8
         Top             =   8520
         Width           =   795
      End
      Begin VB.Label lblCriData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   7
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label lblCriData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   6
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label lblCriData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   3360
         TabIndex        =   5
         Top             =   5400
         Width           =   915
      End
      Begin VB.Label lblCriData 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   3360
         TabIndex        =   4
         Top             =   6960
         Width           =   915
      End
      Begin VB.Label lblCriDatadddd 
         Caption         =   "Min/Max"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Tag             =   "5"
         Top             =   120
         Width           =   915
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   8295
      Left            =   4440
      OleObjectBlob   =   "FrmChart.frx":0442
      TabIndex        =   0
      Top             =   1320
      Width           =   12015
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "19:12:54"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9600
      TabIndex        =   74
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Caption         =   "2011-01-01"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7200
      TabIndex        =   73
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblProgram 
      Alignment       =   2  'Center
      Caption         =   "P:BASERED"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7200
      TabIndex        =   72
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      Caption         =   "YARDWAY LTD."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7200
      TabIndex        =   71
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "K0035 - OK"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7200
      TabIndex        =   70
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblLocation 
      Caption         =   "LOCATION:CRETE ILL"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9960
      TabIndex        =   69
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblUnit 
      Caption         =   "UNIT:K922SN99-U101136(CW632)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6000
      TabIndex        =   68
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "FrmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MODEl_COMMON As Integer = 0
Const MODEL_SMALL As Integer = 1

Enum ModelConstants
    COMMON
    SMALL
End Enum

Dim model As ModelConstants

Dim dataForm As FrmDataGrid


Dim buf() As WeldData

Const SUCCEED_COLOR As Long = &HC000&
Const FAIL_COLOR As Long = &H80000008

Dim fr As FileR
Public Sub Load(filename As String)
' Resource
PlcRes.LoadResFor Me
   
    fr = PlcWld.LoadData(filename)
    
    model = COMMON
        

Dim i As Integer

ReDim buf(fr.header.RecordCount - 1)

For i = 0 To fr.header.RecordCount - 1
    buf(i) = fr.data(i).data
Next


    Call setChart(buf)
    
    lblCompany.Caption = Trim(fr.header.CompanyName)
    lblParam.Caption = Trim(fr.header.filename)
    lblProgram.Caption = Trim(fr.header.BaseRed)
    
    lblDate.Caption = Trim(fr.header.Date)
    lblTime.Caption = Trim(fr.header.Time)
    
    lblUnit.Caption = "UNIT:" & Trim(fr.header.UnitName)
    lblLocation.Caption = "LOCATION:" & Trim(fr.header.Location)
    
    Call anaylize(fr.analysisDefine, fr.analysisResult)

End Sub



Private Function setChartSmall(EmulateData() As WeldData)


Dim i As Integer
Dim bOk As Boolean
Dim pos As Long
Dim posStart As Long
bOk = False

For i = 0 To UBound(EmulateData)
    If EmulateData(i).WeldStage = BOOST_STAGE Then
        posStart = i
        bOk = True
        Exit For
    End If
Next

If bOk = False Then
    Exit Function
End If

Dim sTime As Single
sTime = EmulateData(posStart).Time

For i = posStart To UBound(EmulateData)
    If EmulateData(i).Time - sTime > 15 Then
        pos = i
        bOk = True
        Exit For
    End If
Next

If bOk = False Then
    Exit Function
End If




Dim count As Integer
count = pos - posStart + 1

ReDim MyData(0 To 4, 0 To count)

MyData(1, 0) = "Force"
MyData(2, 0) = "Volt"
MyData(3, 0) = "Amp"
MyData(4, 0) = "Dist"




i = posStart
While i <= UBound(EmulateData) And i <= pos
    MyData(0, i - posStart + 1) = CInt(EmulateData(i).Time - sTime) & Space(1) '注意一定要后面的space(1)，这样做的目的是为了自动显示成标签(字符串类型)
    i = i + 1
Wend

For i = posStart To pos

    MyData(1, i + 1 - posStart) = PlcAnalysiser.toForce(EmulateData(i).PsiUpset, EmulateData(i).PsiOpen)
    MyData(2, i + 1 - posStart) = EmulateData(i).Volt
    MyData(3, i + 1 - posStart) = EmulateData(i).Amp
    MyData(4, i + 1 - posStart) = EmulateData(i).Dist
Next
MSChart1.Plot.DataSeriesInRow = True '设置图形按行读取数据
With MSChart1.Plot.Axis(VtChAxisIdY).ValueScale
    .Minimum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMin", 0))
    .Maximum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMax", 1000))
    .MajorDivision = (.Maximum - .Minimum) / CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVIncr", 100))
End With
With MSChart1.Plot.Axis(VtChAxisIdY2).ValueScale
    .Minimum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMin", 0))
    .Maximum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMax", 160))
    .MajorDivision = (.Maximum - .Minimum) / CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFIncr", 16))
End With
With MSChart1.Plot.Axis(VtChAxisIdX).ValueScale
    .Minimum = 0
    .Maximum = 15
    .MajorDivision = 8
End With
With MSChart1.Plot.Axis(VtChAxisIdX).ValueScale
    .Minimum = 0
    .Maximum = 15
    .MajorDivision = 1
End With
With MSChart1.Plot.Axis(VtChAxisIdX).CategoryScale
    .DivisionsPerLabel = (pos - posStart) / 3
    .DivisionsPerTick = (pos - posStart) / 3
End With

MSChart1.ChartData = MyData
End Function
Private Function setChart(EmulateData() As WeldData)
Dim TimeMax As Integer

TimeMax = CInt(GetSetting(App.EXEName, "WeldChartSetting", "TimeMaxCycleTime", 200))


Dim count As Integer
count = CInt(UBound(EmulateData) * TimeMax / EmulateData(UBound(EmulateData) - 1).Time)

ReDim MyData(0 To 4, 0 To count + 1)

MyData(1, 0) = "Force"
MyData(2, 0) = "Volt"
MyData(3, 0) = "Amp"
MyData(4, 0) = "Dist"

Dim i As Integer
i = 0
While i <= UBound(EmulateData) And i <= count
    MyData(0, i + 1) = CInt(EmulateData(i).Time) & Space(1)  '注意一定要后面的space(1)，这样做的目的是为了自动显示成标签(字符串类型)
    i = i + 1
Wend

For i = UBound(EmulateData) + 1 To count
MyData(0, i + 1) = CInt(((i / count) * TimeMax)) & Space(1)   '注意一定要后面的space(1)，这样做的目的是为了自动显示成标签(字符串类型)
Next

Dim lastv As Integer


' issue from american 7
Dim step As Integer
step = 2

i = 0
lastv = 0
While i <= UBound(EmulateData) And i <= count
    If i = CInt(i / step) * step Then
        lastv = i
        MyData(1, i + 1) = PlcAnalysiser.toForce(EmulateData(i).PsiUpset, EmulateData(i).PsiOpen)
        MyData(2, i + 1) = EmulateData(i).Volt
        MyData(3, i + 1) = EmulateData(i).Amp
        MyData(4, i + 1) = EmulateData(i).Dist
    Else
        MyData(1, i + 1) = PlcAnalysiser.toForce(EmulateData(lastv).PsiUpset, EmulateData(lastv).PsiOpen)
        MyData(2, i + 1) = EmulateData(lastv).Volt
        MyData(3, i + 1) = EmulateData(lastv).Amp
        MyData(4, i + 1) = EmulateData(i).Dist
    End If
    i = i + 1
Wend

MSChart1.Plot.DataSeriesInRow = True '设置图形按行读取数据
With MSChart1.Plot.Axis(VtChAxisIdY).ValueScale
    .Minimum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMin", 0))
    .Maximum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVMax", 1000))
    .MajorDivision = (.Maximum - .Minimum) / CInt(GetSetting(App.EXEName, "WeldChartSetting", "AVIncr", 100))
End With
With MSChart1.Plot.Axis(VtChAxisIdY2).ValueScale
    .Minimum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMin", 0))
    .Maximum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFMax", 160))
    .MajorDivision = (.Maximum - .Minimum) / CInt(GetSetting(App.EXEName, "WeldChartSetting", "DFIncr", 16))
End With
With MSChart1.Plot.Axis(VtChAxisIdX).ValueScale
    .Minimum = 0
    .Maximum = CInt(GetSetting(App.EXEName, "WeldChartSetting", "TimeMaxCycleTime", 200))
    .MajorDivision = (.Maximum - .Minimum) / CInt(GetSetting(App.EXEName, "WeldChartSetting", "TimeIncr", 10))
End With
With MSChart1.Plot.Axis(VtChAxisIdX).CategoryScale
    .DivisionsPerLabel = UBound(EmulateData) * 20 / EmulateData(UBound(EmulateData)).Time
    .DivisionsPerTick = .DivisionsPerLabel
End With
MSChart1.ChartData = MyData

End Function

Private Function anaylize(analysisDefine As WeldAnalysisDefineType, r As WeldAnalysisResultType)

'Pre flash
lblItemData(0).Caption = r.PreFlashVoltage
lblItemData(1).Caption = r.PreFlashCurrent
lblItemData(2).Caption = Format(r.PreFlashRailUsed, "0.0")
lblItemData(3).Caption = CInt(r.PreFlashDuration)

'flash
lblItemData(4).Caption = r.FlashVoltage
lblItemData(5).Caption = r.FlashCurrent
lblItemData(6).Caption = Format(r.FlashRailUsed, "0.0")
lblItemData(7).Caption = Format(r.FlashSpeed, "0.00")
If analysisDefine.FlashEnable Then
    updateCueWithCri 7, r.FlashSpeedSucceed
    lblCriData(7).Caption = Format(analysisDefine.FlashMin, "0.00") & " / " & Format(analysisDefine.FlashMax, "0.00")
End If

lblItemData(8).Caption = CInt(r.FlashDuration)

'boost
lblItemData(9).Caption = r.BoostVoltage
lblItemData(10).Caption = r.BoostCurrent
lblItemData(11).Caption = Format(r.BoostRailUsed, "0.0")
lblItemData(12).Caption = Format(r.BoostSpeed, "0.00")
If analysisDefine.BoostEnable Then
    updateCueWithCri 12, r.BoostSpeedSucceed
    lblCriData(12).Caption = Format(analysisDefine.BoostMin, "0.00") & " / " & Format(analysisDefine.BoostMax, "0.00")
End If

If analysisDefine.CurrentInterruptEnable Then
    If r.HasCurrentInterruptinBoost Then
        lblItemData(13).Caption = "N"
        updateCue 13, True
    Else
        lblItemData(13).Caption = "Y"
        updateCue 13, False
    End If
End If

If analysisDefine.ShortCircuitEnable Then
    If Not r.HasShortCircuitinBoost Then
        lblItemData(14).Caption = "N"
        updateCue 14, True
    Else
        lblItemData(14).Caption = "Y"
        updateCue 14, False
    End If
End If

lblItemData(15).Caption = CInt(r.BoostDuration)


'upset
lblItemData(16).Caption = Format(r.UpsetRailUsage, "0.00")
If analysisDefine.UpsetEnable Then
    updateCueWithCri 16, r.UpsetRailUsageSucceed
    lblCriData(16).Caption = Format(analysisDefine.UpsetMin, "0.0") & " / " & Format(analysisDefine.UpsetMax, "0.0")
End If
 
 

If analysisDefine.SlippageEnable Then
    If Not r.HasSlippage Then
        lblItemData(17).Caption = "N"
        updateCue 17, True
    Else
        lblItemData(17).Caption = "Y"
        updateCue 17, False
    End If
End If

lblItemData(18).Caption = r.UpsetMaxCurrent
lblItemData(19).Caption = Format(r.UpsetCurrentOnTime, "0.00")
lblItemData(20).Caption = Format(r.UpsetDuration, "0.00")

'Forge
lblItemData(21).Caption = r.ForgeAverageForce
If analysisDefine.ForgeEnable Then
    updateCueWithCri 21, r.ForgeForceSucceed
    lblCriData(21).Caption = Format(analysisDefine.ForgeMin, "0") & " / " & Format(analysisDefine.ForgeMax, "0")
End If

lblItemData(22).Caption = Format(r.ForgeDuration, "0.00")

'Overall
lblItemData(23).Caption = Format(r.OverallImpedance, "0.0")
lblItemData(24).Caption = Format(r.TotalRailUsage, "0.0")
If analysisDefine.TotalRailUsageEnable Then
    updateCue 24, r.TotalRailUsageSucceed
End If

lblItemData(25).Caption = r.HoldingTime
lblItemData(26).Caption = CInt(r.TotalDuration)


End Function


Private Function updateCueWithCri(index As Integer, succeed As Boolean)
    If succeed Then
        lblItem(index).ForeColor = SUCCEED_COLOR
        lblItemData(index).ForeColor = SUCCEED_COLOR
        lblCriData(index).ForeColor = SUCCEED_COLOR
    Else
        lblItem(index).ForeColor = FAIL_COLOR
        lblItemData(index).ForeColor = FAIL_COLOR
        lblCriData(index).ForeColor = FAIL_COLOR
    End If

    lblItem(index).FontBold = True
    lblItemData(index).FontBold = True
    lblCriData(index).FontBold = True
End Function

Private Function updateCue(index As Integer, succeed As Boolean)
    If succeed Then
        lblItem(index).ForeColor = SUCCEED_COLOR
        lblItemData(index).ForeColor = SUCCEED_COLOR
    Else
        lblItem(index).ForeColor = FAIL_COLOR
        lblItemData(index).ForeColor = FAIL_COLOR
    End If

    lblItem(index).FontBold = True
    lblItemData(index).FontBold = True
End Function

Private Sub cmdViewDataDetail_Click()
Me.MousePointer = MousePointerConstants.vbHourglass


'    If dataForm Is Nothing Then
        Call setDetailData(buf)
  '  End If
    dataForm.Show vbModal
Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub MSChart1_DblClick()
    If model = COMMON Then
        model = SMALL
        Call setChartSmall(buf)
    Else
        model = COMMON
        Call setChart(buf)
    End If
End Sub



Private Function setDetailData(EmulateData() As WeldData)
    
Dim sa() As String
ReDim sa(UBound(EmulateData))
Dim i As Integer
Dim data As WeldData

Dim entry As String

For i = 0 To UBound(EmulateData)

    data = EmulateData(i)
    If 0 <= data.WeldStage And data.WeldStage <= 6 Then
        entry = PLCDrv.WeldStages(data.WeldStage)
    Else
        entry = data.WeldStage
    End If
    entry = entry & vbTab & data.PlcStage
    entry = entry & vbTab & Format(data.Dist, "##0.00")
    entry = entry & vbTab & data.Amp
    entry = entry & vbTab & data.Volt
    entry = entry & vbTab & data.PsiUpset
    entry = entry & vbTab & data.PsiOpen
    entry = entry & vbTab & Format(PlcAnalysiser.toForce(data.PsiUpset, data.PsiOpen), "##0")
    entry = entry & vbTab & Format(data.Time, "##0.00")
    sa(i) = entry
        
'    MyData(1, i + 1) = (EmulateData(i).PsiUpset - EmulateData(i).PsiOpen) / 25.4
'    MyData(2, i + 1) = EmulateData(i).Volt
'    MyData(3, i + 1) = EmulateData(i).Amp
'    MyData(4, i + 1) = EmulateData(i).Dist
Next


    Set dataForm = New FrmDataGrid
    Call dataForm.setData(sa)

End Function



