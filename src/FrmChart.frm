VERSION 5.00
Begin VB.Form FrmChart 
   Caption         =   "Form1"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "FrmChart.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Tag             =   "11000"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DrawWidth       =   6
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   4440
      ScaleHeight     =   8535
      ScaleWidth      =   10695
      TabIndex        =   75
      Top             =   1440
      Width           =   10695
   End
   Begin VB.CommandButton cmdShowMode 
      Caption         =   "Data Filter"
      Height          =   495
      Left            =   13680
      TabIndex        =   74
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewDataDetail 
      Caption         =   "View Detail"
      Height          =   495
      Left            =   13680
      TabIndex        =   73
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   240
      TabIndex        =   7
      Tag             =   "11100"
      Top             =   360
      Width           =   4455
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
         Left            =   0
         TabIndex        =   72
         Tag             =   "100"
         Top             =   120
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
         Left            =   120
         TabIndex        =   71
         Tag             =   "110"
         Top             =   360
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
         Left            =   120
         TabIndex        =   70
         Tag             =   "120"
         Top             =   600
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
         Left            =   120
         TabIndex        =   69
         Tag             =   "130"
         Top             =   840
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
         Left            =   120
         TabIndex        =   68
         Tag             =   "140"
         Top             =   1080
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
         Left            =   2400
         TabIndex        =   67
         Top             =   360
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
         Left            =   2400
         TabIndex        =   66
         Top             =   600
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
         Left            =   2400
         TabIndex        =   65
         Top             =   840
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
         Left            =   2400
         TabIndex        =   64
         Top             =   1080
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
         Left            =   0
         TabIndex        =   63
         Tag             =   "200"
         Top             =   1440
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
         Left            =   120
         TabIndex        =   62
         Tag             =   "210"
         Top             =   1680
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
         Left            =   120
         TabIndex        =   61
         Tag             =   "220"
         Top             =   1920
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
         Left            =   120
         TabIndex        =   60
         Tag             =   "230"
         Top             =   2160
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
         Left            =   120
         TabIndex        =   59
         Tag             =   "240"
         Top             =   2400
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
         Left            =   2400
         TabIndex        =   58
         Top             =   1680
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   57
         Top             =   1920
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
         Left            =   2400
         TabIndex        =   56
         Top             =   2160
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
         Left            =   2400
         TabIndex        =   55
         Top             =   2400
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
         Left            =   120
         TabIndex        =   54
         Tag             =   "250"
         Top             =   2640
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
         Left            =   2400
         TabIndex        =   53
         Top             =   2640
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
         Left            =   0
         TabIndex        =   52
         Tag             =   "300"
         Top             =   3000
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
         Left            =   120
         TabIndex        =   51
         Tag             =   "310"
         Top             =   3240
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
         Left            =   120
         TabIndex        =   50
         Tag             =   "320"
         Top             =   3480
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
         Left            =   120
         TabIndex        =   49
         Tag             =   "330"
         Top             =   3720
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
         Left            =   120
         TabIndex        =   48
         Tag             =   "340"
         Top             =   3960
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
         Left            =   2400
         TabIndex        =   47
         Top             =   3240
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
         Left            =   2400
         TabIndex        =   46
         Top             =   3480
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
         Left            =   2400
         TabIndex        =   45
         Top             =   3720
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
         Left            =   2400
         TabIndex        =   44
         Top             =   3960
         Width           =   795
      End
      Begin VB.Label lblItem 
         Caption         =   "Current Interrupt(Y/N):"
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
         Left            =   120
         TabIndex        =   43
         Tag             =   "350"
         Top             =   4200
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
         Left            =   2400
         TabIndex        =   42
         Top             =   4200
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
         Left            =   120
         TabIndex        =   41
         Tag             =   "360"
         Top             =   4440
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
         Left            =   2400
         TabIndex        =   40
         Top             =   4440
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
         Left            =   120
         TabIndex        =   39
         Tag             =   "370"
         Top             =   4680
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
         Left            =   2400
         TabIndex        =   38
         Top             =   4680
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
         Left            =   0
         TabIndex        =   37
         Tag             =   "400"
         Top             =   5040
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
         Left            =   120
         TabIndex        =   36
         Tag             =   "410"
         Top             =   5280
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
         Left            =   120
         TabIndex        =   35
         Tag             =   "420"
         Top             =   5520
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
         Left            =   120
         TabIndex        =   34
         Tag             =   "430"
         Top             =   5760
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
         Left            =   120
         TabIndex        =   33
         Tag             =   "440"
         Top             =   6000
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
         Left            =   2400
         TabIndex        =   32
         Top             =   5280
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
         Left            =   2400
         TabIndex        =   31
         Top             =   5520
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
         Left            =   2400
         TabIndex        =   30
         Top             =   5760
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
         Left            =   2400
         TabIndex        =   29
         Top             =   6000
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
         Left            =   120
         TabIndex        =   28
         Tag             =   "450"
         Top             =   6240
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
         Left            =   2400
         TabIndex        =   27
         Top             =   6240
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
         Left            =   0
         TabIndex        =   26
         Tag             =   "500"
         Top             =   6600
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
         Left            =   120
         TabIndex        =   25
         Tag             =   "510"
         Top             =   6840
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
         Left            =   120
         TabIndex        =   24
         Tag             =   "520"
         Top             =   7080
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
         Left            =   2400
         TabIndex        =   23
         Top             =   6840
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
         Left            =   2400
         TabIndex        =   22
         Top             =   7080
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
         Left            =   0
         TabIndex        =   21
         Tag             =   "600"
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label lblItem 
         Caption         =   "Impedance(Ohm):"
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
         Left            =   120
         TabIndex        =   20
         Tag             =   "610"
         Top             =   7680
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
         Left            =   120
         TabIndex        =   19
         Tag             =   "620"
         Top             =   7920
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
         Left            =   120
         TabIndex        =   18
         Tag             =   "630"
         Top             =   8160
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
         Left            =   120
         TabIndex        =   17
         Tag             =   "640"
         Top             =   8400
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
         Left            =   2400
         TabIndex        =   16
         Top             =   7680
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
         Left            =   2400
         TabIndex        =   15
         Top             =   7920
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
         Left            =   2400
         TabIndex        =   14
         Top             =   8160
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
         Left            =   2400
         TabIndex        =   13
         Top             =   8400
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
         Left            =   3240
         TabIndex        =   12
         Top             =   2400
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
         Left            =   3240
         TabIndex        =   11
         Top             =   3960
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
         Left            =   3240
         TabIndex        =   10
         Top             =   5280
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
         Left            =   3240
         TabIndex        =   9
         Top             =   6840
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
         Left            =   3240
         TabIndex        =   8
         Tag             =   "5"
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Label lblTime 
      Caption         =   "19:12:54"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   840
      Width           =   3600
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Caption         =   "2011-01-01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   840
      Width           =   3600
   End
   Begin VB.Label lblProgram 
      Alignment       =   2  'Center
      Caption         =   "PULSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   3600
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      Caption         =   "YARDWAY LTD."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   3600
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "K0035 - OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      Top             =   360
      Width           =   3600
   End
   Begin VB.Label lblLocation 
      Caption         =   "LOCATION:CRETE ILL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   10320
      TabIndex        =   1
      Top             =   1080
      Width           =   3600
   End
   Begin VB.Label lblUnit 
      Alignment       =   1  'Right Justify
      Caption         =   "UNIT:K922SN99-U101136(CW632)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   1080
      Width           =   3600
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
    PURE
    SMALLCOMMON
    SMALLPURE
End Enum

Public model As ModelConstants

Dim dataForm As FrmDataGrid


Dim buf() As WeldData

Const SUCCEED_COLOR As Long = &HC000&
Const FAIL_COLOR As Long = &HFF&
Const NOTUSED_COLOR As Long = &H80000012

Dim lastTime As Long

Public weldFileName As String

Dim fr As FileR

Dim lastX As Integer
Dim lastY As Integer
Dim inArea As Boolean
Const MASK_COLOR As Long = vbWhite

Public Sub Load(FileName As String)
' Resource
PlcRes.LoadResFor Me
    fr = PlcWld.LoadData(FileName)
    weldFileName = FileName
    
    model = COMMON

    Dim i As Integer

    ReDim buf(fr.header2.RecordCount - 1)

    Call setChart(fr, COMMON)
    
    lblCompany.Caption = Trim(fr.header1.CompanyName)
    
    Dim WeldNumberDriver As IWeldNumber
    Dim displayName As String
        
    Select Case fr.header2.WeldNumberMode
        Case GeneralMode:
            Set WeldNumberDriver = New GeneralWeldNumber
        Case EngMode:
            Set WeldNumberDriver = New EngWeldNumber
        Case JinanMode:
            Set WeldNumberDriver = New JinanWeldNumber
        Case Else:
            Set WeldNumberDriver = New GeneralWeldNumber
    End Select
    
    Dim unitName As String
    Dim operator As String
    
    unitName = fr.header1.unitName
    operator = fr.header1.operator
        
    displayName = WeldNumberDriver.ToDisplay(CDate(fr.header2.Date), WeldNumberDriver.FromCompact(Trim(fr.header2.CompactedWeldNumber)))
    
    Select Case fr.header2.WeldNumberMode
        Case EngMode:
            displayName = Trim(fr.header1.unitName) & displayName & Trim(fr.header1.operator)
        Case JinanMode:
            displayName = Trim(fr.header1.unitName) & displayName
    End Select
        
    If fr.analysisResult.Succeed = OK Then
        lblParam.Caption = displayName & "-OK"
    ElseIf fr.analysisResult.Succeed = NO Then
        lblParam.Caption = displayName & "-NO"
    Else
        lblParam.Caption = displayName & "-INT"
    End If
    
    Dim paramType As String
    Select Case fr.header2.ParamSettingMode
        Case "R":
            paramType = "R"
        Case "P":
            paramType = "P"
        Case Else:
            paramType = "P"
    End Select
    
    lblProgram.Caption = paramType & ":" & Trim(fr.header2.ParamSettingName)
    
    lblDate.Caption = Trim(fr.header2.Date)
    lblTime.Caption = Trim(fr.header2.Time)
    
    lblUnit.Caption = "UNIT:" & Trim(fr.header1.unitName)
    lblLocation.Caption = "LOCATION:" & Trim(fr.header1.Location)
    
    
    updateCueControl lblCompany, fr.analysisResult.Succeed
    updateCueControl lblParam, fr.analysisResult.Succeed
    updateCueControl lblProgram, fr.analysisResult.Succeed
    updateCueControl lblDate, fr.analysisResult.Succeed
    updateCueControl lblTime, fr.analysisResult.Succeed
    updateCueControl lblUnit, fr.analysisResult.Succeed
    updateCueControl lblLocation, fr.analysisResult.Succeed
    
    Call anaylize(fr.analysisDefine, fr.analysisResult)

End Sub

Private Function setChartSmall(fr As FileR)
    setChart fr, SMALLCOMMON
End Function

Private Function setChart(fr As FileR, model As ModelConstants)
    Dim prnt As PictureBox
    Set prnt = P
    
    Dim i As Integer
    prnt.Cls

    Select Case model
    Case ModelConstants.COMMON:
        buf = LoadDataAll(fr.data, fr.header2.RecordCount)
        Call PrepareDraw(P, 600, P.height - 400, P.width - 1200, -(P.height - 600), buf(0).Time)
        DrawChartAll P, buf, fr.analysisDefine
        
    Case ModelConstants.PURE:
        buf = LoadDataAllThin(fr.data, fr.header2.RecordCount)
        Call PrepareDraw(P, 600, P.height - 400, P.width - 1200, -(P.height - 600), buf(0).Time)
        DrawChartAll P, buf, fr.analysisDefine
        
    Case ModelConstants.SMALLCOMMON, ModelConstants.SMALLPURE:
        buf = LoadDataUpset(fr.data, fr.header2.RecordCount)
        Call PrepareDraw(P, 600, P.height - 400, P.width - 1200, -(P.height - 600), buf(0).Time)
        DrawChartUpset P, buf, fr.analysisDefine
        
    End Select
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
    If r.HasCurrentInterruptinBoost = OK Then
        lblItemData(13).Caption = "N"
        updateCue 13, r.HasCurrentInterruptinBoost
    ElseIf r.HasCurrentInterruptinBoost = NO Then
        lblItemData(13).Caption = "Y"
        updateCue 13, r.HasCurrentInterruptinBoost
    Else
        lblItemData(13).Caption = "-"
        updateCue 13, r.HasCurrentInterruptinBoost
    End If
End If

If analysisDefine.ShortCircuitEnable Then
    If r.HasShortCircuitinBoost = OK Then
        lblItemData(14).Caption = "N"
        updateCue 14, r.HasShortCircuitinBoost
    ElseIf r.HasShortCircuitinBoost = NO Then
        lblItemData(14).Caption = "Y"
        updateCue 14, r.HasShortCircuitinBoost
    Else
        lblItemData(14).Caption = "-"
        updateCue 14, r.HasShortCircuitinBoost
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
    If r.HasSlippage = OK Then
        lblItemData(17).Caption = "N"
        updateCue 17, r.HasSlippage
    ElseIf r.HasSlippage = NO Then
        lblItemData(17).Caption = "Y"
        updateCue 17, r.HasSlippage
    Else
        lblItemData(17).Caption = "-"
        updateCue 17, r.HasSlippage
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


Private Function updateCueWithCri(index As Integer, Succeed As Integer)
    If Succeed = OK Then
        lblItem(index).ForeColor = SUCCEED_COLOR
        lblItemData(index).ForeColor = SUCCEED_COLOR
        lblCriData(index).ForeColor = SUCCEED_COLOR
        lblItem(index).FontBold = True
        lblItemData(index).FontBold = True
        lblCriData(index).FontBold = True
    ElseIf Succeed = NO Then
        lblItem(index).ForeColor = FAIL_COLOR
        lblItemData(index).ForeColor = FAIL_COLOR
        lblCriData(index).ForeColor = FAIL_COLOR
        lblItem(index).FontBold = True
        lblItemData(index).FontBold = True
        lblCriData(index).FontBold = True
    Else
        lblItem(index).ForeColor = NOTUSED_COLOR
        lblItemData(index).ForeColor = NOTUSED_COLOR
        lblCriData(index).ForeColor = NOTUSED_COLOR
        lblItem(index).FontBold = False
        lblItemData(index).FontBold = False
        lblCriData(index).FontBold = False
    End If

End Function

Private Function updateCue(index As Integer, Succeed As Integer)
    If Succeed = OK Then
        lblItem(index).ForeColor = SUCCEED_COLOR
        lblItemData(index).ForeColor = SUCCEED_COLOR
        lblItem(index).FontBold = True
        lblItemData(index).FontBold = True
    ElseIf Succeed = NO Then
        lblItem(index).ForeColor = FAIL_COLOR
        lblItemData(index).ForeColor = FAIL_COLOR
        lblItem(index).FontBold = True
        lblItemData(index).FontBold = True
    Else
        lblItem(index).ForeColor = NOTUSED_COLOR
        lblItemData(index).ForeColor = NOTUSED_COLOR
        lblItem(index).FontBold = False
        lblItemData(index).FontBold = False
    End If

End Function
Private Function updateCueControl(con As Control, Succeed As Integer)
    If Succeed = OK Then
        con.ForeColor = SUCCEED_COLOR
        con.FontBold = True
    ElseIf Succeed = NO Then
        con.ForeColor = FAIL_COLOR
        con.FontBold = True
    Else
        con.ForeColor = NOTUSED_COLOR
        con.FontBold = False
    End If

End Function

Private Sub cmdShowMode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = MousePointerConstants.vbHourglass

End Sub

Private Sub cmdViewDataDetail_Click()
Me.MousePointer = MousePointerConstants.vbHourglass

'    If dataForm Is Nothing Then
        Call setDetailData(buf)
  '  End If
    dataForm.Show vbModal
Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub Form_Load()
    P.AutoRedraw = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub cmdShowMode_Click()
    Me.MousePointer = VtMousePointer.VtMousePointerHourGlass
    If model = COMMON Then
        model = PURE
        Call setChart(fr, model)
    ElseIf model = PURE Then
        model = COMMON
        Call setChart(fr, model)
    End If
    Me.MousePointer = VtMousePointer.VtMousePointerArrow
End Sub

Private Sub P_DblClick()

    If model = COMMON Then
        model = SMALLCOMMON
        Call setChartSmall(fr)
        cmdShowMode.Enabled = False
    ElseIf model = PURE Then
        model = SMALLPURE
        Call setChartSmall(fr)
        cmdShowMode.Enabled = False
    ElseIf model = SMALLCOMMON Then
        model = COMMON
        Call setChart(fr, model)
        cmdShowMode.Enabled = True
    
    ElseIf model = SMALLPURE Then
        model = PURE
        Call setChart(fr, model)
        cmdShowMode.Enabled = True
    End If
End Sub

Private Function setDetailData(EmulateData() As WeldData)
    
Dim sa() As String
ReDim sa(UBound(EmulateData))
Dim i As Integer
Dim data As WeldData

Dim entry As String

entry = ""
For i = 0 To UBound(EmulateData)

    data = EmulateData(i)
    
    Dim ws As Long
    ws = data.WeldStage
    
    If 1 < data.PlcStage And data.PlcStage <= 12 Then
        If fr.header2.ParamSettingMode = "R" Then
            If PLCDrv.RegulareStageMapping(data.PlcStage) > ws Then
                ws = PLCDrv.RegulareStageMapping(data.PlcStage)
            End If
        Else
            If PLCDrv.PulseStageMapping(data.PlcStage) < ws Then
                ws = PLCDrv.PulseStageMapping(data.PlcStage)
            End If
            
            If PLCDrv.PulseStageMapping(data.PlcStage) <> data.WeldStage Then
                Debug.Print data.PlcStage & " " & PLCDrv.PulseStageMapping(data.PlcStage) & " - " & data.WeldStage
            End If
        End If
    End If
    
    If 0 <= ws And ws <= 6 Then
        entry = PLCDrv.WeldStages(ws)
    Else
        entry = ws
    End If
    entry = entry & vbTab & data.PlcStage
    entry = entry & vbTab & Format(data.Dist, "##0.00")
    entry = entry & vbTab & data.Amp
    entry = entry & vbTab & data.Volt
    entry = entry & vbTab & data.PsiUpset
    entry = entry & vbTab & data.PsiOpen
    entry = entry & vbTab & Format(PlcAnalysiser.toForce(data.PsiUpset, data.PsiOpen, fr.analysisDefine), "##0")
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

Private Sub P_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    P.DrawMode = 7
    P.DrawStyle = 0
    
    If inArea Then
        P.Line (0, lastY)-(P.ScaleWidth, lastY), MASK_COLOR
        P.Line (lastX, 0)-(lastX, P.ScaleHeight), MASK_COLOR
        inArea = False
    End If
    
    If Button = vbRightButton Then
        lastX = x
        lastY = y
        P.Line (0, lastY)-(P.ScaleWidth, lastY), MASK_COLOR
        P.Line (lastX, 0)-(lastX, P.ScaleHeight), MASK_COLOR
        inArea = True
    End If
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    P.DrawMode = 7
    P.DrawStyle = 0
    
    If inArea Then
        P.Line (0, lastY)-(P.ScaleWidth, lastY), MASK_COLOR
        P.Line (lastX, 0)-(lastX, P.ScaleHeight), MASK_COLOR
        inArea = False
    End If

    If Button = vbRightButton Then
        lastX = x
        lastY = y
        P.Line (0, lastY)-(P.ScaleWidth, lastY), MASK_COLOR
        P.Line (lastX, 0)-(lastX, P.ScaleHeight), MASK_COLOR
        inArea = True
    End If
End Sub

Private Sub P_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    P.DrawMode = 7
    P.DrawStyle = 0
    
    If inArea Then
        P.Line (0, lastY)-(P.ScaleWidth, lastY), MASK_COLOR
        P.Line (lastX, 0)-(lastX, P.ScaleHeight), MASK_COLOR
        inArea = False
    End If
End Sub
