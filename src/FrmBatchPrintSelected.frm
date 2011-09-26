VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   6885
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptChartType 
      Caption         =   "Option2"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.OptionButton OptChartType 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2940
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
