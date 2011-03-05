VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   8670
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   7695
      Left            =   240
      ScaleHeight     =   7635
      ScaleMode       =   0  'User
      ScaleWidth      =   2.53143e5
      TabIndex        =   2
      Top             =   1200
      Width           =   20000
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   4935
         Left            =   960
         OleObjectBlob   =   "frmPrint.frx":0000
         TabIndex        =   3
         Top             =   840
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         Height          =   6495
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   6855
      End
      Begin VB.Line Line1 
         X1              =   12190.48
         X2              =   68571.43
         Y1              =   1200
         Y2              =   6240
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
    Dim f As Form
    
    Set f = WeldMDIForm.ActiveForm
    f.PrintForm
    
End Sub
