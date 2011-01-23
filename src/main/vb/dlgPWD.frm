VERSION 5.00
Begin VB.Form dlgPWD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChangePwd 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtPWD 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlgPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public pass As Boolean


Private Sub CancelButton_Click()
    pass = False
    Me.Hide
End Sub

Private Sub cmdChangePwd_Click()
    dlgChangePwd.Show vbModal, Me
End Sub

Private Sub OKButton_Click()
    pass = True
    Me.Hide
End Sub
