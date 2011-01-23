VERSION 5.00
Begin VB.Form dlgChangePWD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change password"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Tag             =   "25000"
   Begin VB.TextBox txtConfirmNewPwd 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtPWD 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Tag             =   "25050"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Tag             =   "25040"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm new password:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Tag             =   "25030"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "New password:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Tag             =   "25020"
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Old password:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Tag             =   "25010"
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "dlgChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim pass As Boolean

Public Sub clear()
    pass = False
    Me.txtPWD.Text = ""
    Me.txtNewPwd.Text = ""
    Me.txtConfirmNewPwd = ""
    
End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
PlcRes.LoadResFor Me
End Sub

Private Sub OKButton_Click()
Dim oldPwd As String


    oldPwd = GetSetting(App.EXEName, "Setting", "PWD", "123456")
    If oldPwd = Me.txtPWD.Text Then
        If Me.txtNewPwd.Text = Me.txtConfirmNewPwd.Text Then
            Call SaveSetting(App.EXEName, "Setting", "PWD", Me.txtNewPwd.Text)
            pass = True
            Me.Hide
            Exit Sub
        End If
    End If
End Sub
