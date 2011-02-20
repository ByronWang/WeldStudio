VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..."
   ClientHeight    =   1680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1920
      Top             =   120
   End
   Begin VB.Frame frmProgress 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Label lblProgress 
         BackColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const seperator As Integer = 400

Dim step As Integer

Private Sub Form_Load()
    step = (frmProgress.Width - seperator) / 50
    lblProgress.Width = 0
End Sub


Private Sub Timer1_Timer()
    If lblProgress.Width < frmProgress.Width - seperator Then
        lblProgress.Width = lblProgress.Width + step
    Else
        Me.Hide
        Unload Me
    End If
End Sub
