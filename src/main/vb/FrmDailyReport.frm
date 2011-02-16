VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDailyReport 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Data From"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function setData(sa() As String)
Dim i As Integer
For i = 0 To UBound(sa)
    MSFlexGrid1.AddItem sa(i)
Next
MSFlexGrid1.RemoveItem (1)

MSFlexGrid1.TextMatrix(0, 0) = "STAGE"
MSFlexGrid1.ColWidth(0) = 1200
MSFlexGrid1.TextMatrix(0, 1) = "PLC STAGE"
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.TextMatrix(0, 2) = "DIST"
MSFlexGrid1.ColWidth(2) = 700
MSFlexGrid1.TextMatrix(0, 3) = "AMP"
MSFlexGrid1.ColWidth(3) = 600
MSFlexGrid1.TextMatrix(0, 4) = "VOLT"
MSFlexGrid1.ColWidth(4) = 600
MSFlexGrid1.TextMatrix(0, 5) = "PSI(Upset)"
MSFlexGrid1.ColWidth(5) = 1100
MSFlexGrid1.TextMatrix(0, 6) = "PSI(Open)"
MSFlexGrid1.ColWidth(6) = 1000
MSFlexGrid1.TextMatrix(0, 7) = "FORCE"
MSFlexGrid1.ColWidth(7) = 700
MSFlexGrid1.TextMatrix(0, 8) = "TIMER"
MSFlexGrid1.ColWidth(8) = 800
'
'
'WeldSstage 0-init, 1-preflash 2-flash 3-boost 4-upset 5-forge 6-shear
'PLC stage
'DIST scaled reading in mm * 100
'AMP scaled reading in A
'VOLT scaled reading in V
'PSI scaled reading in psi
'PSI2 scaled reading in psi
'Force = (PSI -PSI2) / 25.42    注：25.27~25.47 具体数值不清楚
'
'

'With MSFlexGrid1
'.AllowBigSelection = True
'For i = 0 To .Rows - 1
'.Row = i: .Col = .FixedCols
'.ColSel = .Cols() - .FixedCols - 1
'If i Mod 2 = 0 Then
'.CellBackColor = &HC0C0C0
'Else
'.CellBackColor = vbBlue
'End If
'Next i
'End With
End Function

Private Sub Form_Resize()
    Me.MSFlexGrid1.Width = Me.Width - 120
    Me.MSFlexGrid1.Height = Me.Height - 380 - MSFlexGrid1.Top
    
End Sub

