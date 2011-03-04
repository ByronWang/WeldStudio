VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDailyReport 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   9135
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Load(filename As String)
 Dim data() As DailyReport
  
data = PlcDailyReport.LoadData(filename)

Dim path As String
path = Left(filename, InStrRev(filename, "\"))


Dim sa() As String
ReDim sa(UBound(data))
Dim f As FileR
Dim WeldFile As String
Dim entry As String
  Dim i As Integer
  For i = LBound(data) To UBound(data)
            WeldFile = CStr(data(i).Sequence)
            WeldFile = path & data(i).Serial & Left("0000", 4 - Len(WeldFile)) & WeldFile & ".wld"
            
            f = PlcWld.LoadData(WeldFile)
            

'   Result
If f.analysisResult.succeed = 1 Then
    entry = "OK"
ElseIf f.analysisResult.succeed = 2 Then
    entry = "NO"
ElseIf f.analysisResult.succeed = 3 Then
    entry = "INT"
Else
    MsgBox "qqqq"
End If

'   Time
entry = entry & vbTab & f.header2.Time
'   Duration
entry = entry & vbTab & f.analysisResult.TotalDuration
'   UPSET
entry = entry & vbTab & f.analysisResult.UpsetRailUsage
'   max.Current
entry = entry & vbTab & f.analysisResult.UpsetMaxCurrent
'   Impedance
entry = entry & vbTab & f.analysisResult.OverallImpedance
'   Rail Usage
entry = entry & vbTab & f.analysisResult.TotalRailUsage
'   FLASH Speed
entry = entry & vbTab & f.analysisResult.FlashSpeed
'   BOOST Speed
entry = entry & vbTab & f.analysisResult.BoostSpeed
'   FORGE force
entry = entry & vbTab & f.analysisResult.ForgeAverageForce
'   Slippage
If f.analysisResult.HasSlippage = 1 Then
    entry = entry & vbTab & "Y"
ElseIf f.analysisResult.HasSlippage = 2 Then
    entry = entry & vbTab & "N"
ElseIf f.analysisResult.HasSlippage = 3 Then
    entry = entry & vbTab & "-"
Else
    entry = entry & vbTab & "-"
End If

'entry = entry & vbTab & f.analysisResult.HasSlippage
'   Weld Program  ---

'   Chainage  ---
sa(i) = entry

  Next i
    
    
   Call setData(sa)
    
    'lblDate.Caption = Trim(fr.header.Date)
    'lblTime.Caption = Trim(fr.header.Time)
    'lblParam.Caption = "UNIT:" & Trim(fr.header.ParamName)

End Sub

Public Function setData(sa() As String)
Dim i As Integer
For i = LBound(sa) To UBound(sa)
    MSFlexGrid1.AddItem sa(i)
Next

i = 0
MSFlexGrid1.RemoveItem (1)
MSFlexGrid1.TextMatrix(0, i) = "Result"
MSFlexGrid1.ColWidth(i) = 800
MSFlexGrid1.ColAlignment(i) = AlignmentSettings.flexAlignCenterCenter
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Time"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Duration"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Upset"
MSFlexGrid1.ColWidth(i) = 800
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Max.Current"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Impedance"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "RailUsage"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "FlashSpeed"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "BoostSpeed"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "ForgeForce"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Slippage"
MSFlexGrid1.ColWidth(i) = 1000
MSFlexGrid1.ColAlignment(i) = AlignmentSettings.flexAlignCenterCenter
i = i + 1
'MSFlexGrid1.TextMatrix(0, i) = "WeldProgram"
'MSFlexGrid1.ColWidth(i) = 1000
'i = i + 1
'MSFlexGrid1.TextMatrix(0, i) = "Chainage"
'MSFlexGrid1.ColWidth(i) = 1000
'i = i + 1

End Function

Private Sub Form_Resize()
    Me.MSFlexGrid1.Width = Me.Width - 120
    Me.MSFlexGrid1.Height = Me.Height - 500 - MSFlexGrid1.Top
    
End Sub


