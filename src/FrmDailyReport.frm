VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDailyReport 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   13560
   WindowState     =   2  'Maximized
   Begin VB.Frame frmSum 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   5295
      Begin VB.Label labelAccepted 
         Alignment       =   1  'Right Justify
         Caption         =   "Accepted welds:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3600
      End
      Begin VB.Label labelRejected 
         Alignment       =   1  'Right Justify
         Caption         =   "Rejected welds:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label labelTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total welds:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   3600
      End
      Begin VB.Label lblAccepted 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblReject 
         Alignment       =   1  'Right Justify
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      SelectionMode   =   1
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   3600
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   3600
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   3600
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3600
   End
End
Attribute VB_Name = "FrmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SUCCEED_COLOR As Long = &HFF00&
Const FAIL_COLOR As Long = &HFF&
Const NOTUSED_COLOR As Long = &HFFFFFF

Public DailyReportFileName As String

Public Sub Load(FileName As String)
Dim data() As DailyReport

Dim acceptedWelds As Integer
Dim rejectedWelds As Integer
Dim totalWelds As Integer


data = PlcDailyReport.LoadData(FileName)
Me.DailyReportFileName = FileName

Dim path As String
path = Left(FileName, InStrRev(FileName, "\"))

Dim sa() As String
ReDim sa(UBound(data))
Dim f As FileR
Dim WeldFile As String
Dim entry As String
Dim cellcolors() As Long
ReDim cellcolors(UBound(data))

  Dim i As Integer
  For i = LBound(data) To UBound(data)
            WeldFile = CStr(data(i).Sequence)
            WeldFile = path & data(i).Serial & Left("0000", 4 - Len(WeldFile)) & WeldFile & ".wld"
            
            f = PlcWld.LoadData(WeldFile)
            


    lblCompany.Caption = Trim(f.header1.CompanyName)
        
    lblDate.Caption = "DailyReport: " & Trim(f.header2.Date)
    
    lblUnit.Caption = "UNIT:" & Trim(f.header1.unitName)
    lblLocation.Caption = "LOCATION:" & Trim(f.header1.Location)
    

Dim WeldNumberDriver As IWeldNumber
Dim displayName As String
    
Select Case f.header2.WeldNumberMode
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

unitName = f.header1.unitName
operator = f.header1.operator
    
displayName = WeldNumberDriver.ToDisplay(CDate(f.header2.Date), WeldNumberDriver.FromCompact(Trim(f.header2.CompactedWeldNumber)))
    
Select Case f.header2.WeldNumberMode
    Case EngMode:
        displayName = Trim(f.header1.unitName) & displayName & Trim(f.header1.operator)
    Case JinanMode:
        displayName = Trim(f.header1.unitName) & displayName
End Select
    
entry = displayName

'   Result
If f.analysisResult.Succeed = PlcDeclare.OK Then
    acceptedWelds = acceptedWelds + 1
    entry = entry & vbTab & "OK"
    cellcolors(i) = SUCCEED_COLOR
ElseIf f.analysisResult.Succeed = PlcDeclare.NO Then
    rejectedWelds = rejectedWelds + 1
    entry = entry & vbTab & "NO"
    cellcolors(i) = FAIL_COLOR
ElseIf f.analysisResult.Succeed = PlcDeclare.INTERRUPT Then
    rejectedWelds = rejectedWelds + 1
    entry = entry & vbTab & "INT"
    cellcolors(i) = FAIL_COLOR
Else
    rejectedWelds = rejectedWelds + 1
    entry = entry & vbTab & " - "
    cellcolors(i) = NOTUSED_COLOR
End If

totalWelds = totalWelds + 1



'   Time
entry = entry & vbTab & f.header2.Time
'   Duration
entry = entry & vbTab & Format(f.analysisResult.TotalDuration, "##0")
'   UPSET
entry = entry & vbTab & Format(f.analysisResult.UpsetRailUsage, "##0.00")
'   max.Current
entry = entry & vbTab & Format(f.analysisResult.UpsetMaxCurrent, "##0")
'   Impedance
entry = entry & vbTab & Format(f.analysisResult.OverallImpedance, "##0.0")
'   Rail Usage
entry = entry & vbTab & Format(f.analysisResult.TotalRailUsage, "##0.0")
'   FLASH Speed
entry = entry & vbTab & Format(f.analysisResult.FlashSpeed, "##0.00")
'   BOOST Speed
entry = entry & vbTab & Format(f.analysisResult.BoostSpeed, "##0.00")
'   FORGE force
entry = entry & vbTab & Format(f.analysisResult.ForgeAverageForce, "##0")
'   Slippage
If f.analysisResult.HasSlippage = 1 Then
    entry = entry & vbTab & "N"
ElseIf f.analysisResult.HasSlippage = 2 Then
    entry = entry & vbTab & "Y"
ElseIf f.analysisResult.HasSlippage = 3 Then
    entry = entry & vbTab & "-"
Else
    entry = entry & vbTab & "-"
End If

Dim paramType As String
Select Case f.header2.ParamSettingMode
    Case "R":
        paramType = "R"
    Case "P":
        paramType = "P"
    Case Else:
        paramType = "P"
End Select

entry = entry & vbTab & paramType & ":" & Trim(f.header2.ParamSettingName)

'entry = entry & vbTab & f.analysisResult.HasSlippage
'   Weld Program  ---

'   Chainage  ---
sa(i) = entry

  Next i
    
    
   Call setData(sa)
   
   Dim j As Integer
   Dim color As Long
   
   
   With MSFlexGrid1
   For i = 1 To .Rows - 1
        .Row = i
        color = cellcolors(i - 1)
        For j = 0 To .Cols - 1
            .Col = j
            .CellBackColor = color
        Next j
        
        
   Next
   End With

    Me.lblAccepted.Caption = acceptedWelds
    Me.lblReject.Caption = rejectedWelds
    Me.lblTotal.Caption = totalWelds
    
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

MSFlexGrid1.TextMatrix(0, i) = "Weld#"
MSFlexGrid1.ColWidth(i) = 1500
MSFlexGrid1.ColAlignment(i) = AlignmentSettings.flexAlignLeftCenter
i = i + 1

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
MSFlexGrid1.ColWidth(i) = 1400
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Impedance"
MSFlexGrid1.ColWidth(i) = 1200
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "RailUsage"
MSFlexGrid1.ColWidth(i) = 1000
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "FlashSpeed"
MSFlexGrid1.ColWidth(i) = 1200
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "BoostSpeed"
MSFlexGrid1.ColWidth(i) = 1200
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "ForgeForce"
MSFlexGrid1.ColWidth(i) = 1200
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Slippage"
MSFlexGrid1.ColWidth(i) = 1000
MSFlexGrid1.ColAlignment(i) = AlignmentSettings.flexAlignCenterCenter
i = i + 1
MSFlexGrid1.TextMatrix(0, i) = "Parameter"
MSFlexGrid1.ColWidth(i) = 1500
MSFlexGrid1.ColAlignment(i) = AlignmentSettings.flexAlignLeftCenter
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
    Me.MSFlexGrid1.Height = Me.Height - 1500 - MSFlexGrid1.Top
    Me.frmSum.Top = Me.Height - frmSum.Height - 700
    
End Sub


