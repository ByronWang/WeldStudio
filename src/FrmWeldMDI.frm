VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm WeldMDIForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Weld Monitoring Studio"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11400
   Icon            =   "FrmWeldMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Tag             =   "10000"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Tag             =   "10100"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         HelpContextID   =   1000
         Tag             =   "10110"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Tag             =   "10120"
      End
      Begin VB.Menu mnuParameters 
         Caption         =   "&Parameters"
         Tag             =   "10130"
         Begin VB.Menu mnuRegularProcess 
            Caption         =   "&Regular Process"
            HelpContextID   =   1000
            Tag             =   "10131"
         End
         Begin VB.Menu mnuPulseProcess 
            Caption         =   "&Pulse Process"
            Tag             =   "10134"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuBatchPrint 
         Caption         =   "Batch Print"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Tag             =   "10140"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tool"
      Tag             =   "10200"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Option"
         Tag             =   "10250"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Chart"
      Tag             =   "10300"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Tag             =   "10500"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shut&down"
         Tag             =   "10510"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Tag             =   "10400"
      Begin VB.Menu menuUserGuide 
         Caption         =   "&User's Guide"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Tag             =   "10410"
      End
   End
End
Attribute VB_Name = "WeldMDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Htmlhelp Lib "hhctrl.ocx " Alias "HtmlHelpA " (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long

Dim fGraph As FrmGraph

Private Sub MDIForm_Load()

' Resource
PlcRes.LoadResFor Me

    App.HelpFile = App.path & "\weld.chm "
    
    PLCDrv.InitSystem
    
    If ReadOnly Then
        Me.mnuConnect.Enabled = False
        'Me.mnuTools.Enabled = False
        Me.mnuParameters.Enabled = False
        Exit Sub
    End If
    
    If GetSetting(App.EXEName, "UserData", "CompanyName", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Unit", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Location", "") = "" Then
        
        Call mnuOptions_Click
    End If
    
    If GetSetting(App.EXEName, "UserData", "CompanyName", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Unit", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Location", "") = "" Then
        Unload Me
    End If
    
    
    Dim onlineStartUp As Integer
    
    onlineStartUp = GetSetting(App.EXEName, "General", "OnlineOnStartUp", 0)
    If onlineStartUp = 1 Then
        mnuConnect_Click
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub menuUserGuide_Click()
    Shell "hh.exe " & App.path & "\WMS.chm ", vbNormalFocus
End Sub

Private Sub mnuBatchPrint_Click()
    Dim dlg As New DlgPrinterPrepare
    dlg.Show vbModal, Me
    Unload dlg
End Sub

Private Sub mnuFile_click()
    If WeldMDIForm.ActiveForm Is Nothing Then
        mnuPrint.Enabled = False
    ElseIf TypeOf WeldMDIForm.ActiveForm Is FrmChart Or TypeOf WeldMDIForm.ActiveForm Is FrmDailyReport Then
        mnuPrint.Enabled = True
    Else
        mnuPrint.Enabled = False
    End If
End Sub

Private Sub mnuPrint_Click()
    Dim f As Form
    
    Set f = WeldMDIForm.ActiveForm
    If f Is Nothing Then
        Exit Sub
    ElseIf TypeOf f Is FrmChart Then
        Me.CommonDialog1.PrinterDefault = True
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlPDNoPageNums
        
        On Error Resume Next
        Me.CommonDialog1.ShowPrinter
        If Err.Number = cdlCancel Then
            Exit Sub
        End If
        On Error GoTo 0
        DoEvents
        
        Dim fc As FrmChart
        Set fc = f
        For i = 1 To CommonDialog1.Copies
            Call PrintChart(fc)
            Call PrintGraph(Printer, fc.weldFileName, fc.model = COMMON)
            Printer.EndDoc
        Next i
    
    ElseIf TypeOf f Is FrmDailyReport Then
        Me.CommonDialog1.PrinterDefault = True
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlPDNoPageNums
        
        On Error Resume Next
        Me.CommonDialog1.ShowPrinter
        If Err.Number = cdlCancel Then
            Exit Sub
        End If
        On Error GoTo 0
        DoEvents
        
        Dim fd As FrmDailyReport
        Set fd = f
        For i = 1 To CommonDialog1.Copies
            Call PrintDailyReport(fd)
            Printer.EndDoc
        Next i
    End If
    
    Exit Sub
ERRORHANDLE:
Select Case Err.Number
Case cdlCancel
'User clicked Cancel button on Print dialog box
Case Else
MsgBox Err.Description
End Select

End Sub




Private Sub mnuShutdown_Click()
    Shell "Shutdown -s -f -t 1"
End Sub

Private Sub mnuAbout_Click()
    FrmAbout.Show vbModal, Me
End Sub

Private Sub mnuConnect_Click()

    frmProgress.LoadMode = PlcDeclare.LOAD_ALL_PARAMETER
    frmProgress.ParamName = name
    frmProgress.Show vbModal, Me
    
    If Not fGraph Is Nothing Then
        Unload fGraph
        Set fGraph = Nothing
    End If
    
    If frmProgress.status = 0 Then
        Set fGraph = New FrmGraph
        Load fGraph
        fGraph.alive = True
        fGraph.Show
    End If
    
'    Dim i As Integer
'    Dim f As Form
'    For i = LBound(Forms) To UBound(Forms)
'        Set f = Forms(i)
'    End If
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    CommonDialog1.Filter = "Data File (*.WLD) |*.wld|Daily Report(*.DLY)|*.DLY"
    CommonDialog1.InitDir = ".\data\"
    
    Dim i As Integer
    Dim f As Form
    Dim fc As FrmChart
    Dim fd As FrmDailyReport
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" And UCase(Right(CommonDialog1.FileName, 4)) = ".WLD" Then
        
        For i = 0 To Forms.count - 1
            Set f = Forms(i)
            If TypeOf f Is FrmChart Then
                Set fc = f
                If UCase(fc.weldFileName) = UCase(CommonDialog1.FileName) Then
                    fc.SetFocus
                    Exit Sub
                End If
            End If
        Next i
                
        Set fc = New FrmChart
        fc.Load CommonDialog1.FileName
        fc.Caption = CommonDialog1.FileName
        fc.Show
    ElseIf CommonDialog1.FileName <> "" And UCase(Right(CommonDialog1.FileName, 4)) = ".DLY" Then
        For i = 0 To Forms.count - 1
            Set f = Forms(i)
            If TypeOf f Is FrmDailyReport Then
                Set fd = f
                If UCase(fd.DailyReportFileName) = UCase(CommonDialog1.FileName) Then
                    fd.SetFocus
                    Exit Sub
                End If
            End If
        Next i
        
        Set fd = New FrmDailyReport
        fd.Load CommonDialog1.FileName
        fd.Caption = CommonDialog1.FileName
        fd.Show
    End If
End Sub

Private Sub mnuOptions_Click()
    FrmOption.Show vbModal, Me
    PLCDrv.InitSystem
End Sub

Private Sub mnuPulseProcess_Click()
Dim fpwd As New FrmPWD
    fpwd.clear
    fpwd.Show vbModal, Me
    
    If fpwd.pass Then
        Dim frm As New FrmPulseSetting
        frm.Show vbModal, Me
    End If
End Sub

Private Sub mnuRegularProcess_Click()
Dim fpwd As New FrmPWD
    fpwd.clear
    fpwd.Show vbModal, Me
    
    If fpwd.pass Then
        Dim frm As New FrmRegularSetting
        frm.Show vbModal, Me
    End If
End Sub

Private Sub mnuStartEmulate_Click()
    PlcDemo.StartDemo
End Sub

