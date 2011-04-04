VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm WeldMDIForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Weld Monitoring Studio"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13890
   Icon            =   "FrmWeldMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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

Private Sub MDIForm_Load()

' Resource
PlcRes.LoadResFor Me

    App.HelpFile = App.path & "\weld.chm "
    
    If GetSetting(App.EXEName, "UserData", "CompanyName", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Unit", "") = "" Or _
        GetSetting(App.EXEName, "UserData", "Location", "") = "" Then
        
        Call mnuOptions_Click
    End If
    
    
End Sub

Private Sub menuUserGuide_Click()
    Shell "hh.exe " & App.path & "\WMS.chm ", vbNormalFocus
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
            Call printChart(fc)
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
            Call printDailyReport(fd)
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


Private Function printChart(fc As FrmChart)

        Printer.Orientation = vbPRORLandscape
                
        fc.MSChart1.EditCopy
        DoEvents   ' may be needed for large datasets
        DoEvents   ' may be needed for large datasets
        Printer.Print " "
        'Printer.Print " ------------------------------- "
        Printer.Print " "
        Printer.PaintPicture Clipboard.GetData(), 3500, 2200
        
        
        Dim i As Integer
        Dim j As Integer
        Dim gSep As Single
        Dim iSep As Single
        Dim gLeft As Integer
        Dim iLeft As Integer
        Dim idLeft As Integer
        
        gLeft = 800
        iLeft = 1100
        idLeft = 3100
        
        gSep = 100
        iSep = 50
        
        Printer.CurrentY = 2300
        
        Dim lTop As Integer
        
        With fc
    
            i = 0
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
            
            For j = 0 To 3
                Printer.FontBold = False
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Call setFrom(.lblItemData(j))
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
            
            i = 1
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
            
            For j = 4 To 8
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Call setFrom(.lblItemData(j))
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
            
            i = 2
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
                            
            For j = 9 To 15
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Call setFrom(.lblItemData(j))
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
            
            
            i = 3
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
            
            For j = 16 To 20
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Call setFrom(.lblItemData(j))
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
            
            
            i = 4
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
            
            For j = 21 To 22
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Printer.Print .lblItemData(j).Caption
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
                            
            
            i = 5
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + lineSep
            
            For j = 23 To 24
                Printer.CurrentX = iLeft
                lTop = Printer.CurrentY
                Call setFrom(.lblItem(j))
                Printer.CurrentY = lTop
                
                Printer.CurrentX = idLeft
                Call setFrom(.lblItemData(j))
                Printer.CurrentY = Printer.CurrentY + iSep
            Next
            
            
            
            
            Call navControl(fc.lblCompany)
            Call navControl(fc.lblParam)
            Call navControl(fc.lblProgram)
            Call navControl(fc.lblDate)
            Call navControl(fc.lblTime)
            
            Call navControl(fc.lblUnit)
            Call navControl(fc.lblLocation)
            
        End With
        
        Printer.EndDoc
End Function

Private Function printDailyReport(f As FrmDailyReport)
Printer.Orientation = vbPRORLandscape
    
Dim x, y As Long
x = 1200
y = 2200

Call navControlForDailyReport(f.lblCompany)
Call navControlForDailyReport(f.lblDate)
Call navControlForDailyReport(f.lblLocation)
Call navControlForDailyReport(f.lblUnit)


For j = 0 To f.MSFlexGrid1.Cols - 1

    Printer.CurrentY = y
    
    For i = 0 To 0
        Printer.CurrentX = x
        Printer.Print f.MSFlexGrid1.TextMatrix(i, j)
    Next i
    
    Printer.CurrentY = Printer.CurrentY + 200
    
    
    For i = 1 To f.MSFlexGrid1.Rows - 1
        Printer.CurrentY = Printer.CurrentY + 100
        Printer.CurrentX = x + 100
        Printer.Print f.MSFlexGrid1.TextMatrix(i, j)
    Next i
    x = x + f.MSFlexGrid1.ColWidth(j) * 1.2
Next j
Printer.EndDoc

    
End Function

Private Function navControlForDailyReport(con As Label)
    Printer.CurrentX = con.Left + 2000
    Printer.CurrentY = con.Top + 1100
    
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    Printer.Print con.Caption
End Function



Private Function navControl(con As Label)
    Printer.CurrentX = con.Left
    Printer.CurrentY = con.Top + 1100
    
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    Printer.Print con.Caption
End Function

Private Function setFrom(con As Control)
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    Printer.Print con.Caption
End Function


Private Sub mnuShutdown_Click()
    Shell "Shutdown -s -f -t 1"
End Sub

Private Sub mnuAbout_Click()
    FrmAbout.Show vbModal, Me
End Sub

Private Sub mnuConnect_Click()
On Error GoTo ERROR_HANDLE

    frmProgress.LoadMode = PlcDeclare.LOAD_ALL_PARAMETER
    frmProgress.ParamName = name
    frmProgress.Show vbModal, Me
    If frmProgress.Status = 0 Then
        Call FrmGraph.Show
    End If
    
Exit Sub
ERROR_HANDLE:
    MsgBox PlcRes.LoadMsgResString(99000 + Err.Number) & vbCrLf & PLCDrv.g_Error_String, vbCritical
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    'CommonDialog1.Filter = "Weld Data File (*.wdd) | *.wdd |Old Data File (*.wld) | *.wld"
    CommonDialog1.Filter = "Data File (*.WLD) |*.wld|Daily Report(*.DLY)|*.DLY"
    CommonDialog1.InitDir = ".\data\"
    
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" And UCase(Right(CommonDialog1.filename, 4)) = ".WLD" Then
        Dim f As New FrmChart
        f.Load CommonDialog1.filename
        f.Caption = CommonDialog1.filename
        f.Show
    ElseIf CommonDialog1.filename <> "" And UCase(Right(CommonDialog1.filename, 4)) = ".DLY" Then
        Dim frmDR  As New FrmDailyReport
        frmDR.Load CommonDialog1.filename
        frmDR.Caption = CommonDialog1.filename
        frmDR.Show
    End If
End Sub

Private Sub mnuOptions_Click()
    FrmOption.Show vbModal, Me
End Sub

Private Sub mnuPulseProcess_Click()
Dim fpwd As New FrmPWD
    fpwd.clear
    fpwd.Show vbModal, Me
    
    If fpwd.pass Then
        FrmPulseSetting.Show , Me
    End If
End Sub

Private Sub mnuRegularProcess_Click()
Dim fpwd As New FrmPWD
    fpwd.clear
    fpwd.Show vbModal, Me
    
    If fpwd.pass Then
       FrmRegularSetting.Show , Me
    End If
End Sub

Private Sub mnuStartEmulate_Click()
    PlcDemo.StartDemo
End Sub

