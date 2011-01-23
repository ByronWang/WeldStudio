VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm WeldMDIForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm"
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
      Caption         =   "&Window"
      Tag             =   "10300"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Tag             =   "10400"
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


Private Sub MDIForm_Load()
' Resource
PlcRes.LoadResFor Me


    PLCDrv.InitPLCConnection
    mnuConnect.Enabled = PLCDrv.beActive
    PLCDrv.UninitPLCConection
    
    
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    PLCDrv.UninitPLCConection
End Sub

Private Sub mnuAbout_Click()
    FrmAbout.Show vbModal, Me
End Sub

Private Sub mnuConnect_Click()
    FrmGraph.WindowState = FormWindowStateConstants.vbMaximized
    Call FrmGraph.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    'CommonDialog1.Filter = "Weld Data File (*.wdd) | *.wdd |Old Data File (*.wld) | *.wld"
    CommonDialog1.Filter = "Old Data File (*.wld) | *.wld"
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
        Dim f As New FrmChart
        f.Load CommonDialog1.filename
        f.Caption = CommonDialog1.filename
        f.Show
    End If
End Sub

Private Sub mnuOptions_Click()
    FrmOption.Show vbModal, Me
End Sub

Private Sub mnuPulseProcess_Click()
    dlgPWD.clear
    dlgPWD.Show vbModal, Me
    
    If dlgPWD.pass Then
        FrmPulseSetting.Show vbModal, Me
    End If
End Sub

Private Sub mnuRegularProcess_Click()
    dlgPWD.clear
    dlgPWD.Show vbModal, Me
    
    If dlgPWD.pass Then
       FrmRegularSetting.Show vbModal, Me
    End If
End Sub

Private Sub mnuStartEmulate_Click()
    PlcDemo.StartDemo
    
End Sub

