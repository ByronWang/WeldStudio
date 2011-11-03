VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DlgPrinterPrepare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Print"
   ClientHeight    =   7005
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton OptMode 
      Caption         =   "Upset Area of Weld"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.OptionButton OptMode 
      Caption         =   "Full Weld Cycle"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   5295
   End
   Begin VB.ComboBox CboPrinter 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox LstFiles 
      Height          =   2985
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Print"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "File List:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Printer:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "DlgPrinterPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim printing As Boolean

Private Sub CancelButton_Click()
    If printing Then
        printing = False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ERROR_HANDLER
    Dim path As String
    CommonDialog1.Filter = "Data File (*.WLD,*.DLY) |*.wld; *.dly"
    
    If Me.CommonDialog1.FileName <> "" Then
        path = left(Me.CommonDialog1.FileName, InStrRev(Me.CommonDialog1.FileName, "\"))
        Me.CommonDialog1.InitDir = path
    Else
        Me.CommonDialog1.InitDir = "./data/"
    End If
    Me.CommonDialog1.flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNFileMustExist
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    
    
    Dim fs() As String
    fs = Split(Me.CommonDialog1.FileName, " ")
    
    Dim fname As String
    Dim i As Integer
    
    If UBound(fs) - LBound(fs) > 1 Then
        path = fs(LBound(fs))
        For i = LBound(fs) + 1 To UBound(fs)
            fname = path & fs(i)
            addFile fname
        Next
    Else
        fname = fs(LBound(fs))
        addFile fname
    End If
Exit Sub
ERROR_HANDLER:
    
End Sub

Private Sub addFile(fname As String)
    Dim i As Integer
    
    For i = 0 To LstFiles.ListCount - 1
        If LstFiles.List(i) = fname Then
            LstFiles.ListIndex = i
            Exit Sub
        End If
    Next
    
    LstFiles.AddItem fname
    LstFiles.ListIndex = LstFiles.ListCount - 1

End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    For i = LstFiles.ListCount - 1 To 0 Step -1
        If LstFiles.Selected(i) Then
            LstFiles.RemoveItem i
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim P As Printer
    Dim defaultPrinter As String
    
    defaultPrinter = GetSetting(App.EXEName, "Printer", "DefaultPrinter", "")

    For i = 0 To Printers.count - 1
        Set P = Printers(i)
        CboPrinter.AddItem P.DeviceName
        If P.DeviceName = defaultPrinter Then
            Me.CboPrinter.ListIndex = i
        End If
    Next
    
    If Me.CboPrinter.ListIndex <= 0 Then
        Me.CboPrinter.ListIndex = 0
    End If
    
    Me.OptMode(0).Value = True
    
End Sub

Private Sub OKButton_Click()
On Error GoTo ERROR_HANDLE
    Dim i As Integer
    Dim fname As String
    Dim page As Integer
    
    Me.cmdAdd.Enabled = False
    Me.cmdRemove.Enabled = False
    Me.OKButton.Enabled = False
    Me.CboPrinter.Enabled = False
    
    printing = True
    
    For i = 0 To LstFiles.ListCount - 1
        
        If Not printing Then
            GoTo FINISH
        End If
        
        fname = LstFiles.List(i)
        
        LstFiles.ListIndex = i
        LstFiles.List(i) = "Printing - " & fname
        DoEvents
        LstFiles.Refresh
        DoEvents
        
        page = page + 1
        If page > 1 Then
            Printer.NewPage
        End If
        Printer.Orientation = vbPRORLandscape
        
        If UCase(Right(fname, 4)) = ".WLD" Then
            Dim f As New FrmChart
            f.Load fname
            PrintChart f
            PrintGraph Printer, fname, OptMode(0).Value
            Unload f
             
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.ForeColor = vbBlack
            Printer.CurrentX = Printer.ScaleWidth * 0.94: Printer.CurrentY = Printer.ScaleHeight * 0.94: Printer.Print page
        Else
            Dim fdr As New FrmDailyReport
            fdr.Load fname
            page = PrintDailyReport(fdr, page)
            Unload fdr
            
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.ForeColor = vbBlack

        End If
        
        DoEvents
        LstFiles.List(i) = "Done - " & fname
        DoEvents
    Next i
    
    Printer.EndDoc
    Unload Me
Exit Sub
ERROR_HANDLE:
    MsgBox "Print Error ,Please contact administrator!"
    

FINISH:
    Printer.EndDoc
    If fname <> "" Then
        LstFiles.List(i) = fname
        For i = i - 1 To 0 Step -1
            LstFiles.RemoveItem i
        Next
    End If
    
    Me.cmdAdd.Enabled = True
    Me.cmdRemove.Enabled = True
    Me.OKButton.Enabled = True
    Me.CboPrinter.Enabled = True
End Sub
