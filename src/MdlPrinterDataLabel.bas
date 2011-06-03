Attribute VB_Name = "MdlPrinterDataLabel"
Option Explicit

Public Function PrintChart(fc As FrmChart)
Printer.Orientation = vbPRORLandscape
        
'        fc.MSChart1.EditCopy
'        DoEvents   ' may be needed for large datasets
'        DoEvents   ' may be needed for large datasets
'        Printer.Print " "
'        'Printer.Print " ------------------------------- "
'        Printer.Print " "
'        Printer.PaintPicture Clipboard.GetData(), 3500, 2200
        
        
        Dim i, j As Integer
        Dim gSep As Single
        Dim iSep As Single
        
        Dim gLeft, iLeft, idLeft As Integer
        
        gLeft = 800
        iLeft = 1100
        idLeft = 3200
        
        gSep = 100 ' group sep
        iSep = 50 ' item sep
        
        Printer.CurrentY = 1600
        
        Dim lTop As Integer
        
        With fc
    
            i = 0
            Printer.CurrentX = gLeft
            Call setFrom(.lblGroup(i))
            Printer.CurrentY = Printer.CurrentY + gSep
            
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
            Printer.CurrentY = Printer.CurrentY + gSep
            
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
            Printer.CurrentY = Printer.CurrentY + gSep
                            
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
            Printer.CurrentY = Printer.CurrentY + gSep
            
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
            Printer.CurrentY = Printer.CurrentY + gSep
            
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
            Printer.CurrentY = Printer.CurrentY + gSep
            
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
End Function

Public Function PrintDailyReport(f As FrmDailyReport)
Printer.Orientation = vbPRORLandscape
    
Dim x, y, k, i, j As Long
Dim pagelines As Integer
Dim stepTo As Integer

pagelines = 26

For k = 1 To f.MSFlexGrid1.Rows - 1 Step pagelines
    If k + pagelines < f.MSFlexGrid1.Rows Then
        stepTo = k + pagelines
    Else
        stepTo = f.MSFlexGrid1.Rows
    End If
        
    x = 1000
    y = 1600
        
    Call navControlForDailyReport(f.lblCompany)
    Call navControlForDailyReport(f.lblDate)
    Call navControlForDailyReport(f.lblLocation)
    Call navControlForDailyReport(f.lblUnit)
    

    For j = 0 To f.MSFlexGrid1.Cols - 1
        
        Printer.CurrentY = y
        
        Printer.FontBold = True
        Printer.FontSize = 10
        For i = 0 To 0
            Printer.CurrentX = x
            Printer.Print f.MSFlexGrid1.TextMatrix(i, j)
        Next i
        
        Printer.CurrentY = Printer.CurrentY + 100
                
        Printer.FontBold = False
        Printer.FontSize = 10
        For i = k + 0 To stepTo - 1
            Printer.CurrentY = Printer.CurrentY + 100
            Printer.CurrentX = x - 20
            Printer.Print f.MSFlexGrid1.TextMatrix(i, j)
        Next i
        x = x + f.MSFlexGrid1.ColWidth(j) * 1
    Next j
    
    Printer.CurrentY = 10800
    Printer.CurrentX = 15360
    Printer.Print CInt((k + pagelines - 1) \ pagelines) & " / " & CInt((f.MSFlexGrid1.Rows - 1 + pagelines - 1) \ pagelines)
    
    If f.MSFlexGrid1.Rows > k + pagelines Then
        Printer.NewPage
    End If
Next k

Call navControlForDailyReportBottom(f.frmSum, f.labelAccepted)
Call navControlForDailyReportBottom(f.frmSum, f.labelRejected)
Call navControlForDailyReportBottom(f.frmSum, f.labelTotal)
Call navControlForDailyReportBottom(f.frmSum, f.lblAccepted)
Call navControlForDailyReportBottom(f.frmSum, f.lblReject)
Call navControlForDailyReportBottom(f.frmSum, f.lblTotal)
    
Printer.EndDoc

    
End Function

Private Function navControlForDailyReport(con As Label)
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    
    Dim sca As Single
    sca = 1
    If con.width > 3000 Then
        sca = 3600 / con.width
    End If
    
    Printer.CurrentY = con.top + 700
    
    Select Case con.Alignment
        Case vbLeftJustify:
            Printer.CurrentX = con.left * sca + 0
        Case vbRightJustify:
            Printer.CurrentX = con.left * sca + con.width * sca - Printer.TextWidth(con.Caption)
        Case vbCenter:
            Printer.CurrentX = con.left * sca + (con.width * sca - Printer.TextWidth(con.Caption)) / 2
    End Select

    Printer.Print con.Caption
End Function

Private Function navControlForDailyReportBottom(fm As Frame, con As Label)
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    
    Dim sca As Single
    sca = 1
    If con.width > 3000 Then
        sca = 3600 / con.width
    End If
    
    Dim pLeftOffset As String
    pLeftOffset = 0
    
    Printer.CurrentY = 9600 + con.top + 700
    Select Case con.Alignment
        Case vbLeftJustify:
            Printer.CurrentX = pLeftOffset + con.left * sca + 0
        Case vbRightJustify:
            Printer.CurrentX = pLeftOffset + con.left * sca + con.width * sca - Printer.TextWidth(con.Caption)
        Case vbCenter:
            Printer.CurrentX = pLeftOffset + con.left * sca + (con.width * sca - Printer.TextWidth(con.Caption)) / 2
    End Select
    
    Printer.Print con.Caption
End Function


Private Function navControl(con As Label)
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    
    Dim sca As Single
    sca = 1
    If con.width > 3000 Then
        sca = 3600 / con.width
    End If
    
    Dim pLeftOffset As String
    pLeftOffset = 600
    Printer.CurrentY = con.top * sca + 100
    
    Select Case con.Alignment
        Case vbLeftJustify:
            Printer.CurrentX = pLeftOffset + con.left * sca + 0
        Case vbRightJustify:
            Printer.CurrentX = pLeftOffset + con.left * sca + con.width * sca - Printer.TextWidth(con.Caption)
        Case vbCenter:
            Printer.CurrentX = pLeftOffset + con.left * sca + (con.width * sca - Printer.TextWidth(con.Caption)) / 2
    End Select
    
    Printer.Print con.Caption
End Function

Private Function setFrom(con As Control)
    Printer.FontSize = con.FontSize
    Printer.FontBold = con.FontBold
    Printer.ForeColor = con.ForeColor
    Printer.Print con.Caption
End Function
