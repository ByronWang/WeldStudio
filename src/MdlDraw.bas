Attribute VB_Name = "MdlDraw"
Option Explicit

Dim pWidth As Single
Dim pHeight As Single
'
'
Const GRID_COLOR As Long = &HBBBBBB
Const Aix_COLOR As Long = &H0

Dim V_Y1_Start, V_Y1_Max, V_Y1_Major As Single
Dim V_Y2_Start, V_Y2_Max, V_Y2_Major As Single
Dim V_X_Start, V_X_Max, V_X_Major As Single
'
Dim xScale, y1Scale, y2Scale As Single
Dim buf() As WeldData

Dim pLeft, pTop As Single

Public Sub PrintGraph(canvas As Printer, fname As String, fullWeldCycle As Boolean) ', posX As Single, posY As Single, maxX As Single, maxY As Single)

    Dim fr As FileR
    
    fr = PlcWld.LoadData(fname)

    Dim i As Integer
    
    If fullWeldCycle Then
        buf = LoadDataAll(fr.data, fr.header2.RecordCount)
        Call PrepareDraw(canvas, canvas.ScaleWidth * 2.8 / 10, canvas.ScaleHeight * 8.5 / 10, canvas.ScaleWidth * 6.5 / 10, -canvas.ScaleHeight * 7 / 10, buf(0).Time)
        DrawChartAll canvas, buf, fr.analysisDefine
    Else
        buf = LoadDataUpset(fr.data, fr.header2.RecordCount)
        Call PrepareDraw(canvas, canvas.ScaleWidth * 2.8 / 10, canvas.ScaleHeight * 8.5 / 10, canvas.ScaleWidth * 6.5 / 10, -canvas.ScaleHeight * 7 / 10, buf(0).Time)
        DrawChartUpset canvas, buf, fr.analysisDefine
    End If
   
    
End Sub


Public Function LoadDataAll(rs() As Record, count As Integer) As WeldData()
    Dim i As Integer
    Dim buf() As WeldData
    ReDim buf(count - 1)
    For i = 0 To count - 1
        buf(i) = rs(i).data
    Next
    LoadDataAll = buf
End Function

Public Function LoadDataUpset(rs() As Record, count As Integer) As WeldData()
    Dim buf() As WeldData
    Dim i, posStart, posEnd As Integer
    Dim sTime As Single
    
    posStart = 0
    
    For i = 0 To count - 1
        If rs(i).data.WeldStage = BOOST_STAGE Then
            posStart = i
            Exit For
        End If
    Next
    
    If posStart = 0 Then
        GoTo ERROR_HANDLE
    End If

    sTime = rs(i).data.Time
    For i = posStart To count - 1
        If rs(i).data.Time - sTime > 25 Then
            Exit For
        End If
    Next
    
    posEnd = i - 1
    
    If posEnd - posStart <= 0 Then
        GoTo ERROR_HANDLE
    End If
    
    ReDim buf(posEnd - posStart)
    
    For i = posStart To posEnd
        buf(i - posStart) = rs(i).data
    Next

    LoadDataUpset = buf
Exit Function
ERROR_HANDLE:
    ReDim buf(0)
    LoadDataUpset = buf
End Function

Public Function DrawData(canvas, ByVal rStart As Single, ByVal rScale As Single, ByVal dStart As Single, ByVal dScale As Single, color As Long, ref() As Single, d() As Single) As Integer
    Dim i As Integer
    i = LBound(d)
    
    Dim fx, fy, tx, ty As Single
    
    canvas.DrawWidth = 1
    canvas.DrawMode = 13
    
    Dim dv As Single
    
    fx = 0
    dv = d(i)
    If dv < 0 Then
        dv = 0
    End If
    fy = (dv - dStart) * dScale
        
    For i = i + 1 To UBound(d)
        dv = d(i)
        If dv < 0 Then
            dv = 0
        End If
        
        tx = (ref(i) - rStart) * rScale
        ty = (dv - dStart) * dScale
        
        
         canvas.Line (pLeft + fx, pTop + fy)-(pLeft + tx, pTop + ty), color ', &HFF0000 * Rnd(1) + &HFF00 * Rnd(1) + &HFF * Rnd(1)
         fx = tx
         fy = ty
    Next i
    
End Function

Public Sub PrepareDraw(canvas, left As Single, top As Single, width As Single, height, startX As Single)
    pLeft = left
    pTop = top
    pWidth = width
    pHeight = height

    V_Y1_Start = CSng(GetSetting(App.EXEName, "WeldChartSetting", "AVMin", 0))
    V_Y1_Max = CSng(GetSetting(App.EXEName, "WeldChartSetting", "AVMax", 1000))
    V_Y1_Major = CSng(GetSetting(App.EXEName, "WeldChartSetting", "AVIncr", 100))
    
    V_Y2_Start = CSng(GetSetting(App.EXEName, "WeldChartSetting", "DFMin", 0))
    V_Y2_Max = CSng(GetSetting(App.EXEName, "WeldChartSetting", "DFMax", 160))
    V_Y2_Major = CSng(GetSetting(App.EXEName, "WeldChartSetting", "DFIncr", 16))
    
    If startX > 1 Then
        V_X_Start = CInt(startX)
        V_X_Max = V_X_Start + 20
        V_X_Major = 5
    Else
        V_X_Start = 0
        V_X_Max = CSng(GetSetting(App.EXEName, "WeldChartSetting", "TicanvasMaxCycleTicanvas", 200))
        V_X_Major = CSng(GetSetting(App.EXEName, "WeldChartSetting", "TicanvasIncr", 10))
    End If
    
    Dim tXLabelOffset, tYLabelOffset As Single
    
    canvas.DrawWidth = 1
    canvas.DrawStyle = 0
    
    tXLabelOffset = -canvas.TextWidth(0) - canvas.TextWidth(0)
    tYLabelOffset = canvas.TextHeight(0) / 2
    
'    max = 1000
        
    Dim pos As Single
        
    'XXXXX
    canvas.DrawWidth = 1
    'Printer.FontSize = 12
    Printer.FontBold = False
    Printer.ForeColor = vbBlack
    
    xScale = pWidth / (V_X_Max - V_X_Start)
    y1Scale = pHeight / (V_Y1_Max - V_Y1_Start)
    y2Scale = pHeight / (V_Y2_Max - V_Y2_Start)
    
    
    For pos = V_X_Start To V_X_Max Step V_X_Major
        canvas.Line (pLeft + (pos - V_X_Start) * xScale, pTop + 0)-(pLeft + (pos - V_X_Start) * xScale, pTop + pHeight), GRID_COLOR
    Next pos
    
    For pos = V_X_Start To V_X_Max Step V_X_Major
        canvas.Line (pLeft + (pos - V_X_Start) * xScale, pTop + 0)-(pLeft + (pos - V_X_Start) * xScale, pTop + tYLabelOffset), GRID_COLOR
    Next pos
'
    For pos = V_X_Start To V_X_Max Step V_X_Major
        canvas.CurrentX = pLeft + (pos - V_X_Start) * xScale - canvas.TextWidth(0) / 1.5 - canvas.TextWidth(pos) / 2: canvas.CurrentY = pTop + tYLabelOffset: canvas.Print pos
    Next pos

    'Y1
    For pos = V_Y1_Start To V_Y1_Max Step V_Y1_Major
        canvas.Line (pLeft + 0, pTop + (pos - V_Y1_Start) * y1Scale)-(pLeft + pWidth, pTop + (pos - V_Y1_Start) * y1Scale), GRID_COLOR
    Next pos

    For pos = V_Y1_Start To V_Y1_Max Step V_Y1_Major
        canvas.Line (pLeft + 0, pTop + (pos - V_Y1_Start) * y1Scale)-(pLeft + tXLabelOffset / 3, pTop + (pos - V_Y1_Start) * y1Scale), GRID_COLOR
    Next pos

    For pos = V_Y1_Start To V_Y1_Max Step V_Y1_Major
        canvas.CurrentX = pLeft + tXLabelOffset - canvas.TextWidth(pos): canvas.CurrentY = pTop + (pos - V_Y1_Start) * y1Scale - canvas.TextHeight(pos) / 2: canvas.Print pos
    Next pos
    
    'Y2
    canvas.DrawStyle = 1
    For pos = V_Y2_Start To V_Y2_Max Step V_Y2_Major
        canvas.Line (pLeft + 0, pTop + (pos - V_Y2_Start) * y2Scale)-(pLeft + pWidth, pTop + (pos - V_Y2_Start) * y2Scale), GRID_COLOR
    Next pos

    canvas.DrawStyle = 0
    For pos = V_Y2_Start To V_Y2_Max Step V_Y2_Major
        canvas.Line (pLeft + pWidth, pTop + (pos - V_Y2_Start) * y2Scale)-(pLeft + pWidth - tXLabelOffset / 3, pTop + (pos - V_Y2_Start) * y2Scale), GRID_COLOR
    Next pos

    For pos = V_Y2_Start To V_Y2_Max Step V_Y2_Major
        canvas.CurrentX = pLeft + pWidth: canvas.CurrentY = pTop + (pos - V_Y2_Start) * y2Scale - canvas.TextHeight(pos) / 2: canvas.Print pos
    Next pos
    
    canvas.DrawWidth = 1
    canvas.Line (pLeft + 0, pTop + 0)-(pLeft + 0, pTop + pHeight), Aix_COLOR
    canvas.Line (pLeft + 0, pTop + 0)-(pLeft + pWidth, pTop + 0), Aix_COLOR
    canvas.Line (pLeft + pWidth, pTop + 0)-(pLeft + pWidth, pTop + pHeight), Aix_COLOR
    
End Sub

Public Sub DrawChartAll(canvas, data() As WeldData, analysisDefine As WeldAnalysisDefineType)

    Dim i As Integer

    Dim tim() As Single
    Dim psi() As Single
    Dim Volt() As Single
    Dim Amp() As Single
    Dim Dist() As Single

    Dim count As Integer
    count = UBound(data)

    ReDim tim(count - 1)
    ReDim psi(count - 1)
    ReDim Volt(count - 1)
    ReDim Amp(count - 1)
    ReDim Dist(count - 1)

    For i = 0 To count - 1
        tim(i) = data(i).Time
        psi(i) = PlcAnalysiser.toForce(data(i).PsiUpset, data(i).PsiOpen, analysisDefine)
        Volt(i) = data(i).Volt
        Amp(i) = data(i).Amp
        Dist(i) = data(i).Dist
    Next

    Call DrawData(canvas, V_X_Start, xScale, V_Y1_Start, y1Scale, &HC000&, tim(), Amp())
    Call DrawData(canvas, V_X_Start, xScale, V_Y1_Start, y1Scale, &HFF&, tim(), Volt())

    Call DrawData(canvas, V_X_Start, xScale, V_Y2_Start, y2Scale, &H80&, tim(), psi())
    Call DrawData(canvas, V_X_Start, xScale, V_Y2_Start, y2Scale, &HFF0000, tim(), Dist())
End Sub

Public Sub DrawChartUpset(canvas, data() As WeldData, analysisDefine As WeldAnalysisDefineType)
    If UBound(data) < 1 Then
        Exit Sub
    End If
    
    
    Dim i As Integer
    
    Dim tim() As Single
    Dim psi() As Single
    Dim Volt() As Single
    Dim Amp() As Single
    Dim Dist() As Single
    
    Dim count As Integer
    count = UBound(data)
    
    ReDim tim(count - 1)
    ReDim psi(count - 1)
    ReDim Volt(count - 1)
    ReDim Amp(count - 1)
    ReDim Dist(count - 1)
    
    For i = 0 To count - 1
        tim(i) = data(i).Time
        psi(i) = PlcAnalysiser.toForce(data(i).PsiUpset, data(i).PsiOpen, analysisDefine)
        Volt(i) = data(i).Volt
        Amp(i) = data(i).Amp
        Dist(i) = data(i).Dist
    Next

    Call DrawData(canvas, data(0).Time, xScale, V_Y1_Start, y1Scale, &HC000&, tim(), Amp())
    Call DrawData(canvas, data(0).Time, xScale, V_Y1_Start, y1Scale, &HFF&, tim(), Volt())
    
    Call DrawData(canvas, data(0).Time, xScale, V_Y2_Start, y2Scale, &H80&, tim(), psi())
    Call DrawData(canvas, data(0).Time, xScale, V_Y2_Start, y2Scale, &HFF0000, tim(), Dist())
End Sub
