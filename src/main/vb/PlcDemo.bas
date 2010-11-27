Attribute VB_Name = "PlcDemo"
Option Explicit

Private start As Boolean

Dim data As WeldMinitor
Public EmulateData() As Record

Dim startTime  As Long
Dim lastTime  As Single
Dim lastPos As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function StartDemo() As Integer()
    EmulateData = PlcWld.LoadData("E:\WeldData\27-Sep-2008\T0039.WLD")
    start = True
    lastPos = 0
    lastTime = 0
    
        If lastTime = 0 Then
            startTime = timeGetTime
        End If
End Function




Public Function readPcMonitor() As WeldMinitor
        Dim i As Long
        
    If start Then
        
        If lastPos < UBound(EmulateData) Then
        For i = lastPos To UBound(EmulateData) - 1
            If lastTime < EmulateData(i).data.Time Then
                Exit For
            End If
        Next
        
        Dim nowTime As Long
        nowTime = timeGetTime
        
        Dim et As Long
        et = (EmulateData(i).data.Time - lastTime) * 1000
        If et < 0 Then
            et = 0
        End If
        
        Sleep (et)
        
        lastTime = EmulateData(i).data.Time
        readPcMonitor = EmulateData(i).data
        lastPos = i
        
        Else
        data.WeldStage = 0
        data.PlcStage = 0
        readPcMonitor = data
        End If
    Else
        data.WeldStage = 0
        data.PlcStage = 0
        readPcMonitor = data
    End If
    
    
End Function

