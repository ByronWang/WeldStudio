VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JinanWeldNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWeldNumber

Const MS_0C As String = "123456789ABC"

Dim unit As String
Dim weldSerailNumber As Long

Dim base As IWeldNumber

Private Sub Class_Initialize()
    unit = GetSetting(App.EXEName, "UserData", "Unit", "")
    weldSerailNumber = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    
    If Not CInt(weldSerailNumber / 100) = Month(Date) * 100 + Day(Date) Then
        weldSerailNumber = Month(Date) * 10000 + Day(Date) * 100 + 1
        Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    End If
End Sub

Private Function IWeldNumber_Compacted() As String
    IWeldNumber_Compacted = IWeldNumber_ToCompact(weldSerailNumber)
End Function

Private Function IWeldNumber_Display() As String
    IWeldNumber_Display = unit & IWeldNumber_ToDisplay(Date, weldSerailNumber)
End Function

Private Function IWeldNumber_MoveNext() As Variant
    Dim wn As Long
    unit = GetSetting(App.EXEName, "UserData", "Unit", "")
    wn = GetSetting(App.EXEName, "WELD", "LastSerialNumber", 1)
    
    If wn = weldSerailNumber Then
        weldSerailNumber = weldSerailNumber + 1
        Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    Else
        If Not CInt(wn / 100) = Month(Date) * 100 + Day(Date) Then
            weldSerailNumber = Month(Date) * 10000 + Day(Date) * 100 + 1
            Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", IWeldNumber_Compacted)
        End If
    End If
End Function

Private Function IWeldNumber_Update(str As String) As Boolean
    If Len(str) <> 5 Then
        GoTo CHECK_ERROR
    End If
    
    If InStr(1, MS_0C, left(str, 1)) < 0 Then
        GoTo CHECK_ERROR
    End If
    
    If InStr(1, MS_0C, left(str, 1)) <> Month(Date) Or CInt(Mid(str, 2, 2)) <> Day(Date) Then
        GoTo CHECK_ERROR
    End If
    
    weldSerailNumber = IWeldNumber_FromCompact(str)
    Call SaveSetting(App.EXEName, "WELD", "LastSerialNumber", weldSerailNumber)
    IWeldNumber_Update = True
    
Exit Function
CHECK_ERROR:
        IWeldNumber_Update = False
End Function

Private Function IWeldNumber_FromCompact(s As String) As Long
    Dim leadNumber As Long
    Dim leadChar As String
    Dim leaveNumber As Long
    
    leadChar = left(s, 1)
    leadNumber = InStr(1, MS_0C, leadChar)

    leaveNumber = CInt(Mid(s, 2, 4))
    
    IWeldNumber_FromCompact = leadNumber * 10000 + leaveNumber
End Function

Private Function IWeldNumber_ToCompact(n As Long) As String
    Dim leadNumber As Long
    Dim leadChar As String
    Dim leaveNumber As Long
    
    leadNumber = PlcDeclare.Floor(n, 10000)
    
    leaveNumber = n - leadNumber * 10000#
    
    If leadNumber > 12 Then
        leadNumber = 1
    End If
    
    leadChar = Mid(MS_0C, leadNumber, 1)
    
    Dim showString As String
    showString = CStr(leaveNumber)
    
    IWeldNumber_ToCompact = "" & leadChar & Format(leaveNumber, "0000")
End Function

Private Function IWeldNumber_FromDisplay(sFull As String) As Long
    Dim leadNumber As Long
    Dim leadChar As String
    Dim leaveNumber As Long
    
    Dim s As String
    s = Mid(sFull, Len(sFull) - 6)
    
    leadChar = left(s, 2)
    leadNumber = CInt(leadChar)

    leaveNumber = CInt(Mid(s, 3, 4))
    
    IWeldNumber_FromDisplay = leadNumber * 10000 + leaveNumber
End Function

Private Function IWeldNumber_ToDisplay(dat As Date, n As Long) As String
    Dim leadNumber As Long
    Dim leadChar As String
    Dim leaveNumber As Long
    
    leadNumber = PlcDeclare.Floor(n, 10000)
    leaveNumber = n - leadNumber * 10000#
    
    If leadNumber > 12 Then
        leadNumber = 1
    End If
    
    leadChar = Mid(Format(Year(dat), "0000"), 3) & Format(leadNumber, "00")
        
    Dim showString As String
    showString = Format(leaveNumber, "0000")
    
    IWeldNumber_ToDisplay = "" & leadChar & Format(leaveNumber, "0000")
End Function
